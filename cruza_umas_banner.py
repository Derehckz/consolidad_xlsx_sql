import subprocess
import time
import pandas as pd
import pyodbc
import os
import sys
import threading
from tqdm import tqdm
from openpyxl import load_workbook, Workbook
import logging
import signal
import configparser

# --- CONFIGURACIÓN PERSONALIZADA ---
config = configparser.ConfigParser()
config.read(os.path.join(os.path.dirname(__file__), 'config.ini')) # Cargar configuración desde un archivo ini
if not config.sections():
    print("Error: No se pudo cargar el archivo de configuración 'config.ini'.")
    sys.exit(1)

BASE_DIR = config['Paths']['base_dir']
OPENVPN_PATH = config['VPN']['openvpn_path']
OVPN_FILE = config['VPN']['ovpn_file']
AUTH_FILE = config['VPN']['auth_file']

SQL_DRIVER = config['SQL']['driver']
SERVER = config['SQL']['server']
USERNAME = config['SQL']['username']
PASSWORD = config['SQL']['password']

CONSULTAS_SQL = dict(config['Consultas'])
EXCEL_ENTRADA = os.path.join(BASE_DIR, config['Paths']['input_excel'])
HOJA_EXCEL = config['Paths'].getint('excel_sheet', fallback=0)
SALIDA = os.path.join(BASE_DIR, config['Paths']['output_excel'])


vpn_proceso = None
vpn_conectada = False

# --- CONFIGURACIÓN DE LOGGING ---
logging.basicConfig(
    filename=os.path.join(BASE_DIR, "consolidacion.log"),
    filemode='a',
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    encoding='utf-8'
)

# --- COLORES PARA CONSOLA ---
class Colors:
    GREEN = "\033[92m"
    YELLOW = "\033[93m"
    RED = "\033[91m"
    CYAN = "\033[96m"
    RESET = "\033[0m"

def msg_info(text):
    print(f"{Colors.CYAN}[ℹ️ INFO]{Colors.RESET} {text}")
    logging.info(text)

def msg_success(text):
    print(f"{Colors.GREEN}[✅ ÉXITO]{Colors.RESET} {text}")
    logging.info(text)

def msg_warn(text):
    print(f"{Colors.YELLOW}[⚠️ ADVERTENCIA]{Colors.RESET} {text}")
    logging.warning(text)

def msg_error(text):
    print(f"{Colors.RED}[❌ ERROR]{Colors.RESET} {text}", file=sys.stderr)
    logging.error(text)

def print_step(step_num, total_steps, description):
    msg_info(f"➡ Paso {step_num}/{total_steps}: {description}")

# --- SPINNER PARA INDICAR PROGRESO EN CONSULTAS LARGAS ---
class Spinner:
    def __init__(self, mensaje="Procesando"):
        self.mensaje = mensaje
        self._running = False
        self._thread = None
        self.chars = ['|', '/', '-', '\\']
    
    def _spin(self):
        i = 0
        while self._running:
            sys.stdout.write(f"\r{self.mensaje}... {self.chars[i % len(self.chars)]}")
            sys.stdout.flush()
            time.sleep(0.1)
            i += 1
        sys.stdout.write("\r" + " " * (len(self.mensaje) + 5) + "\r")
    
    def start(self):
        self._running = True
        self._thread = threading.Thread(target=self._spin)
        self._thread.start()
    
    def stop(self):
        self._running = False
        if self._thread:
            self._thread.join()

# --- FUNCIONES MODULARES ---
def verificar_archivos():
    print_step(1, 9, "Validando archivos y rutas necesarias...")
    archivos = {
        "Ejecutable OpenVPN": OPENVPN_PATH,
        "Archivo configuración VPN": OVPN_FILE,
        "Archivo credenciales VPN": AUTH_FILE,
        "Archivo Excel de entrada": EXCEL_ENTRADA
    }
    errores = []
    for nombre, ruta in archivos.items():
        if not os.path.exists(ruta):
            errores.append(f"{nombre} NO encontrado en: {ruta}")
            msg_error(errores[-1])
        else:
            msg_success(f"{nombre} encontrado en: {ruta}")
    if errores:
        return False
    return True

def monitor_vpn_output(proc):
    global vpn_conectada
    while True:
        line = proc.stdout.readline()
        if not line:
            break
        line_str = line.decode(errors='ignore').strip()
        if any(keyword in line_str for keyword in ["Initialization Sequence Completed", "AUTH_FAILED", "ERROR"]):
            msg_info(f"[VPN] {line_str}")
        if "Initialization Sequence Completed" in line_str:
            msg_success("VPN conectada correctamente.")
            vpn_conectada = True
        if "AUTH_FAILED" in line_str:
            msg_error("Autenticación VPN fallida. Revisa credenciales.")
            break

def iniciar_vpn():
    global vpn_proceso, vpn_conectada
    print_step(2, 9, "Iniciando conexión VPN...")
    cmd = [
        OPENVPN_PATH,
        "--config", OVPN_FILE,
        "--auth-user-pass", AUTH_FILE,
        "--data-ciphers", "DEFAULT:AES-256-CBC"
    ]
    print(f"{Colors.CYAN}▶ Ejecutando comando VPN:{Colors.RESET}")
    print(" ".join(f'"{arg}"' if " " in arg else arg for arg in cmd))
    vpn_proceso = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    threading.Thread(target=monitor_vpn_output, args=(vpn_proceso,), daemon=True).start()

    msg_info("🕒 Esperando hasta 60 segundos para que la VPN se conecte...")
    for _ in tqdm(range(60), desc="Conectando VPN", ncols=70):
        if vpn_conectada:
            return True
        time.sleep(1)
    raise TimeoutError("⛔ Tiempo de espera agotado: VPN no se conectó.")

def ruta_ya_agregada():
    resultado = subprocess.run(["route", "print"], stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, text=True)
    return "192.168.110.0" in resultado.stdout

def agregar_ruta_estatica():
    if ruta_ya_agregada():
        msg_info("Ruta ya existente, no se agrega de nuevo.")
        return
    print_step(3, 9, "Agregando ruta estática para red SQL...")
    comando = ["route", "add", "192.168.110.0", "mask", "255.255.255.0", "10.0.200.1"]
    resultado = subprocess.run(comando, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    if resultado.returncode == 0:
        msg_success("Ruta agregada correctamente.")
    else:
        msg_warn(f"No se pudo agregar la ruta. Error: {resultado.stderr.decode().strip()}")

def verificar_ping(host, intentos=6, espera=5):
    print_step(4, 9, f"Verificando conectividad con {host}...")
    for i in tqdm(range(intentos), desc="Ping servidor SQL", ncols=70):
        resultado = subprocess.run(["ping", "-n", "1", host], stdout=subprocess.DEVNULL)
        if resultado.returncode == 0:
            msg_success("Conectividad establecida.")
            return True
        msg_info(f"⏳ Intento {i+1}/{intentos}, esperando {espera} segundos...")
        time.sleep(espera)
    msg_error("No se pudo establecer conexión con el servidor.")
    return False

def conectar_sql(database):
    try:
        conn = pyodbc.connect(
            f"DRIVER={{{SQL_DRIVER}}};SERVER={SERVER};DATABASE={database};UID={USERNAME};PWD={PASSWORD};TrustServerCertificate=yes;Encrypt=no",
            timeout=5
        )
        msg_success(f"Conexión a base de datos {database} establecida.")
        return conn
    except pyodbc.Error as e:
        msg_error(f"Fallo conexión a {database}: {e}")
        raise

def fusionar_con_base(df_excel, conn, sql_query, origen):
    msg_info(f"Consultando base {origen}...")
    spinner = Spinner(f"{Colors.CYAN}Consultando {origen}{Colors.RESET}")
    spinner.start()
    try:
        df_sql = pd.read_sql(sql_query, conn)
    finally:
        spinner.stop()
    if df_sql.empty:
        msg_warn(f"La consulta a {origen} devolvió 0 registros.")
    else:
        msg_success(f"Registros obtenidos de {origen}: {len(df_sql)}")
    df_sql["ORIGEN_BD"] = origen

    if "COD_UNICO_ST" not in df_excel.columns or "COD_UNICO_ST" not in df_sql.columns:
        msg_error("Columna 'COD_UNICO_ST' no encontrada en uno de los DataFrames.")
        return pd.DataFrame()  # Devuelve vacío para evitar errores

    df_merged = pd.merge(df_excel, df_sql, on="COD_UNICO_ST", how="inner")
    msg_success(f"Registros combinados con {origen}: {len(df_merged)}")
    return df_merged

def guardar_resultados_en_excel(df_original, df_ipll, df_cftll, ruta_salida):
    spinner = Spinner(f"{Colors.CYAN}Guardando archivo Excel{Colors.RESET}")
    try:
        msg_info(f"➡ Iniciando guardado de resultados en Excel: {ruta_salida}...")
        spinner.start()

        with pd.ExcelWriter(ruta_salida, engine='openpyxl') as writer:
            df_original.to_excel(writer, sheet_name='Original', index=False)
            if not df_ipll.empty:
                df_ipll.to_excel(writer, sheet_name='ST_IPLL_Combinado', index=False)
            if not df_cftll.empty:
                df_cftll.to_excel(writer, sheet_name='ST_CFTLL_Combinado', index=False)

    finally:
        spinner.stop()

    msg_success(f"✅ Archivo guardado correctamente en: {ruta_salida}")

def desconectar_vpn():
    global vpn_proceso
    if vpn_proceso and vpn_proceso.poll() is None:
        msg_info("Terminando conexión VPN...")
        vpn_proceso.terminate()
        vpn_proceso.wait()
        msg_success("VPN desconectada correctamente.")

def manejar_salida(signal_received=None, frame=None):
    msg_info("Finalizando proceso...")
    desconectar_vpn()
    msg_info("Proceso finalizado. Cerrando conexión VPN.")
    sys.exit(0)

def resumen_final(registros_combinados, tiempo_inicio):
    tiempo_total = time.time() - tiempo_inicio
    msg_success("\n===== Resumen de la Ejecución =====")
    if registros_combinados:
        total = sum(len(df) for df in registros_combinados if df is not None)
        msg_success(f"Total registros combinados: {total}")
    else:
        msg_warn("No se combinaron registros desde ninguna base.")
    msg_info(f"Tiempo total de ejecución: {tiempo_total:.2f} segundos")
    print("="*35)

# --- MAIN ---
def main():
    signal.signal(signal.SIGINT, manejar_salida)  # Ctrl+C

    tiempo_inicio = time.time()
    total_steps = 9

    if not verificar_archivos():
        msg_error("Errores en archivos requeridos. Corrige e intenta nuevamente.")
        sys.exit(1)

    try:
        iniciar_vpn()
        agregar_ruta_estatica()

        if not verificar_ping(SERVER):
            raise ConnectionError("Servidor SQL no responde. Revisa VPN o IP.")

        print_step(5, total_steps, "Cargando Excel recibido...")
        df_excel = pd.read_excel(EXCEL_ENTRADA, sheet_name=HOJA_EXCEL)
        msg_success(f"Excel cargado con {len(df_excel)} registros.")

        combinados_totales = []

        for idx, (bd, consulta) in enumerate(CONSULTAS_SQL.items(), start=6):
            print_step(idx, total_steps, f"Consultando base de datos {bd}...")
            conn = conectar_sql(bd)  # Conexión manual para cierre explícito
            try:
                df_resultado = fusionar_con_base(df_excel, conn, consulta, bd)
                combinados_totales.append((bd, df_resultado))
            except Exception as e:
                msg_error(f"Error consultando {bd}: {e}")
            finally:
                conn.close()

        # Extraer dataframes por base
        df_ipll = next((df for bd, df in combinados_totales if bd == "ST_IPLL"), pd.DataFrame())
        df_cftll = next((df for bd, df in combinados_totales if bd == "ST_CFTLL"), pd.DataFrame())

        # Guardar resultados
        guardar_resultados_en_excel(df_excel, df_ipll, df_cftll, SALIDA)

        resumen_final([df for _, df in combinados_totales], tiempo_inicio)

    except TimeoutError as e:
        msg_error(f"Timeout: {e}")
    except ConnectionError as e:
        msg_error(f"Error de conexión: {e}")
    except Exception as e:
        msg_error(f"Error general: {e}")
        msg_info("Sugerencia: Verifica conexión VPN, credenciales y archivos de entrada.")
    finally:
        manejar_salida()

if __name__ == "__main__":
    main()
