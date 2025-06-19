# Consolidar Avance Académico 📊

Este script en Python automatiza la conexión a una red privada mediante VPN, consulta datos académicos desde dos bases de datos SQL distintas y cruza esa información con un archivo Excel de entrada para consolidarla en un único archivo Excel de salida.

---

## 🚀 Funcionalidades

- ✅ Conexión automática a VPN usando OpenVPN
- ✅ Verificación de conectividad y rutas estáticas de red
- ✅ Consultas seguras a SQL Server con credenciales externas
- ✅ Combinación de datos desde múltiples fuentes por `COD_UNICO_ST`
- ✅ Guardado de resultados con formato en múltiples hojas de Excel
- ✅ Progreso visual en consola con spinner, íconos y colores
- ✅ Manejo de errores y resumen de ejecución

---

## 🧾 Requisitos

- Python 3.9 o superior
- Librerías: `pandas`, `openpyxl`, `pyodbc`, `tqdm`
- Cliente OpenVPN instalado
- Acceso habilitado al servidor SQL mediante VPN
- Archivo Excel de entrada con columna `COD_UNICO_ST`

---