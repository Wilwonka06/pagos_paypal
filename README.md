# Sistema de Automatización de Pagos PayPal (v1.0.0)

Este proyecto es una solución integral en **Python** diseñada para automatizar el ciclo completo de gestión de pagos de PayPal, desde la extracción de datos en SAP hasta la validación de soportes documentales.

## Funcionalidades Principales
- **Automatización de SAP:** Navegación automática en la transacción `FAGLL03`, login y descarga de reportes contables en Excel [6, 7].
- **Gestión Documental:** Creación automática de estructuras de carpetas por número de pago (`Pago #N`) [6, 7].
- **Procesamiento de Datos:** Reorganización de columnas de SAP y consolidación automática con un archivo maestro de Courier [6, 7].
- **Validación de Soportes:** Búsqueda inteligente de facturas y guías en formato PDF, verificando fechas y cruzando información con el reporte [4, 6].
- **Interfaz Moderna:** GUI desarrollada con `CustomTkinter` que incluye barra de progreso y logs en tiempo real [1, 6].

## Tecnologías Utilizadas
- **Lenguaje:** Python 3.11+ [8].
- **Automatización Web:** Selenium (Firefox + geckodriver) [8].
- **Análisis de Datos:** Pandas, Openpyxl, XLrd [8].
- **Manipulación PDF:** PyMuPDF (fitz) [8].
- **Interfaz:** CustomTkinter [8].

## Requisitos Previos
1. **Python 3.11** o superior instalado en Windows [9].
2. **Mozilla Firefox** y el controlador `geckodriver` configurado en el PATH [9].
3. Acceso a **SAP** con credenciales válidas en el archivo de configuración [9].
4. Archivo **Maestro de Courier** disponible en la ruta especificada [9].

## Instalación
1. Clonar el repositorio.
2. Crear y activar un entorno virtual:
   ```bash
   python -m venv venv
   .\venv\Scripts\activate
Instalar dependencias:


📖 Modo de Uso
Interfaz Gráfica (GUI)
Ejecute el siguiente comando para iniciar la aplicación visual:
python interfaz.py
El sistema sugerirá el número de pago automáticamente
.
Presione "EJECUTAR PROCESO COMPLETO" y siga el progreso en pantalla
.
Modo Consola
Para una ejecución rápida y automática del siguiente pago pendiente:
python main.py
Generación de Ejecutable
Para crear una versión .exe distribuible, ejecute el script:
.\build.bat
Podrá definir una ruta de salida personalizada para el ejecutable generado
.
📂 Estructura del Proyecto
main.py: Lógica central y modo consola
.
interfaz.py: Código de la interfaz gráfica y orquestación de hilos
.
config_paypal.ini: Parámetros de SAP, rutas y configuraciones generales
.
build.bat: Automatización de compilación con PyInstaller
.
📝 Notas de Versión
Implementación de detección de pagos parciales (etiquetados como "Próximo pago")
.
Sistema de logs automático en carpeta /logs
.
Automatización de comandos de teclado (Mayús+F4) para exportación en SAP
.

--------------------------------------------------------------------------------
