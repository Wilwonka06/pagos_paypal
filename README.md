# SISTEMA DE AUTOMATIZACIÓN DE PAGOS PAYPAL  
Versión **1.0.0**

## Descripción general
Aplicación en **Python** que automatiza el flujo de pagos PayPal:

- Descarga el reporte contable desde **SAP** usando Selenium.
- Organiza la estructura de carpetas por pago (`Pago #N`).
- Procesa el Excel descargado (SAP) y genera una **segunda hoja** consolidada.
- Busca y valida los documentos **PDF** de soporte (facturas y guías).
- Marca observaciones (faltan documentos, fechas, “Próximo pago”, etc.).
- Ofrece una **interfaz gráfica** moderna (CustomTkinter) y un modo consola.

## Tecnologías y lenguaje
- Lenguaje: **Python 3.11+**
- GUI: `customtkinter` (sobre `tkinter`)
- Automatización web: `selenium` (Firefox + geckodriver)
- Excel: `pandas`, `openpyxl`, `xlrd`
- PDF: `PyMuPDF` (`fitz`)
- Build ejecutable: `pyinstaller`, `pyinstaller-hooks-contrib`
- Utilidades: `python-dateutil`

Todas las dependencias están listadas en [requirements.txt](file:///c:/Proyectos%20Comodin/pagos_paypal/requirements.txt).

## Estructura principal del proyecto
- [main.py](file:///c:/Proyectos%20Comodin/pagos_paypal/main.py)  
  Lógica central del proceso (SAP, Excel, PDFs, carpetas, logging).

- [interfaz.py](file:///c:/Proyectos%20Comodin/pagos_paypal/interfaz.py)  
  Interfaz gráfica `PaymentApp` que orquesta todo el flujo paso a paso.

- [build.bat](file:///c:/Proyectos%20Comodin/pagos_paypal/build.bat)  
  Script para instalar dependencias y generar el ejecutable con PyInstaller.

- [requirements.txt](file:///c:/Proyectos%20Comodin/pagos_paypal/requirements.txt)  
  Lista de librerías necesarias.

## Clases y funciones principales (detalle)
Archivo [main.py](file:///c:/Proyectos%20Comodin/pagos_paypal/main.py):

- `class Config`  
  - `BASE_PAYPAL` / `RUTA_DESCARGAS` / `RUTA_MAESTRO` / `RUTAS_PDF`: rutas principales de trabajo.  
  - `SAP_URL` / `SAP_USER` / `SAP_PASSWORD` / `SAP_TRANSACCION` / `SAP_CUENTA` / `SAP_SOCIEDAD`: parámetros para conectarse a SAP.  
  - `COLUMNAS_SEGUNDA_HOJA`: orden y nombres de las columnas de la segunda hoja de Excel.  
  - `TIMEOUT_SAP`, `TIMEOUT_DOWNLOAD`: tiempos máximos de espera para SAP y descargas.  
  - `ACTIVAR_LOG_ARCHIVO`: activa o desactiva la generación de archivo de log.

- `configurar_logging()`  
  - Configura el sistema de logs.  
  - Siempre envía mensajes a la consola.  
  - Si `Config.ACTIVAR_LOG_ARCHIVO` es `True`, crea carpeta `logs` y archivo de log con fecha y hora.

- `class GestorCarpetas`  
  - `__init__(base_path)`  
    - Recibe la ruta base (`BASE_PAYPAL`) y  la guarda para operar sobre ella.  
  - `obtener_pago_pendiente_o_siguiente()`  
    - Escanea las carpetas existentes `Pago #N`.  
    - Devuelve el número de pago pendiente o el siguiente número disponible.  
  - `crear_estructura_pago(numero_pago)`  
    - Crea (si no existen) las carpetas:  
      - `Pago #N` (principal).  
      - carpeta de soportes (por ejemplo `Pago #N\Soporte`).  
    - Devuelve las rutas de ambas carpetas.

- `class DescargadorSAP`  
  - `__init__()`  
    - Inicializa el logger y la ruta de descargas (`Config.RUTA_DESCARGAS`).  
  - `configurar_firefox()`  
    - Crea un perfil de Firefox con:  
      - Carpeta de descargas configurada.  
      - Tipos MIME de Excel marcados para “guardar sin preguntar”.  
    - Devuelve una instancia de `webdriver.Firefox`.  
  - `esperar_descarga(timeout, patron_alternativo)`  
    - Vigila la carpeta de descargas hasta encontrar un archivo Excel válido.  
    - Acepta archivos tipo `EXPORT_*.xlsx` o que coincidan con `patron_alternativo` (por ejemplo `pago 12*`).  
    - Verifica que el archivo no esté incompleto (`.part`) y que el tamaño sea razonable.  
  - `_esperar_y_hacer(wait, by, selector, accion, valor, descripcion)`  
    - Envoltura para `WebDriverWait` con mensajes de error descriptivos.  
    - Puede esperar presencia o “clickable” y opcionalmente escribir texto.  
  - `descargar_reporte_sap(numero_pago)`  
    - Abre SAP en el navegador.  
    - Hace login con `SAP_USER` y `SAP_PASSWORD`.  
    - Navega a la transacción `FAGLL03` y llena campos de cuenta y sociedad.  
    - Ejecuta la consulta (botón F8).  
    - Envía **Mayus + F4** para abrir la exportación a hoja de cálculo.  
    - Escribe el nombre del archivo (`pago N`) en el campo de nombre.  
    - Pulsa el botón “Generar” del cuadro de exportación.  
    - Intenta detectar y pulsar los distintos diálogos de confirmación de descarga.  
    - Llama a `esperar_descarga()` y devuelve la ruta del archivo descargado, o `None` si falla.

- `class ProcesadorExcel`  
  - `__init__()`  
    - Inicializa el logger para operaciones sobre Excel.  
  - `buscar_archivo_pago_en_descargas(numero_pago)`  
    - Busca en `RUTA_DESCARGAS` un archivo cuyo nombre contenga `pago N` en distintas variantes (`"pago 12"`, `"Pago#12"`, etc.).  
    - Devuelve la ruta al archivo encontrado o `None`.  
  - `mover_y_renombrar_descarga(archivo_descarga, carpeta_destino, numero_pago)`  
    - Renombra el archivo descargado con un formato estándar `EXPORT_YYYYMMDD_Pago#N.xlsx`.  
    - Lo mueve desde Descargas a la carpeta `Pago #N`.  
  - `reorganizar_columnas_primera_hoja(archivo_principal)`  
    - Reordena las columnas de la primera hoja (SAP) para dejar “Referencia” al inicio y limpiar encabezados.  
  - `crear_segunda_hoja(archivo_principal, ruta_maestro)`  
    - Lee el maestro en la hoja configurada.  
    - Filtra los registros por mes/año actual según `Fecha_pago`.  
    - Normaliza los nombres de columnas del maestro y los mapea contra `COLUMNAS_SEGUNDA_HOJA`.  
    - Llena columnas como `Date`, `Currency`, `Gross`, `Fee`, `Net`, `Flete`, `Valor mcia`, `Invoice Numbers`, etc.  
    - Detecta pagos parciales (Net > 0 con Gross/Fee = 0) y marca `"Proximo pago"` en `Observaciones`.  
    - Devuelve un `DataFrame` con la segunda hoja ya armada.  
  - `calcular_mon_grupo_y_diferencia(archivo_principal, df_segunda_hoja)`  
    - Lee de la primera hoja el valor “Mon.grupo/Valoración grupo” por `Referencia`.  
    - Lo mapea a la segunda hoja usando `Invoice Numbers`.  
    - Calcula `Valoración flete` y la columna de `Diferencia` entre lo cobrado y lo registrado en SAP.  
  - `guardar_excel_con_dos_hojas(archivo_principal, df_segunda_hoja)`  
    - Abre el Excel principal.  
    - Actualiza/crea la segunda hoja con el `DataFrame` recibido.  
    - Asegura tipos correctos (por ejemplo `Order Id Paypal` como texto para evitar notación científica).

- `class GestorPDFs`  
  - `__init__(rutas_pdf)`  
    - Recibe una lista de carpetas donde buscar los PDFs de facturas y guías.  
  - `procesar_documentos_soporte(df, carpeta_soporte)`  
    - Por cada registro de la segunda hoja:  
      - Busca PDFs de factura y guía según los números (invoice, guía).  
      - Verifica fechas de envío para detectar soportes antiguos.  
      - Escribe en `Observaciones` mensajes como:  
        - “Faltan ambos documentos”.  
        - “Falta la guia de transporte”.  
        - “Falta la factura comercial y Fecha anterior registrada ...”.  
        - “Soportes OK”.  
    - Respeta los registros ya marcados como “Proximo pago” para no sobrescribir esa observación.

- `main()`  
  - Punto de entrada en **modo consola**.  
  - Fases:
    1. Gestión de carpetas (`GestorCarpetas`).  
    2. Descarga de SAP (`DescargadorSAP`).  
    3. Procesamiento de Excel (primera hoja) (`ProcesadorExcel`).  
    4. Creación de segunda hoja y detección de pagos parciales (`ProcesadorExcel`).  
    5. Procesamiento de PDFs (`GestorPDFs`).  
    6. Guardado del Excel final con dos hojas.  
  - Registra en el log un resumen final con ruta del archivo y estado de soportes.

Archivo [interfaz.py](file:///c:/Proyectos%20Comodin/pagos_paypal/interfaz.py):

- `class PaymentApp(ctk.CTk)`  
  - Constructor `__init__()`  
    - Inicializa ventana, tema y colores.  
    - Configura logging usando `configurar_logging()`.  
    - Define los pasos del workflow (carpetas, SAP, Excel, PDFs, maestro).  
    - Crea los tres “estados” de pantalla: inicio, ejecución y completado.  
    - Crea la barra de estado inferior con mensaje y barra de progreso.  
    - Llama a `load_initial_state()` para precargar el número de pago.  
  - `create_idle_content()`  
    - Construye la pantalla inicial de bienvenida.  
    - Muestra la descripción del sistema.  
    - Campo para ingresar el número de pago.  
    - Botón grande “ EJECUTAR PROCESO COMPLETO”.  
  - `create_running_content()`  
    - Construye la pantalla de ejecución.  
    - Barra de progreso grande.  
    - Lista de pasos con iconos de estado (pendiente, en proceso, completado).  
    - Área de logs en tiempo real.  
  - `create_completed_content()`  
    - Construye la pantalla de resultado.  
    - Muestra ícono de éxito, mensaje y resumen del pago ejecutado.  
    - Incluye botones `CONTINUAR` (siguiente pago) y `FINALIZAR`.  
  - `show_state(state)`  
    - Oculta/mostrará los frames `idle`, `running` o `completed`.  
    - Actualiza la barra de estado según el estado actual.  
  - `update_status(message, progress)`  
    - Cambia el texto de la barra de estado.  
    - Actualiza el porcentaje de la barra inferior.  
  - `load_initial_state()`  
    - Llama a `GestorCarpetas` para obtener el siguiente pago pendiente.  
    - Llena el campo de número de pago en la pantalla inicial.  
  - `start_workflow()`  
    - Valida que no haya otra operación en curso.  
    - Valida que el número de pago sea numérico.  
    - Cambia al estado `RUNNING`.  
    - Limpia y reinicia la lista de pasos.  
    - Lanza un hilo (`_workflow_thread`) para no bloquear la GUI.  
  - `_workflow_thread()`  
    - Hilo que ejecuta secuencialmente:
      - `_verify_folders()`  
      - `_download_from_sap()`  
      - `_process_excel()`  
      - `_search_pdfs()`  
      - `_update_master()`  
    - Al terminar, llama a `_on_workflow_completed()`.  
  - `run_step(step_id, step_function)`  
    - Marca un paso como “en progreso”, ejecuta la función y luego lo marca como “completado”.  
    - Actualiza la barra de progreso y el porcentaje.  
  - `log_message(message)`  
    - Añade una línea al área de logs con marca de tiempo.  
  - `_verify_folders()`  
    - Paso 1: muestra cuántas carpetas `Pago #N` existen y registra en el log.  
  - `_download_from_sap()`  
    - Paso 2: usa `DescargadorSAP` para descargar el Excel desde SAP para el número de pago actual.  
    - Informa en logs si el archivo se encontró o si hay que revisar Descargas.  
  - `_process_excel()`  
    - Paso 3: busca el archivo del pago en Descargas.  
    - Crea la estructura de carpetas del pago.  
    - Mueve y renombra el archivo.  
    - Reorganiza columnas de la primera hoja.  
    - Crea la segunda hoja a partir del maestro.  
    - Calcula `Valoración flete` y `Diferencia`.  
    - Guarda un Excel con las dos hojas (versión preliminar).  
  - `_search_pdfs()`  
    - Paso 4: si existe la segunda hoja y la carpeta de soportes, llama a `GestorPDFs`.  
    - Actualiza `Observaciones` según los PDFs encontrados.  
    - Guarda el Excel final actualizado.  
  - `_update_master()`  
    - Paso 5: placeholder para futura lógica de actualización directa del maestro global.  
    - Actualmente informa en logs la ruta del archivo listo.  
  - `_on_workflow_completed()`  
    - Marca el proceso como terminado.  
    - Pone la barra de progreso al 100 %.  
    - Genera un resumen del pago ejecutado y cambia a la pantalla de “Proceso Completado”.  
  - `continue_workflow()`  
    - Incrementa el número de pago.  
    - Vuelve al estado inicial para ejecutar otro ciclo.  
  - `finish_workflow()` / `on_close()`  
    - Gestionan el cierre limpio de la aplicación, preguntando si hay una operación en curso.

## Requisitos previos
- **Python 3.11 o superior** instalado en Windows.
- **Firefox** y **geckodriver** instalados y accesibles en PATH.
- Acceso a SAP con el usuario configurado en `Config.SAP_USER`.
- Archivo maestro de Courier en la ruta definida en `Config.RUTA_MAESTRO`.

## Instalación del entorno
Desde la carpeta del proyecto `c:\Proyectos Comodin\pagos_paypal`:

1. Crear entorno virtual:

   ```bash
   python -m venv venv
   ```

2. Activar entorno virtual (PowerShell):

   ```bash
   .\venv\Scripts\Activate.ps1
   ```

3. Instalar dependencias:

   ```bash
   python -m pip install --upgrade pip
   python -m pip install -r requirements.txt
   ```

## Uso: interfaz gráfica (recomendado)

1. Activar el entorno virtual:

   ```bash
   cd "c:\Proyectos Comodin\pagos_paypal"
   .\venv\Scripts\Activate.ps1
   ```

2. Ejecutar la interfaz:

   ```bash
   python interfaz.py
   ```

3. Flujo desde la GUI:
   - Verifica que aparece el número de pago sugerido (tomado de las carpetas existentes).  
   - Ajusta el número de pago si es necesario.  
   - Pulsa **“🚀 EJECUTAR PROCESO COMPLETO”**.  
   - Observa el avance de los pasos:
     1. Verificar carpetas  
     2. Descargar de SAP  
     3. Procesar Excel  
     4. Buscar PDFs  
     5. Actualizar maestro (pendiente de integración directa)  
   - Al finalizar, revisa el resumen y la carpeta `Pago #N` generada/actualizada.

## Uso: modo consola (opcional)

1. Activar el entorno virtual.
2. Ejecutar:

   ```bash
   python main.py
   ```

El script detecta el siguiente pago pendiente, descarga, procesa y genera el Excel final sin interfaz gráfica.

## Generar ejecutable (.exe) con PyInstaller

El script [build.bat](file:///c:/Proyectos%20Comodin/pagos_paypal/build.bat) automatiza la creación del ejecutable de la interfaz.

Pasos:

1. Doble clic en `build.bat` o desde PowerShell:

   ```bash
   cd "c:\Proyectos Comodin\pagos_paypal"
   .\build.bat
   ```

2. Cuando el script lo solicite, ingresa la **ruta de salida** para la carpeta que contendrá el exe  
   (por ejemplo `C:\Finanzas\Apps\PayPalPagosDist`).  
   Si dejas vacío, usará `dist` dentro del proyecto.

3. Al finalizar, tendrás algo como:

   ```text
   C:\Finanzas\Apps\PayPalPagosDist\PayPalPagos\PayPalPagos.exe
   ```

Ese exe abre directamente la interfaz gráfica.

## Comandos útiles

- Instalar dependencias:

  ```bash
  python -m pip install -r requirements.txt
  ```

- Ejecutar GUI:

  ```bash
  python interfaz.py
  ```

- Ejecutar modo consola:

  ```bash
  python main.py
  ```

- Construir ejecutable:

  ```bash
  .\build.bat
  ```

## Notas de versión 1.0.0

- Se unificó el flujo de SAP, Excel y PDFs en un solo sistema.  
- Se añadió interfaz gráfica moderna con seguimiento de pasos.  
- Se automatizó el envío de **Mayus+F4** y el nombrado del archivo `pago N` en SAP.  
- Se implementó detección de pagos parciales y observaciones en la segunda hoja.  
- Se agregó script de build con salida configurable para el ejecutable.
