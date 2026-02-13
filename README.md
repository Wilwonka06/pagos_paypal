# SISTEMA DE DESCARGA Y VALIDACIÓN DE CAMBIARIOS EN PAGOS PAYPAL

## Descripción
Este proyecto automatiza el proceso de descarga y validación de cambios en pagos realizados a través de PayPal. Utiliza Selenium para interactuar con la interfaz web de PayPal, descargando archivos Excel con los detalles de los pagos y luego validando estos cambios en una base de datos local.

## Características
- Descarga automática de archivos Excel desde PayPal.
- Validación de cambios en los archivos descargados.
- Almacenamiento de datos en una base de datos SQLite local.
- Interfaz gráfica de usuario (GUI) para facilitar el uso.

## Requisitos
- Python 3.11
- Bibliotecas listadas en `requirements.txt`

## Instalación
1. Clona este repositorio:
   ```bash
   git clone https://github.com/tu_usuario/pagos_paypal.git
   cd pagos_paypal
   ```

2. Instala las dependencias:
   ```bash
   pip install -r requirements.txt
   ```

3. Ejecuta la aplicación:
   ```bash
   python main.py       
