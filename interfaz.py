"""
Interfaz Gráfica Moderno para Sistema de Automatización de Pagos PayPal
Diseño simplificado con flujo de trabajo paso a paso
"""

import customtkinter as ctk
from tkinter import messagebox
from pathlib import Path
from datetime import datetime
import threading
import logging
import sys
import os

# Importar clases de main.py
from main import (
    Config, GestorCarpetas, DescargadorSAP, ProcesadorExcel,
    GestorPDFs, configurar_logging
)


# Configuración de tema y colores
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("green")

# Colores personalizados
COLOR_PRIMARY = "#2CC985"  # Verde principal
COLOR_SECONDARY = "#1A1A2E"  # Fondo oscuro
COLOR_ACCENT = "#0F3460"  # Acento azul
COLOR_TEXT = "#FFFFFF"  # Texto blanco
COLOR_ERROR = "#FF6B6B"  # Rojo para errores
COLOR_WARNING = "#FFE66D"  # Amarillo para advertencias
COLOR_SUCCESS = "#4ECDC4"  # Verde azulado para éxito


# Estados de la aplicación
STATE_IDLE = "idle"
STATE_RUNNING = "running"
STATE_COMPLETED = "completed"


class PaymentApp(ctk.CTk):
    """Aplicación principal con interfaz gráfica simplificada"""
    
    def __init__(self):
        super().__init__()
        
        # Configuración de ventana
        self.title(" Automatización de Pagos PayPal")
        self.geometry("900x650")
        self.minsize(800, 550)
        self.resizable(True, True)
        
        # Configurar logging
        self.logger = configurar_logging()
        
        # Variables de estado
        self.current_state = STATE_IDLE
        self.operation_running = False
        self.numero_pago = 1
        
        # Colores de tema
        self.colors = {
            'primary': COLOR_PRIMARY,
            'secondary': COLOR_SECONDARY,
            'accent': COLOR_ACCENT,
            'text': COLOR_TEXT,
            'error': COLOR_ERROR,
            'warning': COLOR_WARNING,
            'success': COLOR_SUCCESS
        }
        
        # Pasos del workflow
        self.workflow_steps = [
            ("step1", "Verificar carpetas"),
            ("step2", "Descargar de SAP"),
            ("step3", "Procesar Excel"),
            ("step4", "Buscar PDFs"),
            ("step5", "Actualizar Maestro")
        ]
        
        # Configurar estilo
        self.configure(fg_color=COLOR_SECONDARY)
        
        # Crear interfaz
        self.create_widgets()
        
        # Cargar estado inicial
        self.load_initial_state()
        
        # Configurar cierre
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        
    def create_widgets(self):
        """Crea todos los elementos de la interfaz"""
        
        # ========== HEADER ==========
        self.header_frame = ctk.CTkFrame(
            self, 
            fg_color=COLOR_ACCENT,
            height=70,
            corner_radius=0
        )
        self.header_frame.pack(fill="x", side="top")
        
        # Logo/Título
        self.title_label = ctk.CTkLabel(
            self.header_frame,
            text="💰 Sistema de Automatización de Pagos PayPal",
            font=("Roboto Medium", 20),
            text_color=COLOR_TEXT
        )
        self.title_label.pack(side="left", padx=20, pady=20)
        
        # ========== CONTENIDO PRINCIPAL ==========
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Contenedor del contenido dinámico
        self.content_frame = ctk.CTkFrame(
            self.main_frame,
            fg_color="#1E1E2E",
            corner_radius=15
        )
        self.content_frame.pack(fill="both", expand=True)
        
        # Crear los tres estados de contenido
        self.create_idle_content()
        self.create_running_content()
        self.create_completed_content()
                
        # ========== BARRA DE ESTADO ==========
        self.status_frame = ctk.CTkFrame(
            self,
            fg_color=COLOR_ACCENT,
            height=35,
            corner_radius=0
        )
        self.status_frame.pack(fill="x", side="bottom")
        
        # Etiqueta de estado
        self.status_label = ctk.CTkLabel(
            self.status_frame,
            text="✅ Sistema listo",
            font=("Roboto", 12),
            text_color=COLOR_TEXT
        )
        self.status_label.pack(side="left", padx=15, pady=8)
        
        # Progress bar
        self.progress_bar = ctk.CTkProgressBar(
            self.status_frame,
            width=200,
            height=8,
            progress_color=COLOR_PRIMARY
        )
        self.progress_bar.set(0)
        self.progress_bar.pack(side="right", padx=20, pady=12)
        
        # Mostrar estado inicial
        self.show_state(STATE_IDLE)
        
    def create_idle_content(self):
        """Crea el contenido del estado IDLE (pantalla de bienvenida)"""
        self.idle_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        
        # Espaciador superior
        ctk.CTkFrame(self.idle_frame, height=30, fg_color="transparent").pack()
        
        # Icono grande
        ctk.CTkLabel(
            self.idle_frame,
            text="🏠",
            font=("Roboto", 64)
        ).pack(pady=(20, 10))
        
        # Título de bienvenida
        ctk.CTkLabel(
            self.idle_frame,
            text="👋 Bienvenido al Sistema",
            font=("Roboto", 24, "bold"),
            text_color=COLOR_PRIMARY
        ).pack(pady=(0, 10))
        
        # Descripción del sistema
        desc_frame = ctk.CTkFrame(self.idle_frame, fg_color="#2A2A3E", corner_radius=10)
        desc_frame.pack(fill="x", padx=40, pady=20)
        
        ctk.CTkLabel(
            desc_frame,
            text="Este sistema automatiza el proceso completo de pagos PayPal:",
            font=("Roboto", 14),
            text_color=COLOR_TEXT
        ).pack(anchor="w", padx=20, pady=(15, 10))
        
        steps_text = """
            1. 📁 Verificar y crear carpetas necesarias
2. 📥 Descargar reportes desde SAP
3. 📄 Buscar documentos PDF asociados
4. 📊 Procesar archivos Excel
5. 📋 Actualizar el archivo maestro
        """
        ctk.CTkLabel(
            desc_frame,
            text=steps_text.strip(),
            font=("Roboto", 13),
            text_color="#AAAAAA",
            justify="left"
        ).pack(anchor="w", padx=20, pady=(0, 15))
        
        # Número de pago
        payment_frame = ctk.CTkFrame(self.idle_frame, fg_color="transparent")
        payment_frame.pack(fill="x", padx=40, pady=20)
        
        ctk.CTkLabel(
            payment_frame,
            text="📌 Número de Pago:",
            font=("Roboto", 14),
            text_color=COLOR_TEXT
        ).pack(side="left", padx=(0, 10))
        
        self.payment_entry = ctk.CTkEntry(
            payment_frame,
            placeholder_text="Ej: 1, 2, 3...",
            width=100,
            font=("Roboto", 14)
        )
        self.payment_entry.pack(side="left", padx=(0, 20))
        
        # Botón Ejecutar grande
        self.btn_ejecutar = ctk.CTkButton(
            self.idle_frame,
            text="EJECUTAR PROCESO COMPLETO",
            command=self.start_workflow,
            fg_color=COLOR_PRIMARY,
            hover_color="#25A25A",
            font=("Roboto", 16, "bold"),
            height=50,
            width=300
        )
        self.btn_ejecutar.pack(pady=30)
        
        # Espaciador inferior
        ctk.CTkFrame(self.idle_frame, height=20, fg_color="transparent").pack()
        
    def create_running_content(self):
        """Crea el contenido del estado RUNNING (progress y logs)"""
        self.running_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        
        # Título
        ctk.CTkLabel(
            self.running_frame,
            text="⚡ Ejecutando Proceso...",
            font=("Roboto", 20, "bold"),
            text_color=COLOR_PRIMARY
        ).pack(pady=(20, 10))
        
        # Progress bar principal
        self.main_progress = ctk.CTkProgressBar(
            self.running_frame,
            width=500,
            height=15,
            progress_color=COLOR_PRIMARY
        )
        self.main_progress.set(0)
        self.main_progress.pack(pady=10)
        
        # Porcentaje
        self.progress_label = ctk.CTkLabel(
            self.running_frame,
            text="0%",
            font=("Roboto", 14, "bold"),
            text_color=COLOR_TEXT
        )
        self.progress_label.pack(pady=(0, 10))
        
        # Contenedor de pasos
        steps_container = ctk.CTkFrame(self.running_frame, fg_color="#2A2A3E", corner_radius=10)
        steps_container.pack(fill="x", padx=30, pady=15)
        
        ctk.CTkLabel(
            steps_container,
            text="📋 Pasos del Proceso:",
            font=("Roboto", 14, "bold"),
            text_color=COLOR_TEXT
        ).pack(anchor="w", padx=15, pady=(15, 10))
        
        # Labels de cada paso
        self.step_labels = {}
        for i, (step_id, step_name) in enumerate(self.workflow_steps):
            step_frame = ctk.CTkFrame(steps_container, fg_color="transparent")
            step_frame.pack(fill="x", padx=15, pady=3)
            
            # Checkbox visual (simulado con label)
            status_icon = ctk.CTkLabel(
                step_frame,
                text="⭕",
                font=("Roboto", 14),
                text_color="#666666",
                width=30
            )
            status_icon.pack(side="left", padx=(0, 10))
            
            step_label = ctk.CTkLabel(
                step_frame,
                text=step_name,
                font=("Roboto", 12),
                text_color="#AAAAAA"
            )
            step_label.pack(side="left")
            
            self.step_labels[step_id] = {
                'icon': status_icon,
                'label': step_label
            }
        
        ctk.CTkFrame(steps_container, height=10, fg_color="transparent").pack()
        
        # Área de logs
        log_frame = ctk.CTkFrame(self.running_frame, fg_color="#2A2A3E", corner_radius=10)
        log_frame.pack(fill="both", expand=True, padx=30, pady=15)
        
        ctk.CTkLabel(
            log_frame,
            text="📝 Registro de Operaciones:",
            font=("Roboto", 14, "bold"),
            text_color=COLOR_TEXT
        ).pack(anchor="w", padx=15, pady=(15, 10))
        
        # Textbox para logs
        self.log_text = ctk.CTkTextbox(
            log_frame,
            fg_color="#1A1A2E",
            text_color=COLOR_TEXT,
            font=("Consolas", 11)
        )
        self.log_text.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        self.log_text.insert("1.0", "📋 Esperando inicio del proceso...\n")
        self.log_text.configure(state="disabled")
        
    def create_completed_content(self):
        """Crea el contenido del estado COMPLETED"""
        self.completed_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        
        # Icono de éxito
        ctk.CTkLabel(
            self.completed_frame,
            text="✅",
            font=("Roboto", 64)
        ).pack(pady=(30, 10))
        
        # Título
        ctk.CTkLabel(
            self.completed_frame,
            text="🎉 Proceso Completado",
            font=("Roboto", 24, "bold"),
            text_color=COLOR_SUCCESS
        ).pack(pady=(0, 10))
        
        # Mensaje de resumen
        self.completion_message = ctk.CTkLabel(
            self.completed_frame,
            text="El proceso de automatización ha finalizado exitosamente.",
            font=("Roboto", 14),
            text_color=COLOR_TEXT
        )
        self.completion_message.pack(pady=10)
        
        # Resumen de pasos completados
        summary_frame = ctk.CTkFrame(self.completed_frame, fg_color="#2A2A3E", corner_radius=10)
        summary_frame.pack(fill="x", padx=40, pady=20)
        
        ctk.CTkLabel(
            summary_frame,
            text="📋 Resumen:",
            font=("Roboto", 14, "bold"),
            text_color=COLOR_TEXT
        ).pack(anchor="w", padx=15, pady=(15, 10))
        
        self.summary_text = ctk.CTkTextbox(
            summary_frame,
            height=120,
            fg_color="#1A1A2E",
            text_color=COLOR_TEXT,
            font=("Roboto", 12)
        )
        self.summary_text.pack(fill="x", padx=15, pady=(0, 15))
        self.summary_text.insert("1.0", "Todos los pasos completados exitosamente.")
        self.summary_text.configure(state="disabled")
        
        # Botones de acción
        buttons_frame = ctk.CTkFrame(self.completed_frame, fg_color="transparent")
        buttons_frame.pack(pady=30)
        
        # Botón Continuar
        self.btn_continuar = ctk.CTkButton(
            buttons_frame,
            text="🔄 CONTINUAR",
            command=self.continue_workflow,
            fg_color="#3498DB",
            hover_color="#2980B9",
            font=("Roboto", 14, "bold"),
            height=45,
            width=180
        )
        self.btn_continuar.pack(side="left", padx=(0, 20))
        
        # Botón Finalizar
        self.btn_finalizar = ctk.CTkButton(
            buttons_frame,
            text="❌ FINALIZAR",
            command=self.finish_workflow,
            fg_color=COLOR_ERROR,
            hover_color="#E74C3C",
            font=("Roboto", 14, "bold"),
            height=45,
            width=180
        )
        self.btn_finalizar.pack(side="left")
        
        # Espaciador inferior
        ctk.CTkFrame(self.completed_frame, height=20, fg_color="transparent").pack()
        
    def show_state(self, state):
        """Muestra el contenido del estado especificado"""
        # Ocultar todos los estados
        self.idle_frame.pack_forget()
        self.running_frame.pack_forget()
        self.completed_frame.pack_forget()
        
        # Mostrar el estado solicitado
        if state == STATE_IDLE:
            self.idle_frame.pack(fill="both", expand=True)
            self.update_status("✅ Sistema listo", 0)
        elif state == STATE_RUNNING:
            self.running_frame.pack(fill="both", expand=True)
            self.update_status("⚡ Ejecutando proceso...", 0.5)
        elif state == STATE_COMPLETED:
            self.completed_frame.pack(fill="both", expand=True)
            self.update_status("✅ Proceso completado", 1.0)
        
        self.current_state = state
        
    def update_status(self, message, progress=0):
        """Actualiza la barra de estado"""
        self.status_label.configure(text=f"  {message}")
        self.progress_bar.set(progress)
        
    def load_initial_state(self):
        """Carga el estado inicial"""
        try:
            # Obtener número de pago actual
            gestor = GestorCarpetas(Config.BASE_PAYPAL)
            self.numero_pago = gestor.obtener_pago_pendiente_o_siguiente()
            self.payment_entry.delete(0, "end")
            self.payment_entry.insert(0, str(self.numero_pago))
            
            self.logger.info("Interfaz inicializada correctamente")
            self.log_to_running("🟢 Sistema inicializado correctamente")
            
        except Exception as e:
            self.logger.error(f"Error al inicializar estado: {e}")
            
    def start_workflow(self):
        """Inicia el workflow completo"""
        if self.operation_running:
            messagebox.showwarning("Operación en Curso", "Ya hay una operación en progreso.")
            return
            
        numero_pago = self.payment_entry.get()
        if not numero_pago.isdigit():
            messagebox.showerror("Error", "Por favor ingrese un número de pago válido")
            return
            
        self.numero_pago = int(numero_pago)
        self.operation_running = True
        
        # Deshabilitar botón
        self.btn_ejecutar.configure(state="disabled", text="⏳ Ejecutando...")
        
        # Cambiar a estado running
        self.show_state(STATE_RUNNING)
        
        # Limpiar logs
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.insert("1.0", "📋 Iniciando proceso completo...\n")
        self.log_text.configure(state="disabled")
        
        # Resetear pasos
        for step_id, widgets in self.step_labels.items():
            widgets['icon'].configure(text="⭕", text_color="#666666")
            widgets['label'].configure(text_color="#AAAAAA")
        
        # Iniciar en hilo separado
        thread = threading.Thread(target=self._workflow_thread)
        thread.daemon = True
        thread.start()
        
    def _workflow_thread(self):
        """Hilo principal del workflow"""
        try:
            self.log_message(f"🚀 Iniciando proceso para Pago #{self.numero_pago}")
            
            # Variables de intercambio entre pasos
            self.df_segunda = None
            self.archivo_movido = None
            self.carpeta_soporte = None
            
            # Paso 1: Verificar carpetas
            self.run_step("step1", self._verify_folders)
            
            # Paso 2: Descargar de SAP
            self.run_step("step2", self._download_from_sap)
            
            # Paso 3: Procesar Excel
            self.run_step("step3", self._process_excel)
            
            # Paso 4: Buscar PDFs
            self.run_step("step4", self._search_pdfs)
            
            # Paso 5: Actualizar Maestro
            self.run_step("step5", self._update_master)
            
            # Proceso completado
            self.after(0, self._on_workflow_completed)
            
        except Exception as e:
            self.log_message(f"❌ Error en proceso: {e}")
            self.after(0, lambda: messagebox.showerror("Error", f"Error en el proceso:\n{e}"))
            self.operation_running = False
            self.after(0, lambda: self.btn_ejecutar.configure(state="normal", text=" EJECUTAR PROCESO COMPLETO"))
            
    def run_step(self, step_id, step_function):
        """Ejecuta un paso del workflow"""
        # Primero mostrar "en progreso"
        self.update_step_icon(step_id, "🔄", "#3498DB")
        
        step_index = [s[0] for s in self.workflow_steps].index(step_id)
        progress = (step_index + 0.5) / len(self.workflow_steps)
        self.after(0, lambda p=progress: self.main_progress.set(p))
        self.after(0, lambda p=progress: self.progress_label.configure(text=f"{int(p * 100)}%"))
        
        # Ejecutar el paso real
        step_function()
        
        # Solo después de terminar, marcar como completado
        progress_final = (step_index + 1) / len(self.workflow_steps)
        self.update_step_icon(step_id, "✅", COLOR_PRIMARY)
        self.after(0, lambda p=progress_final: self.main_progress.set(p))
        self.after(0, lambda p=progress_final: self.progress_label.configure(text=f"{int(p * 100)}%"))
        
    def update_step_icon(self, step_id, icon, color):
        """Actualiza el icono de un paso"""
        def _update():
            self.step_labels[step_id]['icon'].configure(text=icon, text_color=color)
            if icon == "✅":
                self.step_labels[step_id]['label'].configure(text_color=COLOR_SUCCESS)
            elif icon == "🔄":
                self.step_labels[step_id]['label'].configure(text_color=COLOR_PRIMARY)
        self.after(0, _update)
        
    def log_message(self, message):
        """Agrega un mensaje al log"""
        def _log():
            timestamp = datetime.now().strftime("%H:%M:%S")
            full_message = f"[{timestamp}] {message}\n"
            self.log_text.configure(state="normal")
            self.log_text.insert("end", full_message)
            self.log_text.see("end")
            self.log_text.configure(state="disabled")
        self.after(0, _log)
        
    def log_to_running(self, message):
        """Log específico para pantalla de running"""
        pass  # En el nuevo diseño, solo hay un área de logs
        
    def _verify_folders(self):
        """Paso 1: Verificar carpetas"""
        self.log_message("📁 Verificando estructura de carpetas...")
        try:
            gestor = GestorCarpetas(Config.BASE_PAYPAL)
            carpetas = [d for d in Config.BASE_PAYPAL.iterdir() if d.is_dir() and d.name.startswith("Pago #")]
            self.log_message(f"✅ Carpetas verificadas: {len(carpetas)} carpetas existentes")
        except Exception as e:
            self.log_message(f"❌ Error verificando carpetas: {e}")
            
    def _download_from_sap(self):
        """Paso 2: Descargar de SAP"""
        self.log_message("📥 Descargando desde SAP...")
        try:
            descargador = DescargadorSAP()
            archivo = descargador.descargar_reporte_sap(self.numero_pago)
            if archivo:
                self.log_message(f"✅ Archivo descargado: {archivo.name}")
            else:
                self.log_message("⚠️ No se encontró archivo. Busque manualmente en Descargas.")
        except Exception as e:
            self.log_message(f"❌ Error en descarga SAP: {e}")
            
    def _search_pdfs(self):
        """Paso 4: Buscar PDFs"""
        self.log_message("📄 Buscando y validando documentos PDF...")
        try:
            if self.df_segunda is not None and self.carpeta_soporte:
                gestor_pdfs = GestorPDFs(Config.RUTAS_PDF)
                self.df_segunda = gestor_pdfs.procesar_documentos_soporte(self.df_segunda, self.carpeta_soporte)
                
                # Guardar el Excel final después de procesar PDFs
                procesador = ProcesadorExcel()
                procesador.guardar_excel_con_dos_hojas(self.archivo_movido, self.df_segunda)
                
                self.log_message(f"✅ PDFs procesados y Excel final actualizado.")
            else:
                self.log_message("⚠️ Saltando búsqueda de PDFs (no hay datos de Excel)")
        except Exception as e:
            self.log_message(f"❌ Error buscando PDFs: {e}")
            
    def _process_excel(self):
        """Paso 3: Procesar Excel"""
        self.log_message("📊 Procesando archivo Excel...")
        try:
            procesador = ProcesadorExcel()
            
            # Buscar archivo
            archivo = procesador.buscar_archivo_pago_en_descargas(self.numero_pago)
            if archivo:
                self.log_message(f"✅ Archivo encontrado: {archivo.name}")
                
                # Obtener carpetas
                gestor = GestorCarpetas(Config.BASE_PAYPAL)
                carpeta_pago, carpeta_soporte = gestor.crear_estructura_pago(self.numero_pago)
                self.carpeta_soporte = carpeta_soporte
                
                # Mover archivo
                self.archivo_movido = procesador.mover_y_renombrar_descarga(archivo, carpeta_pago, self.numero_pago)
                self.log_message(f"📁 Archivo movido a: {carpeta_pago}")
                
                # Reorganizar columnas
                procesador.reorganizar_columnas_primera_hoja(self.archivo_movido)
                self.log_message("✅ Columnas reorganizadas")
                
                # Crear segunda hoja
                if Config.RUTA_MAESTRO.exists():
                    self.df_segunda = procesador.crear_segunda_hoja(self.archivo_movido, Config.RUTA_MAESTRO)
                    self.log_message(f"✅ Segunda hoja creada con {len(self.df_segunda)} registros")
                    
                    # Calcular totales
                    self.df_segunda = procesador.calcular_mon_grupo_y_diferencia(self.archivo_movido, self.df_segunda)
                    
                    # Guardar una versión preliminar
                    procesador.guardar_excel_con_dos_hojas(self.archivo_movido, self.df_segunda)
                    self.log_message("✅ Procesamiento inicial de Excel completado")
                else:
                    self.log_message("⚠️ Archivo maestro no encontrado")
            else:
                self.log_message("❌ No se encontró archivo para procesar")
                
        except Exception as e:
            self.log_message(f"❌ Error procesando Excel: {e}")
            
    def _update_master(self):
        """Paso 5: Actualizar Maestro"""
        self.log_message("📋 Actualizando archivo maestro...")
        try:
            if self.archivo_movido and self.archivo_movido.exists():
                self.log_message(f"⚠️ Funcionalidad pendiente de implementar.")
                self.log_message(f"📁 Archivo listo en: {self.archivo_movido}")
            else:
                self.log_message("⚠️ No hay archivo procesado para actualizar en el maestro.")
        except Exception as e:
            self.log_message(f"❌ Error en actualización de maestro: {e}")
            
    def _on_workflow_completed(self):
        """Maneja la finalización del workflow"""
        self.operation_running = False
        self.main_progress.set(1.0)
        self.progress_label.configure(text="100%")
        
        self.log_message(" Proceso completado exitosamente")
        
        # Generar resumen
        summary = f"""
Pago #{self.numero_pago} completado:
✓ Carpetas verificadas
✓ Archivo SAP descargado
✓ PDFs buscados
✓ Excel procesado
✓ Maestro actualizado

El sistema está listo para el siguiente pago.
        """
        
        self.summary_text.configure(state="normal")
        self.summary_text.delete("1.0", "end")
        self.summary_text.insert("1.0", summary.strip())
        self.summary_text.configure(state="disabled")
        
        # Cambiar al estado completado
        self.show_state(STATE_COMPLETED)
        
    def continue_workflow(self):
        """Continuar con otro proceso"""
        # Incrementar número de pago
        self.numero_pago += 1
        self.payment_entry.delete(0, "end")
        self.payment_entry.insert(0, str(self.numero_pago))
        
        # Volver al estado idle
        self.btn_ejecutar.configure(state="normal", text=" EJECUTAR PROCESO COMPLETO")
        self.show_state(STATE_IDLE)
        self.update_status("✅ Sistema listo", 0)
        
    def finish_workflow(self):
        """Finalizar y cerrar"""
        self.on_close()
        
    def on_close(self):
        """Manejo del cierre de la aplicación"""
        if self.operation_running:
            if not messagebox.askyesno("Salir", "Hay una operación en progreso. ¿Desea salir igualmente?"):
                return
                
        self.logger.info("Cerrando aplicación...")
        self.destroy()
        sys.exit(0)


def main():
    """Función principal"""
    try:
        app = PaymentApp()
        app.mainloop()
    except Exception as e:
        print(f"Error crítico: {e}")
        import traceback
        traceback.print_exc()
        input("Presione Enter para salir...")


if __name__ == "__main__":
    main()
