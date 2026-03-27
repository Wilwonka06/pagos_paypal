import customtkinter as ctk
from tkinter import messagebox
from pathlib import Path
from datetime import datetime 
from config_manager import ConfiguradorRutasPayPal
import threading
import logging
import sys
import os

# Importar clases de main.py
from main import (
    Config, GestorCarpetas, DescargadorSAP, ProcesadorExcel,
    GestorPDFs, configurar_logging, resolver_rutas_swift_dinamicas
)

# NUEVO: Importar verificador/actualizador
from scripts.verificacion import VerificadorActualizadorSoportes, ResultadoVerificacion

# Configuración de tema y colores
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

# Colores personalizados (Formato: (Light, Dark))

COLOR_PRIMARY = ("#08129B")      # Azul principal (Claro / Oscuro)
COLOR_SECONDARY = ("#FFFFFF")    # Fondo principal ventana (Blanco / Gris neutro oscuro para descanso visual)
COLOR_ACCENT = ("#F0F2F5")       # Fondo de frames principales (Gris claro / Gris oscuro solicitado)
COLOR_ACCENT_LIGHT = ("#E4E6EB") # Acento para cards/subframes (Gris suave / Azul noche profundo)
COLOR_TEXT = ("#1C1E21")         # Texto (Casi negro / Blanco)
COLOR_TEXT_DIM = ("#65676B")     # Texto atenuado (Gris oscuro / Gris claro)

COLOR_ERROR = ("#D32F2F")        # Rojo
COLOR_WARNING = ("#F57C00")      # Naranja/Dorado (Mantener para advertencias críticas)
COLOR_SUCCESS = ("#388E3C")      # Verde
COLOR_BLUE = ("#1976D2")         # Azul info

# El color naranja anterior se reemplaza por el azul primario en los elementos de acción
COLOR_ACTION = COLOR_PRIMARY


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
        
        # Obtener dimensiones de pantalla para un inicio proporcional pero grande
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        
        # Usar el 90% de la pantalla pero no más de 1200x800 como base inicial
        width = min(1200, int(screen_width * 0.9))
        height = min(800, int(screen_height * 0.9))
        
        self.geometry(f"{width}x{height}+50+50")
        self.minsize(800, 600)
        self.resizable(True, True)
        
        # Forzar el estado maximizado después de que la ventana se haya inicializado un poco
        self.after(100, lambda: self.state('zoomed'))
        
        # Protocolo de cierre de ventana (X)
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        
        # Configurar logging
        self.logger = configurar_logging()
        
        # Variables de estado
        self.current_state = STATE_IDLE
        self.operation_running = False
        self.numero_pago = 1
        self.cancel_requested = False  # NUEVO: Bandera para cancelar proceso
        
        # NUEVO: Variables para verificación
        self.modo_verificacion = False
        self.pago_verificando = None
        self.resultados_verificacion = []
        
        # Variables para proceso completo
        self.archivo_movido = None
        self.df_segunda = None
        self.carpeta_soporte = None
        
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
        
        # Pasos del workflow principal
        self.workflow_steps = [
            ("step1", "Verificar carpetas"),
            ("step2", "Descargar de SAP"),
            ("step3", "Procesar Excel"),
            ("step4", "Buscar PDFs"),
            ("step5", "Actualizar Maestro")
        ]

        # Pasos del workflow de verificación
        self.verification_steps = [
            ("v_step1", "Identificar archivos"),
            ("v_step2", "Buscar y Copiar PDFs"),
            ("v_step3", "Analizar Observaciones"),
            ("v_step4", "Guardar Resultados")
        ]
        
        # Configurar estilo
        self.configure(fg_color=COLOR_SECONDARY)

        # ── Configuración de rutas ──────────────────────────────────────
        self.configurador = ConfiguradorRutasPayPal()

        if not self.configurador.cargar_config():
            # Primera ejecución: mostrar config obligatoria
            self._config_pendiente = True
        else:
            rutas = self.configurador.obtener_rutas()
            Config.cargar_desde_ini(rutas)
            self._config_pendiente = False
        
        # Crear interfaz
        self.create_widgets()
        
        # Cargar estado inicial
        self.load_initial_state()
        
        # Configurar cierre
        self.protocol("WM_DELETE_WINDOW", self.on_close)
    
    def load_initial_state(self):
        """Carga el estado inicial"""
        try:
            # El número de pago se generará automáticamente al ejecutar el proceso
            # No es necesario mostrarlo en la vista principal
            pass
        except:
            pass
    
    def create_widgets(self):
        """Crea todos los elementos de la interfaz con un diseño centrado y limpio"""
        
        # ========== HEADER PRINCIPAL ==========
        self.header_frame = ctk.CTkFrame(
            self, 
            fg_color="#FFFFFF", # Blanco
            height=80,
            corner_radius=0
        )
        self.header_frame.pack(fill="x", side="top")
        
        self.header_content = ctk.CTkFrame(self.header_frame, fg_color="transparent")
        self.header_content.pack(expand=True)
        
        self.brand_label = ctk.CTkLabel(
            self.header_content,
            text="SISTEMA DE PAGOS PAYPAL",
            font=("Roboto", 24, "bold"), # Aumentado un poco el tamaño
            text_color=COLOR_PRIMARY
        )
        self.brand_label.pack(side="left", padx=15, pady=15)

        # Botón de configuración en el header
        ctk.CTkButton(
            self.header_content,
            text="⚙️",
            command=lambda: self.show_state("config_rutas"),
            width=40,
            height=40,
            fg_color="transparent",
            text_color=COLOR_PRIMARY,
            hover_color=COLOR_ACCENT,
            font=("Roboto", 18)
        ).pack(side="right", padx=10)

        # ========== CONTENIDO PRINCIPAL ==========
        self.main_container = ctk.CTkFrame(self, fg_color="transparent")
        self.main_container.pack(fill="both", expand=True, padx=40, pady=30)
        
        # Contenedor dinámico central
        self.content_frame = ctk.CTkFrame(
            self.main_container,
            fg_color=COLOR_ACCENT,
            corner_radius=20,
            border_width=1,
            border_color=COLOR_ACCENT_LIGHT
        )
        self.content_frame.pack(fill="both", expand=True)
        
        # Crear los estados de contenido
        self.create_idle_content()
        self.create_running_content()
        self.create_completed_content()
        self.create_verificar_soportes_content()
        self.create_verificando_content()
        self.create_resultado_verificacion_content()
                
        # ========== BARRA DE ESTADO ==========
        self.status_frame = ctk.CTkFrame(
            self,
            fg_color="transparent",
            height=30
        )
        self.status_frame.pack(fill="x", side="bottom", padx=40, pady=(0, 20))
        
        self.progress_bar = ctk.CTkProgressBar(
            self.status_frame,
            width=250,
            height=8,
            progress_color=COLOR_PRIMARY
        )
        self.progress_bar.set(0)
        self.progress_bar.pack(side="right")
        
        # Mostrar estado inicial
        self.show_state(STATE_IDLE)
        # Mostrar config si es primera ejecución
        if getattr(self, '_config_pendiente', False):
            self.after(100, self._mostrar_config_obligatoria)
    
    def create_idle_content(self):
        """Crea el contenido del estado IDLE con un diseño moderno y scroll"""
        self.idle_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        
        # Grid principal para centrar el scrollable
        self.idle_frame.grid_columnconfigure(0, weight=1)
        self.idle_frame.grid_rowconfigure(0, weight=1)
        
        # Scroll container para asegurar visibilidad en cualquier tamaño
        scroll_idle = ctk.CTkScrollableFrame(self.idle_frame, fg_color="transparent")
        scroll_idle.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        
        # Grid para organizar el contenido dentro del scroll
        scroll_idle.grid_columnconfigure(0, weight=1)
        scroll_idle.grid_columnconfigure(1, weight=1)
        
        # LADO IZQUIERDO: Bienvenida e Información
        info_side = ctk.CTkFrame(scroll_idle, fg_color="transparent")
        info_side.grid(row=0, column=0, sticky="nsew", padx=30, pady=30)
        
        ctk.CTkLabel(
            info_side,
            text="👋 Bienvenido,",
            font=("Roboto", 32, "bold"),
            text_color=COLOR_TEXT
        ).pack(anchor="w")
        
        ctk.CTkLabel(
            info_side,
            text="Sistema de Gestión de Pagos",
            font=("Roboto", 16),
            text_color=COLOR_TEXT # Cambiado de COLOR_PRIMARY para mejor contraste
        ).pack(anchor="w", pady=(0, 30))
        
        # Cards de beneficios
        def create_card(parent, title, desc, icon):
            card = ctk.CTkFrame(parent, fg_color=COLOR_ACCENT_LIGHT, corner_radius=10)
            card.pack(fill="x", pady=10)
            
            ctk.CTkLabel(card, text=icon, font=("Roboto", 24)).pack(side="left", padx=15, pady=15)
            
            txt_frame = ctk.CTkFrame(card, fg_color="transparent")
            txt_frame.pack(side="left", fill="both", expand=True, pady=10)
            
            ctk.CTkLabel(txt_frame, text=title, font=("Roboto", 13, "bold"), text_color=COLOR_TEXT).pack(anchor="w")
            ctk.CTkLabel(txt_frame, text=desc, font=("Roboto", 11), text_color=COLOR_TEXT_DIM).pack(anchor="w")

        create_card(info_side, "Automatización SAP", "Descarga de reportes sin intervención manual.", "📥")
        create_card(info_side, "Procesamiento Inteligente", "Cálculos y organización de datos en Excel.", "📊")
        create_card(info_side, "Gestión Documental", "Búsqueda y validación de PDFs en OneDrive.", "📄")

        # LADO DERECHO: Acciones
        action_side = ctk.CTkFrame(scroll_idle, fg_color=COLOR_ACCENT_LIGHT, corner_radius=15)
        action_side.grid(row=0, column=1, sticky="nsew", padx=30, pady=30)
        
        ctk.CTkLabel(
            action_side,
            text="Acciones Rápidas",
            font=("Roboto", 18, "bold"),
            text_color=COLOR_TEXT
        ).pack(pady=(30, 10))

        # Selectores de Periodo (Mes y Año)
        ctk.CTkLabel(action_side, text="Periodo del Reporte:", font=("Roboto", 11, "bold"), text_color=COLOR_TEXT_DIM).pack(pady=(0, 5))
        date_frame = ctk.CTkFrame(action_side, fg_color="transparent")
        date_frame.pack(fill="x", padx=30, pady=(0, 20))
        
        meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", 
                 "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
        self.mes_select = ctk.CTkComboBox(date_frame, values=meses, width=110, height=35, state="readonly")
        self.mes_select.set(meses[datetime.now().month - 1])
        self.mes_select.pack(side="left", expand=True, padx=(0, 5))
        
        años = [str(y) for y in range(datetime.now().year - 1, datetime.now().year + 2)]
        self.año_select = ctk.CTkComboBox(date_frame, values=años, width=70, height=35, state="readonly")
        self.año_select.set(str(datetime.now().year))
        self.año_select.pack(side="left", expand=True, padx=(5, 0))
        
        # Botón Principal
        self.btn_ejecutar = ctk.CTkButton(
            action_side,
            text="▶️ Comenzar",
            command=self.start_workflow,
            fg_color=COLOR_PRIMARY,
            hover_color=("#060D6F", "#4A52A7"),
            font=("Roboto", 14, "bold"),
            height=50
        )
        self.btn_ejecutar.pack(fill="x", padx=30, pady=10)
        
        ctk.CTkLabel(
            action_side,
            text="Inicia descarga SAP, procesa Excel y busca PDFs.",
            font=("Roboto", 11),
            text_color=COLOR_TEXT_DIM
        ).pack(pady=(0, 20))
        
        # Botón Secundario
        self.btn_verificar = ctk.CTkButton(
            action_side,
            text=" Actualizar Soportes",
            command=self.show_verificar_soportes,
            fg_color="transparent",
            border_color=COLOR_PRIMARY,
            border_width=2,
            text_color=COLOR_PRIMARY,
            hover_color=("#E4E6EB", "#2A3357"),
            font=("Roboto", 14, "bold"),
            height=50
        )
        self.btn_verificar.pack(fill="x", padx=30, pady=10)
        
        ctk.CTkLabel(
            action_side,
            text="Solo busca y actualiza documentos faltantes.",
            font=("Roboto", 11),
            text_color=COLOR_TEXT_DIM
        ).pack(pady=(0, 20))
        
    def create_running_content(self):
        """Crea el contenido del estado RUNNING con un diseño compacto y adaptativo"""
        self.running_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        
        # Grid principal para RUNNING
        self.running_frame.grid_columnconfigure(0, weight=1)
        self.running_frame.grid_rowconfigure(0, weight=1)
        
        # Contenedor central con scroll por si la ventana es pequeña
        scroll_container = ctk.CTkScrollableFrame(self.running_frame, fg_color="transparent")
        scroll_container.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)
        
        # Título y Estado
        header_section = ctk.CTkFrame(scroll_container, fg_color="transparent")
        header_section.pack(fill="x", pady=(0, 20))
        
        self.payment_number_display = ctk.CTkLabel(
            header_section,
            text="Procesando Pago #---",
            font=("Roboto", 22, "bold"),
            text_color=COLOR_TEXT # Blanco para mejor contraste
        )
        self.payment_number_display.pack()
        
        self.current_step_label = ctk.CTkLabel(
            header_section,
            text="Iniciando...",
            font=("Roboto", 13),
            text_color=COLOR_TEXT_DIM
        )
        self.current_step_label.pack(pady=2)

        # Barra de progreso principal
        progress_section = ctk.CTkFrame(scroll_container, fg_color=COLOR_ACCENT_LIGHT, corner_radius=15)
        progress_section.pack(fill="x", pady=10)
        
        self.main_progress = ctk.CTkProgressBar(
            progress_section,
            width=400,
            height=10,
            progress_color=COLOR_PRIMARY
        )
        self.main_progress.set(0)
        self.main_progress.pack(padx=30, pady=(20, 5))
        
        self.progress_label = ctk.CTkLabel(
            progress_section,
            text="0%",
            font=("Roboto", 14, "bold"),
            text_color=COLOR_TEXT
        )
        self.progress_label.pack(pady=(0, 15))

        # Panel de Pasos Visual (Limpio)
        steps_panel = ctk.CTkFrame(scroll_container, fg_color="transparent")
        steps_panel.pack(fill="x", pady=10)
        
        self.step_labels = {}
        steps_list_frame = ctk.CTkFrame(steps_panel, fg_color="transparent")
        steps_list_frame.pack(expand=True)

        for step_id, step_name in self.workflow_steps:
            f = ctk.CTkFrame(steps_list_frame, fg_color="transparent")
            f.pack(side="left", padx=10)
            
            icon = ctk.CTkLabel(f, text="○", font=("Roboto", 18), text_color=COLOR_TEXT_DIM)
            icon.pack()
            
            lbl = ctk.CTkLabel(f, text=step_name.replace(" ", "\n"), font=("Roboto", 9), text_color=COLOR_TEXT_DIM)
            lbl.pack()
            
            self.step_labels[step_id] = {'icon': icon, 'label': lbl}

        # Área de Información (En lugar de consola)
        info_panel = ctk.CTkFrame(scroll_container, fg_color=COLOR_ACCENT_LIGHT, corner_radius=10)
        info_panel.pack(fill="both", expand=True, pady=10)
        
        ctk.CTkLabel(
            info_panel,
            text="📋 Actividad reciente:",
            font=("Roboto", 12, "bold"),
            text_color=COLOR_TEXT
        ).pack(anchor="w", padx=20, pady=(10, 2))
        
        self.last_log_label = ctk.CTkLabel(
            info_panel,
            text="Preparando el entorno...",
            font=("Roboto", 11),
            text_color=COLOR_SUCCESS,
            justify="left"
        )
        self.last_log_label.pack(anchor="w", padx=20, pady=(0, 15))
        
        # Mantenemos el log_text oculto
        self.log_text = ctk.CTkTextbox(self.running_frame, width=1, height=1) 
        self.log_text.grid(row=1, column=1)
        self.log_text.lower()

        # Botón Cancelar (Asegurando visibilidad)
        self.btn_cancelar = ctk.CTkButton(
            scroll_container,
            text="✕ Cancelar Operación",
            command=self.cancel_process,
            fg_color="transparent",
            border_color=COLOR_ERROR,
            border_width=1,
            text_color=COLOR_ERROR,
            hover_color=("#FFEBEE", "#331111"),
            height=35,
            width=200
        )
        self.btn_cancelar.pack(pady=(10, 0))
    
    def create_verificando_content(self):
        """Crea el contenido del estado de verificación en progreso"""
        self.verificando_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        
        self.verificando_frame.grid_columnconfigure(0, weight=1)
        self.verificando_frame.grid_rowconfigure(0, weight=1)
        
        scroll_container = ctk.CTkScrollableFrame(self.verificando_frame, fg_color="transparent")
        scroll_container.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)
        
        header_section = ctk.CTkFrame(scroll_container, fg_color="transparent")
        header_section.pack(fill="x", pady=(0, 20))
        
        self.v_payment_display = ctk.CTkLabel(
            header_section,
            text="Actualizando Soportes Pago #---",
            font=("Roboto", 22, "bold"),
            text_color=COLOR_TEXT
        )
        self.v_payment_display.pack()
        
        self.v_current_step_label = ctk.CTkLabel(
            header_section,
            text="Iniciando...",
            font=("Roboto", 13),
            text_color=COLOR_TEXT_DIM
        )
        self.v_current_step_label.pack(pady=2)

        progress_section = ctk.CTkFrame(scroll_container, fg_color=COLOR_ACCENT_LIGHT, corner_radius=15)
        progress_section.pack(fill="x", pady=10)
        
        self.v_main_progress = ctk.CTkProgressBar(
            progress_section,
            width=400,
            height=10,
            progress_color=COLOR_PRIMARY
        )
        self.v_main_progress.set(0)
        self.v_main_progress.pack(padx=30, pady=(20, 5))
        
        self.v_progress_label = ctk.CTkLabel(
            progress_section,
            text="0%",
            font=("Roboto", 14, "bold"),
            text_color=COLOR_TEXT
        )
        self.v_progress_label.pack(pady=(0, 15))

        # Pasos específicos de verificación
        v_steps_panel = ctk.CTkFrame(scroll_container, fg_color="transparent")
        v_steps_panel.pack(fill="x", pady=10)
        
        self.v_step_labels = {}
        v_steps_list_frame = ctk.CTkFrame(v_steps_panel, fg_color="transparent")
        v_steps_list_frame.pack(expand=True)

        for step_id, step_name in self.verification_steps:
            f = ctk.CTkFrame(v_steps_list_frame, fg_color="transparent")
            f.pack(side="left", padx=15)
            
            icon = ctk.CTkLabel(f, text="○", font=("Roboto", 20), text_color=COLOR_TEXT_DIM)
            icon.pack()
            
            lbl = ctk.CTkLabel(f, text=step_name.replace(" ", "\n"), font=("Roboto", 10), text_color=COLOR_TEXT_DIM)
            lbl.pack()
            
            self.v_step_labels[step_id] = {'icon': icon, 'label': lbl}

        info_panel = ctk.CTkFrame(scroll_container, fg_color=COLOR_ACCENT_LIGHT, corner_radius=10)
        info_panel.pack(fill="both", expand=True, pady=10)
        
        ctk.CTkLabel(
            info_panel,
            text="📋 Actividad reciente:",
            font=("Roboto", 12, "bold"),
            text_color=COLOR_TEXT
        ).pack(anchor="w", padx=20, pady=(10, 2))
        
        self.v_last_log_label = ctk.CTkLabel(
            info_panel,
            text="Iniciando verificación...",
            font=("Roboto", 11),
            text_color=COLOR_SUCCESS,
            justify="left"
        )
        self.v_last_log_label.pack(anchor="w", padx=20, pady=(0, 15))
        
        self.v_btn_cancelar = ctk.CTkButton(
            scroll_container,
            text="✕ Cancelar Verificación",
            command=self.cancel_process,
            fg_color="transparent",
            border_color=COLOR_ERROR,
            border_width=1,
            text_color=COLOR_ERROR,
            hover_color=("#FFEBEE", "#331111"),
            height=35,
            width=200
        )
        self.v_btn_cancelar.pack(pady=(10, 0))

    def create_completed_content(self):
        """Crea el contenido del estado COMPLETED con un diseño visualmente atractivo y scroll"""
        self.completed_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        
        # Grid para centrar el scrollable
        self.completed_frame.grid_columnconfigure(0, weight=1)
        self.completed_frame.grid_rowconfigure(0, weight=1)
        
        # Scroll container
        scroll_completed = ctk.CTkScrollableFrame(self.completed_frame, fg_color="transparent")
        scroll_completed.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        
        container = ctk.CTkFrame(scroll_completed, fg_color="transparent")
        container.pack(expand=True, fill="both", padx=50, pady=40)
        
        # Icono de Éxito Grande
        ctk.CTkLabel(
            container,
            text="✨",
            font=("Roboto", 64)
        ).pack(pady=(0, 10))
        
        ctk.CTkLabel(
            container,
            text="¡Proceso Finalizado con Éxito!",
            font=("Roboto", 24, "bold"),
            text_color=COLOR_SUCCESS
        ).pack()
        
        ctk.CTkLabel(
            container,
            text="El reporte ha sido procesado y el archivo maestro actualizado.",
            font=("Roboto", 13),
            text_color=COLOR_TEXT_DIM
        ).pack(pady=(5, 30))
        
        # Panel de Resumen
        summary_panel = ctk.CTkFrame(container, fg_color=COLOR_ACCENT_LIGHT, corner_radius=15)
        summary_panel.pack(fill="both", expand=True, pady=10)
        
        ctk.CTkLabel(
            summary_panel,
            text="� RESUMEN DE LA OPERACIÓN",
            font=("Roboto", 12, "bold"),
            text_color=COLOR_TEXT
        ).pack(anchor="w", padx=25, pady=(20, 10))
        
        self.summary_text = ctk.CTkTextbox(
            summary_panel,
            fg_color=COLOR_ACCENT,
            text_color=COLOR_TEXT,
            font=("Consolas", 11),
            border_width=1,
            border_color=COLOR_ACCENT_LIGHT,
            corner_radius=10
        )
        self.summary_text.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        self.summary_text.configure(state="disabled")
        
        # Botones de Acción (Lado a Lado - Flex)
        btns_container = ctk.CTkFrame(container, fg_color="transparent")
        btns_container.pack(fill="x", pady=(20, 0))
        
        ctk.CTkButton(
            btns_container,
            text="🔄 Procesar Siguiente",
            command=self.continue_workflow,
            fg_color=COLOR_PRIMARY,
            hover_color="#169c46",
            font=("Roboto", 14, "bold"),
            height=45
        ).pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        ctk.CTkButton(
            btns_container,
            text="🏠 Panel Principal",
            command=self.back_to_idle,
            fg_color="transparent",
            border_color=COLOR_ACCENT_LIGHT,
            border_width=1,
            text_color=COLOR_TEXT,
            hover_color=COLOR_ACCENT_LIGHT,
            font=("Roboto", 14, "bold"),
            height=45
        ).pack(side="left", fill="x", expand=True, padx=(10, 0))
    
    def create_verificar_soportes_content(self):
        """Rediseño de la interfaz de actualización de soportes con scroll"""
        self.verificar_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        
        # Grid para organizar el scrollable
        self.verificar_frame.grid_columnconfigure(0, weight=1)
        self.verificar_frame.grid_rowconfigure(0, weight=1)
        
        # Scroll container
        scroll_verif = ctk.CTkScrollableFrame(self.verificar_frame, fg_color="transparent")
        scroll_verif.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        
        # Grid para organizar dentro del scroll
        scroll_verif.grid_columnconfigure(0, weight=1)
        scroll_verif.grid_columnconfigure(1, weight=1)
        
        # LADO IZQUIERDO: Configuración y Selección
        config_side = ctk.CTkFrame(scroll_verif, fg_color="transparent")
        config_side.grid(row=0, column=0, sticky="nsew", padx=30, pady=30)
        
        ctk.CTkLabel(
            config_side,
            text=" Actualizar Soportes",
            font=("Roboto", 24, "bold"),
            text_color=COLOR_TEXT # Cambiado de COLOR_PRIMARY para mejor contraste
        ).pack(anchor="w")
        
        ctk.CTkLabel(
            config_side,
            text="Busca PDFs faltantes y actualiza el Excel.",
            font=("Roboto", 13),
            text_color=COLOR_TEXT_DIM
        ).pack(anchor="w", pady=(0, 30))
        
        # Selector de Pago con estilo mejorado
        selector_card = ctk.CTkFrame(config_side, fg_color=COLOR_ACCENT_LIGHT, corner_radius=15)
        selector_card.pack(fill="x", pady=10)
        
        ctk.CTkLabel(
            selector_card,
            text="Selecciona el número de pago:",
            font=("Roboto", 12, "bold"),
            text_color=COLOR_TEXT
        ).pack(anchor="w", padx=20, pady=(15, 5))

        # --- DEBUG: Mostrar ruta de búsqueda ---
        debug_path_text = "Ruta de pagos (BASE_PAYPAL) no configurada."
        is_configured = False
        try:
            # Usar el método de la clase Config si existe
            if hasattr(Config, 'esta_configurado'):
                is_configured = Config.esta_configurado()
            # Fallback por si el método falla o no existe
            elif hasattr(Config, 'BASE_PAYPAL') and Config.BASE_PAYPAL:
                is_configured = True
        except Exception:
            pass  # Silenciar error si la clase Config es inestable

        if is_configured and hasattr(Config, 'BASE_PAYPAL') and Config.BASE_PAYPAL:
             debug_path_text = f"Buscando carpetas 'Pago #...' en: {Config.BASE_PAYPAL}"
        
        ctk.CTkLabel(selector_card, text=debug_path_text, font=("Roboto", 9), text_color=COLOR_TEXT_DIM, wraplength=350, justify="left").pack(anchor="w", padx=20, pady=(0,5))
        # --- FIN DEBUG ---

        
        # Obtener pagos disponibles de forma segura
        numeros_disponibles = []
        try:
            todas_carpetas = [d for d in Config.BASE_PAYPAL.iterdir() if d.is_dir() and d.name.startswith("Pago #")]
            for carpeta in todas_carpetas:
                try:
                    num = int(carpeta.name.replace("Pago #", ""))
                    numeros_disponibles.append(num)
                except: continue
            numeros_disponibles.sort(reverse=True)
        except Exception as e:
            # Mostrar el error en la interfaz para diagnóstico
            error_msg = f"Error al leer la carpeta de pagos:\n{e}"
            ctk.CTkLabel(selector_card, text=error_msg, text_color=COLOR_ERROR, font=("Roboto", 10), wraplength=300).pack(pady=10)
        
        if numeros_disponibles:
            self.pago_select = ctk.CTkComboBox(
                selector_card,
                values=[str(n) for n in numeros_disponibles],
                state="readonly",
                height=40,
                fg_color=COLOR_ACCENT,
                border_color=COLOR_ACCENT_LIGHT,
                button_color=COLOR_PRIMARY,
                button_hover_color=("#060D6F", "#4A52A7")
            )
            self.pago_select.set(str(numeros_disponibles[0]))
            self.pago_select.pack(fill="x", padx=20, pady=(5, 20))
        else:
            ctk.CTkLabel(selector_card, text="No se encontraron pagos", text_color=COLOR_ERROR).pack(pady=20)

        # Botones de acción lado a lado
        btn_action_frame = ctk.CTkFrame(config_side, fg_color="transparent")
        btn_action_frame.pack(fill="x", pady=(20, 10))
        
        self.btn_verificar_action = ctk.CTkButton(
            btn_action_frame,
            text="▶️ Iniciar",
            command=self.start_verificacion,
            fg_color=COLOR_PRIMARY,
            hover_color=("#060D6F", "#4A52A7"),
            font=("Roboto", 14, "bold"),
            height=45
        )
        self.btn_verificar_action.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        btn_back = ctk.CTkButton(
            btn_action_frame,
            text="↩️ Volver",
            command=self.back_to_idle,
            fg_color="transparent",
            border_color=COLOR_ACCENT_LIGHT,
            border_width=1,
            text_color=COLOR_TEXT_DIM,
            hover_color=COLOR_ACCENT_LIGHT,
            height=45
        )
        btn_back.pack(side="left", fill="x", expand=True, padx=(5, 0))

        # LADO DERECHO: Resumen de Proceso
        info_side = ctk.CTkFrame(scroll_verif, fg_color=COLOR_ACCENT_LIGHT, corner_radius=15)
        info_side.grid(row=0, column=1, sticky="nsew", padx=30, pady=30)
        
        ctk.CTkLabel(
            info_side,
            text="¿Qué hará el sistema?",
            font=("Roboto", 16, "bold"),
            text_color=COLOR_TEXT
        ).pack(pady=(25, 15))
        
        # Lista de pasos con iconos
        steps_info = [
            ("📁", "Escaneo de carpetas en OneDrive"),
            ("📥", "Copia automática de archivos a Soporte"),
            ("📝", "Análisis de observaciones actuales"),
            ("✅", "Actualización automática de 'Soportes OK'"),
            ("📊", "Generación de reporte detallado")
        ]
        
        for icon, text in steps_info:
            step_f = ctk.CTkFrame(info_side, fg_color="transparent")
            step_f.pack(fill="x", padx=25, pady=8)
            ctk.CTkLabel(step_f, text=icon, font=("Roboto", 18)).pack(side="left", padx=(0, 10))
            ctk.CTkLabel(step_f, text=text, font=("Roboto", 11), text_color=COLOR_TEXT).pack(side="left")

    def create_resultado_verificacion_content(self):
        """Crea la interfaz de resultados de verificación con diseño moderno y scroll"""
        self.resultado_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        
        # Grid para centrar el scrollable
        self.resultado_frame.grid_columnconfigure(0, weight=1)
        self.resultado_frame.grid_rowconfigure(0, weight=1)
        
        # Scroll container
        scroll_res_verif = ctk.CTkScrollableFrame(self.resultado_frame, fg_color="transparent")
        scroll_res_verif.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        
        container = ctk.CTkFrame(scroll_res_verif, fg_color="transparent")
        container.pack(expand=True, fill="both", padx=50, pady=40)
        
        # Header de Resultados
        header_section = ctk.CTkFrame(container, fg_color="transparent")
        header_section.pack(fill="x", pady=(0, 20))
        
        self.resultado_titulo = ctk.CTkLabel(
            header_section,
            text="✅ Verificación Finalizada",
            font=("Roboto", 24, "bold"),
            text_color=COLOR_PRIMARY
        )
        self.resultado_titulo.pack()
        
        self.resultado_stats = ctk.CTkLabel(
            header_section,
            text="Se han actualizado las observaciones en el archivo Excel.",
            font=("Roboto", 13),
            text_color=COLOR_TEXT_DIM
        )
        self.resultado_stats.pack(pady=5)
        
        # Panel de Detalles
        details_panel = ctk.CTkFrame(container, fg_color=COLOR_ACCENT_LIGHT, corner_radius=15)
        details_panel.pack(fill="both", expand=True, pady=10)
        
        ctk.CTkLabel(
            details_panel,
            text="📋 DETALLES DE LOS CAMBIOS",
            font=("Roboto", 12, "bold"),
            text_color=COLOR_TEXT
        ).pack(anchor="w", padx=25, pady=(20, 10))
        
        self.resultado_detalles = ctk.CTkTextbox(
            details_panel,
            fg_color=COLOR_ACCENT,
            text_color=COLOR_TEXT,
            font=("Consolas", 11),
            border_width=1,
            border_color=COLOR_ACCENT_LIGHT,
            corner_radius=10
        )
        self.resultado_detalles.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        self.resultado_detalles.configure(state="disabled")
        
        # Botones de Acción (Lado a Lado - Flex)
        btns_container = ctk.CTkFrame(container, fg_color="transparent")
        btns_container.pack(fill="x", pady=(20, 0))
        
        ctk.CTkButton(
            btns_container,
            text="🔍 Verificar Otro Pago",
            command=self.show_verificar_soportes,
            fg_color=COLOR_PRIMARY,
            hover_color=("#060D6F", "#4A52A7"),
            font=("Roboto", 14, "bold"),
            height=45
        ).pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        ctk.CTkButton(
            btns_container,
            text="🏠 Panel Principal",
            command=self.back_to_idle,
            fg_color="transparent",
            border_color=COLOR_ACCENT_LIGHT,
            border_width=1,
            text_color=COLOR_TEXT,
            hover_color=COLOR_ACCENT_LIGHT,
            font=("Roboto", 14, "bold"),
            height=45
        ).pack(side="left", fill="x", expand=True, padx=(10, 0))
    
    def show_verificar_soportes(self):
        """NUEVO: Muestra la interfaz para verificar soportes"""
        self.modo_verificacion = True
        self.show_state("verificar_soportes")
    
    def start_verificacion(self):
        """NUEVO: Inicia la verificación del pago seleccionado"""
        try:
            numero_pago = int(self.pago_select.get())
            self.pago_verificando = numero_pago
            
            self.operation_running = True
            self.show_state("verificando")
            
            # Ejecutar en hilo
            thread = threading.Thread(target=self._run_verificacion, daemon=True)
            thread.start()
        
        except Exception as e:
            messagebox.showerror("Error", f"Error iniciando verificación: {e}")
    
    def _run_verificacion(self):
        """NUEVO: Ejecuta la verificación Y actualización (en hilo separado)"""
        try:
            self.log_message(" Iniciando búsqueda y actualización de soportes...")
            
            # Crear verificador/actualizador
            rutas_busqueda = resolver_rutas_swift_dinamicas(Config.RAIZ_SWIFT_LATAM) or Config.RUTAS_PDF
            verificador = VerificadorActualizadorSoportes(rutas_busqueda)
            
            # Callback para progreso real
            def update_progress(p, msg):
                self.after(0, lambda: self.v_main_progress.set(p))
                self.after(0, lambda: self.v_progress_label.configure(text=f"{int(p * 100)}%"))
                
                # Actualizar pasos según el progreso p (0.0 a 1.0)
                if p < 0.15:
                    self.update_verification_step_icon("v_step1", "🔄", COLOR_PRIMARY)
                elif p < 0.60:
                    self.update_verification_step_icon("v_step1", "✓", COLOR_SUCCESS)
                    self.update_verification_step_icon("v_step2", "🔄", COLOR_PRIMARY)
                elif p < 0.95:
                    self.update_verification_step_icon("v_step2", "✓", COLOR_SUCCESS)
                    self.update_verification_step_icon("v_step3", "🔄", COLOR_PRIMARY)
                elif p < 1.0:
                    self.update_verification_step_icon("v_step3", "✓", COLOR_SUCCESS)
                    self.update_verification_step_icon("v_step4", "🔄", COLOR_PRIMARY)
                else:
                    self.update_verification_step_icon("v_step4", "✓", COLOR_SUCCESS)

                # Actualizar log visual específico
                if hasattr(self, 'v_last_log_label'):
                    self.after(0, lambda: self.v_last_log_label.configure(text=f"• {msg}"))
                if hasattr(self, 'v_current_step_label'):
                    self.after(0, lambda: self.v_current_step_label.configure(text=msg))
                
                self.log_message(msg)

            # Verificar cancelación antes de iniciar
            if self.check_cancel_and_continue():
                return
            
            resultado = verificador.procesar_pago_completo(
                self.pago_verificando, 
                Config.BASE_PAYPAL,
                progress_callback=update_progress
            )
            
            # Verificar cancelación después de proceso
            if self.check_cancel_and_continue():
                return
            
            self.resultados_verificacion = [resultado]
            
            # Actualizar UI con resultado
            self.after(0, self._mostrar_resultado_verificacion, resultado)
            
        except Exception as e:
            self.log_message(f" Error durante verificación: {e}")
            self.logger.error(f"Error en verificación: {e}", exc_info=True)
        finally:
            self.operation_running = False
    
    def _mostrar_resultado_verificacion(self, resultado):
        """NUEVO: Muestra el resultado con CAMBIOS REALIZADOS"""
        self.show_state("resultado_verificacion")
        
        # Encabezado
        titulo = f"Pago #{resultado.numero_pago} - {resultado.estado_general}"
        self.resultado_titulo.configure(text=titulo)
        
        # Estadísticas
        stats_text = (
            f"Registros totales: {resultado.registros_totales}\n"
            f"Con observaciones: {resultado.registros_con_observaciones}\n"
            f"\n"
            f"🔄 CAMBIOS REALIZADOS:\n"
            f"   📁 Archivos copiados: {resultado.documentos_copiados}\n"
            f"   📝 Observaciones actualizadas: {resultado.observaciones_actualizadas}"
        )
        
        self.resultado_stats.configure(text=stats_text)
        
        # Detalles de cambios
        self.resultado_detalles.configure(state="normal")
        self.resultado_detalles.delete("1.0", "end")
        
        # Mostrar archivos copiados
        if resultado.archivos_copiados:
            self.resultado_detalles.insert(
                "end",
                f"\n📁 ARCHIVOS COPIADOS A SOPORTE ({len(resultado.archivos_copiados)}):\n"
                f"{'-'*80}\n"
            )
            for archivo in resultado.archivos_copiados:
                self.resultado_detalles.insert(
                    "end",
                    f" {archivo['nombre']}\n"
                )
        
        # Mostrar cambios en observaciones
        if resultado.cambios_realizados:
            self.resultado_detalles.insert(
                "end",
                f"\n\n📝 OBSERVACIONES ACTUALIZADAS ({len(resultado.cambios_realizados)}):\n"
                f"{'-'*80}\n"
            )
            for cambio in resultado.cambios_realizados:
                texto = (
                    f"\n Fila {cambio['fila']} (Invoice: {cambio['invoice']}):\n"
                    f"  ANTES: {cambio['observacion_anterior']}\n"
                    f"  DESPUÉS: {cambio['observacion_nueva']}\n"
                )
                self.resultado_detalles.insert("end", texto)
        
        # Si no hay cambios
        if not resultado.archivos_copiados and not resultado.cambios_realizados:
            self.resultado_detalles.insert(
                "end",
                "ℹ️ Verificación completada.\n"
                "No hay nuevos documentos ni cambios pendientes."
            )
        
        self.resultado_detalles.configure(state="disabled")
    
    def back_to_idle(self):
        """NUEVO: Vuelve al estado inicial"""
        self.modo_verificacion = False
        self.pago_verificando = None
        self.show_state(STATE_IDLE)
        self.operation_running = False

    def show_state(self, state):
        """Muestra el contenido según el estado actual"""
        # Ocultar todos
        if hasattr(self, 'idle_frame'):
            self.idle_frame.pack_forget()
        if hasattr(self, 'running_frame'):
            self.running_frame.pack_forget()
        if hasattr(self, 'completed_frame'):
            self.completed_frame.pack_forget()
        if hasattr(self, 'verificar_frame'):
            self.verificar_frame.pack_forget()
        if hasattr(self, 'verificando_frame'):
            self.verificando_frame.pack_forget()
        if hasattr(self, 'resultado_frame'):
            self.resultado_frame.pack_forget()
        if hasattr(self, 'config_frame'):
            self.config_frame.pack_forget()
        
        # Mostrar el seleccionado
        if state == STATE_IDLE:
            if not hasattr(self, 'idle_frame'):
                self.create_idle_content()
            self.idle_frame.pack(fill="both", expand=True)
        
        elif state == STATE_RUNNING:
            if not hasattr(self, 'running_frame'):
                self.create_running_content()
            self.running_frame.pack(fill="both", expand=True)
        
        elif state == STATE_COMPLETED:
            if not hasattr(self, 'completed_frame'):
                self.create_completed_content()
            self.completed_frame.pack(fill="both", expand=True)
        
        elif state == "verificar_soportes":
            if not hasattr(self, 'verificar_frame'):
                self.create_verificar_soportes_content()
            self.verificar_frame.pack(fill="both", expand=True)
        
        elif state == "verificando":
            if not hasattr(self, 'verificando_frame'):
                self.create_verificando_content()
            self.verificando_frame.pack(fill="both", expand=True)
            
            # Resetear iconos de pasos de verificación
            for step_id, _ in self.verification_steps:
                self.update_verification_step_icon(step_id, "○", COLOR_TEXT_DIM)
                
            # Actualizar título para verificación
            if hasattr(self, 'v_payment_display'):
                self.v_payment_display.configure(text=f"Actualizando Soportes Pago #{self.pago_verificando}")
            
            # Resetear progreso
            if hasattr(self, 'v_main_progress'):
                self.v_main_progress.set(0)
            if hasattr(self, 'v_progress_label'):
                self.v_progress_label.configure(text="0%")
            
            self.log_message("Buscando y actualizando soportes...")
        
        elif state == "resultado_verificacion":
            if not hasattr(self, 'resultado_frame'):
                self.create_resultado_verificacion_content()
            self.resultado_frame.pack(fill="both", expand=True)
        elif state == "config_rutas":
            if not hasattr(self, 'config_frame'):
                self.create_config_rutas_content()
            # Recrear siempre para reflejar rutas actuales
            else:
                self.config_frame.destroy()
                del self.config_frame
                self.create_config_rutas_content()
            self.config_frame.pack(fill="both", expand=True)
        
        self.current_state = state
    
    def start_workflow(self):
        """Inicia el workflow completo"""
        try:
            # Resetear flag de cancelación
            self.cancel_requested = False

            # Obtener mes y año seleccionados
            meses_dict = {"Enero":1, "Febrero":2, "Marzo":3, "Abril":4, "Mayo":5, "Junio":6, 
                          "Julio":7, "Agosto":8, "Septiembre":9, "Octubre":10, "Noviembre":11, "Diciembre":12}
            self.mes_pago = meses_dict.get(self.mes_select.get(), datetime.now().month)
            self.año_pago = int(self.año_select.get())
            
            # Obtener el siguiente número de pago automáticamente
            gestor = GestorCarpetas(Config.BASE_PAYPAL)
            self.numero_pago = gestor.obtener_pago_pendiente_o_siguiente()
            
            # Actualizar la visualización del número de pago en la vista de ejecución
            if hasattr(self, 'payment_number_display'):
                self.payment_number_display.configure(text=f"#{self.numero_pago}")
            
            self.operation_running = True
            self.show_state(STATE_RUNNING)
            
            # Resetear UI para un nuevo inicio
            if hasattr(self, 'main_progress'):
                self.main_progress.set(0)
            if hasattr(self, 'progress_label'):
                self.progress_label.configure(text="0%")
            if hasattr(self, 'btn_cancelar'):
                self.btn_cancelar.configure(
                    state="normal",
                    text="✕ Cancelar Operación",
                    fg_color="transparent",
                    border_color=COLOR_ERROR,
                    border_width=1,
                    text_color=COLOR_ERROR,
                    hover_color=("#FFEBEE", "#331111"),
                    command=self.cancel_process,
                )
            
            # Resetear iconos de pasos
            for step_id, _ in self.workflow_steps:
                self.update_step_icon(step_id, "○", COLOR_TEXT_DIM)
            
            self.update_status("Ejecutando proceso...", 0)
            
            # Ejecutar en hilo
            thread = threading.Thread(target=self._execute_workflow, daemon=True)
            thread.start()
        
        except Exception as e:
            messagebox.showerror("Error", f"Error al iniciar el proceso: {e}")
            self.operation_running = False
    
    def _execute_workflow(self):
        """Ejecuta el workflow completo en hilo separado controlando la cancelación"""
        try:
            if not self.run_step("step1", self._verify_folders): return
            if not self.run_step("step2", self._download_from_sap): return
            if not self.run_step("step3", self._process_excel): return
            if not self.run_step("step4", self._search_pdfs): return
            if not self.run_step("step5", self._update_master): return
            
            self._on_workflow_completed()
        except Exception as e:
            self.log_message(f" Error en workflow: {e}")
            self.logger.error(f"Error workflow: {e}", exc_info=True)
        finally:
            self.operation_running = False
            self.after(0, lambda: self.btn_ejecutar.configure(state="normal", text=" EJECUTAR PROCESO COMPLETO"))
    
    def update_status(self, message, progress):
        """Actualiza barra de estado"""
        def _update():
            self.status_label.configure(text=message)
            if progress >= 0:
                self.progress_bar.set(progress)
        self.after(0, _update)
    
    def run_step(self, step_id, step_function):
        """Ejecuta un paso del workflow de forma segura y verificando cancelación"""
        # 1. Verificar si se solicitó cancelación ANTES del paso
        if self.check_cancel_and_continue():
            return False
        
        self.update_step_icon(step_id, "🔄", "#3498DB")
        
        step_index = [s[0] for s in self.workflow_steps].index(step_id)
        progress = (step_index + 0.5) / len(self.workflow_steps)
        self.after(0, lambda p=progress: self.main_progress.set(p))
        self.after(0, lambda p=progress: self.progress_label.configure(text=f"{int(p * 100)}%"))
        
        # 2. Ejecutar el paso
        try:
            step_function()
        except Exception as e:
            self.log_message(f"❌ Error en {step_id}: {e}")
            raise e
        
        # 3. Verificar cancelación DESPUÉS del paso
        if self.check_cancel_and_continue():
            return False
        
        progress_final = (step_index + 1) / len(self.workflow_steps)
        self.update_step_icon(step_id, "✓", COLOR_SUCCESS)
        self.after(0, lambda p=progress_final: self.main_progress.set(p))
        self.after(0, lambda p=progress_final: self.progress_label.configure(text=f"{int(p * 100)}%"))
        return True
    
    def update_step_icon(self, step_id, icon, color):
        """Actualiza el icono y estilo de un paso en el timeline"""
        def _update():
            # Cambiar icono y color
            self.step_labels[step_id]['icon'].configure(text=icon, text_color=color)
            
            # Cambiar estilo del texto
            if icon == "✓": # Éxito
                self.step_labels[step_id]['label'].configure(text_color=COLOR_SUCCESS, font=("Roboto", 12, "bold"))
                self.step_labels[step_id]['icon'].configure(text="✓")
            elif icon == "🔄": # En proceso
                self.step_labels[step_id]['label'].configure(text_color=COLOR_PRIMARY, font=("Roboto", 12, "bold"))
            else: # Pendiente u otro
                self.step_labels[step_id]['label'].configure(text_color=COLOR_TEXT_DIM, font=("Roboto", 12))
                
        self.after(0, _update)

    def update_verification_step_icon(self, step_id, icon, color):
        """Actualiza el icono y estilo de un paso en el timeline de verificación"""
        def _update():
            if step_id in self.v_step_labels:
                # Cambiar icono y color
                self.v_step_labels[step_id]['icon'].configure(text=icon, text_color=color)
                
                # Cambiar estilo del texto
                if icon == "✓": # Éxito
                    self.v_step_labels[step_id]['label'].configure(text_color=COLOR_SUCCESS, font=("Roboto", 11, "bold"))
                elif icon == "🔄": # En proceso
                    self.v_step_labels[step_id]['label'].configure(text_color=COLOR_PRIMARY, font=("Roboto", 11, "bold"))
                else: # Pendiente u otro
                    self.v_step_labels[step_id]['label'].configure(text_color=COLOR_TEXT_DIM, font=("Roboto", 10))
        self.after(0, _update)
    
    def log_message(self, message):
        """Agrega un mensaje al log y actualiza el label de actividad"""
        def _log():
            timestamp = datetime.now().strftime("%H:%M:%S")
            full_message = f"[{timestamp}] {message}\n"
            
            # Actualizar el label de actividad (limpio)
            if hasattr(self, 'last_log_label'):
                self.last_log_label.configure(text=f"• {message}")
            
            # Actualizar el label del paso actual si corresponde
            if hasattr(self, 'current_step_label'):
                self.current_step_label.configure(text=message)
            
            # Mantener compatibilidad con el log_text original
            self.log_text.configure(state="normal")
            self.log_text.insert("end", full_message)
            self.log_text.see("end")
            self.log_text.configure(state="disabled")
        self.after(0, _log)
    
    def cancel_process(self):
        """NUEVO: Cancela el proceso en ejecución"""
        if messagebox.askyesno("Confirmar Cancelación", "¿Está seguro de que desea cancelar el proceso en ejecución?"):
            self.cancel_requested = True
            self.log_message("⚠️ Cancelación solicitada por el usuario...")
            self.btn_cancelar.configure(state="disabled", text="⏳ CANCELANDO...")
    
    def check_cancel_and_continue(self):
        """NUEVO: Verifica si se solicitó cancelación"""
        if self.cancel_requested:
            self.log_message(" Proceso cancelado por el usuario")
            self.operation_running = False
            self.cancel_requested = False
            # Mostrar vista de ejecución con estado cancelado
            self._show_cancelled_state()
            return True
        return False
    
    def _show_cancelled_state(self):
        def _update():
            if hasattr(self, 'current_step_label'):
                self.current_step_label.configure(
                    text="⚠️  OPERACIÓN CANCELADA POR EL USUARIO",
                    text_color=COLOR_ERROR
                )
 
            if hasattr(self, 'main_progress'):
                self.main_progress.set(0)
 
            if hasattr(self, 'progress_label'):
                self.progress_label.configure(text="CANCELADO")
 
            # Reutilizar btn_cancelar como botón de retorno — sin inyectar widgets nuevos
            if hasattr(self, 'btn_cancelar'):
                self.btn_cancelar.configure(
                    state="normal",
                    text="↩  Volver al inicio",
                    fg_color=COLOR_PRIMARY,
                    border_color=COLOR_PRIMARY,
                    text_color="#FFFFFF",
                    hover_color=("#060D6F", "#4A52A7"),
                    command=self.back_to_idle_from_cancelled,
                )
 
        self.after(0, _update)
    
    def back_to_idle_from_cancelled(self):
        """NUEVO: Vuelve al estado idle desde proceso cancelado"""
        self.cancel_requested = False
        self.operation_running = False
        self.back_to_idle()
    
    def _verify_folders(self):
        """Paso 1: Verificar carpetas"""
        self.log_message("📁 Verificando estructura de carpetas...")
        try:
            gestor = GestorCarpetas(Config.BASE_PAYPAL)
            carpetas = [d for d in Config.BASE_PAYPAL.iterdir() if d.is_dir() and d.name.startswith("Pago #")]
            self.log_message(f" Carpetas verificadas: {len(carpetas)} carpetas existentes")
        except Exception as e:
            self.log_message(f" Error verificando carpetas: {e}")
    
    def _download_from_sap(self):
        """Paso 2: Descargar de SAP"""
        self.log_message("📥 Descargando desde SAP...")
        try:
            # DEBUG: Mostrar información de contexto
            import sys, os
            from tkinter import messagebox as mb
            meipass = getattr(sys, '_MEIPASS', None)
            driver_path = os.path.join(meipass, 'chromedriver.exe') if meipass else 'N/A'
            driver_ok = os.path.exists(driver_path) if meipass else False
            mb.showinfo("DEBUG SAP",
                f"frozen={getattr(sys, 'frozen', False)}\n"
                f"_MEIPASS={meipass or 'N/A'}\n"
                f"chromedriver={driver_path}\n"
                f"driver_existe={driver_ok}"
            )
            
            descargador = DescargadorSAP()
            archivo = descargador.descargar_reporte_sap(self.numero_pago)
            if archivo:
                self.log_message(f" Archivo descargado: {archivo.name}")
            else:
                self.log_message(" No se encontró archivo. Busque manualmente en Descargas.")
        except Exception as e:
            # DEBUG: Mostrar el error real
            from tkinter import messagebox as mb
            mb.showerror("ERROR SAP", f"Error al descargar:\n\n{type(e).__name__}: {e}")
            self.log_message(f" Error en descarga SAP: {e}")
    def _search_pdfs(self):
        """Paso 4: Buscar PDFs"""
        self.log_message("📄 Buscando y validando documentos PDF...")
        try:
            if self.df_segunda is not None and self.carpeta_soporte:
                gestor_pdfs = GestorPDFs(Config.RUTAS_PDF)
                
                # Callback para progreso real dentro del paso
                step_index = 3 # Índice del paso 4 (0-based)
                total_steps = len(self.workflow_steps)
                
                def update_step_progress(p, msg):
                    # Mapear p (0.0 a 1.0) al rango del paso (index/total a (index+1)/total)
                    real_p = (step_index + p) / total_steps
                    self.after(0, lambda: self.main_progress.set(real_p))
                    self.after(0, lambda: self.progress_label.configure(text=f"{int(real_p * 100)}%"))
                    self.log_message(msg)

                self.df_segunda = gestor_pdfs.procesar_documentos_soporte(
                    self.df_segunda, 
                    self.carpeta_soporte,
                    progress_callback=update_step_progress
                )
                
                procesador = ProcesadorExcel()
                procesador.guardar_excel_con_dos_hojas(self.archivo_movido, self.df_segunda)
                
                self.log_message(f" PDFs procesados y Excel final actualizado.")
            else:
                self.log_message("⚠️ Saltando búsqueda de PDFs (no hay datos de Excel)")
        except Exception as e:
            self.log_message(f" Error buscando PDFs: {e}")
    
    def _process_excel(self):
        """Paso 3: Procesar Excel"""
        self.log_message("📊 Procesando archivo Excel...")
        try:
            procesador = ProcesadorExcel()
            
            archivo = procesador.buscar_archivo_pago_en_descargas(self.numero_pago)
            if archivo:
                self.log_message(f" Archivo encontrado: {archivo.name}")
                
                gestor = GestorCarpetas(Config.BASE_PAYPAL)
                carpeta_pago, carpeta_soporte = gestor.crear_estructura_pago(self.numero_pago)
                self.carpeta_soporte = carpeta_soporte
                
                self.archivo_movido = procesador.mover_y_renombrar_descarga(archivo, carpeta_pago, self.numero_pago)
                self.log_message(f"📁 Archivo movido a: {carpeta_pago}")
                
                procesador.reorganizar_columnas_primera_hoja(self.archivo_movido)
                self.log_message(" Columnas reorganizadas")
                
                if Config.RUTA_MAESTRO.exists():
                    self.df_segunda = procesador.crear_segunda_hoja(
                        self.archivo_movido, 
                        Config.RUTA_MAESTRO,
                        mes_filtro=getattr(self, 'mes_pago', None),
                        año_filtro=getattr(self, 'año_pago', None)
                    )
                    self.log_message(f"Segunda hoja creada con {len(self.df_segunda)} registros")
                    
                    self.df_segunda = procesador.calcular_mon_grupo_y_diferencia(self.archivo_movido, self.df_segunda)
                    
                    procesador.guardar_excel_con_dos_hojas(self.archivo_movido, self.df_segunda)
                    self.log_message(" Procesamiento inicial de Excel completado")
                else:
                    self.log_message("⚠️ Archivo maestro no encontrado")
            else:
                self.log_message(" No se encontró archivo para procesar")
                
        except Exception as e:
            self.log_message(f" Error procesando Excel: {e}")
    
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
            self.log_message(f" Error en actualización de maestro: {e}")
    
    def _on_workflow_completed(self):
        """Maneja la finalización del workflow"""
        self.operation_running = False
        self.main_progress.set(1.0)
        self.progress_label.configure(text="100%")
        
        self.log_message(" Proceso completado exitosamente")
        
        summary = f"""
            Pago #{self.numero_pago} completado:
            Carpetas verificadas
            Archivo SAP descargado
            PDFs buscados
            Excel procesado
            Maestro actualizado

            El sistema está listo para el siguiente pago.
        """
        
        self.summary_text.configure(state="normal")
        self.summary_text.delete("1.0", "end")
        self.summary_text.insert("1.0", summary.strip())
        self.summary_text.configure(state="disabled")
        
        self.show_state(STATE_COMPLETED)
    
    def continue_workflow(self):
        """Continuar con otro proceso"""
        self.numero_pago += 1
        # El número de pago se mostrará automáticamente en la siguiente ejecución
        
        self.btn_ejecutar.configure(state="normal", text=" EJECUTAR PROCESO COMPLETO")
        self.show_state(STATE_IDLE)
    
    def on_close(self):
        """Manejo del cierre de la aplicación"""
        try:
            if self.operation_running:
                if not messagebox.askyesno("Salir", "Hay una operación en progreso. ¿Desea salir igualmente?"):
                    return
            
            # Si estamos en configuración obligatoria y el usuario cierra
            if getattr(self, '_config_pendiente', False):
                self.logger.info("Cerrando aplicación sin completar configuración.")
            else:
                self.logger.info("Cerrando aplicación...")
                
            self.destroy()
            sys.exit(0)
        except Exception:
            # En caso de error al cerrar, forzar salida
            sys.exit(0)
    
    # ════════════════════════════════════════════════════════════════════
    # CONFIGURACIÓN DE RUTAS
    # ════════════════════════════════════════════════════════════════════

    def _mostrar_config_obligatoria(self):
        """Muestra config con mensaje de bienvenida en primera ejecución."""
        from tkinter import messagebox as mb
        respuesta = mb.askokcancel(
            "Configuración inicial",
            "Bienvenido al Sistema de Pagos PayPal.\n\n"
            "Es necesario configurar las rutas de trabajo para que la aplicación funcione.\n\n"
            "¿Desea configurar las rutas ahora?"
        )
        if not respuesta:
            self.logger.info("El usuario canceló la configuración inicial. Cerrando...")
            self.destroy()
            sys.exit(0)
            
        self.show_state("config_rutas")

    def create_config_rutas_content(self):
        """
        Pantalla de configuración de rutas.
        - Rutas principales: BASE_PAYPAL y RUTA_MAESTRO
        - Rutas PDF: lista dinámica con + Agregar / ✕ Quitar
        """
        self.config_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")

        self.config_frame.grid_columnconfigure(0, weight=1)
        self.config_frame.grid_rowconfigure(0, weight=1)

        scroll = ctk.CTkScrollableFrame(self.config_frame, fg_color="transparent")
        scroll.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)
        scroll.grid_columnconfigure(0, weight=1)

        # ── Título ────────────────────────────────────────────────────
        ctk.CTkLabel(
            scroll,
            text="⚙️  Configuración de Rutas",
            font=("Roboto", 22, "bold"),
            text_color=COLOR_TEXT
        ).pack(anchor="w", padx=10, pady=(10, 4))

        ctk.CTkLabel(
            scroll,
            text="Configure las carpetas y archivos que usará el sistema.",
            font=("Roboto", 12),
            text_color=COLOR_TEXT_DIM
        ).pack(anchor="w", padx=10, pady=(0, 20))

        # ── Card: Rutas principales ───────────────────────────────────
        card_principal = ctk.CTkFrame(
            scroll, fg_color=COLOR_ACCENT_LIGHT, corner_radius=12
        )
        card_principal.pack(fill="x", padx=10, pady=(0, 16))
        card_principal.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            card_principal,
            text="📁  Rutas Principales",
            font=("Roboto", 14, "bold"),
            text_color=COLOR_TEXT
        ).pack(anchor="w", padx=20, pady=(16, 8))

        # Variables para rutas principales
        rutas_actuales = self.configurador.obtener_rutas() or {}

        self._var_base_paypal = ctk.StringVar(
            value=str(rutas_actuales.get("base_paypal", ""))
        )
        self._var_maestro = ctk.StringVar(
            value=str(rutas_actuales.get("ruta_maestro", ""))
        )

        _raiz_raw = rutas_actuales.get("raiz_swift_latam", None)
        self.var_raiz_swift = ctk.StringVar(
            value=str(_raiz_raw) if _raiz_raw is not None else ""
        )

        self._crear_campo_ruta(
            card_principal,
            "Carpeta BASE PAYPAL",
            self._var_base_paypal,
            es_archivo=False
        )
        self._crear_campo_ruta(
            card_principal,
            "Archivo MAESTRO (.xlsm)",
            self._var_maestro,
            es_archivo=True,
            filetypes=[("Excel con macros", "*.xlsm"), ("Excel", "*.xlsx")]
        )
        self._crear_campo_ruta(
            card_principal,
            "Raíz de Swift Latam",
            self.var_raiz_swift,
            es_archivo=False
        )

        # Espaciado inferior de la card
        ctk.CTkLabel(card_principal, text="", height=8, fg_color="transparent").pack()

        # ── Botones de acción ─────────────────────────────────────────
        btn_frame = ctk.CTkFrame(scroll, fg_color="transparent")
        btn_frame.pack(fill="x", padx=10, pady=(8, 20))

        ctk.CTkButton(
            btn_frame,
            text="💾  Guardar Configuración",
            command=self._guardar_config_rutas,
            fg_color=COLOR_PRIMARY,
            hover_color=("#060D6F", "#4A52A7"),
            font=("Roboto", 13, "bold"),
            height=46
        ).pack(side="left", fill="x", expand=True, padx=(0, 8))

        # Botón volver solo si ya hay config previa
        is_configured = False
        try:
            if hasattr(Config, 'esta_configurado'):
                is_configured = Config.esta_configurado()
            else:
                # Fallback por si hay problemas con la clase Config
                is_configured = (Config.BASE_PAYPAL is not None and 
                                Config.RUTA_MAESTRO is not None)
        except:
            is_configured = False

        if is_configured:
            ctk.CTkButton(
                btn_frame,
                text="↩️  Cancelar",
                command=lambda: self.show_state(STATE_IDLE),
                fg_color="transparent",
                border_color=COLOR_ACCENT_LIGHT,
                border_width=1,
                text_color=COLOR_TEXT_DIM,
                hover_color=COLOR_ACCENT_LIGHT,
                font=("Roboto", 13, "bold"),
                height=46
            ).pack(side="left", fill="x", expand=True, padx=(8, 0))

    def _crear_campo_ruta(self, parent, label: str, var: ctk.StringVar, es_archivo: bool = False, filetypes: list = None):
        """Crea un campo de texto + botón Buscar para una ruta."""
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="x", padx=20, pady=6)

        ctk.CTkLabel(
            frame,
            text=label,
            font=("Roboto", 11, "bold"),
            text_color=COLOR_TEXT_DIM
        ).pack(anchor="w", pady=(0, 4))

        row = ctk.CTkFrame(frame, fg_color="transparent")
        row.pack(fill="x")
        row.grid_columnconfigure(0, weight=1)

        entry = ctk.CTkEntry(
            row,
            textvariable=var,
            font=("Roboto", 11),
            height=36,
            fg_color=COLOR_ACCENT,
            border_color=COLOR_ACCENT_LIGHT,
            text_color=COLOR_TEXT
        )
        entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        def _buscar():
            from tkinter import filedialog
            if es_archivo:
                ft = filetypes or [("Todos", "*.*")]
                path = filedialog.askopenfilename(filetypes=ft)
            else:
                path = filedialog.askdirectory()
            if path:
                var.set(path)

        ctk.CTkButton(
            row,
            text="Buscar...",
            command=_buscar,
            width=90,
            height=36,
            fg_color=COLOR_ACCENT_LIGHT,
            text_color=COLOR_TEXT,
            hover_color=COLOR_BORDE if hasattr(self, 'COLOR_BORDE') else "#D0D0D0",
            font=("Roboto", 11)
        ).grid(row=0, column=1)

    def _agregar_fila_pdf(self, valor_inicial: str = ""):
        """Agrega una fila dinámica para una ruta PDF en la lista."""
        var = ctk.StringVar(value=valor_inicial)

        fila = ctk.CTkFrame(self._pdf_container, fg_color=COLOR_ACCENT, corner_radius=8)
        fila.pack(fill="x", pady=4)
        fila.grid_columnconfigure(0, weight=1)

        # Número de orden
        num = len(self._pdf_vars) + 1
        ctk.CTkLabel(
            fila,
            text=f"#{num}",
            font=("Roboto", 11, "bold"),
            text_color=COLOR_TEXT_DIM,
            width=28
        ).grid(row=0, column=0, sticky="w", padx=(10, 4), pady=8)

        entry = ctk.CTkEntry(
            fila,
            textvariable=var,
            font=("Roboto", 11),
            height=34,
            fg_color="white",
            border_color=COLOR_ACCENT_LIGHT,
            text_color=COLOR_TEXT,
            placeholder_text="Ruta de carpeta OneDrive..."
        )
        entry.grid(row=0, column=0, sticky="ew", padx=(40, 8), pady=8)

        def _buscar_pdf():
            from tkinter import filedialog
            path = filedialog.askdirectory(title="Seleccionar carpeta de PDFs")
            if path:
                var.set(path)

        ctk.CTkButton(
            fila,
            text="Buscar...",
            command=_buscar_pdf,
            width=85,
            height=34,
            fg_color=COLOR_ACCENT_LIGHT,
            text_color=COLOR_TEXT,
            font=("Roboto", 11)
        ).grid(row=0, column=1, padx=(0, 4), pady=8)

        def _quitar(f=fila, v=var):
            """Elimina esta fila de la lista."""
            f.destroy()
            if v in [item[1] for item in self._pdf_vars]:
                self._pdf_vars = [item for item in self._pdf_vars if item[1] is not v]
            self._renumerar_filas_pdf()

        ctk.CTkButton(
            fila,
            text="✕",
            command=_quitar,
            width=34,
            height=34,
            fg_color="transparent",
            border_color=COLOR_ERROR,
            border_width=1,
            text_color=COLOR_ERROR,
            hover_color="#FFEBEE",
            font=("Roboto", 12, "bold")
        ).grid(row=0, column=2, padx=(0, 10), pady=8)

        self._pdf_vars.append((fila, var))

    def _renumerar_filas_pdf(self):
        """Actualiza los números de orden tras eliminar una fila."""
        for i, (fila, _) in enumerate(self._pdf_vars):
            # El primer child del frame es el Label de número
            for child in fila.winfo_children():
                if isinstance(child, ctk.CTkLabel) and child.cget("width") == 28:
                    child.configure(text=f"#{i + 1}")
                    break

    def _guardar_config_rutas(self):
        """Valida y persiste la configuración de rutas."""
        from tkinter import messagebox as mb

        base = self._var_base_paypal.get().strip()
        maestro = self._var_maestro.get().strip()

        # Validación básica de campos obligatorios
        if not base or not maestro:
            mb.showerror(
                "Campos requeridos",
                "La carpeta BASE PAYPAL y el archivo MAESTRO son obligatorios."
            )
            return

        # Advertencia si alguna ruta no existe (puede ser red desconectada)
        rutas_inexistentes = []
        from pathlib import Path as _Path
        if not _Path(base).exists():
            rutas_inexistentes.append(f"• BASE PAYPAL: {base}")
        if not _Path(maestro).exists():
            rutas_inexistentes.append(f"• MAESTRO: {maestro}")

        if rutas_inexistentes:
            detalle = "\n".join(rutas_inexistentes)
            continuar = mb.askyesno(
                "Rutas no encontradas",
                f"Las siguientes rutas no son accesibles ahora:\n\n{detalle}\n\n"
                "¿Desea guardar de todas formas?\n"
                "(Puede ser una unidad de red temporalmente desconectada)"
            )
            if not continuar:
                return

        # Guardar
        ok = self.configurador.guardar_config(base, maestro, rutas_pdf= [], raiz_swift_latam = self.var_raiz_swift.get().strip())
        if not ok:
            mb.showerror("Error", "No se pudo guardar la configuración.")
            return

        # Inyectar en Config
        Config.cargar_desde_ini(self.configurador.obtener_rutas())

        mb.showinfo("Guardado", "✅ Configuración guardada correctamente.")
        self.show_state(STATE_IDLE)


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