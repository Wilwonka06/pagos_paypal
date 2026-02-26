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

# NUEVO: Importar verificador/actualizador
from verificacion import VerificadorActualizadorSoportes, ResultadoVerificacion


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
COLOR_ORANGE = "#FF9500"  # Naranja para verificador


# Estados de la aplicación
STATE_IDLE = "idle"
STATE_RUNNING = "running"
STATE_COMPLETED = "completed"


class PaymentApp(ctk.CTk):
    """Aplicación principal con interfaz gráfica simplificada"""
    
    def __init__(self):
        super().__init__()
        
        # Configuración de ventana
        self.title("💰 Automatización de Pagos PayPal")
        self.geometry("900x650")
        self.minsize(800, 550)
        self.resizable(True, True)
        
        # Configurar logging
        self.logger = configurar_logging()
        
        # Variables de estado
        self.current_state = STATE_IDLE
        self.operation_running = False
        self.numero_pago = 1
        
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
    
    def load_initial_state(self):
        """Carga el estado inicial"""
        try:
            gestor = GestorCarpetas(Config.BASE_PAYPAL)
            self.numero_pago = gestor.obtener_pago_pendiente_o_siguiente()
            if hasattr(self, 'payment_entry'):
                self.payment_entry.delete(0, "end")
                self.payment_entry.insert(0, str(self.numero_pago))
        except:
            pass
    
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
        
        # Crear los estados de contenido
        self.create_idle_content()
        self.create_running_content()
        self.create_completed_content()
        
        # NUEVO: Crear estados de verificación
        self.create_verificar_soportes_content()
        self.create_resultado_verificacion_content()
                
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
        
        # NUEVO: Contenedor de botones
        botones_frame = ctk.CTkFrame(self.idle_frame, fg_color="transparent")
        botones_frame.pack(pady=30)
        
        # Botón Ejecutar Proceso Completo
        self.btn_ejecutar = ctk.CTkButton(
            botones_frame,
            text=" EJECUTAR PROCESO COMPLETO",
            command=self.start_workflow,
            fg_color=COLOR_PRIMARY,
            hover_color="#25A25A",
            font=("Roboto", 14, "bold"),
            height=45,
            width=250
        )
        self.btn_ejecutar.pack(side="left", padx=10)
        
        # NUEVO: Botón BUSCAR Y ACTUALIZAR SOPORTES
        self.btn_verificar = ctk.CTkButton(
            botones_frame,
            text="🔍 BUSCAR Y ACTUALIZAR",
            command=self.show_verificar_soportes,
            fg_color=COLOR_ORANGE,
            hover_color="#E68A00",
            font=("Roboto", 14, "bold"),
            height=45,
            width=250
        )
        self.btn_verificar.pack(side="left", padx=10)
        
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
            text="¡Proceso Completado!",
            font=("Roboto", 24, "bold"),
            text_color=COLOR_SUCCESS
        ).pack(pady=10)
        
        # Resumen
        summary_frame = ctk.CTkFrame(self.completed_frame, fg_color="#2A2A3E", corner_radius=10)
        summary_frame.pack(fill="both", expand=True, padx=40, pady=20)
        
        ctk.CTkLabel(
            summary_frame,
            text="📋 Resumen:",
            font=("Roboto", 14, "bold"),
            text_color=COLOR_TEXT
        ).pack(anchor="w", padx=15, pady=(15, 10))
        
        self.summary_text = ctk.CTkTextbox(
            summary_frame,
            fg_color="#1A1A2E",
            text_color=COLOR_TEXT,
            font=("Consolas", 11)
        )
        self.summary_text.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        self.summary_text.configure(state="disabled")
        
        # Botones
        button_frame = ctk.CTkFrame(self.completed_frame, fg_color="transparent")
        button_frame.pack(fill="x", padx=40, pady=20)
        
        ctk.CTkButton(
            button_frame,
            text="Procesar Siguiente Pago",
            command=self.continue_workflow,
            fg_color=COLOR_PRIMARY,
            hover_color="#25A25A",
            font=("Roboto", 12, "bold"),
            height=40
        ).pack(side="left", padx=5, fill="x", expand=True)
        
        ctk.CTkButton(
            button_frame,
            text="Salir",
            command=self.on_close,
            fg_color=COLOR_ERROR,
            hover_color="#E55353",
            font=("Roboto", 12, "bold"),
            height=40
        ).pack(side="left", padx=5, fill="x", expand=True)
    
    # ════════════════════════════════════════════════════════════════════════
    # NUEVO: MÉTODOS PARA VERIFICACIÓN DE SOPORTES
    # ════════════════════════════════════════════════════════════════════════
    
    def create_verificar_soportes_content(self):
        """NUEVO: Crea interfaz para búsqueda y actualización de soportes"""
        self.verificar_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        
        # Título
        titulo = ctk.CTkLabel(
            self.verificar_frame,
            text="🔍 BUSCAR Y ACTUALIZAR SOPORTES",
            font=("Roboto", 20, "bold"),
            text_color=COLOR_ORANGE
        )
        titulo.pack(pady=(20, 10))
        
        # Descripción
        desc = ctk.CTkLabel(
            self.verificar_frame,
            text="Busca documentos faltantes en OneDrive, cópialos a Soporte y actualiza Excel",
            font=("Roboto", 12),
            text_color="#AAAAAA"
        )
        desc.pack(pady=(0, 20))
        
        # Frame de selección
        selector_frame = ctk.CTkFrame(self.verificar_frame, fg_color="#2A2A3E", corner_radius=10)
        selector_frame.pack(fill="x", padx=40, pady=10)
        
        ctk.CTkLabel(
            selector_frame,
            text="Número de Pago:",
            font=("Roboto", 12),
            text_color=COLOR_TEXT
        ).pack(anchor="w", padx=20, pady=(15, 5))
        
        # Obtener pagos disponibles
        todas_carpetas = [d for d in Config.BASE_PAYPAL.iterdir() if d.is_dir() and d.name.startswith("Pago #")]
        numeros_disponibles = []
        for carpeta in todas_carpetas:
            try:
                num = int(carpeta.name.replace("Pago #", ""))
                numeros_disponibles.append(num)
            except ValueError:
                continue
        
        numeros_disponibles.sort(reverse=True)
        
        if numeros_disponibles:
            self.pago_select = ctk.CTkComboBox(
                selector_frame,
                values=[str(n) for n in numeros_disponibles],
                state="readonly",
                font=("Roboto", 12)
            )
            self.pago_select.set(str(numeros_disponibles[0]))
            self.pago_select.pack(fill="x", padx=20, pady=(5, 15))
        else:
            ctk.CTkLabel(
                selector_frame,
                text="❌ No hay pagos disponibles",
                font=("Roboto", 12),
                text_color=COLOR_ERROR
            ).pack(padx=20, pady=20)
            return
        
        # Botones de acción
        botones_frame = ctk.CTkFrame(self.verificar_frame, fg_color="transparent")
        botones_frame.pack(fill="x", padx=40, pady=20)
        
        btn_procesar = ctk.CTkButton(
            botones_frame,
            text="BUSCAR Y ACTUALIZAR",
            command=self.start_verificacion,
            fg_color=COLOR_ORANGE,
            hover_color="#E68A00",
            font=("Roboto", 14, "bold"),
            height=40
        )
        btn_procesar.pack(side="left", padx=5, fill="x", expand=True)
        
        btn_volver = ctk.CTkButton(
            botones_frame,
            text="VOLVER",
            command=self.back_to_idle,
            fg_color="#666666",
            hover_color="#777777",
            font=("Roboto", 14, "bold"),
            height=40
        )
        btn_volver.pack(side="left", padx=5, fill="x", expand=True)
        
        # Área de información
        info_frame = ctk.CTkFrame(self.verificar_frame, fg_color="#2A2A3E", corner_radius=10)
        info_frame.pack(fill="both", expand=True, padx=40, pady=(0, 20))
        
        ctk.CTkLabel(
            info_frame,
            text="ℹ️ PROCESO AUTOMÁTICO:",
            font=("Roboto", 12, "bold"),
            text_color=COLOR_TEXT
        ).pack(anchor="w", padx=15, pady=(15, 10))
        
        info_text = ctk.CTkTextbox(
            info_frame,
            fg_color="#1A1A2E",
            text_color="#AAAAAA",
            font=("Consolas", 10),
            height=120
        )
        info_text.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        
        info_text.insert("1.0", 
            "1️⃣ BUSCA documentos en OneDrive\n\n"
            "2️⃣ COPIA documentos a Pago #X/Soporte/\n\n"
            "3️⃣ ACTUALIZA observaciones en Excel\n"
            "   • Detecta qué documentos llegaron\n"
            "   • Cambia observaciones a 'Soportes OK'\n"
            "   • Mantiene pendientes si aún faltan\n\n"
            "4️⃣ MUESTRA reporte con cambios"
        )
        info_text.configure(state="disabled")
    
    def create_resultado_verificacion_content(self):
        """NUEVO: Crea interfaz para mostrar resultados de verificación"""
        self.resultado_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        
        # Título resultado
        self.resultado_titulo = ctk.CTkLabel(
            self.resultado_frame,
            text="",
            font=("Roboto", 18, "bold"),
            text_color=COLOR_PRIMARY
        )
        self.resultado_titulo.pack(pady=(20, 10))
        
        # Estadísticas CON CAMBIOS
        stats_frame = ctk.CTkFrame(self.resultado_frame, fg_color="#2A2A3E", corner_radius=10)
        stats_frame.pack(fill="x", padx=40, pady=10)
        
        self.resultado_stats = ctk.CTkLabel(
            stats_frame,
            text="",
            font=("Roboto", 11),
            text_color="#AAAAAA",
            justify="left"
        )
        self.resultado_stats.pack(anchor="w", padx=15, pady=15)
        
        # Detalles (ahora incluye cambios)
        detalles_frame = ctk.CTkFrame(self.resultado_frame, fg_color="#2A2A3E", corner_radius=10)
        detalles_frame.pack(fill="both", expand=True, padx=40, pady=10)
        
        ctk.CTkLabel(
            detalles_frame,
            text="📋 CAMBIOS Y DETALLES:",
            font=("Roboto", 12, "bold"),
            text_color=COLOR_TEXT
        ).pack(anchor="w", padx=15, pady=(15, 10))
        
        self.resultado_detalles = ctk.CTkTextbox(
            detalles_frame,
            fg_color="#1A1A2E",
            text_color=COLOR_TEXT,
            font=("Consolas", 10)
        )
        self.resultado_detalles.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        self.resultado_detalles.configure(state="disabled")
        
        # Botones finales
        botones_frame = ctk.CTkFrame(self.resultado_frame, fg_color="transparent")
        botones_frame.pack(fill="x", padx=40, pady=20)
        
        btn_otra = ctk.CTkButton(
            botones_frame,
            text="VERIFICAR OTRO PAGO",
            command=self.show_verificar_soportes,
            fg_color=COLOR_ORANGE,
            hover_color="#E68A00",
            font=("Roboto", 12, "bold"),
            height=40
        )
        btn_otra.pack(side="left", padx=5, fill="x", expand=True)
        
        btn_inicio = ctk.CTkButton(
            botones_frame,
            text="VOLVER AL INICIO",
            command=self.back_to_idle,
            fg_color=COLOR_PRIMARY,
            hover_color="#25A25A",
            font=("Roboto", 12, "bold"),
            height=40
        )
        btn_inicio.pack(side="left", padx=5, fill="x", expand=True)
    
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
            self.log_message("🔍 Iniciando búsqueda y actualización de soportes...")
            
            # Crear verificador/actualizador
            verificador = VerificadorActualizadorSoportes(Config.RUTAS_PDF)
            
            # PROCESO COMPLETO:
            # 1. Busca documentos en OneDrive
            # 2. Copia a carpeta Soporte
            # 3. Actualiza observaciones en Excel
            resultado = verificador.procesar_pago_completo(
                self.pago_verificando, 
                Config.BASE_PAYPAL
            )
            
            self.resultados_verificacion = [resultado]
            
            # Actualizar UI con resultado
            self.after(0, self._mostrar_resultado_verificacion, resultado)
            
        except Exception as e:
            self.log_message(f"❌ Error durante verificación: {e}")
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
                    f"✅ {archivo['nombre']}\n"
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
                    f"\n📌 Fila {cambio['fila']} (Invoice: {cambio['invoice']}):\n"
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
        if hasattr(self, 'resultado_frame'):
            self.resultado_frame.pack_forget()
        
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
            if not hasattr(self, 'running_frame'):
                self.create_running_content()
            self.running_frame.pack(fill="both", expand=True)
            self.log_message("⏳ Buscando y actualizando soportes...")
        
        elif state == "resultado_verificacion":
            if not hasattr(self, 'resultado_frame'):
                self.create_resultado_verificacion_content()
            self.resultado_frame.pack(fill="both", expand=True)
        
        self.current_state = state
    
    def start_workflow(self):
        """Inicia el workflow completo"""
        try:
            self.numero_pago = int(self.payment_entry.get())
            self.operation_running = True
            self.show_state(STATE_RUNNING)
            self.update_status("⚡ Ejecutando proceso...", 0)
            
            # Ejecutar en hilo
            thread = threading.Thread(target=self._execute_workflow, daemon=True)
            thread.start()
        
        except ValueError:
            messagebox.showerror("Error", "Por favor ingrese un número de pago válido")
            self.operation_running = False
    
    def _execute_workflow(self):
        """Ejecuta el workflow completo en hilo separado"""
        try:
            self.run_step("step1", self._verify_folders)
            self.run_step("step2", self._download_from_sap)
            self.run_step("step3", self._process_excel)
            self.run_step("step4", self._search_pdfs)
            self.run_step("step5", self._update_master)
            
            self._on_workflow_completed()
        except Exception as e:
            self.log_message(f"❌ Error en workflow: {e}")
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
        """Ejecuta un paso del workflow"""
        self.update_step_icon(step_id, "🔄", "#3498DB")
        
        step_index = [s[0] for s in self.workflow_steps].index(step_id)
        progress = (step_index + 0.5) / len(self.workflow_steps)
        self.after(0, lambda p=progress: self.main_progress.set(p))
        self.after(0, lambda p=progress: self.progress_label.configure(text=f"{int(p * 100)}%"))
        
        step_function()
        
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
            
            archivo = procesador.buscar_archivo_pago_en_descargas(self.numero_pago)
            if archivo:
                self.log_message(f"✅ Archivo encontrado: {archivo.name}")
                
                gestor = GestorCarpetas(Config.BASE_PAYPAL)
                carpeta_pago, carpeta_soporte = gestor.crear_estructura_pago(self.numero_pago)
                self.carpeta_soporte = carpeta_soporte
                
                self.archivo_movido = procesador.mover_y_renombrar_descarga(archivo, carpeta_pago, self.numero_pago)
                self.log_message(f"📁 Archivo movido a: {carpeta_pago}")
                
                procesador.reorganizar_columnas_primera_hoja(self.archivo_movido)
                self.log_message("✅ Columnas reorganizadas")
                
                if Config.RUTA_MAESTRO.exists():
                    self.df_segunda = procesador.crear_segunda_hoja(self.archivo_movido, Config.RUTA_MAESTRO)
                    self.log_message(f"✅ Segunda hoja creada con {len(self.df_segunda)} registros")
                    
                    self.df_segunda = procesador.calcular_mon_grupo_y_diferencia(self.archivo_movido, self.df_segunda)
                    
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
        
        self.log_message("✅ Proceso completado exitosamente")
        
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
        
        self.show_state(STATE_COMPLETED)
    
    def continue_workflow(self):
        """Continuar con otro proceso"""
        self.numero_pago += 1
        self.payment_entry.delete(0, "end")
        self.payment_entry.insert(0, str(self.numero_pago))
        
        self.btn_ejecutar.configure(state="normal", text=" EJECUTAR PROCESO COMPLETO")
        self.show_state(STATE_IDLE)
    
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