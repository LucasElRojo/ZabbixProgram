"""
Zabbix Metrics Extractor - GUI Module
Modern CustomTkinter interface with filtering and multi-select.
"""

import customtkinter as ctk
from tkinter import messagebox
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
import logging
import os
import sys
from typing import Optional, List, Dict, Any, Callable

from zabbix_client import ZabbixClient
from chart_downloader import ChartDownloader
from trend_analyzer import TrendAnalyzer
from pdf_generator import PDFReportGenerator
from config_storage import ConfigStorage

# Configure logging
logger = logging.getLogger(__name__)


class ConsoleHandler(logging.Handler):
    """Custom logging handler that outputs to the GUI console."""
    
    def __init__(self, callback: Callable[[str], None]):
        super().__init__()
        self.callback = callback
        self.setFormatter(logging.Formatter('%(asctime)s - %(message)s', datefmt='%H:%M:%S'))
    
    def emit(self, record):
        msg = self.format(record)
        self.callback(msg)


class ZabbixExtractorApp(ctk.CTk):
    """Main application window for Zabbix Metrics Extractor."""
    
    def __init__(self):
        super().__init__()
        
        # Configure window
        self.title("Zabbix Metrics Extractor")
        self.geometry("1200x800")
        self.minsize(1100, 700)
        
        # Set appearance
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        
        # Initialize components
        self.zabbix_client = ZabbixClient()
        self.chart_downloader: Optional[ChartDownloader] = None
        
        # Data storage
        self.templates: List[Dict[str, Any]] = []
        self.hosts: List[Dict[str, Any]] = []
        self.all_items: Dict[str, List[Dict[str, Any]]] = {}  # host_id -> items
        
        # Selection tracking
        self.selected_hosts: Dict[str, Dict[str, Any]] = {}  # host_id -> host
        self.selected_items: Dict[str, Dict[str, Any]] = {}  # item_id -> {item, host}
        self.selected_template: Optional[Dict[str, Any]] = None  # Currently selected template
        
        # Accordion UI state for host items
        self.host_accordion_frames: Dict[str, dict] = {}  # host_id -> accordion info
        self.expanded_host_id: Optional[str] = None  # Currently expanded host
        
        # Config storage for saved connections and templates
        self.config_storage = ConfigStorage()
        
        # Build UI
        self._create_widgets()
        self._setup_logging()
        
        # Get base path for downloads
        if getattr(sys, 'frozen', False):
            self.base_path = os.path.dirname(sys.executable)
        else:
            self.base_path = os.path.dirname(os.path.abspath(__file__))
    
    def _create_widgets(self):
        """Create all UI widgets."""
        # Main container
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # === Connection Section ===
        self._create_connection_frame()
        
        # === Main Content (4 columns: Templates/Hosts, Items, Summary, Console) ===
        content_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        content_frame.pack(fill="both", expand=True, pady=(10, 0))
        content_frame.columnconfigure(0, weight=1)
        content_frame.columnconfigure(1, weight=1)
        content_frame.columnconfigure(2, weight=1)
        content_frame.columnconfigure(3, weight=1)  # Console column
        content_frame.rowconfigure(0, weight=1)
        
        # Left column: Templates & Hosts
        self._create_left_column(content_frame)
        
        # Middle column: Items
        self._create_middle_column(content_frame)
        
        # Right column: Summary & Actions
        self._create_right_column(content_frame)
        
        # Far right column: Console
        self._create_console_frame(content_frame)
    
    def _create_connection_frame(self):
        """Create connection input fields."""
        conn_frame = ctk.CTkFrame(self.main_frame)
        conn_frame.pack(fill="x", pady=(0, 5))
        
        # Title
        ctk.CTkLabel(conn_frame, text="🔗 Conexión Zabbix", 
                     font=ctk.CTkFont(size=16, weight="bold")).pack(anchor="w", padx=10, pady=5)
        
        # Input fields container
        inputs_frame = ctk.CTkFrame(conn_frame, fg_color="transparent")
        inputs_frame.pack(fill="x", padx=10, pady=5)
        
        # URL
        ctk.CTkLabel(inputs_frame, text="URL:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.url_entry = ctk.CTkEntry(inputs_frame, width=280, placeholder_text="http://zabbix.ejemplo.com")
        self.url_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        
        # User
        ctk.CTkLabel(inputs_frame, text="Usuario:").grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.user_entry = ctk.CTkEntry(inputs_frame, width=120, placeholder_text="Admin")
        self.user_entry.grid(row=0, column=3, padx=5, pady=5)
        
        # Password
        ctk.CTkLabel(inputs_frame, text="Contraseña:").grid(row=0, column=4, padx=5, pady=5, sticky="w")
        self.password_entry = ctk.CTkEntry(inputs_frame, width=120, show="•", placeholder_text="••••••")
        self.password_entry.grid(row=0, column=5, padx=5, pady=5)
        
        # Connect Button
        self.connect_btn = ctk.CTkButton(inputs_frame, text="Conectar", width=100, 
                                         command=self._on_connect)
        self.connect_btn.grid(row=0, column=6, padx=10, pady=5)
        
        # Status indicator
        self.status_label = ctk.CTkLabel(inputs_frame, text="● Desconectado", text_color="gray")
        self.status_label.grid(row=0, column=7, padx=5, pady=5)
        
        # Saved Connections button
        ctk.CTkButton(inputs_frame, text="📋 Conexiones", width=110, fg_color="gray40",
                      hover_color="gray50", command=self._show_connections_dialog
                      ).grid(row=0, column=8, padx=5, pady=5)
        
        inputs_frame.columnconfigure(1, weight=1)
    
    def _create_left_column(self, parent):
        """Create left column with templates and hosts."""
        left_frame = ctk.CTkFrame(parent)
        left_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        
        # === Templates Section ===
        ctk.CTkLabel(left_frame, text="📋 Templates", 
                     font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=10, pady=(10, 5))
        
        # Template search
        self.template_search = ctk.CTkEntry(left_frame, placeholder_text="🔍 Buscar template...")
        self.template_search.pack(fill="x", padx=10, pady=(0, 5))
        self.template_search.bind("<KeyRelease>", self._on_template_search)
        
        # Template list
        self.template_listbox = ctk.CTkScrollableFrame(left_frame, height=120)
        self.template_listbox.pack(fill="x", padx=10, pady=(0, 10))
        self.template_buttons: Dict[str, ctk.CTkButton] = {}
        
        # === Hosts Section ===
        ctk.CTkLabel(left_frame, text="🖥️ Hosts (selección múltiple)", 
                     font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=10, pady=(10, 5))
        
        # Host search
        self.host_search = ctk.CTkEntry(left_frame, placeholder_text="🔍 Buscar host...")
        self.host_search.pack(fill="x", padx=10, pady=(0, 5))
        self.host_search.bind("<KeyRelease>", self._on_host_search)
        
        # Host list with checkboxes
        self.host_listbox = ctk.CTkScrollableFrame(left_frame, height=200)
        self.host_listbox.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.host_checkboxes: Dict[str, tuple] = {}  # host_id -> (checkbox, var, host)
        
        # Host action buttons
        host_btn_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
        host_btn_frame.pack(fill="x", padx=10, pady=(0, 10))
        ctk.CTkButton(host_btn_frame, text="Seleccionar Todos", width=120,
                      command=self._select_all_hosts).pack(side="left", padx=2)
        ctk.CTkButton(host_btn_frame, text="Deseleccionar", width=100,
                      command=self._deselect_all_hosts).pack(side="left", padx=2)
        self.load_items_btn = ctk.CTkButton(host_btn_frame, text="Cargar Items →", width=110,
                                            command=self._load_items_for_selected_hosts, state="disabled",
                                            fg_color="green", hover_color="darkgreen")
        self.load_items_btn.pack(side="right", padx=2)
    
    def _create_middle_column(self, parent):
        """Create middle column with items."""
        middle_frame = ctk.CTkFrame(parent)
        middle_frame.grid(row=0, column=1, sticky="nsew", padx=5)
        
        # === Common Items Section (for bulk selection) ===
        ctk.CTkLabel(middle_frame, text="⚡ Items Comunes (aplicar a todos)", 
                     font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=10, pady=(10, 5))
        
        # Info label
        self.common_items_info = ctk.CTkLabel(middle_frame, 
            text="Seleccione hosts y click 'Cargar Items' primero",
            text_color="gray60", font=ctk.CTkFont(size=11))
        self.common_items_info.pack(anchor="w", padx=10)
        
        # Search filter for common items
        self.common_items_search = ctk.CTkEntry(middle_frame, placeholder_text="🔍 Buscar item común...")
        self.common_items_search.pack(fill="x", padx=10, pady=(5, 0))
        self.common_items_search.bind("<KeyRelease>", self._on_common_items_search)
        
        # Common items list (compact)
        self.common_items_frame = ctk.CTkScrollableFrame(middle_frame, height=100)
        self.common_items_frame.pack(fill="x", padx=10, pady=(5, 5))
        self.common_item_checkboxes: Dict[str, tuple] = {}  # item_name -> (checkbox, var)
        self.all_common_items: List[str] = []  # Store all common items for filtering
        
        # Button row for common items
        common_btn_frame = ctk.CTkFrame(middle_frame, fg_color="transparent")
        common_btn_frame.pack(fill="x", padx=10, pady=(0, 10))
        
        # Button to load items from template (recommended)
        self.from_template_btn = ctk.CTkButton(common_btn_frame, 
            text="📋 Desde Template", width=120,
            command=self._load_items_from_template,
            fg_color="#1976d2", hover_color="#1565c0", state="disabled")
        self.from_template_btn.pack(side="left", padx=(0, 3))
        
        # Button to calculate common items (lazy - only when clicked)
        self.calc_common_btn = ctk.CTkButton(common_btn_frame, 
            text="⚡ Items Comunes", width=120,
            command=self._calculate_common_items_lazy,
            fg_color="#7b1fa2", hover_color="#6a1b9a", state="disabled")
        self.calc_common_btn.pack(side="left", padx=(0, 3))
        
        # Button to add selected common items to all hosts
        self.add_common_btn = ctk.CTkButton(common_btn_frame, 
            text="➕ Añadir", width=80,
            command=self._add_common_items_to_all_hosts,
            fg_color="#e65100", hover_color="#bf360c", state="disabled")
        self.add_common_btn.pack(side="left")
        
        # Separator
        ctk.CTkFrame(middle_frame, height=2, fg_color="gray40").pack(fill="x", padx=10, pady=5)
        
        # === Individual Items Section ===
        ctk.CTkLabel(middle_frame, text="📊 Items Individuales", 
                     font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=10, pady=(5, 5))
        
        # Item search
        self.item_search = ctk.CTkEntry(middle_frame, placeholder_text="🔍 Buscar item...")
        self.item_search.pack(fill="x", padx=10, pady=(0, 5))
        self.item_search.bind("<KeyRelease>", self._on_item_search)
        
        # Items list with checkboxes
        self.item_listbox = ctk.CTkScrollableFrame(middle_frame)
        self.item_listbox.pack(fill="both", expand=True, padx=10, pady=(0, 5))
        self.item_checkboxes: Dict[str, tuple] = {}  # unique_id -> (checkbox, var, item, host)
        
        # Item action buttons
        item_btn_frame = ctk.CTkFrame(middle_frame, fg_color="transparent")
        item_btn_frame.pack(fill="x", padx=10, pady=(0, 10))
        ctk.CTkButton(item_btn_frame, text="Sel. Todos", width=80,
                      command=self._select_all_items).pack(side="left", padx=2)
        ctk.CTkButton(item_btn_frame, text="Deseleccionar", width=90,
                      command=self._deselect_all_items).pack(side="left", padx=2)
        ctk.CTkButton(item_btn_frame, text="Añadir →", width=80,
                      command=self._add_selected_items, fg_color="green", 
                      hover_color="darkgreen").pack(side="right", padx=2)
    def _create_right_column(self, parent):
        """Create right column with summary and actions - SCROLLABLE."""
        # Outer container for grid placement
        right_container = ctk.CTkFrame(parent)
        right_container.grid(row=0, column=2, sticky="nsew", padx=(5, 0))
        right_container.rowconfigure(0, weight=1)
        right_container.columnconfigure(0, weight=1)
        
        # Scrollable inner frame for all content
        right_frame = ctk.CTkScrollableFrame(right_container)
        right_frame.grid(row=0, column=0, sticky="nsew")
        
        # === Summary Section ===
        ctk.CTkLabel(right_frame, text="📝 Resumen de Descarga", 
                     font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=10, pady=(10, 5))
        
        # Summary text - reduced height to fit better
        self.summary_text = ctk.CTkTextbox(right_frame, height=150, 
                                           font=ctk.CTkFont(family="Consolas", size=11))
        self.summary_text.pack(fill="x", padx=10, pady=(0, 10))
        
        # Clear summary button
        ctk.CTkButton(right_frame, text="🗑️ Limpiar Selección", width=140,
                      command=self._clear_selection).pack(pady=(0, 10))
        
        # === Time Period Section ===
        ctk.CTkLabel(right_frame, text="📅 Período de Tiempo", 
                     font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=10, pady=(10, 5))
        
        time_frame = ctk.CTkFrame(right_frame, fg_color="transparent")
        time_frame.pack(fill="x", padx=10, pady=(0, 10))
        
        self.time_period_var = ctk.StringVar(value="last_30_days")
        
        ctk.CTkRadioButton(time_frame, text="Últimos 30 días", 
                           variable=self.time_period_var, value="last_30_days").pack(anchor="w", pady=2)
        ctk.CTkRadioButton(time_frame, text="Mes anterior (completo)", 
                           variable=self.time_period_var, value="previous_month").pack(anchor="w", pady=2)
        ctk.CTkRadioButton(time_frame, text="Mes actual (hasta hoy)", 
                           variable=self.time_period_var, value="current_month").pack(anchor="w", pady=2)
        
        # === AI Conclusion Section ===
        ctk.CTkLabel(right_frame, text="🤖 Análisis con IA", 
                     font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=10, pady=(15, 5))
        
        # Checkbox to enable conclusion
        self.conclusion_var = ctk.StringVar(value="0")
        self.conclusion_checkbox = ctk.CTkCheckBox(
            right_frame, text="Generar Conclusión con IA",
            variable=self.conclusion_var, onvalue="1", offvalue="0",
            command=self._on_conclusion_toggle)
        self.conclusion_checkbox.pack(anchor="w", padx=10, pady=2)

        # AI provider selector
        self.ai_provider_var = ctk.StringVar(value="deepseek")
        provider_frame = ctk.CTkFrame(right_frame, fg_color="transparent")
        provider_frame.pack(fill="x", padx=10, pady=(2, 2))
        ctk.CTkLabel(provider_frame, text="Proveedor:",
                     font=ctk.CTkFont(size=11)).pack(side="left", padx=(0, 5))
        self.ai_provider_combo = ctk.CTkComboBox(
            provider_frame,
            values=["deepseek", "chatgpt"],
            variable=self.ai_provider_var,
            command=self._on_ai_provider_change,
            width=130
        )
        self.ai_provider_combo.pack(side="left")
        
        # API Key entry (hidden by default)
        self.api_key_frame = ctk.CTkFrame(right_frame, fg_color="transparent")
        self.api_key_frame.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkLabel(self.api_key_frame, text="API Key:", 
                     font=ctk.CTkFont(size=11)).pack(side="left", padx=(0, 5))
        self.api_key_entry = ctk.CTkEntry(self.api_key_frame, show="*", placeholder_text="sk-...", width=180)
        self.api_key_entry.pack(side="left", fill="x", expand=True)
        
        # Info label
        self.ai_info_label = ctk.CTkLabel(right_frame, 
            text="• Genera CSV + estadísticas\n• Conclusión técnica vía LLM",
            text_color="gray60", font=ctk.CTkFont(size=10), justify="left")
        self.ai_info_label.pack(anchor="w", padx=10)
        
        # PDF Report checkbox
        self.pdf_var = ctk.StringVar(value="0")
        self.pdf_checkbox = ctk.CTkCheckBox(
            right_frame, text="📄 Generar Informe PDF Ejecutivo",
            variable=self.pdf_var, onvalue="1", offvalue="0",
            command=self._on_pdf_toggle)
        self.pdf_checkbox.pack(anchor="w", padx=10, pady=(10, 5))
        
        # PDF Config button - always visible, prominent styling
        self.pdf_config_btn = ctk.CTkButton(right_frame, 
            text="📝 Configurar Reporte por Host",
            width=220, height=36, 
            fg_color="#e65100", hover_color="#bf360c",
            font=ctk.CTkFont(size=13, weight="bold"),
            command=self._show_report_config_dialog)
        self.pdf_config_btn.pack(anchor="w", padx=10, pady=(10, 10))
        
        ctk.CTkLabel(right_frame, 
            text="• Incidentes, riesgos por host\n• Uptime + dimensiones globales",
            text_color="gray60", font=ctk.CTkFont(size=10), justify="left").pack(anchor="w", padx=10)
        
        # Per-host report config (incidentes, riesgos, alertas per host)
        self.host_configs = {}  # host_name -> {incidentes, riesgos, alertas}
        
        # Global config (shared across all hosts)
        self.report_config = {
            'uptime_fecha': "",
            'uptime_servidor': "",
            'uptime_bd': "",
            'dim_rendimiento': "",
            'dim_contingencia': "",
            'dim_soporte': "",
            'dim_actualizaciones': "",
            'dim_respaldos': "",
        }
        self.report_defaults = {
            'incidentes': "No se presentan incidentes de servicio.",
            'riesgos': "No se registran riesgos del servicio durante el periodo.",
            'alertas': "No se evidencian alertas que afecten la continuidad operativa.",
            'dim_rendimiento': "Sin observaciones",
            'dim_contingencia': "Sin observaciones",
            'dim_soporte': "Sin observaciones",
            'dim_actualizaciones': "Sin observaciones",
            'dim_respaldos': "Sin observaciones",
        }
        
        # === Download Section ===
        self.download_btn = ctk.CTkButton(right_frame, text="⬇️ DESCARGAR GRÁFICOS", 
                                          height=50, font=ctk.CTkFont(size=16, weight="bold"),
                                          command=self._on_download, state="disabled",
                                          fg_color="#1a73e8", hover_color="#1557b0")
        self.download_btn.pack(fill="x", padx=10, pady=10)
        
        # Template buttons
        template_frame = ctk.CTkFrame(right_frame, fg_color="transparent")
        template_frame.pack(fill="x", padx=10, pady=(0, 5))
        ctk.CTkButton(template_frame, text="💾 Guardar Template", width=120,
                      fg_color="gray40", hover_color="gray50",
                      command=self._show_save_template_dialog).pack(side="left", padx=(0, 5))
        ctk.CTkButton(template_frame, text="📂 Cargar Template", width=120,
                      fg_color="gray40", hover_color="gray50", 
                      command=self._show_load_template_dialog).pack(side="left")
        
        self.progress_bar = ctk.CTkProgressBar(right_frame)
        self.progress_bar.pack(fill="x", padx=10, pady=(0, 10))
        self.progress_bar.set(0)
        
        # Initialize summary after all widgets are created
        self._update_summary()
    
    def _create_console_frame(self, parent):
        """Create console output widget in 4th column."""
        # Console container with border
        self.console_outer_frame = ctk.CTkFrame(parent)
        self.console_outer_frame.grid(row=0, column=3, sticky="nsew", padx=(5, 0))
        
        # Title with toggle buttons
        header_frame = ctk.CTkFrame(self.console_outer_frame, fg_color="transparent")
        header_frame.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkLabel(header_frame, text="📜 Consola de Logs", 
                     font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w")
        
        # Buttons in horizontal row
        btn_frame = ctk.CTkFrame(self.console_outer_frame, fg_color="transparent")
        btn_frame.pack(fill="x", padx=10, pady=(0, 5))
        ctk.CTkButton(btn_frame, text="Limpiar", width=70,
                      command=self._clear_console).pack(side="left", padx=2)
        
        # Console text widget - fills the column
        self.console_text = ctk.CTkTextbox(self.console_outer_frame, 
                                           font=ctk.CTkFont(family="Consolas", size=11))
        self.console_text.pack(fill="both", expand=True, padx=10, pady=(0, 10))
    
    def _minimize_console(self):
        """Minimize console to small size."""
        self.console_text.configure(height=80)
    
    def _maximize_console(self):
        """Maximize console to larger size."""
        self.console_text.configure(height=350)
    
    def _setup_logging(self):
        """Setup logging to console widget."""
        console_handler = ConsoleHandler(self._log_to_console)
        console_handler.setLevel(logging.INFO)
        
        root_logger = logging.getLogger()
        root_logger.setLevel(logging.INFO)
        root_logger.addHandler(console_handler)
    
    def _log_to_console(self, message: str):
        """Append message to console widget (thread-safe)."""
        def append():
            self.console_text.insert("end", message + "\n")
            self.console_text.see("end")
        self.after(0, append)
    
    def _clear_console(self):
        """Clear the console text."""
        self.console_text.delete("1.0", "end")
    
    # =========== CONNECTION ===========
    
    def _on_connect(self):
        """Handle connect button click."""
        url = self.url_entry.get().strip()
        user = self.user_entry.get().strip()
        password = self.password_entry.get()
        
        if not all([url, user, password]):
            messagebox.showerror("Error", "Por favor complete todos los campos de conexión.")
            return
        
        if not url.startswith(('http://', 'https://')):
            messagebox.showerror("Error", "La URL debe comenzar con http:// o https://")
            return
        
        self.connect_btn.configure(state="disabled", text="Conectando...")
        self._log_to_console("🔄 Intentando conexión a Zabbix...")
        self._log_to_console(f"📡 URL: {url}")
        
        def connect_thread():
            error_msg = None
            try:
                self.zabbix_client.connect(url, user, password)
                
                # Create ChartDownloader with web login (not API session)
                self.chart_downloader = ChartDownloader(
                    self.zabbix_client.get_base_url(),
                    user,
                    password
                )
                
                self.templates = self.zabbix_client.get_templates()
                self.after(0, self._on_connect_success)
                
            except Exception as e:
                error_msg = str(e)
                self.after(0, lambda msg=error_msg: self._on_connect_error(msg))
        
        threading.Thread(target=connect_thread, daemon=True).start()
    
    def _on_connect_success(self):
        """Handle successful connection."""
        self.status_label.configure(text="● Conectado", text_color="green")
        self.connect_btn.configure(state="normal", text="Reconectar")
        
        self._populate_templates()
        self._log_to_console(f"✅ Conexión exitosa. {len(self.templates)} templates encontrados.")
        
        # Ask to save connection if not already saved
        self.after(500, self._prompt_save_connection)
    
    def _on_connect_error(self, error: str):
        """Handle connection error."""
        self.status_label.configure(text="● Error", text_color="red")
        self.connect_btn.configure(state="normal", text="Conectar")
        self._log_to_console(f"❌ Error de conexión: {error}")
        messagebox.showerror("Error de Conexión", error)
    
    # =========== TEMPLATES ===========
    
    def _populate_templates(self):
        """Populate template list."""
        for widget in self.template_listbox.winfo_children():
            widget.destroy()
        self.template_buttons.clear()
        
        for template in self.templates:
            btn = ctk.CTkButton(self.template_listbox, text=template['name'],
                               anchor="w", fg_color="transparent", 
                               text_color=("gray10", "gray90"),
                               hover_color=("gray70", "gray30"),
                               command=lambda t=template: self._on_template_selected(t))
            btn.pack(fill="x", pady=1)
            self.template_buttons[template['templateid']] = btn
    
    def _on_template_search(self, event=None):
        """Filter templates based on search."""
        search_text = self.template_search.get().lower()
        
        for template_id, btn in self.template_buttons.items():
            template = next((t for t in self.templates if t['templateid'] == template_id), None)
            if template:
                if search_text == "" or template['name'].lower().startswith(search_text):
                    btn.pack(fill="x", pady=1)
                else:
                    btn.pack_forget()
    
    def _on_template_selected(self, template: Dict[str, Any]):
        """Handle template selection."""
        # Store selected template
        self.selected_template = template
        
        # Highlight selected template
        for tid, btn in self.template_buttons.items():
            if tid == template['templateid']:
                btn.configure(fg_color=("gray70", "gray30"))
            else:
                btn.configure(fg_color="transparent")
        
        # Enable 'Desde Template' button
        self.from_template_btn.configure(state="normal")
        
        self._log_to_console(f"📋 Cargando hosts para: {template['name']}")
        
        def fetch_hosts():
            try:
                self.hosts = self.zabbix_client.get_hosts_by_template(template['templateid'])
                self.after(0, self._populate_hosts)
            except Exception as e:
                error_msg = str(e)
                self.after(0, lambda msg=error_msg: self._log_to_console(f"❌ Error: {msg}"))
        
        threading.Thread(target=fetch_hosts, daemon=True).start()
    
    # =========== HOSTS ===========
    
    def _populate_hosts(self):
        """Populate host list with checkboxes."""
        for widget in self.host_listbox.winfo_children():
            widget.destroy()
        self.host_checkboxes.clear()
        
        if not self.hosts:
            ctk.CTkLabel(self.host_listbox, text="No hay hosts vinculados").pack(pady=10)
            return
        
        for host in self.hosts:
            var = ctk.StringVar(value="0")
            cb = ctk.CTkCheckBox(self.host_listbox, text=host['name'],
                                variable=var, onvalue="1", offvalue="0",
                                command=self._on_host_checkbox_change)
            cb.pack(anchor="w", pady=1)
            self.host_checkboxes[host['hostid']] = (cb, var, host)
        
        self._log_to_console(f"✅ {len(self.hosts)} hosts encontrados.")
    
    def _on_host_search(self, event=None):
        """Filter hosts based on search."""
        search_text = self.host_search.get().lower()
        
        for host_id, (cb, var, host) in self.host_checkboxes.items():
            if search_text == "" or search_text in host['name'].lower():
                cb.pack(anchor="w", pady=1)
            else:
                cb.pack_forget()
    
    def _on_host_checkbox_change(self):
        """Handle host checkbox change."""
        selected_count = sum(1 for _, (_, var, _) in self.host_checkboxes.items() if var.get() == "1")
        if selected_count > 0:
            self.load_items_btn.configure(state="normal", text=f"Cargar Items ({selected_count}) →")
        else:
            self.load_items_btn.configure(state="disabled", text="Cargar Items →")
    
    def _select_all_hosts(self):
        """Select all visible host checkboxes."""
        for host_id, (cb, var, host) in self.host_checkboxes.items():
            if cb.winfo_ismapped():
                var.set("1")
        self._on_host_checkbox_change()
    
    def _deselect_all_hosts(self):
        """Deselect all host checkboxes."""
        for host_id, (cb, var, host) in self.host_checkboxes.items():
            var.set("0")
        self._on_host_checkbox_change()
    
    def _load_items_for_selected_hosts(self):
        """Load items for all selected hosts using parallel threads."""
        selected_hosts = [(host_id, host) for host_id, (_, var, host) in self.host_checkboxes.items() if var.get() == "1"]
        
        if not selected_hosts:
            return
        
        self.load_items_btn.configure(state="disabled", text="Cargando...")
        host_count = len(selected_hosts)
        self._log_to_console(f"📊 Cargando items para {host_count} hosts (paralelo)...")
        
        def fetch_single_host(host_tuple):
            """Fetch items for a single host - runs in thread pool."""
            host_id, host = host_tuple
            try:
                items = self.zabbix_client.get_items_by_host(host_id)
                return (host_id, host, items, None)
            except Exception as e:
                return (host_id, host, [], str(e))
        
        def fetch_all_items_parallel():
            """Load all hosts in parallel using ThreadPoolExecutor."""
            try:
                self.all_items.clear()
                completed = 0
                errors = 0
                
                # Use up to 10 parallel workers
                max_workers = min(10, host_count)
                
                with ThreadPoolExecutor(max_workers=max_workers) as executor:
                    futures = {executor.submit(fetch_single_host, host): host for host in selected_hosts}
                    
                    for future in as_completed(futures):
                        host_id, host, items, error = future.result()
                        completed += 1
                        
                        if error:
                            errors += 1
                            self.after(0, lambda h=host, e=error: self._log_to_console(f"   ⚠️ {h['name']}: {e}"))
                        else:
                            self.all_items[host_id] = items
                            self.after(0, lambda h=host, c=len(items): self._log_to_console(f"   ✓ {h['name']}: {c} items"))
                
                self.after(0, lambda: self._log_to_console(f"✅ Completado: {completed - errors}/{host_count} hosts"))
                self.after(0, lambda: self._populate_items(selected_hosts))
                
                # Enable buttons for common items selection
                self.after(0, lambda: self.calc_common_btn.configure(state="normal"))
                self.after(0, lambda: self.common_items_info.configure(
                    text=f"Click '📋 Desde Template' o '⚡ Items Comunes' para cargar"))
                
            except Exception as e:
                error_msg = str(e)
                self.after(0, lambda msg=error_msg: self._log_to_console(f"❌ Error: {msg}"))
            finally:
                self.after(0, lambda: self.load_items_btn.configure(state="normal", text="Cargar Items →"))
        
        threading.Thread(target=fetch_all_items_parallel, daemon=True).start()
    
    # =========== ITEMS ===========
    
    def _populate_items(self, selected_hosts: List[tuple]):
        """Populate items list with accordion-style expandable hosts."""
        for widget in self.item_listbox.winfo_children():
            widget.destroy()
        self.item_checkboxes.clear()
        self.host_accordion_frames = {}  # host_id -> (header_btn, items_frame, expanded)
        self.expanded_host_id = None  # Only one host expanded at a time
        
        for host_id, host in selected_hosts:
            items = self.all_items.get(host_id, [])
            item_count = len(items)
            
            # Create host accordion header (clickable button)
            header_frame = ctk.CTkFrame(self.item_listbox, fg_color="transparent")
            header_frame.pack(fill="x", pady=2)
            
            header_btn = ctk.CTkButton(
                header_frame,
                text=f"▶ {host['name']} ({item_count} items)",
                anchor="w",
                fg_color="gray25",
                hover_color="gray35",
                font=ctk.CTkFont(weight="bold"),
                command=lambda hid=host_id, h=host: self._toggle_host_accordion(hid, h)
            )
            header_btn.pack(fill="x", padx=5)
            
            # Items frame (initially hidden)
            items_frame = ctk.CTkFrame(self.item_listbox, fg_color="gray20")
            # Don't pack yet - will be packed when expanded
            
            self.host_accordion_frames[host_id] = {
                'header': header_btn,
                'frame': items_frame,
                'host': host,
                'expanded': False,
                'loaded': False
            }
        
        host_count = len(selected_hosts)
        self._log_to_console(f"✅ {host_count} hosts disponibles (click para expandir items)")
    
    def _toggle_host_accordion(self, host_id: str, host: dict):
        """Toggle host accordion - expand/collapse items list."""
        if host_id not in self.host_accordion_frames:
            return
        
        accordion = self.host_accordion_frames[host_id]
        
        # Collapse currently expanded host (if different)
        if self.expanded_host_id and self.expanded_host_id != host_id:
            self._collapse_host_accordion(self.expanded_host_id)
        
        if accordion['expanded']:
            # Collapse this host
            self._collapse_host_accordion(host_id)
        else:
            # Expand this host
            self._expand_host_accordion(host_id, host)
    
    def _expand_host_accordion(self, host_id: str, host: dict):
        """Expand host accordion to show items."""
        accordion = self.host_accordion_frames[host_id]
        
        # Update header to show expanded state
        item_count = len(self.all_items.get(host_id, []))
        accordion['header'].configure(text=f"▼ {host['name']} ({item_count} items)")
        
        # Show items frame
        accordion['frame'].pack(fill="x", padx=10, pady=(0, 5), after=accordion['header'].master)
        
        # Load items if not already loaded
        if not accordion['loaded']:
            self._load_host_items_to_frame(host_id, host, accordion['frame'])
            accordion['loaded'] = True
        
        accordion['expanded'] = True
        self.expanded_host_id = host_id
    
    def _collapse_host_accordion(self, host_id: str):
        """Collapse host accordion to hide items."""
        if host_id not in self.host_accordion_frames:
            return
        
        accordion = self.host_accordion_frames[host_id]
        host = accordion['host']
        item_count = len(self.all_items.get(host_id, []))
        
        # Update header to show collapsed state
        accordion['header'].configure(text=f"▶ {host['name']} ({item_count} items)")
        
        # Hide items frame
        accordion['frame'].pack_forget()
        accordion['expanded'] = False
        
        if self.expanded_host_id == host_id:
            self.expanded_host_id = None
    
    def _load_host_items_to_frame(self, host_id: str, host: dict, frame: ctk.CTkFrame):
        """Load items for a specific host into its frame."""
        items = self.all_items.get(host_id, [])
        
        for item in items:
            var = ctk.StringVar(value="0")
            cb = ctk.CTkCheckBox(
                frame,
                text=f"{item['name']}",
                variable=var,
                onvalue="1",
                offvalue="0",
                command=self._on_item_checkbox_change
            )
            cb.pack(anchor="w", padx=10, pady=1)
            unique_id = f"{host_id}_{item['itemid']}"
            self.item_checkboxes[unique_id] = (cb, var, item, host)
    
    def _on_item_search(self, event=None):
        """Filter items based on search."""
        search_text = self.item_search.get().lower()
        
        for unique_id, (cb, var, item, host) in self.item_checkboxes.items():
            if search_text == "" or search_text in item['name'].lower() or search_text in item.get('key_', '').lower():
                cb.pack(anchor="w", padx=10, pady=1)
            else:
                cb.pack_forget()
    
    def _on_item_checkbox_change(self):
        """Handle item checkbox change."""
        pass  # Could show count if needed
    
    def _select_all_items(self):
        """Select all visible item checkboxes."""
        for unique_id, (cb, var, item, host) in self.item_checkboxes.items():
            if cb.winfo_ismapped():
                var.set("1")
    
    def _deselect_all_items(self):
        """Deselect all item checkboxes."""
        for unique_id, (cb, var, item, host) in self.item_checkboxes.items():
            var.set("0")
    
    def _add_selected_items(self):
        """Add selected items to download queue."""
        added = 0
        for unique_id, (cb, var, item, host) in self.item_checkboxes.items():
            if var.get() == "1":
                if unique_id not in self.selected_items:
                    self.selected_items[unique_id] = {'item': item, 'host': host}
                    added += 1
        
        if added > 0:
            self._log_to_console(f"✅ {added} items añadidos a la cola de descarga.")
            self._update_summary()
            self._deselect_all_items()
    
    # =========== COMMON ITEMS ===========
    
    def _populate_common_items(self):
        """Populate common items section with items shared across all hosts."""
        # Clear existing
        for widget in self.common_items_frame.winfo_children():
            widget.destroy()
        self.common_item_checkboxes.clear()
        
        if not self.all_items:
            self.common_items_info.configure(text="Seleccione hosts y click 'Cargar Items' primero")
            self.add_common_btn.configure(state="disabled")
            return
        
        # Find items that exist in ALL hosts (by name)
        host_item_names = []
        for host_id, items in self.all_items.items():
            item_names = {item['name'] for item in items}
            host_item_names.append(item_names)
        
        if not host_item_names:
            return
        
        # Intersection of all item names
        common_names = host_item_names[0]
        for names in host_item_names[1:]:
            common_names = common_names.intersection(names)
        
        common_names = sorted(common_names)
        
        if not common_names:
            self.common_items_info.configure(text="No hay items comunes entre los hosts seleccionados")
            self.add_common_btn.configure(state="disabled")
            return
        
        # Update info label
        host_count = len(self.all_items)
        self.common_items_info.configure(
            text=f"✨ {len(common_names)} items comunes en {host_count} hosts:")
        
        # Store all common items for search filtering
        self.all_common_items = list(common_names)
        
        # Create checkboxes for common items
        for item_name in common_names:
            var = ctk.StringVar(value="0")
            cb = ctk.CTkCheckBox(self.common_items_frame, text=item_name,
                                variable=var, onvalue="1", offvalue="0",
                                font=ctk.CTkFont(size=12))
            cb.pack(anchor="w", pady=1)
            self.common_item_checkboxes[item_name] = (cb, var)
        
        self.add_common_btn.configure(state="normal")
    
    def _add_common_items_to_all_hosts(self):
        """Add selected common items to ALL hosts."""
        # Get selected common item names
        selected_names = []
        for item_name, (cb, var) in self.common_item_checkboxes.items():
            if var.get() == "1":
                selected_names.append(item_name)
        
        if not selected_names:
            messagebox.showwarning("Aviso", "Seleccione al menos un item común.")
            return
        
        added_count = 0
        hosts_count = 0
        
        # For each host, find items with matching names and add them
        for host_id, items in self.all_items.items():
            # Get host info
            host = None
            for h_id, (_, _, h) in self.host_checkboxes.items():
                if h_id == host_id:
                    host = h
                    break
            
            if not host:
                continue
            
            hosts_count += 1
            
            for item in items:
                if item['name'] in selected_names:
                    unique_id = f"{host_id}_{item['itemid']}"
                    if unique_id not in self.selected_items:
                        self.selected_items[unique_id] = {'item': item, 'host': host}
                        added_count += 1
        
        if added_count > 0:
            self._log_to_console(f"⚡ {added_count} items añadidos ({len(selected_names)} tipos x {hosts_count} hosts)")
            self._update_summary()
            
            # Deselect common checkboxes
            for item_name, (cb, var) in self.common_item_checkboxes.items():
                var.set("0")
        else:
            self._log_to_console("⚠️ No se añadieron items (ya estaban en la cola)")
    
    def _calculate_common_items_lazy(self):
        """Calculate common items across hosts only when user requests it."""
        if not self.all_items:
            messagebox.showwarning("Aviso", "Primero cargue los items de los hosts seleccionados")
            return
        
        self.calc_common_btn.configure(state="disabled", text="Calculando...")
        self._log_to_console("⚡ Calculando items comunes...")
        
        def calculate_in_thread():
            try:
                # Get item names for each host
                host_item_names = []
                for host_id, items in self.all_items.items():
                    item_names = {item['name'] for item in items}
                    host_item_names.append(item_names)
                
                if not host_item_names:
                    self.after(0, lambda: self.common_items_info.configure(text="No hay items cargados"))
                    return
                
                # Calculate intersection
                common_names = host_item_names[0]
                for names in host_item_names[1:]:
                    common_names = common_names.intersection(names)
                
                common_names = sorted(common_names)
                host_count = len(self.all_items)
                
                # Update UI on main thread
                self.after(0, lambda: self._populate_common_items_result(common_names, host_count))
                
            except Exception as e:
                error_msg = str(e)
                self.after(0, lambda msg=error_msg: self._log_to_console(f"❌ Error: {msg}"))
            finally:
                self.after(0, lambda: self.calc_common_btn.configure(state="normal", text="⚡ Items Comunes"))
        
        threading.Thread(target=calculate_in_thread, daemon=True).start()
    
    def _populate_common_items_result(self, common_names: list, host_count: int):
        """Populate common items section with calculated results."""
        # Clear existing
        for widget in self.common_items_frame.winfo_children():
            widget.destroy()
        self.common_item_checkboxes.clear()
        
        if not common_names:
            self.common_items_info.configure(text="No hay items comunes entre los hosts")
            self.add_common_btn.configure(state="disabled")
            self._log_to_console("⚠️ No se encontraron items comunes")
            return
        
        # Update info label
        self.common_items_info.configure(
            text=f"✨ {len(common_names)} items comunes en {host_count} hosts:")
        
        # Store for filtering
        self.all_common_items = list(common_names)
        
        # Create checkboxes
        for item_name in common_names:
            var = ctk.StringVar(value="0")
            cb = ctk.CTkCheckBox(self.common_items_frame, text=item_name,
                                variable=var, onvalue="1", offvalue="0",
                                font=ctk.CTkFont(size=12))
            cb.pack(anchor="w", pady=1)
            self.common_item_checkboxes[item_name] = (cb, var)
        
        self.add_common_btn.configure(state="normal")
        self._log_to_console(f"✅ {len(common_names)} items comunes encontrados")
    
    def _on_common_items_search(self, event=None):
        """Filter common items based on search text (LIKE '%text%' matching)."""
        search_text = self.common_items_search.get().lower().strip()
        
        for item_name, (cb, var) in self.common_item_checkboxes.items():
            # Match if search text is anywhere in item name (case insensitive)
            if search_text == "" or search_text in item_name.lower():
                cb.pack(anchor="w", pady=1)
            else:
                cb.pack_forget()
    
    def _load_items_from_template(self):
        """Load items from the selected Zabbix template and populate common items section."""
        if not self.selected_template:
            messagebox.showwarning("Aviso", "Seleccione un template primero")
            return
        
        template_id = self.selected_template['templateid']
        template_name = self.selected_template['name']
        
        self.from_template_btn.configure(state="disabled", text="Cargando...")
        self._log_to_console(f"📋 Cargando items del template '{template_name}'...")
        
        def fetch_template_items():
            try:
                items = self.zabbix_client.get_items_by_template(template_id)
                item_names = sorted(set(item['name'] for item in items))
                
                self.after(0, lambda: self._populate_template_items(item_names))
                self.after(0, lambda: self._log_to_console(f"✅ {len(item_names)} items cargados desde template"))
            except Exception as e:
                error_msg = str(e)
                self.after(0, lambda msg=error_msg: self._log_to_console(f"❌ Error: {msg}"))
            finally:
                self.after(0, lambda: self.from_template_btn.configure(state="normal", text="📋 Desde Template"))
        
        threading.Thread(target=fetch_template_items, daemon=True).start()
    
    def _populate_template_items(self, item_names: list):
        """Populate common items section with items from template."""
        # Clear existing
        for widget in self.common_items_frame.winfo_children():
            widget.destroy()
        self.common_item_checkboxes.clear()
        
        if not item_names:
            self.common_items_info.configure(text="No se encontraron items en el template")
            self.add_common_btn.configure(state="disabled")
            return
        
        # Update info label
        self.common_items_info.configure(
            text=f"📋 {len(item_names)} items del template:")
        
        # Store all items for search filtering
        self.all_common_items = list(item_names)
        
        # Create checkboxes for items
        for item_name in item_names:
            var = ctk.StringVar(value="0")
            cb = ctk.CTkCheckBox(self.common_items_frame, text=item_name,
                                variable=var, onvalue="1", offvalue="0",
                                font=ctk.CTkFont(size=12))
            cb.pack(anchor="w", pady=1)
            self.common_item_checkboxes[item_name] = (cb, var)
        
        self.add_common_btn.configure(state="normal")
    
    # =========== SUMMARY ===========
    
    def _update_summary(self):
        """Update the download summary."""
        self.summary_text.delete("1.0", "end")
        
        if not self.selected_items:
            self.summary_text.insert("end", "No hay items seleccionados.\n\n")
            self.summary_text.insert("end", "Para añadir items:\n")
            self.summary_text.insert("end", "1. Seleccione un template\n")
            self.summary_text.insert("end", "2. Marque uno o más hosts\n")
            self.summary_text.insert("end", "3. Click 'Cargar Items'\n")
            self.summary_text.insert("end", "4. Seleccione items y click 'Añadir'\n")
            if hasattr(self, 'download_btn'):
                self.download_btn.configure(state="disabled")
            return
        
        # Group by host
        hosts_items: Dict[str, List] = {}
        for unique_id, data in self.selected_items.items():
            host_name = data['host']['name']
            if host_name not in hosts_items:
                hosts_items[host_name] = []
            hosts_items[host_name].append(data['item']['name'])
        
        # Write summary
        self.summary_text.insert("end", f"📊 TOTAL: {len(self.selected_items)} gráficos\n")
        self.summary_text.insert("end", f"🖥️ HOSTS: {len(hosts_items)}\n")
        self.summary_text.insert("end", "─" * 30 + "\n\n")
        
        for host_name, items in hosts_items.items():
            self.summary_text.insert("end", f"▶ {host_name}\n")
            for item_name in items:
                self.summary_text.insert("end", f"   • {item_name}\n")
            self.summary_text.insert("end", "\n")
        
        if hasattr(self, 'download_btn'):
            self.download_btn.configure(state="normal")
    
    def _clear_selection(self):
        """Clear all selected items."""
        self.selected_items.clear()
        self._update_summary()
        self._log_to_console("🗑️ Selección limpiada.")
    
    def _on_conclusion_toggle(self):
        """Handle conclusion checkbox toggle."""
        enabled = self.conclusion_var.get() == "1"
        if enabled:
            provider = self.ai_provider_var.get()
            provider_name = "DeepSeek" if provider == "deepseek" else "ChatGPT"
            self.ai_info_label.configure(text=f"✅ Análisis habilitado ({provider_name})\n• CSV + estadísticas + IA")
        else:
            self.ai_info_label.configure(text="• Genera CSV + estadísticas\n• Conclusión técnica vía LLM")

    def _on_ai_provider_change(self, _choice: str):
        """Refresh AI section status text when provider changes."""
        if self.conclusion_var.get() == "1":
            self._on_conclusion_toggle()
    
    def _on_pdf_toggle(self):
        """Handle PDF checkbox toggle - show/hide config button."""
        enabled = self.pdf_var.get() == "1"
        if enabled:
            self.pdf_config_btn.pack(anchor="w", padx=10, pady=(0, 5))
        else:
            self.pdf_config_btn.pack_forget()
    
    def _show_report_config_dialog(self):
        """Show dialog to configure report manual inputs with per-host support."""
        # Get unique hosts from selected items
        hosts_in_selection = set()
        for unique_id, data in self.selected_items.items():
            hosts_in_selection.add(data['host']['name'])
        
        if not hosts_in_selection:
            messagebox.showwarning("Aviso", "Primero seleccione items para descargar.")
            return
        
        hosts_list = sorted(hosts_in_selection)
        
        # Create dialog window - large and spacious
        dialog = ctk.CTkToplevel(self)
        dialog.title("📝 Configurar Reporte por Host")
        dialog.geometry("900x800")
        dialog.transient(self)
        dialog.grab_set()
        
        # Center it
        dialog.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - 900) // 2
        y = self.winfo_y() + (self.winfo_height() - 800) // 2
        dialog.geometry(f"+{x}+{y}")
        
        # Store current host entries
        self._current_host_entries = {}
        
        # === Host Selector ===
        host_selector_frame = ctk.CTkFrame(dialog)
        host_selector_frame.pack(fill="x", padx=15, pady=(15, 5))
        
        ctk.CTkLabel(host_selector_frame, text="🖥️ Host:", 
                     font=ctk.CTkFont(size=14, weight="bold")).pack(side="left", padx=(10, 10))
        
        self._selected_host_var = ctk.StringVar(value=hosts_list[0])
        host_dropdown = ctk.CTkComboBox(host_selector_frame, values=hosts_list, 
                                         variable=self._selected_host_var, width=400,
                                         command=self._on_host_config_change)
        host_dropdown.pack(side="left", padx=10)
        
        ctk.CTkLabel(host_selector_frame, text=f"({len(hosts_list)} hosts)", 
                     text_color="gray60").pack(side="left")
        
        # Scrollable content
        scroll_frame = ctk.CTkScrollableFrame(dialog)
        scroll_frame.pack(fill="both", expand=True, padx=15, pady=10)
        
        # === Per-Host: Incidentes, Riesgos, Alertas ===
        ctk.CTkLabel(scroll_frame, text="📋 Texto Operativo (por Host)", 
                     font=ctk.CTkFont(size=16, weight="bold")).pack(anchor="w", pady=(0, 10))
        
        ctk.CTkLabel(scroll_frame, text="Incidentes del Servicio:").pack(anchor="w")
        self._host_incidentes = ctk.CTkTextbox(scroll_frame, height=50)
        self._host_incidentes.pack(fill="x", pady=(0, 8))
        
        ctk.CTkLabel(scroll_frame, text="Riesgos del Servicio:").pack(anchor="w")
        self._host_riesgos = ctk.CTkTextbox(scroll_frame, height=50)
        self._host_riesgos.pack(fill="x", pady=(0, 8))
        
        ctk.CTkLabel(scroll_frame, text="Incidentes de Alerta:").pack(anchor="w")
        self._host_alertas = ctk.CTkTextbox(scroll_frame, height=50)
        self._host_alertas.pack(fill="x", pady=(0, 10))
        
        # === Per-Host: Uptime ===
        ctk.CTkLabel(scroll_frame, text="⏱️ Uptime (por Host)", 
                     font=ctk.CTkFont(size=16, weight="bold")).pack(anchor="w", pady=(15, 10))
        
        uptime_frame = ctk.CTkFrame(scroll_frame, fg_color="transparent")
        uptime_frame.pack(fill="x", pady=(0, 10))
        
        ctk.CTkLabel(uptime_frame, text="Fecha Último Inicio:").grid(row=0, column=0, sticky="w", pady=2)
        self._uptime_fecha = ctk.CTkEntry(uptime_frame, placeholder_text="ej: 01/02/2026", width=200)
        self._uptime_fecha.grid(row=0, column=1, padx=10, pady=2)
        
        ctk.CTkLabel(uptime_frame, text="Uptime Servidor:").grid(row=1, column=0, sticky="w", pady=2)
        self._uptime_server = ctk.CTkEntry(uptime_frame, placeholder_text="ej: 99.95%", width=200)
        self._uptime_server.grid(row=1, column=1, padx=10, pady=2)
        
        ctk.CTkLabel(uptime_frame, text="Uptime BD:").grid(row=2, column=0, sticky="w", pady=2)
        self._uptime_bd = ctk.CTkEntry(uptime_frame, placeholder_text="ej: 99.99%", width=200)
        self._uptime_bd.grid(row=2, column=1, padx=10, pady=2)
        ctk.CTkLabel(scroll_frame, text="📊 Dimensiones (Global)", 
                     font=ctk.CTkFont(size=16, weight="bold")).pack(anchor="w", pady=(15, 10))
        
        dim_frame = ctk.CTkFrame(scroll_frame, fg_color="transparent")
        dim_frame.pack(fill="x", pady=(0, 10))
        
        dims = [
            ("Rendimiento:", "dim_rendimiento"),
            ("Contingencia:", "dim_contingencia"),
            ("Soporte:", "dim_soporte"),
            ("Actualizaciones:", "dim_actualizaciones"),
            ("Respaldos:", "dim_respaldos"),
        ]
        
        self._dim_entries = {}
        for i, (label, key) in enumerate(dims):
            ctk.CTkLabel(dim_frame, text=label).grid(row=i, column=0, sticky="w", pady=2)
            entry = ctk.CTkEntry(dim_frame, placeholder_text="Sin observaciones", width=400)
            entry.grid(row=i, column=1, padx=10, pady=2)
            entry.insert(0, self.report_config.get(key, ''))
            self._dim_entries[key] = entry
        
        # Load current host data (after all fields created)
        self._load_host_config(hosts_list[0])
        
        ctk.CTkLabel(scroll_frame, text="💡 Dejar vacío = valores por defecto | Cambiar host guarda automáticamente",
                     text_color="gray60", font=ctk.CTkFont(size=11)).pack(anchor="w", pady=(10, 5))
        
        # Buttons
        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(fill="x", padx=15, pady=10)
        
        def save_all_and_close():
            # Save current host (includes incidentes, riesgos, alertas, uptime)
            self._save_current_host_config()
            # Save global config (only dimensions)
            for key, entry in self._dim_entries.items():
                self.report_config[key] = entry.get().strip()
            
            self._log_to_console(f"✅ Config guardada para {len(self.host_configs)} hosts")
            dialog.destroy()
        
        ctk.CTkButton(btn_frame, text="💾 Guardar y Cerrar", fg_color="green", hover_color="darkgreen",
                      command=save_all_and_close, width=140).pack(side="right", padx=5)
        ctk.CTkButton(btn_frame, text="Cancelar", fg_color="gray50",
                      command=dialog.destroy, width=100).pack(side="right", padx=5)
    
    def _on_host_config_change(self, new_host: str):
        """Handle host dropdown change - save current and load new."""
        self._save_current_host_config()
        self._load_host_config(new_host)
    
    def _save_current_host_config(self):
        """Save current host's text fields to host_configs (incidentes, riesgos, alertas, uptime)."""
        host = self._selected_host_var.get()
        self.host_configs[host] = {
            'incidentes': self._host_incidentes.get("1.0", "end-1c").strip(),
            'riesgos': self._host_riesgos.get("1.0", "end-1c").strip(),
            'alertas': self._host_alertas.get("1.0", "end-1c").strip(),
            'uptime_fecha': self._uptime_fecha.get().strip(),
            'uptime_servidor': self._uptime_server.get().strip(),
            'uptime_bd': self._uptime_bd.get().strip(),
        }
    
    def _load_host_config(self, host_name: str):
        """Load host config into text fields (incidentes, riesgos, alertas, uptime)."""
        config = self.host_configs.get(host_name, {})
        
        # Incidentes, Riesgos, Alertas
        self._host_incidentes.delete("1.0", "end")
        self._host_incidentes.insert("1.0", config.get('incidentes', ''))
        
        self._host_riesgos.delete("1.0", "end")
        self._host_riesgos.insert("1.0", config.get('riesgos', ''))
        
        self._host_alertas.delete("1.0", "end")
        self._host_alertas.insert("1.0", config.get('alertas', ''))
        
        # Uptime fields
        self._uptime_fecha.delete(0, "end")
        self._uptime_fecha.insert(0, config.get('uptime_fecha', ''))
        
        self._uptime_server.delete(0, "end")
        self._uptime_server.insert(0, config.get('uptime_servidor', ''))
        
        self._uptime_bd.delete(0, "end")
        self._uptime_bd.insert(0, config.get('uptime_bd', ''))
    
    # =========== DOWNLOAD ===========
    
    def _on_download(self):
        """Handle download button click."""
        if not self.selected_items:
            messagebox.showwarning("Aviso", "No hay items seleccionados para descargar.")
            return
        
        if not self.chart_downloader:
            messagebox.showerror("Error", "No hay conexión activa con Zabbix.")
            return
        
        self.download_btn.configure(state="disabled", text="Descargando...")
        self.progress_bar.set(0)
        
        period_type = self.time_period_var.get()
        
        def download_thread():
            try:
                # Check if AI conclusion is enabled
                generate_conclusion = self.conclusion_var.get() == "1"
                api_key = self.api_key_entry.get().strip() if generate_conclusion else None
                ai_provider = self.ai_provider_var.get()
                provider_name = "DeepSeek" if ai_provider == "deepseek" else "ChatGPT"
                trend_analyzer = None
                
                # Check if PDF report is enabled
                generate_pdf = self.pdf_var.get() == "1"
                pdf_generator = None
                
                if generate_conclusion:
                    self._log_to_console(f"🤖 Análisis con IA habilitado ({provider_name})")
                    trend_analyzer = TrendAnalyzer(self.zabbix_client.api, api_key, ai_provider)
                
                output_dir = ChartDownloader.create_output_folder(self.base_path)
                self._log_to_console(f"📁 Carpeta de salida: {output_dir}")
                
                # Initialize PDF generator if enabled
                if generate_pdf:
                    self._log_to_console("📄 Generación de PDF habilitada")
                    pdf_generator = PDFReportGenerator(output_dir)
                    pdf_generator.set_report_config(self.report_config, self.report_defaults)
                    pdf_generator.set_host_configs(self.host_configs)
                
                time_from, time_to = ChartDownloader.calculate_time_range(period_type)
                period_names = {
                    'last_30_days': 'Últimos 30 días',
                    'previous_month': 'Mes anterior',
                    'current_month': 'Mes actual'
                }
                self._log_to_console(f"📅 Período: {period_names.get(period_type, period_type)}")
                self._log_to_console(f"⏱️ Rango: {time_from} → {time_to}")
                self._log_to_console(f"🔗 Base URL: {self.chart_downloader.base_url}")
                self._log_to_console(f"🔑 Web login: {'✓ OK' if self.chart_downloader.logged_in else '✗ Failed'}")
                self._log_to_console("─" * 50)
                
                total = len(self.selected_items)
                success_count = 0
                completed = 0
                
                # Process items in parallel using ThreadPoolExecutor
                def process_single_item(args):
                    """Process a single item (download + analysis). Returns success status."""
                    idx, unique_id, data = args
                    item = data['item']
                    host = data['host']
                    
                    try:
                        self.after(0, lambda i=idx, h=host['name'], itm=item['name']: 
                            self._log_to_console(f"⬇️ [{i+1}/{total}] {h}: {itm}"))
                        
                        image_bytes = self.chart_downloader.download_chart(
                            item['itemid'], time_from, time_to
                        )
                        
                        if image_bytes:
                            filename = f"{host['name']}_{item['name']}"
                            chart_path, legend_path = self.chart_downloader.process_image(
                                image_bytes, filename, output_dir
                            )
                            self.after(0, lambda: self._log_to_console(f"   ✅ Gráfico descargado"))
                            
                            # AI Analysis if enabled
                            if (generate_conclusion or generate_pdf) and trend_analyzer:
                                try:
                                    stats, conclusion, trends = trend_analyzer.analyze_item(
                                        item['itemid'], item['name'], host['name'],
                                        time_from, time_to, period_names.get(period_type, period_type),
                                        output_dir
                                    )
                                    
                                    # Add data to PDF generator (thread-safe append)
                                    if pdf_generator and trends:
                                        pdf_generator.add_item_data(
                                            host['name'], item['name'], trends, stats, conclusion
                                        )
                                    return (True, stats, conclusion)
                                except Exception as ae:
                                    self.after(0, lambda e=str(ae): self._log_to_console(f"   ⚠️ Análisis: {e}"))
                            return (True, None, None)
                        else:
                            self.after(0, lambda: self._log_to_console(f"   ⚠️ Sin datos de imagen"))
                            return (False, None, None)
                            
                    except Exception as e:
                        self.after(0, lambda err=str(e): self._log_to_console(f"   ❌ Error: {err}"))
                        return (False, None, None)
                
                # Prepare items with index
                items_with_idx = [(idx, uid, data) for idx, (uid, data) in enumerate(self.selected_items.items())]
                
                # Use ThreadPoolExecutor for parallel downloads (max 5 workers)
                max_workers = min(5, total)
                with ThreadPoolExecutor(max_workers=max_workers) as executor:
                    futures = {executor.submit(process_single_item, args): args for args in items_with_idx}
                    
                    for future in as_completed(futures):
                        result = future.result()
                        if result[0]:  # Success
                            success_count += 1
                        completed += 1
                        progress = completed / total
                        self.after(0, lambda p=progress: self.progress_bar.set(p))
                
                self._log_to_console("")
                self._log_to_console("─" * 50)
                self._log_to_console(f"🎉 Completado: {success_count}/{total} gráficos guardados")
                self._log_to_console(f"📂 Ubicación: {output_dir}")
                
                # Generate PDF report if enabled
                if pdf_generator and pdf_generator.items_data:
                    self._log_to_console("")
                    
                    # Fetch storage data for each unique host
                    unique_hosts = {}
                    for data in self.selected_items.values():
                        host = data['host']
                        unique_hosts[host['hostid']] = host['name']
                    
                    self._log_to_console("💾 Obteniendo datos de almacenamiento...")
                    for host_id, host_name in unique_hosts.items():
                        try:
                            fs_data = self.zabbix_client.get_filesystem_stats(host_id)
                            if fs_data:
                                pdf_generator.add_storage_data(host_name, fs_data)
                                self._log_to_console(f"   📂 {host_name}: {len(fs_data)} particiones")
                        except Exception as fse:
                            self._log_to_console(f"   ⚠️ {host_name}: {str(fse)}")
                    
                    self._log_to_console("📄 Generando informe PDF ejecutivo...")
                    try:
                        pdf_path = pdf_generator.generate_report("informe_ejecutivo")
                        if pdf_path:
                            self._log_to_console(f"✅ PDF generado: {os.path.basename(pdf_path)}")
                        else:
                            self._log_to_console("⚠️ No se pudo generar el PDF")
                    except Exception as pe:
                        self._log_to_console(f"⚠️ Error generando PDF: {str(pe)}")
                
                if success_count > 0:
                    pdf_msg = "\n📄 Informe PDF generado" if (pdf_generator and pdf_generator.items_data) else ""
                    self.after(0, lambda: messagebox.showinfo("Completado", 
                        f"Descarga completada.\n{success_count}/{total} gráficos guardados en:\n{output_dir}{pdf_msg}"))
                else:
                    self.after(0, lambda: messagebox.showwarning("Advertencia",
                        f"No se pudieron descargar gráficos.\nRevise la consola para más detalles."))
                
            except Exception as e:
                import traceback
                error_msg = str(e)
                self._log_to_console(f"❌ Error: {error_msg}")
                self._log_to_console(f"❌ Traceback: {traceback.format_exc()}")
                self.after(0, lambda msg=error_msg: messagebox.showerror("Error", msg))
            
            finally:
                self.after(0, lambda: self.download_btn.configure(state="normal", text="⬇️ DESCARGAR GRÁFICOS"))
        
        threading.Thread(target=download_thread, daemon=True).start()
    
    # ============ SAVED CONNECTIONS ============
    
    def _show_connections_dialog(self):
        """Show dialog to manage saved Zabbix connections."""
        dialog = ctk.CTkToplevel(self)
        dialog.title("📋 Conexiones Guardadas")
        dialog.geometry("700x500")
        dialog.transient(self)
        dialog.grab_set()
        
        # Center dialog
        dialog.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - 700) // 2
        y = self.winfo_y() + (self.winfo_height() - 500) // 2
        dialog.geometry(f"+{x}+{y}")
        
        # Header
        ctk.CTkLabel(dialog, text="📋 Conexiones Guardadas", 
                     font=ctk.CTkFont(size=18, weight="bold")).pack(pady=15)
        
        # Password visibility toggle
        self._show_passwords = ctk.BooleanVar(value=False)
        
        # Scrollable list
        list_frame = ctk.CTkScrollableFrame(dialog, height=300)
        list_frame.pack(fill="both", expand=True, padx=15, pady=10)
        
        def refresh_list():
            # Clear existing
            for widget in list_frame.winfo_children():
                widget.destroy()
            
            connections = self.config_storage.get_connections()
            
            if not connections:
                ctk.CTkLabel(list_frame, text="No hay conexiones guardadas", 
                             text_color="gray").pack(pady=20)
                return
            
            for conn in connections:
                conn_frame = ctk.CTkFrame(list_frame, fg_color="gray25")
                conn_frame.pack(fill="x", pady=5, padx=5)
                
                # Info column
                info_frame = ctk.CTkFrame(conn_frame, fg_color="transparent")
                info_frame.pack(side="left", fill="x", expand=True, padx=10, pady=8)
                
                ctk.CTkLabel(info_frame, text=conn['name'],
                             font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w")
                ctk.CTkLabel(info_frame, text=f"🌐 {conn['url']}", 
                             text_color="gray70", font=ctk.CTkFont(size=11)).pack(anchor="w")
                
                user_text = f"👤 {conn['username']}"
                pwd_display = conn['password'] if self._show_passwords.get() else "••••••••"
                ctk.CTkLabel(info_frame, text=f"{user_text}  |  🔑 {pwd_display}", 
                             text_color="gray60", font=ctk.CTkFont(size=11)).pack(anchor="w")
                
                # Buttons column
                btn_frame = ctk.CTkFrame(conn_frame, fg_color="transparent")
                btn_frame.pack(side="right", padx=10, pady=8)
                
                ctk.CTkButton(btn_frame, text="➡️ Conectar", width=90, 
                              fg_color="green", hover_color="darkgreen",
                              command=lambda c=conn: connect_to(c)).pack(side="left", padx=3)
                
                ctk.CTkButton(btn_frame, text="🗑️", width=40, 
                              fg_color="red", hover_color="darkred",
                              command=lambda c=conn: delete_conn(c)).pack(side="left", padx=3)
        
        def connect_to(conn):
            """Fill connection fields and connect."""
            self.url_entry.delete(0, "end")
            self.url_entry.insert(0, conn['url'])
            self.user_entry.delete(0, "end")
            self.user_entry.insert(0, conn['username'])
            self.password_entry.delete(0, "end")
            self.password_entry.insert(0, conn['password'])
            dialog.destroy()
            self._on_connect()
        
        def delete_conn(conn):
            """Delete a connection after confirmation."""
            if messagebox.askyesno("Confirmar", f"¿Eliminar conexión '{conn['name']}'?"):
                self.config_storage.delete_connection(conn['id'])
                refresh_list()
        
        def toggle_passwords():
            refresh_list()
        
        # Bottom controls
        controls_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        controls_frame.pack(fill="x", padx=15, pady=10)
        
        ctk.CTkCheckBox(controls_frame, text="👁️ Mostrar contraseñas", 
                        variable=self._show_passwords,
                        command=toggle_passwords).pack(side="left")
        
        ctk.CTkButton(controls_frame, text="Cerrar", fg_color="gray50",
                      command=dialog.destroy).pack(side="right")
        
        refresh_list()
    
    def _prompt_save_connection(self):
        """Auto-save current connection if not already saved."""
        if not self.zabbix_client or not self.zabbix_client.is_connected:
            return
        
        url = self.url_entry.get().strip()
        username = self.user_entry.get().strip()
        password = self.password_entry.get().strip()
        
        # Check if already saved
        for conn in self.config_storage.get_connections():
            if conn['url'].lower() == url.lower():
                return  # Already saved
        
        # Get server name from Zabbix API for auto-naming
        try:
            response = self.zabbix_client.api_request('apiinfo.version', {})
            server_name = f"Zabbix {response}"
        except:
            server_name = "Zabbix Server"
        
        # Auto-save without asking
        try:
            self.config_storage.add_connection(server_name, url, username, password)
            self._log_to_console(f"✅ Conexión '{server_name}' guardada automáticamente")
        except ValueError as e:
            self._log_to_console(f"⚠️ No se pudo guardar conexión: {e}")
    
    # ============ ITEM TEMPLATES ============
    
    def _show_save_template_dialog(self):
        """Show dialog to save current item selection as template."""
        if not self.selected_items:
            messagebox.showwarning("Aviso", "Seleccione items antes de guardar un template")
            return
        
        # Get current host URL
        current_url = self.url_entry.get().strip()
        if not current_url:
            messagebox.showwarning("Aviso", "Debe estar conectado a un Zabbix para guardar templates")
            return
        
        # Get unique item names from selection
        item_names = sorted(set(
            data['item']['name'] for data in self.selected_items.values()
        ))
        
        dialog = ctk.CTkToplevel(self)
        dialog.title("💾 Guardar Template")
        dialog.geometry("500x450")
        dialog.transient(self)
        dialog.grab_set()
        
        # Center
        dialog.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - 500) // 2
        y = self.winfo_y() + (self.winfo_height() - 450) // 2
        dialog.geometry(f"+{x}+{y}")
        
        ctk.CTkLabel(dialog, text="💾 Guardar Template de Items",
                     font=ctk.CTkFont(size=16, weight="bold")).pack(pady=15)
        
        # Show linked host
        ctk.CTkLabel(dialog, text=f"🔗 Vinculado a: {current_url}",
                     text_color="gray60", font=ctk.CTkFont(size=11)).pack(anchor="w", padx=20)
        
        # Template name
        name_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        name_frame.pack(fill="x", padx=20, pady=10)
        ctk.CTkLabel(name_frame, text="Nombre:").pack(side="left")
        name_entry = ctk.CTkEntry(name_frame, width=300, placeholder_text="Ej: Métricas CPU")
        name_entry.pack(side="left", padx=10)
        
        # Preview
        ctk.CTkLabel(dialog, text=f"Items a guardar ({len(item_names)}):",
                     font=ctk.CTkFont(size=13)).pack(anchor="w", padx=20)
        
        preview_frame = ctk.CTkScrollableFrame(dialog, height=180)
        preview_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        for name in item_names:
            ctk.CTkLabel(preview_frame, text=f"• {name}", text_color="gray70").pack(anchor="w")
        
        # Buttons
        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=15)
        
        def save():
            template_name = name_entry.get().strip()
            if not template_name:
                messagebox.showerror("Error", "Ingrese un nombre para el template")
                return
            try:
                self.config_storage.add_template(template_name, item_names, current_url)
                self._log_to_console(f"✅ Template '{template_name}' guardado con {len(item_names)} items (host: {current_url})")
                dialog.destroy()
            except ValueError as e:
                messagebox.showerror("Error", str(e))
        
        ctk.CTkButton(btn_frame, text="💾 Guardar", fg_color="green",
                      command=save).pack(side="right", padx=5)
        ctk.CTkButton(btn_frame, text="Cancelar", fg_color="gray50",
                      command=dialog.destroy).pack(side="right", padx=5)
    
    def _show_load_template_dialog(self):
        """Show dialog to load and apply an item template for current host."""
        current_url = self.url_entry.get().strip()
        all_templates = self.config_storage.get_templates()
        
        # Filter templates by current host URL
        templates = [t for t in all_templates 
                     if t.get('host_url', '').lower() == current_url.lower()]
        
        if not templates:
            if all_templates:
                messagebox.showinfo("Info", 
                    f"No hay templates para el host actual.\n\nHay {len(all_templates)} templates para otros hosts.")
            else:
                messagebox.showinfo("Info", "No hay templates guardados")
            return
        
        dialog = ctk.CTkToplevel(self)
        dialog.title("📂 Cargar Template")
        dialog.geometry("600x480")
        dialog.transient(self)
        dialog.grab_set()
        
        # Center
        dialog.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - 600) // 2
        y = self.winfo_y() + (self.winfo_height() - 480) // 2
        dialog.geometry(f"+{x}+{y}")
        
        ctk.CTkLabel(dialog, text="📂 Cargar Template",
                     font=ctk.CTkFont(size=16, weight="bold")).pack(pady=15)
        
        # Show current host
        ctk.CTkLabel(dialog, text=f"🔗 Host actual: {current_url}",
                     text_color="gray60", font=ctk.CTkFont(size=11)).pack(anchor="w", padx=15)
        
        selected_template = [None]
        
        # Template list
        list_frame = ctk.CTkScrollableFrame(dialog, height=200)
        list_frame.pack(fill="both", expand=True, padx=15, pady=10)
        
        template_buttons = []
        
        def select_template(tmpl, btn):
            selected_template[0] = tmpl
            for b in template_buttons:
                b.configure(fg_color="gray30")
            btn.configure(fg_color="blue")
            update_preview(tmpl)
        
        for tmpl in templates:
            btn = ctk.CTkButton(list_frame, text=f"📋 {tmpl['name']} ({len(tmpl['items'])} items)",
                               fg_color="gray30", hover_color="gray40", anchor="w")
            btn.configure(command=lambda t=tmpl, b=btn: select_template(t, b))
            btn.pack(fill="x", pady=2)
            template_buttons.append(btn)
        
        # Preview
        ctk.CTkLabel(dialog, text="Vista previa:", font=ctk.CTkFont(size=13)).pack(anchor="w", padx=15)
        preview_text = ctk.CTkTextbox(dialog, height=100)
        preview_text.pack(fill="x", padx=15, pady=5)
        
        def update_preview(tmpl):
            preview_text.delete("1.0", "end")
            if tmpl:
                preview_text.insert("1.0", "\n".join(tmpl['items']))
        
        # Buttons
        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(fill="x", padx=15, pady=15)
        
        def apply():
            if not selected_template[0]:
                messagebox.showwarning("Aviso", "Seleccione un template")
                return
            self._apply_template(selected_template[0])
            dialog.destroy()
        
        def delete():
            if not selected_template[0]:
                messagebox.showwarning("Aviso", "Seleccione un template")
                return
            if messagebox.askyesno("Confirmar", f"¿Eliminar template '{selected_template[0]['name']}'?"):
                self.config_storage.delete_template(selected_template[0]['id'])
                dialog.destroy()
                self._show_load_template_dialog()  # Reopen
        
        ctk.CTkButton(btn_frame, text="✅ Aplicar", fg_color="green",
                      command=apply).pack(side="right", padx=5)
        ctk.CTkButton(btn_frame, text="🗑️ Eliminar", fg_color="red",
                      command=delete).pack(side="right", padx=5)
        ctk.CTkButton(btn_frame, text="Cerrar", fg_color="gray50",
                      command=dialog.destroy).pack(side="right", padx=5)
    
    def _apply_template(self, template: dict):
        """Apply a template by selecting matching items across all loaded hosts."""
        template_items = set(template['items'])
        found_count = 0
        not_found = []
        
        # Search for matching items in all hosts
        for host_id, items in self.all_items.items():
            host = next((h for h in self.hosts if h['hostid'] == host_id), None)
            if not host:
                continue
            
            for item in items:
                item_name = item.get('name', '')
                if item_name in template_items:
                    unique_id = f"{host_id}_{item['itemid']}"
                    if unique_id not in self.selected_items:
                        self.selected_items[unique_id] = {'item': item, 'host': host}
                        found_count += 1
        
        # Check what wasn't found
        found_names = set(data['item']['name'] for data in self.selected_items.values())
        not_found = [name for name in template_items if name not in found_names]
        
        # Update UI
        self._update_summary()
        
        msg = f"✅ Template aplicado: {found_count} items seleccionados"
        if not_found:
            msg += f"\n⚠️ {len(not_found)} items no encontrados"
        self._log_to_console(msg)
        
        if not_found:
            messagebox.showwarning("Aviso", 
                f"Algunos items del template no se encontraron en los hosts cargados:\n\n" +
                "\n".join(not_found[:5]) + 
                (f"\n...y {len(not_found)-5} más" if len(not_found) > 5 else ""))





def main():
    """Application entry point."""
    app = ZabbixExtractorApp()
    app.mainloop()


if __name__ == "__main__":
    main()
