import os
import sys
import json
import threading
import webbrowser
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import customtkinter as ctk
from tkcalendar import DateEntry
import pandas as pd

from outlook_renovaciones_engine import OutlookRenovacionesEngine


CONFIG_FILE = os.path.join(
    os.path.expanduser("~"),
    ".fonacot_validador_renovaciones_config.json"
)

ITOP_PORTAL_URL = (
    "https://mservicio.fonacot.gob.mx/pages/UI.php"
    "?operation=details&class=UserRequest&id={num_id}&c[menu]=New UserRequest"
)


class FonacotValidadorRenovacionesApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("FONACOT - Validador y Desglose Diario de Cancelación de Renovaciones v1.0")
        self.geometry("1460x900")
        self.minsize(1240, 760)

        ctk.set_appearance_mode("Light")
        ctk.set_default_color_theme("blue")

        self.engine = OutlookRenovacionesEngine()
        self.config_data = self._load_config()

        # Estado de datos
        self.all_rows = []
        self.discarded_row_ids = set()
        self.validated_row_ids = set()
        self.is_extracting = False

        self._build_ui()
        self.after(250, self._init_accounts)

    def _load_config(self) -> dict:
        default_cfg = {
            "account": "",
            "modo_operacion": "CONCENTRADO",  # "CONCENTRADO" o "AGENTE"
            "agente_filter": "2609",
            "auto_exclude_reassigned": False
        }
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    default_cfg.update(data)
            except Exception:
                pass
        return default_cfg

    def _save_config(self):
        try:
            cfg = {
                "account": self.cmb_account.get().strip(),
                "modo_operacion": self.var_modo.get(),
                "agente_filter": self.ent_agente.get().strip(),
                "auto_exclude_reassigned": bool(self.var_auto_exclude_reassigned.get())
            }
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(cfg, f, indent=2, ensure_ascii=False)
        except Exception:
            pass

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        # ================= 1. ENCABEZADO SUPERIOR =================
        header = ctk.CTkFrame(self, fg_color="#0F2942", corner_radius=0, height=62)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(1, weight=1)

        title_box = ctk.CTkFrame(header, fg_color="transparent")
        title_box.grid(row=0, column=0, padx=18, pady=8, sticky="w")

        ctk.CTkLabel(
            title_box,
            text="🛡️ FONACOT | Validador Diario de Cancelación de Renovaciones",
            font=ctk.CTkFont(size=19, weight="bold"),
            text_color="#FFFFFF"
        ).pack(anchor="w")

        ctk.CTkLabel(
            title_box,
            text="Desglose inteligente 1 a N (Múltiples Clientes, Múltiples Créditos Renovados, Tablas CREDERE y Crédito Padre/Hijo)",
            font=ctk.CTkFont(size=12),
            text_color="#A3C2E0"
        ).pack(anchor="w")

        # Selector de cuenta Outlook a la derecha del header
        acc_box = ctk.CTkFrame(header, fg_color="transparent")
        acc_box.grid(row=0, column=2, padx=18, pady=8, sticky="e")

        ctk.CTkLabel(
            acc_box,
            text="Buzón Outlook:",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#E2E8F0"
        ).pack(side="left", padx=(0, 6))

        self.cmb_account = ctk.CTkComboBox(
            acc_box,
            values=["Cargando cuentas..."],
            width=260,
            height=30
        )
        self.cmb_account.pack(side="left")

        # ================= 2. PANEL DE CONTROLES Y MODO DE OPERACIÓN =================
        ctrl_card = ctk.CTkFrame(self, fg_color="#F8FAFC", corner_radius=10, border_width=1, border_color="#CBD5E1")
        ctrl_card.grid(row=1, column=0, padx=14, pady=(10, 6), sticky="ew")
        ctrl_card.grid_columnconfigure(3, weight=1)

        # Columna 0: Selector de Modo (Concentrado General vs Por Compañero)
        mode_frame = ctk.CTkFrame(ctrl_card, fg_color="#EFF6FF", corner_radius=8, border_width=1, border_color="#BFDBFE")
        mode_frame.grid(row=0, column=0, padx=10, pady=8, sticky="nsw")

        ctk.CTkLabel(
            mode_frame,
            text="1. Modo de Validación:",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#1E3A8A"
        ).grid(row=0, column=0, columnspan=2, padx=10, pady=(6, 2), sticky="w")

        self.var_modo = tk.StringVar(value=self.config_data.get("modo_operacion", "CONCENTRADO"))

        self.rb_concentrado = ctk.CTkRadioButton(
            mode_frame,
            text="🌐 Concentrado General del Día (Todos)",
            variable=self.var_modo,
            value="CONCENTRADO",
            command=self._on_mode_change,
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#0F172A"
        )
        self.rb_concentrado.grid(row=1, column=0, padx=10, pady=4, sticky="w")

        self.rb_agente = ctk.CTkRadioButton(
            mode_frame,
            text="👤 Por Compañero (Gafete o Nombre):",
            variable=self.var_modo,
            value="AGENTE",
            command=self._on_mode_change,
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#0F172A"
        )
        self.rb_agente.grid(row=2, column=0, padx=10, pady=(2, 6), sticky="w")

        self.ent_agente = ctk.CTkEntry(
            mode_frame,
            width=165,
            height=28,
            placeholder_text="Ej: 2609, 7353, Nancy..."
        )
        self.ent_agente.grid(row=2, column=1, padx=(0, 10), pady=(2, 6), sticky="w")
        self.ent_agente.insert(0, self.config_data.get("agente_filter", "2609"))

        # Columna 1: Rango de Fechas
        date_frame = ctk.CTkFrame(ctrl_card, fg_color="#FFFFFF", corner_radius=8, border_width=1, border_color="#E2E8F0")
        date_frame.grid(row=0, column=1, padx=6, pady=8, sticky="nsw")

        ctk.CTkLabel(
            date_frame,
            text="2. Fecha de Recepción:",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#1E293B"
        ).grid(row=0, column=0, columnspan=4, padx=10, pady=(6, 2), sticky="w")

        ctk.CTkLabel(date_frame, text="Del:", font=ctk.CTkFont(size=12)).grid(row=1, column=0, padx=(10, 4), pady=4)
        self.cal_start = DateEntry(
            date_frame, width=11, background="#1E40AF", foreground="white",
            borderwidth=1, date_pattern="dd/mm/yyyy", font=("Segoe UI", 10)
        )
        self.cal_start.grid(row=1, column=1, padx=4, pady=4)

        ctk.CTkLabel(date_frame, text="Al:", font=ctk.CTkFont(size=12)).grid(row=1, column=2, padx=(8, 4), pady=4)
        self.cal_end = DateEntry(
            date_frame, width=11, background="#1E40AF", foreground="white",
            borderwidth=1, date_pattern="dd/mm/yyyy", font=("Segoe UI", 10)
        )
        self.cal_end.grid(row=1, column=3, padx=(4, 10), pady=4)

        btn_quick_dates = ctk.CTkFrame(date_frame, fg_color="transparent")
        btn_quick_dates.grid(row=2, column=0, columnspan=4, padx=8, pady=(0, 6), sticky="ew")

        ctk.CTkButton(
            btn_quick_dates, text="📅 Hoy", width=68, height=24,
            fg_color="#E2E8F0", text_color="#0F172A", hover_color="#CBD5E1",
            command=self._set_today
        ).pack(side="left", padx=2)

        ctk.CTkButton(
            btn_quick_dates, text="📆 30 Sept (Caso Ejemplo)", width=155, height=24,
            fg_color="#FEF3C7", text_color="#92400E", hover_color="#FDE68A",
            command=lambda: self._set_specific_date(2026, 9, 30)
        ).pack(side="left", padx=2)

        # Columna 2: Opciones de Reasignación y Botón Extraer
        action_frame = ctk.CTkFrame(ctrl_card, fg_color="transparent")
        action_frame.grid(row=0, column=2, columnspan=2, padx=10, pady=8, sticky="nsew")
        action_frame.grid_columnconfigure(0, weight=1)

        self.var_auto_exclude_reassigned = tk.BooleanVar(
            value=self.config_data.get("auto_exclude_reassigned", False)
        )
        self.chk_exclude_reassigned = ctk.CTkCheckBox(
            action_frame,
            text="Excluir automáticamente tickets reasignados a otro compañero (modo individual)",
            variable=self.var_auto_exclude_reassigned,
            command=self._refresh_views,
            font=ctk.CTkFont(size=12),
            checkbox_width=18,
            checkbox_height=18
        )
        self.chk_exclude_reassigned.grid(row=0, column=0, padx=4, pady=(4, 6), sticky="w")

        btn_row = ctk.CTkFrame(action_frame, fg_color="transparent")
        btn_row.grid(row=1, column=0, sticky="ew")

        self.btn_extract = ctk.CTkButton(
            btn_row,
            text="🔍 EXTRAER Y VALIDAR RENOVACIONES",
            font=ctk.CTkFont(size=14, weight="bold"),
            height=40,
             width=290,
            fg_color="#1D4ED8",
            hover_color="#1E40AF",
            command=self._start_extraction
        )
        self.btn_extract.pack(side="left", padx=(0, 8))

        self.btn_export = ctk.CTkButton(
            btn_row,
            text="📊 Exportar a Excel",
            font=ctk.CTkFont(size=13, weight="bold"),
            height=40,
            width=160,
            fg_color="#15803D",
            hover_color="#166534",
            command=self._export_excel
        )
        self.btn_export.pack(side="left", padx=4)

        self.btn_copy_batch = ctk.CTkButton(
            btn_row,
            text="📋 Copiar Lote para Correo",
            font=ctk.CTkFont(size=13, weight="bold"),
            height=40,
            width=185,
            fg_color="#0F766E",
            hover_color="#115E59",
            command=self._copy_batch_to_clipboard
        )
        self.btn_copy_batch.pack(side="left", padx=4)

        # ================= 3. TARJETAS KPI Y FILTROS RÁPIDOS =================
        kpi_bar = ctk.CTkFrame(self, fg_color="transparent")
        kpi_bar.grid(row=2, column=0, padx=14, pady=(2, 6), sticky="ew")
        for i in range(5):
            kpi_bar.grid_columnconfigure(i, weight=1)

        self.lbl_kpi_tickets = self._create_kpi_card(
            kpi_bar, 0, "🎫 Tickets de Renovación", "0", "Incidentes únicos en iTop", "#1E40AF", "#EFF6FF"
        )
        self.lbl_kpi_renovaciones = self._create_kpi_card(
            kpi_bar, 1, "🔢 Total Créditos a Cancelar", "0", "Renovaciones desglosadas (1 a N)", "#047857", "#ECFDF5"
        )
        self.lbl_kpi_multiples = self._create_kpi_card(
            kpi_bar, 2, "🚨 Tickets con 2+ Renovaciones", "0", "Contienen múltiples créditos/clientes", "#B45309", "#FFFBEB"
        )
        self.lbl_kpi_alertas = self._create_kpi_card(
            kpi_bar, 3, "⚠️ Alertas / Revisar Datos", "0", "Falta dato, discrepancia o reasignado", "#B91C1C", "#FEF2F2"
        )
        self.lbl_kpi_validados = self._create_kpi_card(
            kpi_bar, 4, "☑️ Avance de Validación", "0 / 0", "Doble clic en fila para marcar ☑️", "#6D28D9", "#F5F3FF"
        )

        # ================= 4. ÁREA PRINCIPAL (TABS + TABLAS + DETALLE) =================
        main_paned = ttk.PanedWindow(self, orient=tk.VERTICAL)
        main_paned.grid(row=3, column=0, padx=14, pady=(2, 6), sticky="nsew")

        top_work_frame = ctk.CTkFrame(main_paned, fg_color="#FFFFFF", corner_radius=8)
        bottom_detail_frame = ctk.CTkFrame(main_paned, fg_color="#F8FAFC", corner_radius=8, border_width=1, border_color="#CBD5E1")

        main_paned.add(top_work_frame, weight=4)
        main_paned.add(bottom_detail_frame, weight=2)

        top_work_frame.grid_columnconfigure(0, weight=1)
        top_work_frame.grid_rowconfigure(1, weight=1)

        # Barra de herramientas de filtrado y acciones sobre filas
        filter_bar = ctk.CTkFrame(top_work_frame, fg_color="#F1F5F9", corner_radius=6, height=42)
        filter_bar.grid(row=0, column=0, padx=8, pady=(8, 4), sticky="ew")

        ctk.CTkLabel(filter_bar, text="🔎 Buscar:", font=ctk.CTkFont(size=12, weight="bold")).pack(side="left", padx=(10, 4), pady=6)
        self.var_search = tk.StringVar()
        self.var_search.trace_add("write", lambda *_: self._refresh_views())
        self.ent_search = ctk.CTkEntry(
            filter_bar,
            textvariable=self.var_search,
            width=230,
            height=28,
            placeholder_text="Ticket, Cliente, Crédito, Agente..."
        )
        self.ent_search.pack(side="left", padx=4, pady=6)

        ctk.CTkLabel(filter_bar, text="Filtro Rápido:", font=ctk.CTkFont(size=12, weight="bold")).pack(side="left", padx=(12, 4), pady=6)
        self.cmb_quick_filter = ctk.CTkComboBox(
            filter_bar,
            values=[
                "📋 Todas las Renovaciones Activas",
                "🚨 Solo Múltiples (2+ renovaciones en el mismo ticket)",
                "⚠️ Solo con Alertas (Falta Dato / Discrepancia / Duplicado / Reasignado)",
                "⏳ Pendientes de Validar (☐)",
                "☑️ Ya Validadas (☑️)",
                "🚫 Papelera (Descartadas)"
            ],
            width=360,
            height=28,
            command=lambda _: self._refresh_views()
        )
        self.cmb_quick_filter.set("📋 Todas las Renovaciones Activas")
        self.cmb_quick_filter.pack(side="left", padx=4, pady=6)

        # Botones interactivos a la derecha de la barra
        ctk.CTkButton(
            filter_bar, text="🌐 Abrir en iTop", width=115, height=28,
            fg_color="#2563EB", hover_color="#1D4ED8",
            command=self._open_selected_in_itop
        ).pack(side="right", padx=(4, 8), pady=6)

        ctk.CTkButton(
            filter_bar, text="↩️ Restaurar", width=95, height=28,
            fg_color="#64748B", hover_color="#475569",
            command=self._restore_selected_rows
        ).pack(side="right", padx=4, pady=6)

        ctk.CTkButton(
            filter_bar, text="🚫 Descartar Fila", width=120, height=28,
            fg_color="#DC2626", hover_color="#B91C1C",
            command=self._discard_selected_rows
        ).pack(side="right", padx=4, pady=6)

        ctk.CTkButton(
            filter_bar, text="☑️ Validar Todo", width=115, height=28,
            fg_color="#7C3AED", hover_color="#6D28D9",
            command=self._validate_all_visible
        ).pack(side="right", padx=4, pady=6)

        ctk.CTkButton(
            filter_bar, text="☑️ Marcar/Desmarcar", width=140, height=28,
            fg_color="#4F46E5", hover_color="#4338CA",
            command=self._toggle_selected_validation
        ).pack(side="right", padx=4, pady=6)

        # Notebook con las 3 vistas
        self.tabs = ctk.CTkTabview(top_work_frame, height=380)
        self.tabs.grid(row=1, column=0, padx=8, pady=(0, 8), sticky="nsew")

        self.tab_desglose = self.tabs.add("🧬 1. Desglose de Renovaciones (1 Fila = 1 Crédito a Cancelar)")
        self.tab_tickets = self.tabs.add("🎫 2. Resumen por Ticket (Incidentes iTop)")
        self.tab_agentes = self.tabs.add("👥 3. Resumen por Compañero (Agentes)")

        self._setup_treeview_styles()
        self._build_tab_desglose()
        self._build_tab_tickets()
        self._build_tab_agentes()

        # Panel inferior de inspección de correo y desglose
        self._build_detail_panel(bottom_detail_frame)

        # ================= 5. BARRA DE ESTADO INFERIOR =================
        status_bar = ctk.CTkFrame(self, fg_color="#0F172A", corner_radius=0, height=30)
        status_bar.grid(row=4, column=0, sticky="ew")
        status_bar.grid_columnconfigure(0, weight=1)

        self.lbl_status = ctk.CTkLabel(
            status_bar,
            text="Listo. Selecciona la fecha y haz clic en 'EXTRAER Y VALIDAR RENOVACIONES'.",
            font=ctk.CTkFont(size=12),
            text_color="#E2E8F0"
        )
        self.lbl_status.grid(row=0, column=0, padx=14, pady=4, sticky="w")

        self.progress_bar = ctk.CTkProgressBar(status_bar, width=220, height=12)
        self.progress_bar.grid(row=0, column=1, padx=14, pady=4, sticky="e")
        self.progress_bar.set(0)

        self._on_mode_change()

    def _create_kpi_card(self, parent, col, title, initial_val, subtitle, text_color, bg_color):
        card = ctk.CTkFrame(parent, fg_color=bg_color, corner_radius=8, border_width=1, border_color="#CBD5E1")
        card.grid(row=0, column=col, padx=5, pady=2, sticky="ew")

        ctk.CTkLabel(
            card, text=title, font=ctk.CTkFont(size=12, weight="bold"), text_color="#334155"
        ).pack(anchor="w", padx=12, pady=(6, 0))

        lbl_val = ctk.CTkLabel(
            card, text=initial_val, font=ctk.CTkFont(size=24, weight="bold"), text_color=text_color
        )
        lbl_val.pack(anchor="w", padx=12, pady=0)

        ctk.CTkLabel(
            card, text=subtitle, font=ctk.CTkFont(size=11), text_color="#64748B"
        ).pack(anchor="w", padx=12, pady=(0, 6))

        return lbl_val

    def _setup_treeview_styles(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "Renovaciones.Treeview",
            background="#FFFFFF",
            foreground="#0F172A",
            rowheight=28,
            fieldbackground="#FFFFFF",
            font=("Segoe UI", 10)
        )
        style.configure(
            "Renovaciones.Treeview.Heading",
            background="#1E293B",
            foreground="#FFFFFF",
            font=("Segoe UI", 10, "bold"),
            relief="flat"
        )
        style.map(
            "Renovaciones.Treeview",
            background=[("selected", "#2563EB")],
            foreground=[("selected", "#FFFFFF")]
        )

    def _build_tab_desglose(self):
        self.tab_desglose.grid_columnconfigure(0, weight=1)
        self.tab_desglose.grid_rowconfigure(0, weight=1)

        cols = (
            "validado", "ticket", "parte", "cliente", "credito_hijo",
            "credito_padre", "gafete", "agente", "alerta_reno",
            "alerta_asig", "motivo", "fecha"
        )
        self.tree_desglose = ttk.Treeview(
            self.tab_desglose, columns=cols, show="headings", style="Renovaciones.Treeview", selectmode="extended"
        )

        headers = {
            "validado": ("☑️ Val.", 65, "center"),
            "ticket": ("Ticket iTop", 95, "center"),
            "parte": ("Desglose", 75, "center"),
            "cliente": ("No. Cliente (Cuenta)", 150, "center"),
            "credito_hijo": ("Crédito Renovado (A Cancelar)", 185, "center"),
            "credito_padre": ("Crédito Origen (Padre)", 145, "center"),
            "gafete": ("Gafete", 65, "center"),
            "agente": ("Agente Responsable", 190, "w"),
            "alerta_reno": ("Estado / Alerta de Renovación", 285, "w"),
            "alerta_asig": ("Estado Asignación", 220, "w"),
            "motivo": ("Motivo / Observación", 220, "w"),
            "fecha": ("Fecha / Hora", 120, "center")
        }

        for col, (text, width, anchor) in headers.items():
            self.tree_desglose.heading(col, text=text, command=lambda c=col: self._sort_tree(self.tree_desglose, c, False))
            self.tree_desglose.column(col, width=width, anchor=anchor)

        # Tags de color para distinguir inmediatamente los casos críticos
        self.tree_desglose.tag_configure("multiple", background="#FEF3C7", foreground="#78350F")       # Ámbar: Múltiples renovaciones en 1 ticket
        self.tree_desglose.tag_configure("incompleto", background="#FEE2E2", foreground="#991B1B")     # Rojo claro: Falta cliente o crédito
        self.tree_desglose.tag_configure("duplicado", background="#FCE7F3", foreground="#9D174D")      # Rosa: Crédito repetido en otro ticket
        self.tree_desglose.tag_configure("reasignado", background="#FFEDD5", foreground="#9A3412")     # Naranja: Reasignado a otro agente
        self.tree_desglose.tag_configure("validado", background="#DCFCE7", foreground="#166534")       # Verde: Ya validado (☑️)
        self.tree_desglose.tag_configure("descartado", background="#E2E8F0", foreground="#64748B")     # Gris: Descartado
        self.tree_desglose.tag_configure("normal", background="#FFFFFF", foreground="#0F172A")

        vsb = ttk.Scrollbar(self.tab_desglose, orient="vertical", command=self.tree_desglose.yview)
        hsb = ttk.Scrollbar(self.tab_desglose, orient="horizontal", command=self.tree_desglose.xview)
        self.tree_desglose.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree_desglose.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        self.tree_desglose.bind("<<TreeviewSelect>>", self._on_select_desglose_row)
        self.tree_desglose.bind("<Double-1>", lambda e: self._toggle_selected_validation())
        self.tree_desglose.bind("<space>", lambda e: self._toggle_selected_validation())

    def _build_tab_tickets(self):
        self.tab_tickets.grid_columnconfigure(0, weight=1)
        self.tab_tickets.grid_rowconfigure(0, weight=1)

        cols = (
            "ticket", "cant_renovaciones", "clientes", "creditos_hijos",
            "creditos_padres", "gafete", "agente", "resumen", "alerta_asig", "fecha"
        )
        self.tree_tickets = ttk.Treeview(
            self.tab_tickets, columns=cols, show="headings", style="Renovaciones.Treeview"
        )

        headers = {
            "ticket": ("Ticket iTop", 100, "center"),
            "cant_renovaciones": ("Cant. Renovaciones", 135, "center"),
            "clientes": ("Cliente(s) en Ticket", 180, "w"),
            "creditos_hijos": ("Crédito(s) Renovado(s) a Cancelar", 240, "w"),
            "creditos_padres": ("Crédito(s) Padre", 150, "w"),
            "gafete": ("Gafete", 70, "center"),
            "agente": ("Agente Responsable", 200, "w"),
            "resumen": ("Resumen del Ticket", 270, "w"),
            "alerta_asig": ("Historial / Asignación", 230, "w"),
            "fecha": ("Fecha", 120, "center")
        }
        for col, (text, width, anchor) in headers.items():
            self.tree_tickets.heading(col, text=text, command=lambda c=col: self._sort_tree(self.tree_tickets, c, False))
            self.tree_tickets.column(col, width=width, anchor=anchor)

        self.tree_tickets.tag_configure("multiple", background="#FEF3C7", foreground="#78350F")
        self.tree_tickets.tag_configure("incompleto", background="#FEE2E2", foreground="#991B1B")
        self.tree_tickets.tag_configure("normal", background="#FFFFFF", foreground="#0F172A")

        vsb = ttk.Scrollbar(self.tab_tickets, orient="vertical", command=self.tree_tickets.yview)
        hsb = ttk.Scrollbar(self.tab_tickets, orient="horizontal", command=self.tree_tickets.xview)
        self.tree_tickets.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree_tickets.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        self.tree_tickets.bind("<<TreeviewSelect>>", self._on_select_ticket_row)

    def _build_tab_agentes(self):
        self.tab_agentes.grid_columnconfigure(0, weight=1)
        self.tab_agentes.grid_rowconfigure(0, weight=1)

        cols = ("gafete", "agente", "tickets_unicos", "total_renovaciones", "tickets_multiples", "con_alerta", "validados")
        self.tree_agentes = ttk.Treeview(
            self.tab_agentes, columns=cols, show="headings", style="Renovaciones.Treeview"
        )

        headers = {
            "gafete": ("Gafete", 90, "center"),
            "agente": ("Nombre del Compañero (Agente)", 280, "w"),
            "tickets_unicos": ("🎫 Tickets Asignados", 160, "center"),
            "total_renovaciones": ("🔢 Total Renovaciones (Créditos)", 210, "center"),
            "tickets_multiples": ("🚨 Tickets Múltiples (2+)", 180, "center"),
            "con_alerta": ("⚠️ Filas con Alerta / Faltante", 190, "center"),
            "validados": ("☑️ Renovaciones Validadas", 180, "center")
        }
        for col, (text, width, anchor) in headers.items():
            self.tree_agentes.heading(col, text=text, command=lambda c=col: self._sort_tree(self.tree_agentes, c, False))
            self.tree_agentes.column(col, width=width, anchor=anchor)

        vsb = ttk.Scrollbar(self.tab_agentes, orient="vertical", command=self.tree_agentes.yview)
        self.tree_agentes.configure(yscrollcommand=vsb.set)
        self.tree_agentes.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

    def _build_detail_panel(self, parent):
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(1, weight=1)

        hdr = ctk.CTkFrame(parent, fg_color="#E2E8F0", corner_radius=6, height=32)
        hdr.grid(row=0, column=0, padx=8, pady=(6, 4), sticky="ew")

        self.lbl_detail_header = ctk.CTkLabel(
            hdr,
            text="🔎 Inspección de Renovación: Selecciona una fila arriba para ver el correo completo y todos sus números extraídos.",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#0F172A"
        )
        self.lbl_detail_header.pack(side="left", padx=10, pady=4)

        self.txt_detail = ctk.CTkTextbox(
            parent,
            font=ctk.CTkFont(family="Consolas", size=12),
            fg_color="#FFFFFF",
            text_color="#0F172A",
            wrap="word"
        )
        self.txt_detail.grid(row=1, column=0, padx=8, pady=(0, 8), sticky="nsew")

    # ================= LÓGICA DE INTERFAZ Y EVENTOS =================

    def _init_accounts(self):
        accounts = self.engine.get_accounts()
        if accounts:
            self.cmb_account.configure(values=accounts)
            saved = self.config_data.get("account", "")
            if saved in accounts:
                self.cmb_account.set(saved)
            else:
                self.cmb_account.set(accounts[0])
        else:
            self.cmb_account.configure(values=["Sin conexión a Outlook"])
            self.cmb_account.set("Sin conexión a Outlook")

    def _on_mode_change(self):
        modo = self.var_modo.get()
        if modo == "CONCENTRADO":
            self.ent_agente.configure(state="disabled", fg_color="#E2E8F0")
            self.chk_exclude_reassigned.configure(state="disabled")
        else:
            self.ent_agente.configure(state="normal", fg_color="#FFFFFF")
            self.chk_exclude_reassigned.configure(state="normal")

    def _set_today(self):
        now = datetime.now()
        self.cal_start.set_date(now)
        self.cal_end.set_date(now)

    def _set_specific_date(self, year: int, month: int, day: int):
        dt = datetime(year, month, day)
        self.cal_start.set_date(dt)
        self.cal_end.set_date(dt)

    def _start_extraction(self):
        if self.is_extracting:
            return

        modo = self.var_modo.get()
        agente_filter = self.ent_agente.get().strip()
        if modo == "AGENTE" and not agente_filter:
            messagebox.showwarning(
                "Falta Gafete o Nombre",
                "En el modo 'Por Compañero', ingresa el número de Gafete (ej. 2609, 7353) o el Nombre del agente."
            )
            return

        self._save_config()

        account = self.cmb_account.get().strip()
        if account == "Sin conexión a Outlook" or account == "Cargando cuentas...":
            account = ""

        d_start = datetime.combine(self.cal_start.get_date(), datetime.min.time())
        d_end = datetime.combine(self.cal_end.get_date(), datetime.max.time())

        if d_start > d_end:
            d_start, d_end = d_end, d_start

        self.is_extracting = True
        self.btn_extract.configure(state="disabled", text="⏳ Analizando Renovaciones...")
        self.progress_bar.set(0.1)

        threading.Thread(
            target=self._run_extraction_thread,
            args=(account, d_start, d_end, modo, agente_filter),
            daemon=True
        ).start()

    def _run_extraction_thread(self, account, d_start, d_end, modo, agente_filter):
        def progress_cb(msg, frac=None):
            self.after(0, lambda: self._update_progress(msg, frac))

        try:
            rows = self.engine.extract_renewals(
                account_name=account,
                start_dt=d_start,
                end_dt=d_end,
                mode=modo,
                agent_filter=agente_filter,
                progress_callback=progress_cb
            )
            self.after(0, lambda: self._on_extraction_done(rows, None))
        except Exception as e:
            self.after(0, lambda: self._on_extraction_done([], str(e)))

    def _update_progress(self, msg, frac):
        self.lbl_status.configure(text=msg)
        if frac is not None:
            self.progress_bar.set(max(0.05, min(1.0, float(frac))))

    def _on_extraction_done(self, rows, error_msg):
        self.is_extracting = False
        self.btn_extract.configure(state="normal", text="🔍 EXTRAER Y VALIDAR RENOVACIONES")
        self.progress_bar.set(1.0)

        if error_msg:
            messagebox.showerror("Error al extraer de Outlook", error_msg)
            self.lbl_status.configure(text=f"Error: {error_msg}")
            return

        self.all_rows = rows
        self.discarded_row_ids.clear()
        self.validated_row_ids.clear()
        self._refresh_views()

        # Mensaje claro si hay tickets con múltiples renovaciones
        multi_tickets = {r['Numero_Ticket'] for r in rows if r.get('Es_Multiple')}
        missing_rows = [r for r in rows if r.get('Datos_Incompletos')]
        if multi_tickets:
            self.lbl_status.configure(
                text=f"✅ Extracción completada: {len(rows)} renovaciones desglosadas. ¡ATENCIÓN! Se detectaron {len(multi_tickets)} ticket(s) con MÚLTIPLES renovaciones adentro."
            )

    def _get_active_and_filtered_rows(self):
        auto_excl = bool(self.var_auto_exclude_reassigned.get()) and (self.var_modo.get() == "AGENTE")
        q_filter = self.cmb_quick_filter.get()
        search_txt = self.var_search.get().strip().lower()

        active_rows = []
        for r in self.all_rows:
            rid = r['Row_ID']
            is_disc = rid in self.discarded_row_ids
            if auto_excl and r.get('Reasignado_A_Otro'):
                is_disc = True

            if not is_disc:
                active_rows.append(r)

        # Filtrar para la tabla visible
        visible_rows = []
        for r in self.all_rows:
            rid = r['Row_ID']
            is_disc = (rid in self.discarded_row_ids) or (auto_excl and r.get('Reasignado_A_Otro'))
            is_val = rid in self.validated_row_ids

            if "Papelera" in q_filter:
                if not is_disc:
                    continue
            else:
                if is_disc:
                    continue
                if "Solo Múltiples" in q_filter and not r.get('Es_Multiple'):
                    continue
                if "Solo con Alertas" in q_filter:
                    has_any_alert = (
                        r.get('Datos_Incompletos')
                        or r.get('Nivel_Alerta') in ('incompleto', 'duplicado')
                        or r.get('Requiere_Validacion_Agente')
                        or "DISCREPANCIA" in r.get('Alerta_Renovacion', '')
                    )
                    if not has_any_alert:
                        continue
                if "Pendientes de Validar" in q_filter and is_val:
                    continue
                if "Ya Validadas" in q_filter and not is_val:
                    continue

            if search_txt:
                haystack = (
                    f"{r.get('Numero_Ticket', '')} {r.get('Numero_Cliente', '')} "
                    f"{r.get('Credito_Renovado_Hijo', '')} {r.get('Credito_Origen_Padre', '')} "
                    f"{r.get('Gafete', '')} {r.get('Nombre_Asignado', '')} "
                    f"{r.get('Motivo', '')} {r.get('Resumen_iTop', '')}"
                ).lower()
                if search_txt not in haystack:
                    continue

            visible_rows.append(r)

        return active_rows, visible_rows

    def _refresh_views(self):
        active_rows, visible_rows = self._get_active_and_filtered_rows()

        # 1. Actualizar KPIs sobre filas activas
        unique_tickets = {r['Numero_Ticket'] for r in active_rows}
        multi_tickets = {r['Numero_Ticket'] for r in active_rows if r.get('Es_Multiple')}
        alert_rows = [
            r for r in active_rows
            if r.get('Datos_Incompletos')
            or r.get('Nivel_Alerta') in ('incompleto', 'duplicado')
            or r.get('Requiere_Validacion_Agente')
            or "DISCREPANCIA" in r.get('Alerta_Renovacion', '')
        ]
        val_count = sum(1 for r in active_rows if r['Row_ID'] in self.validated_row_ids)

        self.lbl_kpi_tickets.configure(text=str(len(unique_tickets)))
        self.lbl_kpi_renovaciones.configure(text=str(len(active_rows)))
        self.lbl_kpi_multiples.configure(text=str(len(multi_tickets)))
        self.lbl_kpi_alertas.configure(text=str(len(alert_rows)))
        self.lbl_kpi_validados.configure(text=f"{val_count} / {len(active_rows)}")

        # 2. Poblar Tab 1: Desglose de Renovaciones
        self.tree_desglose.delete(*self.tree_desglose.get_children())
        for r in visible_rows:
            rid = r['Row_ID']
            is_val = rid in self.validated_row_ids
            is_disc = rid in self.discarded_row_ids

            val_icon = "☑️ Sí" if is_val else "☐ No"
            parte_str = f"{r['Indice_En_Ticket']} de {r['Total_En_Ticket']}"

            if is_disc:
                tag = "descartado"
            elif is_val:
                tag = "validado"
            elif r.get('Datos_Incompletos'):
                tag = "incompleto"
            elif r.get('Nivel_Alerta') == "duplicado":
                tag = "duplicado"
            elif r.get('Requiere_Validacion_Agente'):
                tag = "reasignado"
            elif r.get('Es_Multiple') or r.get('Nivel_Alerta') == "multiple":
                tag = "multiple"
            else:
                tag = "normal"

            self.tree_desglose.insert(
                "",
                "end",
                iid=rid,
                values=(
                    val_icon,
                    r['Numero_Ticket'],
                    parte_str,
                    r['Numero_Cliente'],
                    r['Credito_Renovado_Hijo'],
                    r['Credito_Origen_Padre'],
                    r['Gafete'],
                    r['Nombre_Asignado'],
                    r['Alerta_Renovacion'],
                    r['Alerta_Asignacion'],
                    r['Motivo'],
                    r['Fecha']
                ),
                tags=(tag,)
            )

        # 3. Poblar Tab 2: Resumen por Ticket
        self.tree_tickets.delete(*self.tree_tickets.get_children())
        tickets_grouped = {}
        for r in active_rows:
            tnum = r['Numero_Ticket']
            if tnum not in tickets_grouped:
                tickets_grouped[tnum] = {
                    'ticket': tnum,
                    'clientes': [],
                    'creditos_hijos': [],
                    'creditos_padres': [],
                    'gafete': r['Gafete'],
                    'agente': r['Nombre_Asignado'],
                    'resumen': r['Resumen_iTop'],
                    'alerta_asig': r['Alerta_Asignacion'],
                    'fecha': r['Fecha'],
                    'has_missing': False
                }
            g = tickets_grouped[tnum]
            if r['Numero_Cliente'] not in g['clientes']:
                g['clientes'].append(r['Numero_Cliente'])
            g['creditos_hijos'].append(r['Credito_Renovado_Hijo'])
            if r['Credito_Origen_Padre'] and r['Credito_Origen_Padre'] != "-" and r['Credito_Origen_Padre'] not in g['creditos_padres']:
                g['creditos_padres'].append(r['Credito_Origen_Padre'])
            if r.get('Datos_Incompletos'):
                g['has_missing'] = True

        for tnum, g in tickets_grouped.items():
            cant = len(g['creditos_hijos'])
            tag = "incompleto" if g['has_missing'] else ("multiple" if cant > 1 else "normal")
            cant_label = f"🚨 {cant} Renovaciones" if cant > 1 else "1 Renovación"
            self.tree_tickets.insert(
                "",
                "end",
                iid=f"TKT_{tnum}",
                values=(
                    tnum,
                    cant_label,
                    ", ".join(g['clientes']),
                    ", ".join(g['creditos_hijos']),
                    ", ".join(g['creditos_padres']) if g['creditos_padres'] else "-",
                    g['gafete'],
                    g['agente'],
                    g['resumen'],
                    g['alerta_asig'],
                    g['fecha']
                ),
                tags=(tag,)
            )

        # 4. Poblar Tab 3: Resumen por Compañero (Agente)
        self.tree_agentes.delete(*self.tree_agentes.get_children())
        agentes_grouped = {}
        for r in active_rows:
            key = (r['Gafete'], r['Nombre_Asignado'])
            if key not in agentes_grouped:
                agentes_grouped[key] = {
                    'tickets': set(),
                    'multi_tickets': set(),
                    'renovaciones': 0,
                    'alertas': 0,
                    'validados': 0
                }
            ag = agentes_grouped[key]
            ag['tickets'].add(r['Numero_Ticket'])
            if r.get('Es_Multiple'):
                ag['multi_tickets'].add(r['Numero_Ticket'])
            ag['renovaciones'] += 1
            if r.get('Datos_Incompletos') or r.get('Requiere_Validacion_Agente') or r.get('Nivel_Alerta') == 'duplicado':
                ag['alertas'] += 1
            if r['Row_ID'] in self.validated_row_ids:
                ag['validados'] += 1

        for (gaf, nom), ag in sorted(agentes_grouped.items(), key=lambda x: x[1]['renovaciones'], reverse=True):
            self.tree_agentes.insert(
                "",
                "end",
                values=(
                    gaf,
                    nom,
                    len(ag['tickets']),
                    ag['renovaciones'],
                    len(ag['multi_tickets']),
                    ag['alertas'],
                    f"{ag['validados']} / {ag['renovaciones']}"
                )
            )

    def _on_select_desglose_row(self, _event=None):
        sel = self.tree_desglose.selection()
        if not sel:
            return
        rid = sel[0]
        row_obj = next((r for r in self.all_rows if r['Row_ID'] == rid), None)
        if row_obj:
            self._show_row_detail(row_obj)

    def _on_select_ticket_row(self, _event=None):
        sel = self.tree_tickets.selection()
        if not sel:
            return
        tnum = sel[0].replace("TKT_", "")
        row_obj = next((r for r in self.all_rows if r['Numero_Ticket'] == tnum), None)
        if row_obj:
            self._show_row_detail(row_obj)

    def _show_row_detail(self, r: dict):
        tnum = r['Numero_Ticket']
        siblings = [x for x in self.all_rows if x['Numero_Ticket'] == tnum]

        self.lbl_detail_header.configure(
            text=(
                f"🎫 Ticket {tnum}  |  👤 Cliente: {r['Numero_Cliente']}  |  "
                f"💳 Crédito Renovado (A Cancelar): {r['Credito_Renovado_Hijo']}  |  "
                f"🔗 Crédito Origen (Padre): {r['Credito_Origen_Padre']}  |  "
                f"Agente: {r['Gafete']} {r['Nombre_Asignado']}"
            )
        )

        desglose_lines = []
        for s in siblings:
            chk = "☑️" if s['Row_ID'] in self.validated_row_ids else "☐"
            desglose_lines.append(
                f"   {chk} [{s['Indice_En_Ticket']}/{s['Total_En_Ticket']}] "
                f"No. Cliente: {s['Numero_Cliente']}   |   "
                f"Crédito Renovado (Hijo): {s['Credito_Renovado_Hijo']}   |   "
                f"Crédito Origen (Padre): {s['Credito_Origen_Padre']}   "
                f"({s['Alerta_Renovacion']})"
            )

        detail_text = (
            f"========================================================================================\n"
            f"RESUMEN DE EXTRACCIÓN DEL TICKET {tnum} ({len(siblings)} renovación(es) detectada(s) en este correo)\n"
            f"========================================================================================\n"
            + "\n".join(desglose_lines) +
            f"\n\n📌 Historial de Asignaciones en Outlook: {r.get('Historial_Asignaciones', '-')}\n"
            f"📌 Estado de Asignación: {r.get('Alerta_Asignacion', '-')}\n"
            f"📌 Resumen iTop: {r.get('Resumen_iTop', '-')}\n"
            f"----------------------------------------------------------------------------------------\n"
            f"CUERPO / DESCRIPCIÓN ORIGINAL DEL CORREO:\n"
            f"----------------------------------------------------------------------------------------\n"
            f"{r.get('Descripcion_Completa', '')}\n"
        )

        self.txt_detail.configure(state="normal")
        self.txt_detail.delete("1.0", "end")
        self.txt_detail.insert("1.0", detail_text)
        self.txt_detail.configure(state="disabled")

    def _toggle_selected_validation(self):
        sel = self.tree_desglose.selection()
        if not sel:
            return
        for rid in sel:
            if rid in self.validated_row_ids:
                self.validated_row_ids.remove(rid)
            else:
                self.validated_row_ids.add(rid)
        self._refresh_views()
        # Mantener selección
        for rid in sel:
            if self.tree_desglose.exists(rid):
                self.tree_desglose.selection_add(rid)

    def _validate_all_visible(self):
        _, visible_rows = self._get_active_and_filtered_rows()
        if not visible_rows:
            return
        all_already = all(r['Row_ID'] in self.validated_row_ids for r in visible_rows)
        for r in visible_rows:
            if all_already:
                self.validated_row_ids.discard(r['Row_ID'])
            else:
                self.validated_row_ids.add(r['Row_ID'])
        self._refresh_views()

    def _discard_selected_rows(self):
        sel = self.tree_desglose.selection()
        if not sel:
            messagebox.showinfo("Selecciona una fila", "Selecciona una o más filas en la tabla de desglose para descartarlas.")
            return
        for rid in sel:
            self.discarded_row_ids.add(rid)
        self._refresh_views()

    def _restore_selected_rows(self):
        sel = self.tree_desglose.selection()
        if sel:
            for rid in sel:
                self.discarded_row_ids.discard(rid)
        else:
            self.discarded_row_ids.clear()
        self._refresh_views()

    def _open_selected_in_itop(self):
        tnum = None
        sel_d = self.tree_desglose.selection()
        if sel_d:
            row_obj = next((r for r in self.all_rows if r['Row_ID'] == sel_d[0]), None)
            if row_obj:
                tnum = row_obj['Numero_Ticket']
        else:
            sel_t = self.tree_tickets.selection()
            if sel_t:
                tnum = sel_t[0].replace("TKT_", "")

        if not tnum:
            messagebox.showinfo("Abrir en iTop", "Selecciona primero una renovación o un ticket en la tabla.")
            return

        num_id = tnum.upper().replace("I-", "").strip()
        url = ITOP_PORTAL_URL.format(num_id=num_id)
        webbrowser.open(url)

    def _copy_batch_to_clipboard(self):
        active_rows, _ = self._get_active_and_filtered_rows()
        if not active_rows:
            messagebox.showwarning("Sin datos", "No hay renovaciones activas para copiar.")
            return

        lines = [
            "Ticket iTop\tParte\tNo. Cliente\tCrédito Renovado (A Cancelar)\tCrédito Origen (Padre)\tGafete\tAgente\tMotivo\tValidado"
        ]
        for r in active_rows:
            val_str = "SI" if r['Row_ID'] in self.validated_row_ids else "PENDIENTE"
            parte = f"{r['Indice_En_Ticket']}/{r['Total_En_Ticket']}"
            lines.append(
                f"{r['Numero_Ticket']}\t{parte}\t{r['Numero_Cliente']}\t"
                f"{r['Credito_Renovado_Hijo']}\t{r['Credito_Origen_Padre']}\t"
                f"{r['Gafete']}\t{r['Nombre_Asignado']}\t{r['Motivo']}\t{val_str}"
            )

        tsv_text = "\n".join(lines)
        self.clipboard_clear()
        self.clipboard_append(tsv_text)
        self.update()

        messagebox.showinfo(
            "Concentrado Copiado al Portapapeles",
            f"Se copiaron {len(active_rows)} renovaciones desglosadas al portapapeles.\n\n"
            f"Puedes pegarlas directamente (Ctrl+V) en el correo de lote diario o en una hoja de Excel."
        )

    def _export_excel(self):
        active_rows, _ = self._get_active_and_filtered_rows()
        if not active_rows and not self.all_rows:
            messagebox.showwarning("Sin datos", "Primero realiza una extracción de renovaciones antes de exportar.")
            return

        now_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        modo_str = "Concentrado_Diario" if self.var_modo.get() == "CONCENTRADO" else f"Agente_{self.ent_agente.get().strip()}"
        default_name = f"Validador_Renovaciones_{modo_str}_{now_str}.xlsx"

        filepath = filedialog.asksaveasfilename(
            title="Guardar Reporte de Validación de Renovaciones",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Archivo de Excel", "*.xlsx")]
        )
        if not filepath:
            return

        try:
            self._write_excel_report(filepath, active_rows)
            self.lbl_status.configure(text=f"📊 Reporte Excel guardado exitosamente en: {filepath}")
            if messagebox.askyesno("Reporte Generado", f"El archivo Excel se guardó correctamente:\n{filepath}\n\n¿Deseas abrirlo ahora?"):
                os.startfile(filepath)
        except Exception as e:
            messagebox.showerror("Error al exportar Excel", f"No se pudo guardar el archivo Excel:\n{e}")

    def _write_excel_report(self, filepath: str, active_rows: list):
        # 1. Preparar DataFrame de Desglose (1 fila = 1 crédito renovado a cancelar)
        desglose_data = []
        for r in active_rows:
            tnum = r['Numero_Ticket']
            num_id = tnum.upper().replace("I-", "").strip()
            url = ITOP_PORTAL_URL.format(num_id=num_id)
            desglose_data.append({
                "Validado": "SÍ" if r['Row_ID'] in self.validated_row_ids else "PENDIENTE",
                "Fecha / Hora": r['Fecha'],
                "Ticket iTop": tnum,
                "Desglose en Ticket": f"{r['Indice_En_Ticket']} de {r['Total_En_Ticket']}",
                "No. Cliente (Cuenta)": r['Numero_Cliente'],
                "Crédito Renovado (A Cancelar)": r['Credito_Renovado_Hijo'],
                "Crédito Origen (Padre)": r['Credito_Origen_Padre'],
                "Gafete": r['Gafete'],
                "Agente Responsable": r['Nombre_Asignado'],
                "Estado / Alerta de Renovación": r['Alerta_Renovacion'],
                "Estado Asignación": r['Alerta_Asignacion'],
                "Historial Asignaciones": r.get('Historial_Asignaciones', ''),
                "Motivo / Observación": r['Motivo'],
                "Sucursal / Localidad": r.get('Sucursal', ''),
                "Resumen iTop": r.get('Resumen_iTop', ''),
                "Link Portal iTop": url
            })

        # 2. Preparar DataFrame de Resumen por Ticket
        tickets_map = {}
        for r in active_rows:
            tnum = r['Numero_Ticket']
            if tnum not in tickets_map:
                tickets_map[tnum] = {
                    "Ticket iTop": tnum,
                    "Fecha / Hora": r['Fecha'],
                    "Cantidad de Renovaciones": 0,
                    "Clientes": [],
                    "Créditos Renovados a Cancelar": [],
                    "Créditos Padre": [],
                    "Gafete": r['Gafete'],
                    "Agente Responsable": r['Nombre_Asignado'],
                    "Estado Asignación": r['Alerta_Asignacion'],
                    "Resumen iTop": r.get('Resumen_iTop', '')
                }
            tm = tickets_map[tnum]
            tm["Cantidad de Renovaciones"] += 1
            if r['Numero_Cliente'] not in tm["Clientes"]:
                tm["Clientes"].append(r['Numero_Cliente'])
            tm["Créditos Renovados a Cancelar"].append(r['Credito_Renovado_Hijo'])
            if r['Credito_Origen_Padre'] != "-" and r['Credito_Origen_Padre'] not in tm["Créditos Padre"]:
                tm["Créditos Padre"].append(r['Credito_Origen_Padre'])

        tickets_data = []
        for tnum, tm in tickets_map.items():
            tickets_data.append({
                "Ticket iTop": tm["Ticket iTop"],
                "Fecha / Hora": tm["Fecha / Hora"],
                "Cantidad de Renovaciones": tm["Cantidad de Renovaciones"],
                "No. Cliente(s)": ", ".join(tm["Clientes"]),
                "Crédito(s) Renovado(s) a Cancelar": ", ".join(tm["Créditos Renovados a Cancelar"]),
                "Crédito(s) Origen (Padre)": ", ".join(tm["Créditos Padre"]) if tm["Créditos Padre"] else "-",
                "Gafete": tm["Gafete"],
                "Agente Responsable": tm["Agente Responsable"],
                "Estado Asignación": tm["Estado Asignación"],
                "Resumen iTop": tm["Resumen iTop"]
            })

        # 3. Preparar DataFrame de Resumen por Agente
        agentes_map = {}
        for r in active_rows:
            key = (r['Gafete'], r['Nombre_Asignado'])
            if key not in agentes_map:
                agentes_map[key] = {"tickets": set(), "multi": set(), "renovaciones": 0, "alertas": 0}
            agentes_map[key]["tickets"].add(r['Numero_Ticket'])
            if r.get('Es_Multiple'):
                agentes_map[key]["multi"].add(r['Numero_Ticket'])
            agentes_map[key]["renovaciones"] += 1
            if r.get('Datos_Incompletos') or r.get('Requiere_Validacion_Agente'):
                agentes_map[key]["alertas"] += 1

        agentes_data = [
            {
                "Gafete": gaf,
                "Nombre del Agente": nom,
                "Tickets de Renovación": len(v["tickets"]),
                "Total Créditos Renovados a Cancelar": v["renovaciones"],
                "Tickets con Múltiples Renovaciones": len(v["multi"]),
                "Filas con Alerta": v["alertas"]
            }
            for (gaf, nom), v in sorted(agentes_map.items(), key=lambda x: x[1]["renovaciones"], reverse=True)
        ]

        # 4. Descartados
        auto_excl = bool(self.var_auto_exclude_reassigned.get()) and (self.var_modo.get() == "AGENTE")
        descartados_data = [
            {
                "Fecha / Hora": r['Fecha'],
                "Ticket iTop": r['Numero_Ticket'],
                "No. Cliente": r['Numero_Cliente'],
                "Crédito Renovado": r['Credito_Renovado_Hijo'],
                "Gafete": r['Gafete'],
                "Agente": r['Nombre_Asignado'],
                "Motivo Descarte / Estado": r['Alerta_Asignacion']
            }
            for r in self.all_rows
            if (r['Row_ID'] in self.discarded_row_ids) or (auto_excl and r.get('Reasignado_A_Otro'))
        ]

        with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
            df1 = pd.DataFrame(desglose_data) if desglose_data else pd.DataFrame(columns=["Sin datos"])
            df2 = pd.DataFrame(tickets_data) if tickets_data else pd.DataFrame(columns=["Sin datos"])
            df3 = pd.DataFrame(agentes_data) if agentes_data else pd.DataFrame(columns=["Sin datos"])

            df1.to_excel(writer, sheet_name="1_Desglose_Renovaciones", index=False)
            df2.to_excel(writer, sheet_name="2_Resumen_Por_Ticket", index=False)
            df3.to_excel(writer, sheet_name="3_Resumen_Por_Agente", index=False)
            if descartados_data:
                pd.DataFrame(descartados_data).to_excel(writer, sheet_name="4_Descartados_Excluidos", index=False)

            # Dar formato profesional y agregar hipervínculos a iTop
            from openpyxl.styles import PatternFill, Font, Alignment
            wb = writer.book
            hdr_fill = PatternFill(start_color="0F2942", end_color="0F2942", fill_type="solid")
            hdr_font = Font(color="FFFFFF", bold=True)
            link_font = Font(color="1D4ED8", underline="single", bold=True)

            fill_mult = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
            fill_inc = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")

            for ws in wb.worksheets:
                ws.views.sheetView[0].showGridLines = True
                for cell in ws[1]:
                    cell.fill = hdr_fill
                    cell.font = hdr_font
                    cell.alignment = Alignment(horizontal="center", vertical="center")

                # Ajustar ancho de columnas
                for col in ws.columns:
                    col_letter = col[0].column_letter
                    max_len = max((len(str(c.value or "")) for c in col), default=12)
                    ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 48)

            # Convertir columna Ticket iTop en hipervínculo clicable en Hoja 1
            ws1 = wb["1_Desglose_Renovaciones"]
            if desglose_data:
                for row_idx, r_orig in enumerate(active_rows, start=2):
                    t_cell = ws1.cell(row=row_idx, column=3)
                    tnum = str(t_cell.value or "")
                    if tnum.upper().startswith("I-"):
                        num_id = tnum.upper().replace("I-", "").strip()
                        t_cell.hyperlink = ITOP_PORTAL_URL.format(num_id=num_id)
                        t_cell.font = link_font

                    # Colorear fila si es múltiple o incompleta
                    row_fill = None
                    if r_orig.get("Datos_Incompletos"):
                        row_fill = fill_inc
                    elif r_orig.get("Es_Multiple"):
                        row_fill = fill_mult
                    if row_fill:
                        for c_idx in range(1, ws1.max_column + 1):
                            ws1.cell(row=row_idx, column=c_idx).fill = row_fill

    def _sort_tree(self, tree, col, reverse):
        items = [(tree.set(k, col), k) for k in tree.get_children("")]
        items.sort(reverse=reverse)
        for index, (_val, k) in enumerate(items):
            tree.move(k, "", index)
        tree.heading(col, command=lambda: self._sort_tree(tree, col, not reverse))


if __name__ == "__main__":
    app = FonacotValidadorRenovacionesApp()
    app.mainloop()

