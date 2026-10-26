import os
import re
import sys
import json
import threading
import webbrowser
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import tkinter.font as tkfont

import customtkinter as ctk
from tkcalendar import DateEntry
import pandas as pd

from outlook_renovaciones_engine import OutlookRenovacionesEngine, DIAS_ATRAS_LOTE
from sql_renovaciones import (
    conectar as conectar_sql,
    validar_contra_base,
    probar_conexion,
    texto_corto as texto_sql,
    leer_password,
    guardar_password,
    borrar_password,
    servidor_de,
    separar_servidor,
    host_invalido,
    ErrorSQL,
    TABLA_DEFAULT,
    SQL_PROBLEMAS,
    SQL_SUGERENCIA,
    SQL_OTRA_FECHA,
    SQL_APP_PLAZO,
    SQL_NO_PROCEDE,
    LEYENDA_NO_PROCEDE,
)
from calendario_cedulas import actualizar_calendario, marcar_bloqueos
from core_renovaciones import (
    build_itop_url,
    build_itop_search_url,
    parse_hora_corte,
    previous_business_day,
    parse_dias_inhabiles,
    set_dias_inhabiles,
    HORA_CORTE_DEFAULT,
    ASUNTO_CONCENTRADO_DEFAULT,
    LOTE_HOY,
    LOTE_REZAGADA,
    LOTE_ENVIADA,
    LOTE_EN_CONCENTRADO,
    LOTE_SIGUIENTE,
    LOTE_CERRADA_SIN_ENVIO,
    LOTE_SIN_COMPROBAR,
    resumen_cierre
)
from lote_utils import (
    EstadoValidaciones,
    row_key,
    build_lote_tsv,
    build_ticket_unico_tsv,
    build_ticket_unico_html,
    set_clipboard_html,
    plain_text,
    lote_corto,
    INTRO_TICKET_UNICO_DEFAULT
)


APP_VERSION = "2.7"

CONFIG_FILE = os.path.join(
    os.path.expanduser("~"),
    ".fonacot_validador_renovaciones_config.json"
)
# Validaciones / descartes guardados entre escaneos y reinicios de la aplicación
ESTADO_FILE = os.path.join(
    os.path.expanduser("~"),
    ".fonacot_validador_renovaciones_estado.json"
)

# Tipografía y paleta (estilo sobrio tipo Fluent: superficies blancas, grises neutros y un solo azul de acento)
FONT = "Segoe UI"
FONT_SB = "Segoe UI Semibold"
FONT_ICON = "Segoe MDL2 Assets"

COL = {
    "bg": "#F4F6F9",
    "surface": "#FFFFFF",
    "subtle": "#F8FAFC",
    "track": "#EEF1F5",
    "border": "#E2E8F0",
    "border_strong": "#CBD5E1",
    "text": "#0F172A",
    "text_2": "#334155",
    "muted": "#64748B",
    "faint": "#94A3B8",
    "primary": "#0B5CAD",
    "primary_h": "#094B8F",
    "primary_soft": "#EAF2FB",
    "select": "#DBEAFE",
    "success": "#15803D",
    "success_soft": "#ECFDF5",
    "warning": "#B45309",
    "warning_soft": "#FFFBEB",
    "danger": "#B91C1C",
    "danger_soft": "#FEF2F2",
}

# Estados de lote que NO forman parte del lote a enviar hoy (se ven en su filtro y se pueden "Restaurar")
LOTE_EXCLUIDOS = (LOTE_ENVIADA, LOTE_SIGUIENTE, LOTE_CERRADA_SIN_ENVIO, LOTE_SIN_COMPROBAR)

FILTRO_TODAS = "Todas"
FILTRO_PENDIENTES = "Pendientes"
FILTRO_ALERTAS = "Con alertas"
FILTRO_MULTIPLES = "Múltiples"
FILTRO_REZAGADAS = "Rezagadas"
FILTRO_NUEVAS = "Nuevas"
FILTRO_VALIDADAS = "Validadas"
FILTRO_ENVIADAS = "Ya enviadas (excluidas)"
FILTRO_SIGUIENTE = "Para el siguiente lote"
FILTRO_PAPELERA = "Descartadas"
FILTRO_CERRADAS = "Cerradas en iTop"
FILTRO_NO_ENVIADAS = "No enviadas en su día"
FILTRO_TARDE = "Enviadas tarde"
FILTRO_SIN_COMPROBAR = "Sin comprobar"
FILTRO_ERRORES = "Datos equivocados en el ticket único"
FILTRO_SQL = "No cuadran con la base (SQL)"
FILTRO_EMISION = "Ya no se pueden cancelar (emisión)"
OTRAS_VISTAS = "Otras vistas"

FILTROS_PRINCIPALES = [FILTRO_TODAS, FILTRO_PENDIENTES, FILTRO_ALERTAS, FILTRO_MULTIPLES, FILTRO_REZAGADAS, FILTRO_NUEVAS]
FILTROS_SECUNDARIOS = [FILTRO_EMISION, FILTRO_NO_ENVIADAS, FILTRO_ERRORES, FILTRO_SQL, FILTRO_TARDE, FILTRO_SIN_COMPROBAR, FILTRO_CERRADAS, FILTRO_VALIDADAS,
                       FILTRO_ENVIADAS, FILTRO_SIGUIENTE, FILTRO_PAPELERA]
# Vistas que muestran filas aunque estén fuera del lote de hoy (revisión de envíos)
FILTROS_REVISION = (FILTRO_ENVIADAS, FILTRO_SIGUIENTE, FILTRO_NO_ENVIADAS, FILTRO_ERRORES, FILTRO_SQL, FILTRO_TARDE,
                    FILTRO_SIN_COMPROBAR)
HINT_EXTRAER = "Lee Outlook y arma el lote del día  ·  F5"
TIPOS_REVISION = ("Lote del día", "Revisión semanal")

MODOS = {"CONCENTRADO": "Concentrado", "AGENTE": "Por compañero"}

# Mismo asunto que el correo que ya se envía: "Cancelación de renovaciones – ticket único 22-10-2026"
ASUNTO_BORRADOR_DEFAULT = "Cancelación de renovaciones – ticket único {fecha}"
ASUNTO_BORRADOR_ANTERIOR = "TICKET UNICO {fecha}"

# Columnas que se ocultan de inicio para no saturar la tabla (se activan en «Columnas» o con clic derecho en el encabezado)
COLUMNAS_OCULTAS_DEFAULT = {
    "desglose": ["credito_padre", "alerta_asig", "motivo"],
    "tickets": ["creditos_padres", "alerta_asig"],
}
COLUMNAS_FIJAS = ("validado", "ticket")
VISTA_DEFAULT = {"indicadores": True, "detalle": True, "configuracion": True}
VISTA_ZONAS = (
    ("indicadores", "Indicadores (tarjetas)"),
    ("detalle", "Detalle del ticket y correo"),
    ("configuracion", "Configuración (panel izquierdo)"),
)

ATAJOS = [
    ("F5", "Extraer y validar renovaciones"),
    ("Ctrl + F", "Ir al buscador (Esc lo limpia)"),
    ("Espacio / doble clic", "Marcar o desmarcar como validada"),
    ("Enter", "Abrir el ticket seleccionado en iTop"),
    ("Supr", "Descartar las filas seleccionadas"),
    ("Ctrl + C", "Copiar las filas seleccionadas"),
    ("Ctrl + Shift + C", "Copiar el lote completo (tabla para correo)"),
    ("Ctrl + M", "Crear el correo 'Ticket Único' en Outlook"),
    ("Ctrl + E", "Exportar a Excel"),
    ("Clic derecho", "Más acciones: copiar cliente o crédito, buscar en iTop…"),
    ("Vista / clic derecho en el encabezado", "Elegir columnas y ocultar indicadores, detalle o configuración"),
    ("Clic en un indicador", "Filtra la tabla por ese indicador"),
]


def F(size: int = 12, bold: bool = False) -> ctk.CTkFont:
    return ctk.CTkFont(family=FONT_SB if bold else FONT, size=size)


class TooltipCeldas:
    """Al pasar el mouse muestra el texto completo de la celda cuando no cabe o en la tabla va resumido."""

    def __init__(self, tree: ttk.Treeview, completos: dict):
        self.tree = tree
        self.completos = completos      # {(iid, columna): texto completo}
        self.font = tkfont.Font(family=FONT, size=10)
        self.tip = None
        self._celda = None
        self._after = None
        tree.bind("<Motion>", self._mover, add="+")
        tree.bind("<Leave>", self._ocultar, add="+")
        tree.bind("<ButtonPress>", self._ocultar, add="+")

    def _mover(self, e):
        iid, col = self.tree.identify_row(e.y), self.tree.identify_column(e.x)
        if (iid, col) == self._celda:
            return
        self._ocultar()
        self._celda = (iid, col)
        if not iid or not col or col == "#0":
            return
        nombre = self.tree.column(col, "id")
        mostrado = str(self.tree.set(iid, nombre))
        texto = self.completos.get((iid, nombre)) or mostrado
        if not texto.strip() or (texto == mostrado and self.font.measure(mostrado) <= self.tree.column(nombre, "width") - 14):
            return
        self._after = self.tree.after(350, lambda: self._mostrar(texto, e.x_root, e.y_root))

    def _mostrar(self, texto, x, y):
        self.tip = tk.Toplevel(self.tree)
        self.tip.wm_overrideredirect(True)
        self.tip.attributes("-topmost", True)
        tk.Label(
            self.tip, text=texto, justify="left", wraplength=520, font=(FONT, 10), bg="#FFFFFF", fg=COL["text"],
            relief="solid", borderwidth=1, padx=8, pady=5
        ).pack()
        self.tip.wm_geometry(f"+{x + 14}+{y + 18}")

    def _ocultar(self, _e=None):
        if self._after:
            self.tree.after_cancel(self._after)
            self._after = None
        if self.tip:
            self.tip.destroy()
            self.tip = None
        if _e is not None and getattr(_e, "type", None) == tk.EventType.Leave:
            self._celda = None


def _aplicar_tema():
    """Ajusta el tema de CustomTkinter antes de crear widgets: Segoe UI, esquinas discretas y colores neutros."""
    ctk.set_appearance_mode("Light")
    ctk.set_default_color_theme("blue")
    t = ctk.ThemeManager.theme

    def c(color):
        return [color, color]

    t["CTkFont"]["family"] = FONT
    t["CTk"]["fg_color"] = c(COL["bg"])
    t["CTkToplevel"]["fg_color"] = c(COL["surface"])
    t["CTkButton"].update(corner_radius=6, fg_color=c(COL["primary"]), hover_color=c(COL["primary_h"]))
    t["CTkEntry"].update(
        corner_radius=6, border_width=1, fg_color=c(COL["surface"]), border_color=c(COL["border_strong"]),
        text_color=c(COL["text"]), placeholder_text_color=c(COL["faint"])
    )
    t["CTkComboBox"].update(
        corner_radius=6, border_width=1, fg_color=c(COL["surface"]), border_color=c(COL["border_strong"]),
        button_color=c(COL["border_strong"]), button_hover_color=c(COL["faint"]), text_color=c(COL["text"])
    )
    t["CTkOptionMenu"].update(
        corner_radius=6, fg_color=c(COL["track"]), button_color=c(COL["track"]),
        button_hover_color=c(COL["border"]), text_color=c(COL["text_2"])
    )
    t["DropdownMenu"].update(fg_color=c(COL["surface"]), hover_color=c(COL["primary_soft"]), text_color=c(COL["text"]))
    t["CTkCheckBox"].update(
        corner_radius=4, border_width=2, fg_color=c(COL["primary"]), hover_color=c(COL["primary_h"]),
        border_color=c(COL["faint"]), text_color=c(COL["text_2"])
    )
    t["CTkSwitch"].update(
        fg_color=c(COL["border_strong"]), progress_color=c(COL["primary"]),
        button_color=c(COL["surface"]), button_hover_color=c(COL["surface"]), text_color=c(COL["text_2"])
    )
    t["CTkSegmentedButton"].update(
        corner_radius=6, border_width=3, fg_color=c(COL["track"]), selected_color=c(COL["surface"]),
        selected_hover_color=c(COL["surface"]), unselected_color=c(COL["track"]),
        unselected_hover_color=c(COL["border"]), text_color=c(COL["text"])
    )
    t["CTkProgressBar"].update(fg_color=c(COL["track"]), progress_color=c(COL["primary"]))
    t["CTkScrollbar"].update(button_color=c(COL["border"]), button_hover_color=c(COL["border_strong"]))
    t["CTkTextbox"].update(
        corner_radius=8, border_width=1, fg_color=c(COL["surface"]), border_color=c(COL["border"]),
        text_color=c(COL["text"]), scrollbar_button_color=c(COL["border"]),
        scrollbar_button_hover_color=c(COL["border_strong"])
    )


def _sort_key(value):
    """Orden natural: fechas dd/mm/yyyy HH:MM, folios I-123, '1 de 3' y números como número; lo demás como texto."""
    v = str(value).strip()
    try:
        return (0, datetime.strptime(v, "%d/%m/%Y %H:%M").timestamp())
    except ValueError:
        pass
    m = re.fullmatch(r'(\d+)\s*(?:de|/)\s*\d+', v)
    if m:
        return (0, float(m.group(1)))
    m = re.fullmatch(r'[A-Za-z]{0,2}-?\s*(\d+(?:\.\d+)?)', v)
    if m:
        return (0, float(m.group(1)))
    m = re.match(r'\D{0,4}(\d+)', v)
    if m and len(v) < 25:
        return (0, float(m.group(1)))
    return (1, v.lower())


class SelectorFiltro:
    """Filtros rápidos (botones segmentados) + vistas secundarias (menú) que se leen como un solo valor con get()/set()."""

    def __init__(self, seg: ctk.CTkSegmentedButton, menu: ctk.CTkOptionMenu, on_change):
        self.seg, self.menu, self.on_change = seg, menu, on_change
        self._valor = None
        seg.configure(command=self._desde_seg)
        menu.configure(command=self._desde_menu)
        self.set(FILTRO_TODAS)

    def get(self) -> str:
        return self._valor

    def set(self, valor: str):
        self._valor = valor
        if valor in FILTROS_PRINCIPALES:
            self.seg.set(valor)
            self.menu.set(OTRAS_VISTAS)
            self.menu.configure(fg_color=COL["track"], button_color=COL["track"], text_color=COL["text_2"])
        else:
            self.seg.set("")
            self.menu.set(valor)
            self.menu.configure(fg_color=COL["primary_soft"], button_color=COL["primary_soft"], text_color=COL["primary"])

    def _desde_seg(self, valor):
        self.set(valor)
        self.on_change()

    def _desde_menu(self, valor):
        self.set(FILTRO_TODAS if valor == OTRAS_VISTAS else valor)
        self.on_change()


class PanelDesplazable(ctk.CTkScrollableFrame):
    """
    CTkScrollableFrame que avanza 60 px por golpe de rueda (customtkinter avanza 20 en Windows y con un panel tan
    largo parecía que no se movía) y que también responde a los giros chicos de los touchpads.
    """

    def _pasos(self, delta: int) -> int:
        return -int(delta / 2) or (-10 if delta > 0 else 10)

    def _mouse_wheel_all(self, event):
        if (sys.platform.startswith("win") and not self._shift_pressed and event.delta
                and self._check_if_valid_scroll(event.widget)):
            self._parent_canvas.yview("scroll", self._pasos(event.delta), "units")
            return
        super()._mouse_wheel_all(event)

    def desplazar(self, event):
        """Para widgets de fuera del panel (el bloque de «Extraer y validar»)."""
        if event.delta:
            self._parent_canvas.yview("scroll", self._pasos(event.delta), "units")


class PestanasSubrayadas:
    """Pestañas con subrayado (en lugar del CTkTabview de botones) y contador en la etiqueta."""

    def __init__(self, header, body):
        self.header, self.body = header, body
        self._tabs = {}
        self._actual = None
        body.grid_columnconfigure(0, weight=1)
        body.grid_rowconfigure(0, weight=1)

    def add(self, nombre: str, etiqueta: str) -> ctk.CTkFrame:
        box = ctk.CTkFrame(self.header, fg_color="transparent")
        box.grid(row=0, column=len(self._tabs), padx=(0, 6), sticky="s")
        btn = ctk.CTkButton(
            box, text=etiqueta, font=F(13, True), height=38, width=10, corner_radius=6,
            fg_color="transparent", hover_color=COL["track"], text_color=COL["muted"],
            command=lambda: self.set(nombre)
        )
        btn.pack(padx=2, ipadx=6)
        linea = tk.Frame(box, height=3, bg=COL["surface"], highlightthickness=0)
        linea.pack(fill="x", padx=8)
        frame = ctk.CTkFrame(self.body, fg_color="transparent", corner_radius=0)
        frame.grid(row=0, column=0, sticky="nsew")
        frame.grid_remove()
        self._tabs[nombre] = (btn, linea, frame)
        if self._actual is None:
            self.set(nombre)
        return frame

    def set(self, nombre: str):
        for n, (btn, linea, frame) in self._tabs.items():
            activo = n == nombre
            btn.configure(text_color=COL["text"] if activo else COL["muted"])
            linea.configure(bg=COL["primary"] if activo else COL["surface"])
            if activo:
                frame.grid()
            else:
                frame.grid_remove()
        self._actual = nombre

    def get(self) -> str:
        return self._actual

    def set_label(self, nombre: str, etiqueta: str):
        self._tabs[nombre][0].configure(text=etiqueta)


class FonacotValidadorRenovacionesApp(ctk.CTk):
    def __init__(self):
        _aplicar_tema()
        super().__init__()

        self.title("Validador de Renovaciones · FONACOT")
        self._fit_to_screen()
        self.minsize(1100, 640)
        self.configure(fg_color=COL["bg"])

        self.engine = OutlookRenovacionesEngine()
        self.config_data = self._load_config()
        self.estado = EstadoValidaciones(ESTADO_FILE)

        # Estado de datos
        self.all_rows = []
        self.discarded_row_ids = set()
        self.validated_row_ids = set()
        self.forced_included_row_ids = set()  # Filas ENVIADA / SIGUIENTE_LOTE que el usuario decidió meter al lote
        self._key_by_rid = {}
        self.lote_info = None
        self.validaciones = {}  # {'ocr': resumen de imágenes leídas, 'sql': resumen de la validación en la base}
        self.is_extracting = False

        # Detección de renovaciones nuevas entre escaneos (auto-actualización)
        self._prev_keys = None
        self._prev_scan_signature = None
        self._ultima_busqueda = None  # (revisión, desde, hasta) de la última búsqueda pedida con el botón o F5
        self.new_keys = set()

        self._auto_job = None
        self._search_job = None
        self._aviso_corte_dia = None
        self._sort_state = {}
        self._headers = {}
        self._celdas_completas = {}   # {(iid, columna): texto completo} para el tooltip de celdas resumidas
        self._vars_columnas = []

        self._build_ui()
        self._build_context_menus()
        self._bind_shortcuts()

        self.after(50, self._maximize)
        self.after(300, self._init_accounts)
        self.after(500, self._tick_clock)
        self.after(800, self._schedule_auto_refresh)

    # ================= CONFIGURACIÓN =================

    def _load_config(self) -> dict:
        default_cfg = {
            "account": "",
            "modo_operacion": "CONCENTRADO",  # "CONCENTRADO" o "AGENTE"
            "agente_filter": "2609",
            "auto_exclude_reassigned": False,
            "aplicar_corte": True,
            "hora_corte": HORA_CORTE_DEFAULT,
            "asunto_concentrado": ASUNTO_CONCENTRADO_DEFAULT,
            # Días inhábiles además de los feriados oficiales ("02/11/2026, 12/12/2026")
            "dias_inhabiles": "",
            "feriados_oficiales": True,
            # Días hábiles hacia atrás que revisa el lote del día para no perder lo que no se envió
            "dias_atras_lote": DIAS_ATRAS_LOTE,
            # Capturas de pantalla en los incidentes (OCR de Windows) y validación en la base (Oracle)
            "leer_imagenes": True,
            "sql_activo": False,
            "sql_servidor": "",
            "sql_usuario": "",
            "sql_tabla": TABLA_DEFAULT,
            "sql_recordar": False,
            "sql_cliente_oracle": "",  # carpeta del Oracle Instant Client, solo si la base no acepta el modo thin
            # Correo del lote ("Crear correo Ticket Único"); {fecha} se reemplaza por dd/mm/aaaa
            "asunto_borrador": ASUNTO_BORRADOR_DEFAULT,
            "correo_para": "",
            "correo_cc": "",
            # Texto antes de la tabla; admite {total}. Vacío = el aprendido del último concentrado o el predeterminado.
            "texto_correo": "",
            "texto_correo_aprendido": "",
            "auto_refresh": False,
            "auto_refresh_min": "10",
            # Columnas ocultas por pestaña ("desglose", "tickets"); se eligen en el botón «Columnas»
            "columnas_ocultas": {k: list(v) for k, v in COLUMNAS_OCULTAS_DEFAULT.items()},
            # Zonas de la pantalla que se pueden ocultar desde «Vista» para dar más espacio a la tabla
            "vista": dict(VISTA_DEFAULT),
            # Plantilla opcional para otro servidor de iTop; admite {id}, {ref} y {clase}. Vacío = servidor FONACOT.
            "itop_url_template": "",
            # Marca local de la última renovación incluida en un lote copiado/exportado
            "ultimo_lote": None
        }
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    default_cfg.update(data)
                if default_cfg.get("asunto_borrador") == ASUNTO_BORRADOR_ANTERIOR:
                    default_cfg["asunto_borrador"] = ASUNTO_BORRADOR_DEFAULT
            except Exception:
                pass
        return default_cfg

    def _save_config(self):
        try:
            self.config_data.update({
                "account": self.cmb_account.get().strip(),
                "modo_operacion": self.var_modo.get(),
                "agente_filter": self.ent_agente.get().strip(),
                "auto_exclude_reassigned": bool(self.var_auto_exclude_reassigned.get()),
                "aplicar_corte": bool(self.var_aplicar_corte.get()),
                "hora_corte": parse_hora_corte(self.ent_hora_corte.get()).strftime("%H:%M"),
                "asunto_concentrado": self.ent_asunto_conc.get().strip() or ASUNTO_CONCENTRADO_DEFAULT,
                "dias_inhabiles": self.ent_inhabiles.get().strip(),
                "feriados_oficiales": bool(self.var_oficiales.get()),
                "leer_imagenes": bool(self.var_ocr.get()),
                "sql_activo": bool(self.var_sql.get()),
                "sql_host": self.ent_sql_host.get().strip(),
                "sql_puerto": self.ent_sql_puerto.get().strip() or "1521",
                "sql_servicio": self.ent_sql_servicio.get().strip(),
                "sql_servidor": self._sql_servidor(),
                "sql_usuario": self.ent_sql_usuario.get().strip(),
                "sql_tabla": self.ent_sql_tabla.get().strip() or TABLA_DEFAULT,
                "sql_recordar": bool(self.var_sql_recordar.get()),
                "asunto_borrador": self.ent_asunto_borrador.get().strip() or ASUNTO_BORRADOR_DEFAULT,
                "correo_para": self.ent_correo_para.get().strip(),
                "correo_cc": self.ent_correo_cc.get().strip(),
                "auto_refresh": bool(self.var_auto.get()),
                "auto_refresh_min": self.cmb_auto_min.get().strip() or "10"
            })
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.config_data, f, indent=2, ensure_ascii=False)
        except Exception:
            pass
        self._guardar_password_sql()

    def _guardar_password_sql(self):
        """La contraseña nunca va al archivo de configuración: si se pide recordarla, va al Administrador de credenciales."""
        try:
            usuario, servidor = self.ent_sql_usuario.get().strip(), self._sql_servidor()
            password = self.ent_sql_password.get()
        except Exception:
            return
        if not (usuario and servidor):
            return
        if self.var_sql_recordar.get() and password:
            guardar_password(usuario, servidor, password)
        elif not self.var_sql_recordar.get():
            borrar_password(usuario, servidor)

    def _scaling(self) -> float:
        try:
            return float(ctk.ScalingTracker.get_window_scaling(self)) or 1.0
        except Exception:
            return 1.0

    def _fit_to_screen(self):
        """Tamaño inicial que cabe en la pantalla aun con escala de Windows 125-150% (geometry() va en unidades lógicas)."""
        scale = self._scaling()
        screen_w = self.winfo_screenwidth() / scale
        screen_h = self.winfo_screenheight() / scale
        w = int(min(1500, screen_w - 20))
        h = int(min(920, screen_h - 90))
        self.geometry(f"{w}x{h}+0+0")

    def _maximize(self):
        try:
            self.state("zoomed")
        except Exception:
            pass

    # ================= CONSTRUCCIÓN DE LA INTERFAZ =================

    def _build_ui(self):
        self._setup_styles()
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self._build_topbar()

        body = ctk.CTkFrame(self, fg_color="transparent", corner_radius=0)
        body.grid(row=1, column=0, sticky="nsew")
        body.grid_columnconfigure(2, weight=1)
        body.grid_rowconfigure(0, weight=1)

        self._build_sidebar(body)

        main = ctk.CTkFrame(body, fg_color="transparent", corner_radius=0)
        main.grid(row=0, column=2, sticky="nsew", padx=24, pady=(18, 14))
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(3, weight=1)

        self._build_page_header(main)
        self._build_notice(main)
        self._build_kpis(main)
        self._build_work_area(main)
        self._build_status_bar()

        self._on_mode_change()
        self._refresh_views()

    def _hline(self, parent) -> ctk.CTkFrame:
        return ctk.CTkFrame(parent, height=1, corner_radius=0, fg_color=COL["border"])

    def _btn(self, parent, text: str, command, kind: str = "secondary", **kw) -> ctk.CTkButton:
        estilos = {
            "primary": dict(fg_color=COL["primary"], hover_color=COL["primary_h"], text_color="#FFFFFF", border_width=0),
            "secondary": dict(
                fg_color=COL["surface"], hover_color=COL["track"], text_color=COL["text"],
                border_width=1, border_color=COL["border_strong"]
            ),
            "danger": dict(
                fg_color=COL["surface"], hover_color=COL["danger_soft"], text_color=COL["danger"],
                border_width=1, border_color="#FECACA"
            ),
            "ghost": dict(fg_color="transparent", hover_color=COL["track"], text_color=COL["text_2"], border_width=0),
            "link": dict(fg_color="transparent", hover_color=COL["primary_soft"], text_color=COL["primary"], border_width=0),
        }
        opts = dict(height=32, width=10, corner_radius=6, font=F(12, kind == "primary"))
        opts.update(estilos[kind])
        opts.update(kw)
        return ctk.CTkButton(parent, text=text, command=command, **opts)

    def _build_topbar(self):
        bar = ctk.CTkFrame(self, fg_color=COL["surface"], corner_radius=0)
        bar.grid(row=0, column=0, sticky="ew")
        bar.grid_columnconfigure(1, weight=1)

        brand = ctk.CTkFrame(bar, fg_color="transparent")
        brand.grid(row=0, column=0, padx=(20, 0), pady=11, sticky="w")
        ctk.CTkLabel(
            brand, text="F", width=34, height=34, corner_radius=8,
            fg_color=COL["primary"], text_color="#FFFFFF", font=F(17, True)
        ).pack(side="left")
        names = ctk.CTkFrame(brand, fg_color="transparent")
        names.pack(side="left", padx=(12, 0))
        ctk.CTkLabel(
            names, text="Validador de Renovaciones", font=F(15, True), text_color=COL["text"], height=20
        ).pack(anchor="w")
        ctk.CTkLabel(
            names, text="FONACOT  ·  Cancelación y reversa de renovaciones",
            font=F(11), text_color=COL["muted"], height=16
        ).pack(anchor="w")

        right = ctk.CTkFrame(bar, fg_color="transparent")
        right.grid(row=0, column=2, padx=20, sticky="e")
        # Reloj y estado del ticket único de hoy
        self.lbl_clock = ctk.CTkLabel(
            right, text="--:--", font=F(12, True), height=28, corner_radius=14,
            fg_color=COL["track"], text_color=COL["text_2"]
        )
        self.lbl_clock.pack(side="left", padx=(0, 16), ipadx=12)
        ctk.CTkFrame(right, width=1, height=28, fg_color=COL["border"], corner_radius=0).pack(side="left", padx=(0, 16))
        ctk.CTkLabel(right, text="Buzón de Outlook", font=F(12), text_color=COL["muted"]).pack(side="left", padx=(0, 8))
        self.cmb_account = ctk.CTkComboBox(
            right, values=["Cargando cuentas..."], width=260, height=32, font=F(12), dropdown_font=F(12)
        )
        self.cmb_account.pack(side="left")

        self._hline(self).grid(row=0, column=0, sticky="sew")

    def _section(self, parent, titulo: str, first: bool = False) -> ctk.CTkFrame:
        if not first:
            self._hline(parent).pack(fill="x", padx=12, pady=(18, 0))
        ctk.CTkLabel(
            parent, text=titulo.upper(), font=F(11, True), text_color=COL["muted"], anchor="w", height=16
        ).pack(fill="x", padx=12, pady=(16 if not first else 4, 8))
        inner = ctk.CTkFrame(parent, fg_color="transparent")
        inner.pack(fill="x", padx=12)
        inner.grid_columnconfigure(0, weight=1)
        return inner

    def _field(self, parent, row: int, label: str, value: str = "", placeholder: str = "", width: int = 0) -> ctk.CTkEntry:
        """Etiqueta arriba y campo abajo (ocupa las filas row y row+1)."""
        ctk.CTkLabel(parent, text=label, font=F(12), text_color=COL["text_2"], anchor="w", height=18).grid(
            row=row, column=0, columnspan=2, sticky="w", pady=(8, 3)
        )
        ent = ctk.CTkEntry(parent, height=32, font=F(12), placeholder_text=placeholder, **({"width": width} if width else {}))
        ent.grid(row=row + 1, column=0, columnspan=2, sticky="w" if width else "ew")
        if value:
            ent.insert(0, value)
        return ent

    def _build_sidebar(self, parent):
        wrap = ctk.CTkFrame(parent, fg_color=COL["surface"], corner_radius=0)
        wrap.grid(row=0, column=0, sticky="ns")
        wrap.grid_rowconfigure(1, weight=1)
        wrap.grid_columnconfigure(0, weight=1)
        ctk.CTkFrame(parent, width=1, corner_radius=0, fg_color=COL["border"]).grid(row=0, column=1, sticky="ns")

        top = ctk.CTkFrame(wrap, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 6))
        self.btn_extract = ctk.CTkButton(top, text="Extraer y validar", font=F(13, True), height=40, command=self._start_extraction)
        self.btn_extract.pack(fill="x")
        self.lbl_extract_hint = ctk.CTkLabel(
            top, text=HINT_EXTRAER, font=F(11), text_color=COL["muted"], justify="left", anchor="w", wraplength=230
        )
        self.lbl_extract_hint.pack(anchor="w", pady=(8, 0))

        side = PanelDesplazable(
            wrap, width=262, fg_color=COL["surface"], corner_radius=0,
            scrollbar_button_color=COL["border_strong"], scrollbar_button_hover_color=COL["muted"]
        )
        side.grid(row=1, column=0, sticky="nsew", padx=(8, 0), pady=(0, 8))
        self.sidebar_config = side
        for w in (top, self.btn_extract, self.lbl_extract_hint):
            w.bind("<MouseWheel>", side.desplazar, add="+")

        # Qué se revisa: el lote de hoy o los envíos de una semana
        sec = self._section(side, "¿Qué quieres revisar?", first=True)
        self.seg_tipo = ctk.CTkSegmentedButton(sec, values=list(TIPOS_REVISION), font=F(12), height=32,
                                               command=self._on_tipo_revision)
        self.seg_tipo.grid(row=0, column=0, columnspan=2, sticky="ew")
        self.seg_tipo.set(TIPOS_REVISION[0])
        semanas = ctk.CTkFrame(sec, fg_color="transparent")
        semanas.grid(row=1, column=0, columnspan=2, pady=(6, 0), sticky="w")
        for txt, atras in (("Esta semana", 0), ("Semana pasada", 1)):
            self._btn(semanas, txt, lambda a=atras: self._revisar_semana(a), "link", height=26).pack(
                side="left", padx=(0, 4), ipadx=2)
        ctk.CTkLabel(
            sec, text="La revisión semanal junta todo lo que salió en los tickets únicos de la semana y te dice "
                      "qué renovaciones no se mandaron o salieron con datos equivocados.",
            font=F(11), text_color=COL["faint"], justify="left", anchor="w", wraplength=250
        ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(4, 0))

        # Modo
        sec = self._section(side, "Modo de validación")
        self.var_modo = tk.StringVar(value=self.config_data.get("modo_operacion", "CONCENTRADO"))
        self.seg_modo = ctk.CTkSegmentedButton(
            sec, values=list(MODOS.values()), font=F(12), height=32, command=self._on_modo_seg
        )
        self.seg_modo.grid(row=0, column=0, columnspan=2, sticky="ew")
        self.seg_modo.set(MODOS.get(self.var_modo.get(), MODOS["CONCENTRADO"]))
        self.ent_agente = self._field(sec, 1, "Gafete o nombre del compañero", self.config_data.get("agente_filter", "2609"),
                                      "Ej. 2609, 7353, Nancy")
        self.var_auto_exclude_reassigned = tk.BooleanVar(value=self.config_data.get("auto_exclude_reassigned", False))
        self.chk_exclude_reassigned = ctk.CTkCheckBox(
            sec, text="Excluir reasignados a otro compañero", variable=self.var_auto_exclude_reassigned,
            command=self._refresh_views, font=F(12), checkbox_width=18, checkbox_height=18
        )
        self.chk_exclude_reassigned.grid(row=3, column=0, columnspan=2, pady=(10, 0), sticky="w")

        # Periodo
        sec = self._section(side, "Periodo de recepción")
        sec.grid_columnconfigure((0, 1), weight=1, uniform="fecha")
        cal_opts = dict(
            width=11, date_pattern="dd/mm/yyyy", font=(FONT, 10), borderwidth=0,
            background=COL["primary"], foreground="white", headersbackground=COL["subtle"],
            headersforeground=COL["muted"], selectbackground=COL["primary"], selectforeground="white",
            normalbackground="white", weekendbackground="white", othermonthbackground=COL["subtle"],
            othermonthwebackground=COL["subtle"], othermonthforeground=COL["faint"],
            othermonthweforeground=COL["faint"], bordercolor=COL["border"]
        )
        for col, txt in ((0, "Del"), (1, "Al")):
            ctk.CTkLabel(sec, text=txt, font=F(12), text_color=COL["text_2"], height=18).grid(
                row=0, column=col, sticky="w", pady=(0, 3), padx=(0, 6) if col == 0 else (6, 0)
            )
        self.cal_start = DateEntry(sec, **cal_opts)
        self.cal_start.grid(row=1, column=0, sticky="ew", padx=(0, 6), ipady=3)
        self.cal_end = DateEntry(sec, **cal_opts)
        self.cal_end.grid(row=1, column=1, sticky="ew", padx=(6, 0), ipady=3)
        self.cal_start.bind("<<DateEntrySelected>>", lambda e: self._sincronizar_fechas(self.cal_start, self.cal_end))
        self.cal_end.bind("<<DateEntrySelected>>", lambda e: self._sincronizar_fechas(self.cal_end, self.cal_start))
        self.cal_start.bind("<FocusOut>", lambda e: self._sincronizar_fechas(self.cal_start, self.cal_end), add="+")
        self.cal_end.bind("<FocusOut>", lambda e: self._sincronizar_fechas(self.cal_end, self.cal_start), add="+")
        quick = ctk.CTkFrame(sec, fg_color="transparent")
        quick.grid(row=2, column=0, columnspan=2, pady=(6, 0), sticky="w")
        for txt, cmd in (("Hoy", self._set_today), ("Día hábil anterior", self._set_prev_business_day)):
            self._btn(quick, txt, cmd, "link", height=26).pack(side="left", padx=(0, 4), ipadx=2)
        ctk.CTkLabel(
            sec, text="Elige el periodo y pulsa «Extraer y validar». En «Lote del día» las dos fechas se mueven juntas.",
            font=F(11), text_color=COL["faint"], justify="left", anchor="w", wraplength=250
        ).grid(row=3, column=0, columnspan=2, sticky="w", pady=(4, 0))

        # Corte
        sec = self._section(side, "Envío del ticket único")
        self.var_aplicar_corte = tk.BooleanVar(value=self.config_data.get("aplicar_corte", True))
        ctk.CTkSwitch(
            sec, text="Sumar rezagadas del día anterior", variable=self.var_aplicar_corte,
            onvalue=True, offvalue=False, font=F(12), switch_width=36, switch_height=18
        ).grid(row=0, column=0, columnspan=2, sticky="w")
        self.ent_asunto_conc = self._field(
            sec, 1, "Asunto del correo concentrado", self.config_data.get("asunto_concentrado", ASUNTO_CONCENTRADO_DEFAULT)
        )
        self.ent_hora_corte = self._field(sec, 3, "Hora habitual de envío", self.config_data.get("hora_corte", HORA_CORTE_DEFAULT),
                                          "17:50", width=90)
        self.ent_hora_corte.bind("<FocusOut>", lambda e: self._tick_clock(reschedule=False))
        ctk.CTkLabel(
            sec, text="El corte de cada día es la hora en que salió su ticket único. La hora habitual solo se usa "
                      "para recordarte y en días sin ticket único encontrado.",
            font=F(11), text_color=COL["muted"], justify="left", anchor="w", wraplength=230
        ).grid(row=5, column=0, columnspan=2, sticky="w", pady=(6, 0))
        self.ent_inhabiles = self._field(sec, 6, "Días inhábiles (además de fines de semana)",
                                         self.config_data.get("dias_inhabiles", ""), "Ej. 02/11/2026, 12/12/2026")
        self.var_oficiales = tk.BooleanVar(value=self.config_data.get("feriados_oficiales", True))
        ctk.CTkCheckBox(
            sec, text="Incluir feriados oficiales", variable=self.var_oficiales,
            font=F(12), checkbox_width=18, checkbox_height=18
        ).grid(row=8, column=0, columnspan=2, pady=(8, 0), sticky="w")
        self._aplicar_inhabiles()

        # Validaciones automáticas: capturas de pantalla y base de datos
        sec = self._section(side, "Validaciones automáticas")
        self.var_ocr = tk.BooleanVar(value=self.config_data.get("leer_imagenes", True))
        ctk.CTkCheckBox(
            sec, text="Leer capturas de pantalla (OCR)", variable=self.var_ocr,
            font=F(12), checkbox_width=18, checkbox_height=18
        ).grid(row=0, column=0, columnspan=2, sticky="w")
        ctk.CTkLabel(
            sec, text="Si un incidente manda una imagen en vez de escribir Cliente y Crédito, se leen de la imagen "
                      "y se marcan para que los revises.",
            font=F(11), text_color=COL["muted"], justify="left", anchor="w", wraplength=230
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 0))
        self.var_sql = tk.BooleanVar(value=self.config_data.get("sql_activo", False))
        ctk.CTkSwitch(
            sec, text="Validar en la base de datos", variable=self.var_sql, onvalue=True, offvalue=False,
            font=F(12), switch_width=36, switch_height=18
        ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(14, 0))
        # Como en SQL Developer (tipo de conexión «Básico»): Nombre del host, Puerto y Nombre del servicio
        host, puerto, servicio = (self.config_data.get("sql_host", ""), self.config_data.get("sql_puerto", ""),
                                  self.config_data.get("sql_servicio", ""))
        if not host and self.config_data.get("sql_servidor"):
            host, puerto, servicio = separar_servidor(self.config_data["sql_servidor"])
        if host.strip().isdigit():
            # La v2.6 pudo guardar el puerto como host («1521»): así no conecta (getaddrinfo failed)
            host, puerto = "", puerto or host.strip()
        self.ent_sql_host = self._field(sec, 3, "Nombre del host", host, "servidor.fonacot.gob.mx")
        self.ent_sql_puerto = self._field(sec, 5, "Puerto", puerto or "1521", "1521", width=90)
        self.ent_sql_servicio = self._field(sec, 7, "Nombre del servicio", servicio, "Como en SQL Developer")
        self.ent_sql_usuario = self._field(sec, 9, "Usuario", self.config_data.get("sql_usuario", ""), "Usuario de la base")
        self.ent_sql_password = self._field(sec, 11, "Contraseña", "", "No se guarda en la configuración")
        self.ent_sql_password.configure(show="•")
        self.var_sql_recordar = tk.BooleanVar(value=self.config_data.get("sql_recordar", False))
        if self.var_sql_recordar.get():
            guardada = (leer_password(self.config_data.get("sql_usuario", ""), self._sql_servidor())
                        or leer_password(self.config_data.get("sql_usuario", ""), self.config_data.get("sql_servidor", "")))
            if guardada:
                self.ent_sql_password.insert(0, guardada)
        ctk.CTkCheckBox(
            sec, text="Recordar contraseña en Windows", variable=self.var_sql_recordar,
            font=F(12), checkbox_width=18, checkbox_height=18
        ).grid(row=13, column=0, columnspan=2, pady=(8, 0), sticky="w")
        self.ent_sql_tabla = self._field(sec, 14, "Tabla", self.config_data.get("sql_tabla") or TABLA_DEFAULT, TABLA_DEFAULT)
        self._btn(sec, "Probar conexión", self._probar_sql, "secondary", height=28, font=F(12)).grid(
            row=16, column=0, sticky="w", pady=(10, 0))
        self.lbl_sql_estado = ctk.CTkLabel(
            sec, text="Solo consulta (SELECT): CR_RENOVACIONES y, para las de otro día, SALDOS, CR_DISPOSICIONES y "
                      "CR_CEDULAS_DETALLE.", font=F(11), text_color=COL["muted"],
            justify="left", anchor="w", wraplength=230
        )
        self.lbl_sql_estado.grid(row=17, column=0, columnspan=2, sticky="w", pady=(6, 0))

        # Correo del lote
        sec = self._section(side, "Correo del lote")
        self.ent_asunto_borrador = self._field(
            sec, 0, "Asunto", self.config_data.get("asunto_borrador", ASUNTO_BORRADOR_DEFAULT), ASUNTO_BORRADOR_DEFAULT
        )
        self.ent_correo_para = self._field(sec, 2, "Para", self.config_data.get("correo_para", ""), "correo@fonacot.gob.mx")
        self.ent_correo_cc = self._field(sec, 4, "CC", self.config_data.get("correo_cc", ""), "Opcional")
        ctk.CTkLabel(
            sec, text="{fecha} se cambia por dd-mm-aaaa. El texto del correo se toma del último concentrado.",
            font=F(11), text_color=COL["muted"], justify="left", anchor="w", wraplength=230
        ).grid(row=6, column=0, columnspan=2, sticky="w", pady=(6, 0))

        # Auto-actualizar
        sec = self._section(side, "Actualización automática")
        self.var_auto = tk.BooleanVar(value=self.config_data.get("auto_refresh", False))
        row = ctk.CTkFrame(sec, fg_color="transparent")
        row.grid(row=0, column=0, columnspan=2, sticky="w")
        ctk.CTkSwitch(
            row, text="Escanear cada", variable=self.var_auto, onvalue=True, offvalue=False,
            command=self._on_auto_toggle, font=F(12), switch_width=36, switch_height=18
        ).pack(side="left")
        self.cmb_auto_min = ctk.CTkComboBox(
            row, values=["5", "10", "15", "30", "60"], width=66, height=28, font=F(12), dropdown_font=F(12),
            command=lambda _: self._on_auto_toggle()
        )
        self.cmb_auto_min.set(str(self.config_data.get("auto_refresh_min", "10")))
        self.cmb_auto_min.pack(side="left", padx=(4, 6))
        ctk.CTkLabel(row, text="min", font=F(12), text_color=COL["text_2"]).pack(side="left")
        self.lbl_auto_next = ctk.CTkLabel(
            sec, text="Desactivada", font=F(11), text_color=COL["muted"], anchor="w", justify="left", wraplength=250
        )
        self.lbl_auto_next.grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))

        self._hline(side).pack(fill="x", padx=12, pady=(18, 8))
        foot = ctk.CTkFrame(side, fg_color="transparent")
        foot.pack(fill="x", padx=6, pady=(0, 12))
        self._btn(foot, "Atajos de teclado", self._show_shortcuts, "ghost", height=28, font=F(12)).pack(side="left")
        ctk.CTkLabel(foot, text=f"Versión {APP_VERSION}", font=F(11), text_color=COL["faint"]).pack(side="right", padx=6)

    def _build_page_header(self, parent):
        hdr = ctk.CTkFrame(parent, fg_color="transparent")
        hdr.grid(row=0, column=0, sticky="ew")
        hdr.grid_columnconfigure(0, weight=1)

        self.lbl_page_title = ctk.CTkLabel(hdr, text="Lote del día", font=F(22, True), text_color=COL["text"], anchor="w")
        self.lbl_page_title.grid(row=0, column=0, sticky="w")
        self.lbl_lote_info = ctk.CTkLabel(
            hdr, text=self._texto_marca_guardada(), font=F(12), text_color=COL["muted"],
            justify="left", anchor="w", wraplength=800
        )
        self.lbl_lote_info.grid(row=1, column=0, sticky="ew", pady=(2, 0))

        actions = ctk.CTkFrame(hdr, fg_color="transparent")
        actions.grid(row=0, column=1, rowspan=2, sticky="ne", padx=(16, 0))
        self._btn(actions, "Copiar lote", self._copy_batch_to_clipboard).pack(side="left", padx=(0, 8), ipadx=8)
        self._btn(actions, "Exportar a Excel", self._export_excel).pack(side="left", padx=(0, 8), ipadx=8)
        self._btn(actions, "Crear correo Ticket Único", self._crear_correo_lote, "primary").pack(side="left", ipadx=8)

        # e.width llega en píxeles reales; wraplength de CTk va en unidades lógicas
        hdr.bind(
            "<Configure>",
            lambda e: self.lbl_lote_info.configure(wraplength=max(300, int(e.width / self._scaling()) - 460))
        )

    def _build_notice(self, parent):
        """
        Aviso bajo el título: rojo si hay renovaciones que no se enviaron en su día (clic = verlas),
        ámbar si no se encontró el correo concentrado del día anterior o hay días sin comprobar.
        """
        self.notice = ctk.CTkFrame(parent, fg_color=COL["warning_soft"], corner_radius=8, border_width=1, border_color="#FDE68A")
        self.notice.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        self.notice_bar = ctk.CTkFrame(self.notice, width=3, height=1, corner_radius=2, fg_color="#F59E0B")
        self.notice_bar.pack(side="left", fill="y", padx=(10, 0), pady=10)
        self.lbl_notice = ctk.CTkLabel(
            self.notice, text="", font=F(12), text_color="#92400E", justify="left", anchor="w", wraplength=900
        )
        self.lbl_notice.pack(side="left", fill="x", expand=True, padx=12, pady=9)
        self.notice.bind(
            "<Configure>",
            lambda e: self.lbl_notice.configure(wraplength=max(300, int(e.width / self._scaling()) - 60))
        )
        self._notice_filtro = None
        for w in (self.notice, self.lbl_notice):
            w.bind("<Button-1>", lambda e: self._notice_filtro and self._apply_filter(self._notice_filtro))
        self.notice.grid_remove()

    def _mostrar_aviso(self, texto: str, rojo: bool = False, filtro: str = None):
        if rojo:
            self.notice.configure(fg_color=COL["danger_soft"], border_color="#FECACA")
            self.notice_bar.configure(fg_color="#DC2626")
            self.lbl_notice.configure(text_color="#991B1B")
        else:
            self.notice.configure(fg_color=COL["warning_soft"], border_color="#FDE68A")
            self.notice_bar.configure(fg_color="#F59E0B")
            self.lbl_notice.configure(text_color="#92400E")
        self._notice_filtro = filtro
        cursor = "hand2" if filtro else ""
        self.notice.configure(cursor=cursor)
        self.lbl_notice.configure(text=texto, cursor=cursor)
        self.notice.grid()

    def _build_kpis(self, parent):
        self.kpi_bar = kpi_bar = ctk.CTkFrame(parent, fg_color="transparent")
        kpi_bar.grid(row=2, column=0, sticky="ew", pady=(16, 16))
        for i in range(6):
            kpi_bar.grid_columnconfigure(i, weight=1, uniform="kpi")

        self.lbl_kpi_tickets = self._kpi(kpi_bar, 0, "Tickets", "Incidentes únicos en iTop", "#0B5CAD", FILTRO_TODAS)
        self.lbl_kpi_renovaciones = self._kpi(kpi_bar, 1, "Créditos a cancelar", "Renovaciones desglosadas", "#0F766E", FILTRO_TODAS)
        self.lbl_kpi_multiples = self._kpi(kpi_bar, 2, "Tickets múltiples", "2 o más créditos", "#D97706", FILTRO_MULTIPLES)
        self.lbl_kpi_alertas = self._kpi(kpi_bar, 3, "Alertas", "Falta dato, duplicado o cerrado", "#DC2626", FILTRO_ALERTAS)
        self.lbl_kpi_rezagadas = self._kpi(kpi_bar, 4, "Rezagadas", "No enviadas antes, sumadas hoy", "#4F46E5", FILTRO_REZAGADAS)
        self.lbl_kpi_validados = self._kpi(kpi_bar, 5, "Avance", "Clic para ver pendientes", "#16A34A", FILTRO_PENDIENTES,
                                           progreso=True)

    def _kpi(self, parent, col, titulo, subtitulo, acento, filtro, progreso=False):
        card = ctk.CTkFrame(parent, fg_color=COL["surface"], corner_radius=10, border_width=1, border_color=COL["border"])
        card.grid(row=0, column=col, padx=(0, 0 if col == 5 else 12), sticky="nsew")
        head = ctk.CTkFrame(card, fg_color="transparent")
        head.pack(fill="x", padx=16, pady=(14, 0))
        dot = ctk.CTkFrame(head, width=8, height=8, corner_radius=4, fg_color=acento)
        dot.pack(side="left")
        lbl_t = ctk.CTkLabel(head, text=titulo.upper(), font=F(10, True), text_color=COL["muted"], height=14)
        lbl_t.pack(side="left", padx=(8, 0))
        lbl_v = ctk.CTkLabel(card, text="0", font=F(24, True), text_color=COL["text"], anchor="w", height=34)
        lbl_v.pack(fill="x", padx=16, pady=(4, 0))
        widgets = [card, head, dot, lbl_t, lbl_v]
        if progreso:
            self.kpi_progress = ctk.CTkProgressBar(card, height=4, corner_radius=2, progress_color=acento)
            self.kpi_progress.pack(fill="x", padx=16, pady=(4, 2))
            self.kpi_progress.set(0)
            widgets.append(self.kpi_progress)
        lbl_s = ctk.CTkLabel(card, text=subtitulo, font=F(11), text_color=COL["faint"], anchor="w", height=16)
        lbl_s.pack(fill="x", padx=16, pady=(0, 14))
        widgets.append(lbl_s)

        def hover(on, _e=None):
            if not on:
                w = self.winfo_containing(*self.winfo_pointerxy())
                if w is not None and str(w).startswith(str(card)):
                    return
            card.configure(border_color=COL["primary"] if on else COL["border"])

        card.bind("<Enter>", lambda e: hover(True))
        for w in widgets:
            w.bind("<Button-1>", lambda e, f=filtro: self._apply_filter(f))
            w.bind("<Leave>", lambda e: hover(False))
        return lbl_v

    def _build_work_area(self, parent):
        paned = ttk.PanedWindow(parent, orient=tk.VERTICAL)
        paned.grid(row=3, column=0, sticky="nsew")

        top = ctk.CTkFrame(paned, fg_color=COL["surface"], corner_radius=10, border_width=1, border_color=COL["border"])
        bottom = ctk.CTkFrame(paned, fg_color=COL["surface"], corner_radius=10, border_width=1, border_color=COL["border"])
        paned.add(top, weight=5)
        paned.add(bottom, weight=2)
        self.paned, self.detail_frame = paned, bottom

        top.grid_columnconfigure(0, weight=1)
        top.grid_rowconfigure(3, weight=1)

        # Encabezado: pestañas a la izquierda, contador y buscador a la derecha
        head = ctk.CTkFrame(top, fg_color="transparent")
        head.grid(row=0, column=0, sticky="ew", padx=(12, 16), pady=(6, 0))
        head.grid_columnconfigure(1, weight=1)
        tabs_box = ctk.CTkFrame(head, fg_color="transparent")
        tabs_box.grid(row=0, column=0, sticky="sw")

        self.lbl_count = ctk.CTkLabel(head, text="", font=F(12), text_color=COL["muted"])
        self.lbl_count.grid(row=0, column=2, padx=(0, 12))
        search = ctk.CTkFrame(head, fg_color=COL["surface"], corner_radius=6, border_width=1, border_color=COL["border_strong"])
        search.grid(row=0, column=3, pady=6)
        ctk.CTkLabel(search, text="", font=ctk.CTkFont(family=FONT_ICON, size=12), text_color=COL["muted"], width=16).pack(
            side="left", padx=(10, 2)
        )
        self.ent_search = ctk.CTkEntry(
            search, width=250, height=28, border_width=0, font=F(12),
            placeholder_text="Buscar ticket, cliente, crédito o agente"
        )
        self.ent_search.pack(side="left", padx=(0, 4), pady=2)
        self.ent_search.bind("<KeyRelease>", self._on_search_key)
        self.ent_search.bind("<Escape>", lambda e: self._clear_search())

        self._hline(top).grid(row=1, column=0, sticky="ew", padx=1)

        # Barra de filtros y acciones sobre filas
        bar = ctk.CTkFrame(top, fg_color="transparent")
        bar.grid(row=2, column=0, sticky="ew", padx=16, pady=12)
        self._btn(bar, "Restaurar", self._restore_selected_rows, "ghost").pack(side="right", ipadx=6)
        self._btn(bar, "Descartar", self._discard_selected_rows, "danger").pack(side="right", padx=(8, 4), ipadx=6)
        self._btn(bar, "Validar visibles", self._validate_all_visible).pack(side="right", padx=(8, 0), ipadx=6)
        self._btn(bar, "Validar", self._toggle_selected_validation, "primary").pack(side="right", padx=(8, 0), ipadx=10)

        seg = ctk.CTkSegmentedButton(bar, values=FILTROS_PRINCIPALES, font=F(12), height=32)
        seg.pack(side="left")
        menu = ctk.CTkOptionMenu(
            bar, values=[OTRAS_VISTAS] + FILTROS_SECUNDARIOS, width=190, height=32, font=F(12),
            dropdown_font=F(12), dynamic_resizing=False
        )
        menu.pack(side="left", padx=(10, 0))
        self.cmb_quick_filter = SelectorFiltro(seg, menu, self._refresh_views)
        self.btn_columnas = self._btn(bar, "Vista  ▾", self._abrir_menu_columnas)
        self.btn_columnas.pack(side="left", padx=(8, 0), ipadx=6)

        body = ctk.CTkFrame(top, fg_color="transparent", corner_radius=0)
        body.grid(row=3, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.tabs = PestanasSubrayadas(tabs_box, body)
        self.tabs_names = {"desglose": "desglose", "tickets": "tickets", "agentes": "agentes", "enviado": "enviado",
                           "dias": "dias"}
        self.tab_desglose = self.tabs.add("desglose", "Desglose")
        self.tab_tickets = self.tabs.add("tickets", "Por ticket")
        self.tab_agentes = self.tabs.add("agentes", "Por compañero")
        self.tab_enviado = self.tabs.add("enviado", "Tickets únicos")
        self.tab_dias = self.tabs.add("dias", "Por día")

        self._build_tab_desglose()
        self._build_tab_tickets()
        self._build_tab_agentes()
        self._build_tab_enviado()
        self._build_tab_dias()
        self._build_detail_panel(bottom)
        for tree in (self.tree_desglose, self.tree_tickets):
            self._aplicar_columnas(tree)
            TooltipCeldas(tree, self._celdas_completas)
        self._aplicar_vista()

    def _build_status_bar(self):
        bar = ctk.CTkFrame(self, fg_color=COL["surface"], corner_radius=0)
        bar.grid(row=2, column=0, sticky="ew")
        bar.grid_columnconfigure(0, weight=1)
        self._hline(bar).grid(row=0, column=0, columnspan=2, sticky="ew")
        self.lbl_status = ctk.CTkLabel(
            bar, text="Listo. Elige el periodo y presiona «Extraer y validar» (F5).",
            font=F(11), text_color=COL["muted"], anchor="w", height=18
        )
        self.lbl_status.grid(row=1, column=0, padx=20, pady=6, sticky="ew")
        self.progress_bar = ctk.CTkProgressBar(bar, width=180, height=4, corner_radius=2)
        self.progress_bar.grid(row=1, column=1, padx=20, pady=6, sticky="e")
        self.progress_bar.set(0)

    def _setup_styles(self):
        s = self._scaling()
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "Renovaciones.Treeview",
            background=COL["surface"], fieldbackground=COL["surface"], foreground=COL["text"],
            rowheight=int(32 * s), borderwidth=0, relief="flat", font=(FONT, 10)
        )
        # Sin marco ni anillo de foco alrededor de la tabla
        style.layout("Renovaciones.Treeview", [("Treeview.treearea", {"sticky": "nswe"})])
        style.configure(
            "Renovaciones.Treeview.Heading",
            background=COL["subtle"], foreground=COL["muted"], font=(FONT_SB, 9), relief="flat",
            borderwidth=1, bordercolor=COL["border"], lightcolor=COL["subtle"], darkcolor=COL["border"],
            padding=(10, int(7 * s))
        )
        style.map(
            "Renovaciones.Treeview.Heading",
            background=[("active", COL["track"])], foreground=[("active", COL["text"])]
        )
        style.map(
            "Renovaciones.Treeview",
            background=[("selected", COL["select"])], foreground=[("selected", COL["text"])]
        )
        style.configure("TPanedwindow", background=COL["bg"])
        style.configure(
            "Sash", sashthickness=int(14 * s), gripcount=0,
            background=COL["bg"], lightcolor=COL["bg"], darkcolor=COL["bg"], bordercolor=COL["bg"]
        )
        # DateEntry copia el estilo de TCombobox
        style.configure(
            "TCombobox", fieldbackground=COL["surface"], background=COL["surface"], foreground=COL["text"],
            bordercolor=COL["border_strong"], lightcolor=COL["surface"], darkcolor=COL["surface"],
            arrowcolor=COL["muted"], padding=(6, 2)
        )
        style.map(
            "TCombobox",
            bordercolor=[("focus", COL["primary"])], lightcolor=[("focus", COL["surface"])],
            fieldbackground=[("readonly", COL["surface"])], background=[("active", COL["track"])]
        )

    def _make_tree(self, parent, headers: dict, selectmode="extended") -> ttk.Treeview:
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)
        tree = ttk.Treeview(
            parent, columns=tuple(headers.keys()), show="headings", style="Renovaciones.Treeview", selectmode=selectmode
        )
        for col, (text, width, anchor) in headers.items():
            self._headers[(str(tree), col)] = text
            tree.heading(col, text=text, anchor=anchor, command=lambda c=col, t=tree: self._on_heading_click(t, c))
            # Los anchos de ttk no siguen el zoom de Windows (la fuente sí): se escalan para que el texto quepa
            tree.column(col, width=int(width * self._scaling()), minwidth=50, anchor=anchor, stretch=False)

        # Colores por estado, suaves (los de alerta ganan a la franja alterna porque cada fila lleva un solo tag)
        tree.tag_configure("normal", background="#FFFFFF", foreground=COL["text"])
        tree.tag_configure("normal_alt", background="#FAFBFC", foreground=COL["text"])
        tree.tag_configure("multiple", background="#FFFBEB", foreground="#92400E")      # Ámbar: varias renovaciones en 1 ticket
        tree.tag_configure("incompleto", background="#FEF2F2", foreground="#B91C1C")    # Rojo: falta cliente o crédito
        tree.tag_configure("duplicado", background="#FDF2F8", foreground="#9D174D")     # Rosa: crédito repetido en otro ticket
        tree.tag_configure("reasignado", background="#FFF7ED", foreground="#9A3412")    # Naranja: reasignado
        tree.tag_configure("rezagada", background="#EEF2FF", foreground="#3730A3")      # Índigo: rezagada del día anterior
        tree.tag_configure("fuera_lote", background="#FFFFFF", foreground=COL["faint"])  # Gris: ya enviada / después del corte
        tree.tag_configure("validado", background="#F0FDF4", foreground="#166534")      # Verde: validada
        tree.tag_configure("cerrado", background="#F5F3FF", foreground="#5B21B6")       # Violeta: ya cerrado en iTop
        tree.tag_configure("descartado", background=COL["subtle"], foreground=COL["faint"])  # Gris: descartada
        tree.tag_configure("no_enviada", background="#FEE2E2", foreground="#991B1B")    # Rojo fuerte: no se envió en su día

        vsb = ctk.CTkScrollbar(parent, orientation="vertical", command=tree.yview, width=12)
        hsb = ctk.CTkScrollbar(parent, orientation="horizontal", command=tree.xview, height=12)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        return tree

    def _build_tab_desglose(self):
        self.tree_desglose = self._make_tree(self.tab_desglose, {
            "validado": ("Estado", 130, "w"),
            "ticket": ("Ticket iTop", 105, "center"),
            "parte": ("Parte", 64, "center"),
            "cliente": ("No. cliente", 115, "center"),
            "credito_hijo": ("Crédito a cancelar", 135, "center"),
            "credito_padre": ("Crédito origen", 120, "center"),
            "sql": ("Base de datos", 130, "w"),
            "lote": ("Lote / envío", 300, "w"),
            "cierre": ("Cierre en iTop", 210, "w"),
            "gafete": ("Gafete", 70, "center"),
            "agente": ("Agente responsable", 190, "w"),
            "alerta_reno": ("Alerta de renovación", 290, "w"),
            "alerta_asig": ("Asignación", 240, "w"),
            "motivo": ("Motivo", 230, "w"),
            "fecha": ("Llegada", 135, "center"),
        })
        self.tree_desglose.bind("<<TreeviewSelect>>", self._on_select_desglose_row)
        self.tree_desglose.bind("<Double-1>", lambda e: self._toggle_selected_validation())
        self.tree_desglose.bind("<space>", lambda e: self._toggle_selected_validation())
        self.tree_desglose.bind("<Return>", lambda e: self._open_selected_in_itop())
        self.tree_desglose.bind("<Delete>", lambda e: self._discard_selected_rows())
        self.tree_desglose.bind("<Control-c>", lambda e: self._copy_selected_rows())
        # Estado vacío sobre la tabla
        self.lbl_empty = ctk.CTkLabel(self.tab_desglose, text="", font=F(13), text_color=COL["muted"], fg_color=COL["surface"])

    def _build_tab_tickets(self):
        self.tree_tickets = self._make_tree(self.tab_tickets, {
            "ticket": ("Ticket iTop", 105, "center"),
            "cant_renovaciones": ("Renovaciones", 110, "center"),
            "clientes": ("Cliente(s)", 170, "w"),
            "creditos_hijos": ("Crédito(s) a cancelar", 220, "w"),
            "creditos_padres": ("Crédito(s) origen", 150, "w"),
            "gafete": ("Gafete", 70, "center"),
            "agente": ("Agente responsable", 200, "w"),
            "resumen": ("Resumen del ticket", 320, "w"),
            "alerta_asig": ("Asignación", 240, "w"),
            "cierre": ("Cierre en iTop", 210, "w"),
            "fecha": ("Llegada", 135, "center"),
        }, selectmode="browse")
        self.tree_tickets.bind("<<TreeviewSelect>>", self._on_select_ticket_row)
        self.tree_tickets.bind("<Double-1>", lambda e: self._open_selected_in_itop())
        self.tree_tickets.bind("<Return>", lambda e: self._open_selected_in_itop())

    def _build_tab_agentes(self):
        self.tree_agentes = self._make_tree(self.tab_agentes, {
            "gafete": ("Gafete", 90, "center"),
            "agente": ("Compañero", 280, "w"),
            "tickets_unicos": ("Tickets", 110, "center"),
            "total_renovaciones": ("Créditos a cancelar", 160, "center"),
            "tickets_multiples": ("Tickets múltiples", 140, "center"),
            "con_alerta": ("Con alerta", 110, "center"),
            "rezagadas": ("Rezagadas", 110, "center"),
            "validados": ("Validadas", 130, "center"),
            "pendientes": ("Pendientes", 110, "center"),
        }, selectmode="browse")

    def _build_tab_enviado(self):
        """Todo lo que salió en los tickets únicos del periodo, renglón por renglón, contra los incidentes de Outlook."""
        self.tree_enviado = self._make_tree(self.tab_enviado, {
            "fecha": ("Ticket único del", 120, "center"),
            "ticket": ("Incidente enviado", 130, "center"),
            "cliente": ("No. cliente", 115, "center"),
            "credito": ("Crédito", 110, "center"),
            "incidente": ("Incidente en Outlook", 150, "center"),
            "revision": ("Revisión", 620, "w"),
        }, selectmode="browse")
        self.tree_enviado.bind("<Double-1>", lambda e: self._abrir_enviado_en_itop())
        self.tree_enviado.bind("<Return>", lambda e: self._abrir_enviado_en_itop())
        TooltipCeldas(self.tree_enviado, self._celdas_completas)

    def _build_tab_dias(self):
        """Revisión día por día: qué ticket único salió y cómo quedó lo que debía ir en él."""
        self.tree_dias = self._make_tree(self.tab_dias, {
            "dia": ("Día", 120, "w"),
            "envios": ("Ticket único enviado", 330, "w"),
            "debian": ("Debían ir", 90, "center"),
            "a_tiempo": ("A tiempo", 90, "center"),
            "tarde": ("Tarde", 80, "center"),
            "no_enviadas": ("NO enviadas", 110, "center"),
            "errores": ("Datos equivocados", 140, "center"),
            "sin_comprobar": ("Sin comprobar", 120, "center"),
            "pendientes": ("Pendientes", 100, "center"),
        }, selectmode="browse")
        self.tree_dias.bind("<Double-1>", lambda e: self._ver_dia())
        TooltipCeldas(self.tree_dias, self._celdas_completas)

    def _ver_dia(self):
        """Doble clic en un día: abre su «Lote del día» tal como debía salir."""
        sel = self.tree_dias.selection()
        if not sel:
            return
        dia = datetime.strptime(sel[0].split("_", 1)[1], "%Y%m%d")
        self.seg_tipo.set(TIPOS_REVISION[0])
        self.cal_start.set_date(dia)
        self.cal_end.set_date(dia)
        self._start_extraction()

    def _llenar_dias(self) -> int:
        self.tree_dias.delete(*self.tree_dias.get_children())
        nombres = ("Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom")
        dias = (self.lote_info or {}).get('por_dia') or []
        for d in dias:
            envios = "  ·  ".join(f"{rec:%H:%M} {rem} ({n} renglones)" for rec, rem, n in d['envios'])
            if not envios:
                envios = "Día inhábil" if not d['habil'] else ("No se encontró ticket único" if d['debian'] else "Sin ticket único")
            iid = f"DIA_{d['dia']:%Y%m%d}"
            self._celdas_completas[(iid, "envios")] = envios
            malos = d['no_enviadas'] or d['errores']
            tag = ("no_enviada" if malos else "multiple" if (d['sin_comprobar'] or (d['debian'] and not d['envios']))
                   else "descartado" if not d['habil'] else "validado" if d['debian'] else "normal")
            self.tree_dias.insert("", "end", iid=iid, tags=(tag,), values=(
                f"{nombres[d['dia'].weekday()]} {d['dia']:%d/%m}", envios, d['debian'], d['a_tiempo'], d['tarde'],
                d['no_enviadas'], d['errores'], d['sin_comprobar'], d['pendientes']))
        return len(dias)

    def _abrir_enviado_en_itop(self):
        sel = self.tree_enviado.selection()
        if sel:
            ticket = self.tree_enviado.set(sel[0], "incidente") or self.tree_enviado.set(sel[0], "ticket")
            if ticket.upper().startswith("I-"):
                webbrowser.open(self._itop_url(ticket))

    @staticmethod
    def _revision_enviado(f: dict) -> tuple:
        """(texto, tag) de un renglón del concentrado semanal (con lo que dice la base de datos si no cuadra)."""
        texto, tag = FonacotValidadorRenovacionesApp._revision_outlook(f)
        if f.get('sql_estado') in SQL_PROBLEMAS:
            return f"{texto}  ·  BASE DE DATOS: {f['sql_detalle']}", "no_enviada"
        return texto, tag

    @staticmethod
    def _revision_outlook(f: dict) -> tuple:
        if f.get('error'):
            return f"DATOS EQUIVOCADOS: {f['error']}", "no_enviada"
        if not f.get('incidente'):
            return "No se encontró el correo de este incidente en Outlook (revisa el buzón o el periodo)", "multiple"
        if f.get('duplicada'):
            return f"Bien (solicitud duplicada: también la pidió {', '.join(f['duplicada'])})", "reasignado"
        if f.get('veces', 1) > 1:
            return f"Bien, pero salió en {f['veces']} tickets únicos (repetida)", "reasignado"
        if not f.get('en_periodo'):
            return "Bien (se mandó después del periodo revisado)", "normal"
        return "Bien", "normal"

    def _llenar_enviado(self, search_txt: str) -> int:
        self.tree_enviado.delete(*self.tree_enviado.get_children())
        semana = (self.lote_info or {}).get('semana') or {}
        n = 0
        for i, f in enumerate(semana.get('filas') or []):
            texto, tag = self._revision_enviado(f)
            valores = (f"{f['fecha']:%d/%m/%Y}", f['ticket'] or "-", f['cliente'], f['credito'], f['incidente'] or "-", texto)
            if search_txt and search_txt not in " ".join(valores).lower():
                continue
            iid = f"ENV_{i}"
            self._celdas_completas[(iid, "revision")] = texto
            self.tree_enviado.insert("", "end", iid=iid, values=valores,
                                     tags=(tag if tag != "normal" else ("normal_alt" if n % 2 else "normal"),))
            n += 1
        self._reapply_sort(self.tree_enviado)
        return n

    def _build_detail_panel(self, parent):
        parent.grid_columnconfigure(0, weight=2)
        parent.grid_columnconfigure(1, weight=3)
        parent.grid_rowconfigure(2, weight=1)

        hdr = ctk.CTkFrame(parent, fg_color="transparent")
        hdr.grid(row=0, column=0, columnspan=2, padx=16, pady=(12, 6), sticky="ew")
        self.lbl_detail_header = ctk.CTkLabel(
            hdr, text="Selecciona una renovación para ver el detalle del ticket y el correo original.",
            font=F(13, True), text_color=COL["text"], anchor="w"
        )
        self.lbl_detail_header.pack(side="left", fill="x", expand=True)
        self._btn(
            hdr, "Copiar cliente y crédito", lambda: self._copy_field(('Numero_Cliente', 'Credito_Renovado_Hijo')), height=30
        ).pack(side="right", ipadx=6)
        self._btn(hdr, "Abrir en iTop", self._open_selected_in_itop, height=30).pack(side="right", padx=(0, 8), ipadx=6)

        for col, txt in ((0, "DESGLOSE DEL TICKET"), (1, "CORREO ORIGINAL  ·  números detectados resaltados")):
            ctk.CTkLabel(parent, text=txt, font=F(10, True), text_color=COL["muted"], anchor="w", height=14).grid(
                row=1, column=col, sticky="w", padx=(16, 6) if col == 0 else (6, 16), pady=(0, 6)
            )

        self.txt_detail = ctk.CTkTextbox(parent, font=F(12), fg_color=COL["subtle"], wrap="word")
        self.txt_detail.grid(row=2, column=0, padx=(16, 6), pady=(0, 16), sticky="nsew")
        self.txt_detail.tag_config("titulo", foreground=COL["text"])
        self.txt_detail.tag_config("ok", foreground=COL["success"])
        self.txt_detail.tag_config("warn", foreground=COL["warning"])
        self.txt_detail.tag_config("muted", foreground=COL["muted"])
        self.txt_detail.tag_config("cierre", foreground="#5B21B6")
        self.txt_detail.tag_config("error", foreground="#B91C1C")

        self.txt_body = ctk.CTkTextbox(parent, font=F(12), wrap="word")
        self.txt_body.grid(row=2, column=1, padx=(6, 16), pady=(0, 16), sticky="nsew")
        self.txt_body.tag_config("hl", background="#FEF08A", foreground="#713F12")

    def _build_context_menus(self):
        self._menu_opts = opts = dict(
            tearoff=0, font=(FONT, 10), bg=COL["surface"], fg=COL["text"], activebackground=COL["primary_soft"],
            activeforeground=COL["text"], relief="flat", borderwidth=1
        )
        m = tk.Menu(self, **opts)
        m.add_command(label="Marcar / desmarcar validada", accelerator="Espacio", command=self._toggle_selected_validation)
        m.add_command(label="Abrir en iTop", accelerator="Enter", command=self._open_selected_in_itop)
        m.add_command(label="Buscar el folio en iTop", command=lambda: self._open_selected_in_itop(buscar=True))
        m.add_separator()
        m.add_command(label="Copiar No. cliente", command=lambda: self._copy_field(('Numero_Cliente',)))
        m.add_command(label="Copiar crédito a cancelar", command=lambda: self._copy_field(('Credito_Renovado_Hijo',)))
        m.add_command(label="Copiar ticket", command=lambda: self._copy_field(('Numero_Ticket',)))
        m.add_command(label="Copiar fila(s)", accelerator="Ctrl+C", command=self._copy_selected_rows)
        m.add_separator()
        m.add_command(label="Descartar", accelerator="Supr", command=self._discard_selected_rows)
        m.add_command(label="Restaurar / meter al lote", command=self._restore_selected_rows)
        self.menu_desglose = m

        t = tk.Menu(self, **opts)
        t.add_command(label="Abrir en iTop", accelerator="Enter", command=self._open_selected_in_itop)
        t.add_command(label="Buscar el folio en iTop", command=lambda: self._open_selected_in_itop(buscar=True))
        t.add_command(label="Copiar ticket", command=self._copy_selected_ticket)
        self.menu_tickets = t

        self.tree_desglose.bind("<Button-3>", lambda e: self._popup_menu(e, self.tree_desglose, self.menu_desglose))
        self.tree_tickets.bind("<Button-3>", lambda e: self._popup_menu(e, self.tree_tickets, self.menu_tickets))

    def _bind_shortcuts(self):
        self.bind("<F5>", lambda e: self._start_extraction())
        self.bind("<Control-f>", lambda e: self._focus_search())
        self.bind("<Control-e>", lambda e: self._export_excel())
        self.bind("<Control-m>", lambda e: self._crear_correo_lote())
        self.bind("<Control-C>", lambda e: self._copy_batch_to_clipboard())  # Ctrl+Shift+C

    # ================= EVENTOS GENERALES =================

    def _init_accounts(self):
        accounts = self.engine.get_accounts()
        if accounts:
            self.cmb_account.configure(values=accounts)
            saved = self.config_data.get("account", "")
            self.cmb_account.set(saved if saved in accounts else accounts[0])
        else:
            self.cmb_account.configure(values=["Sin conexión a Outlook"])
            self.cmb_account.set("Sin conexión a Outlook")

    def _on_modo_seg(self, etiqueta: str):
        self.var_modo.set(next((k for k, v in MODOS.items() if v == etiqueta), "CONCENTRADO"))
        self._on_mode_change()

    def _on_mode_change(self):
        if self.var_modo.get() == "CONCENTRADO":
            self.ent_agente.configure(state="disabled", fg_color=COL["subtle"], text_color=COL["faint"])
            self.chk_exclude_reassigned.configure(state="disabled")
        else:
            self.ent_agente.configure(state="normal", fg_color=COL["surface"], text_color=COL["text"])
            self.chk_exclude_reassigned.configure(state="normal")

    def _set_today(self):
        now = datetime.now()
        self.cal_start.set_date(now)
        self.cal_end.set_date(now)
        self._periodo_pendiente()

    def _set_prev_business_day(self):
        prev = previous_business_day(datetime.now().date())
        self.cal_start.set_date(prev)
        self.cal_end.set_date(prev)
        self._periodo_pendiente()

    def _periodo_pendiente(self):
        """Recuerda que lo elegido aún no se ha buscado (cambiar de modo o de fechas ya no busca solo)."""
        if self.is_extracting:
            return
        ini, fin = self.cal_start.get_date(), self.cal_end.get_date()
        if self._es_revision():
            que = f"Revisión del {ini:%d/%m} al {fin:%d/%m}" if ini != fin else f"Revisión del {fin:%d/%m}"
        else:
            que = "Lote de hoy" if fin == datetime.now().date() else f"Lote del {fin:%d/%m}"
        self.lbl_extract_hint.configure(text=f"{que}: pulsa «Extraer y validar» (F5) para buscar",
                                        text_color=COL["primary"])

    def _sql_servidor(self) -> str:
        """«host:puerto/servicio» con lo que está en los campos (o en la configuración, antes de crearlos)."""
        if not hasattr(self, "ent_sql_servicio"):
            return servidor_de(self.config_data.get("sql_host", ""), self.config_data.get("sql_puerto", ""),
                               self.config_data.get("sql_servicio", ""))
        return servidor_de(self.ent_sql_host.get(), self.ent_sql_puerto.get(), self.ent_sql_servicio.get())

    def _sql_conexion(self) -> dict:
        return {
            "servidor": self._sql_servidor(),
            "usuario": self.ent_sql_usuario.get().strip(),
            "password": self.ent_sql_password.get(),
            "tabla": self.ent_sql_tabla.get().strip() or TABLA_DEFAULT,
            "cliente_oracle": self.config_data.get("sql_cliente_oracle", ""),
        }

    def _sql_params(self) -> dict:
        """Datos de conexión si la validación en la base está activa (None si no)."""
        return self._sql_conexion() if self.var_sql.get() else None

    def _probar_sql(self):
        p = self._sql_conexion()
        problema = host_invalido(servidor_de(self.ent_sql_host.get(), self.ent_sql_puerto.get(), self.ent_sql_servicio.get())
                                 or self.ent_sql_puerto.get())
        if problema or not (p["servidor"] and self.ent_sql_servicio.get().strip()):
            self._sql_probado(problema or "Escribe el Nombre del host y el Nombre del servicio (los mismos de SQL Developer).",
                              False)
            return
        self.lbl_sql_estado.configure(text="Probando conexión…", text_color=COL["muted"])

        def trabajo():
            try:
                texto, ok = probar_conexion(**p), True
            except ErrorSQL as e:
                texto, ok = str(e), False
            except Exception as e:
                texto, ok = f"{type(e).__name__}: {e}", False
            self.after(0, lambda: self._sql_probado(texto, ok))
        threading.Thread(target=trabajo, daemon=True).start()

    def _sql_probado(self, texto: str, ok: bool):
        self.lbl_sql_estado.configure(text=texto, text_color=COL["success"] if ok else "#B91C1C")
        if ok:
            self._save_config()

    @staticmethod
    def _validar_sql(rows: list, lote_info: dict, cfg: dict) -> dict:
        """Valida cada renovación (y lo enviado en los tickets únicos) en la base. Un error no detiene la extracción."""
        try:
            conn = conectar_sql(cfg["servidor"], cfg["usuario"], cfg["password"], cfg["cliente_oracle"])
            try:
                filas = ((lote_info or {}).get('semana') or {}).get('filas') or []
                return validar_contra_base(conn, cfg["tabla"], rows, filas)
            finally:
                conn.close()
        except ErrorSQL as e:
            return {'ok': False, 'error': str(e)}
        except Exception as e:
            texto = str(e).strip().splitlines()[0] if str(e).strip() else ""
            return {'ok': False, 'error': f"{type(e).__name__}: {texto}"}

    def _aplicar_inhabiles(self):
        set_dias_inhabiles(parse_dias_inhabiles(self.ent_inhabiles.get()), bool(self.var_oficiales.get()))

    def _es_revision(self) -> bool:
        return self.seg_tipo.get() == TIPOS_REVISION[1]

    def _sincronizar_fechas(self, origen, otro):
        """En «Lote del día» el periodo es un solo día: elegir una fecha mueve la otra."""
        if not self._es_revision():
            otro.set_date(origen.get_date())
        self._periodo_pendiente()

    def _on_tipo_revision(self, tipo: str):
        """Cambiar de modo solo prepara el periodo (las fechas se mueven): no busca hasta «Extraer y validar»."""
        self.seg_tipo.set(tipo)
        if tipo == TIPOS_REVISION[0]:
            # De una semana al lote de hoy; un solo día ya elegido (hoy o un día pasado) se respeta
            dia = self.cal_end.get_date() if self.cal_start.get_date() == self.cal_end.get_date() else datetime.now().date()
            self.cal_start.set_date(dia)
            self.cal_end.set_date(dia)
            self._periodo_pendiente()
        elif self.cal_start.get_date() == self.cal_end.get_date():
            # De un solo día a la semana: la actual, o la pasada si en esta todavía no hay un día hábil antes de hoy
            hoy = datetime.now().date()
            lunes = hoy - timedelta(days=hoy.weekday())
            self._revisar_semana(0 if lunes <= previous_business_day(hoy) else 1)
        else:
            self._periodo_pendiente()

    def _revisar_semana(self, semanas_atras: int = 0):
        """Pone de lunes a viernes de la semana (la actual llega hasta hoy); se busca con «Extraer y validar»."""
        hoy = datetime.now().date()
        lunes = hoy - timedelta(days=hoy.weekday() + 7 * semanas_atras)
        viernes = lunes + timedelta(days=4)
        self.cal_start.set_date(lunes)
        self.cal_end.set_date(min(viernes, hoy) if semanas_atras == 0 else viernes)
        self.seg_tipo.set(TIPOS_REVISION[1])
        self._periodo_pendiente()

    def _focus_search(self):
        self.ent_search.focus_set()
        return "break"

    def _on_search_key(self, _event=None):
        # Pequeña pausa para no redibujar la tabla en cada tecla
        if self._search_job:
            self.after_cancel(self._search_job)
        self._search_job = self.after(180, self._refresh_views)

    def _clear_search(self):
        self.ent_search.delete(0, "end")
        self._refresh_views()

    def _apply_filter(self, filtro: str):
        self.cmb_quick_filter.set(filtro)
        self.tabs.set(self.tabs_names["desglose"])
        self._refresh_views()

    def _show_shortcuts(self):
        win = ctk.CTkToplevel(self)
        win.title("Atajos de teclado")
        win.resizable(False, False)
        win.transient(self)
        ctk.CTkLabel(win, text="Atajos de teclado", font=F(16, True), text_color=COL["text"]).grid(
            row=0, column=0, columnspan=2, sticky="w", padx=24, pady=(20, 12)
        )
        for i, (tecla, accion) in enumerate(ATAJOS, 1):
            ctk.CTkLabel(
                win, text=tecla, font=F(11, True), fg_color=COL["track"], corner_radius=4,
                text_color=COL["text_2"], height=24
            ).grid(row=i, column=0, sticky="w", padx=(24, 16), pady=3, ipadx=8)
            ctk.CTkLabel(win, text=accion, font=F(12), text_color=COL["text"]).grid(row=i, column=1, sticky="w", padx=(0, 24))
        self._btn(win, "Cerrar", win.destroy, "primary").grid(
            row=len(ATAJOS) + 1, column=1, sticky="e", padx=24, pady=(16, 20), ipadx=12
        )
        win.after(150, win.focus_force)

    def _toast(self, msg: str, color: str = COL["primary"], ms: int = 7000):
        """Aviso flotante en la esquina inferior derecha que se cierra solo (o con clic)."""
        try:
            win = ctk.CTkToplevel(self)
            win.overrideredirect(True)
            win.attributes("-topmost", True)
            win.configure(fg_color="#1F2937")
            frame = ctk.CTkFrame(win, fg_color="#1F2937", corner_radius=0)
            frame.pack(fill="both", expand=True)
            strip = ctk.CTkFrame(frame, width=4, corner_radius=2, fg_color=color)
            strip.pack(side="left", fill="y", padx=(12, 0), pady=12)
            lbl = ctk.CTkLabel(
                frame, text=msg, text_color="#F9FAFB", font=F(12), wraplength=360, justify="left", anchor="w"
            )
            lbl.pack(side="left", padx=(12, 18), pady=12)
            win.update_idletasks()
            x = self.winfo_rootx() + self.winfo_width() - win.winfo_width() - 24
            y = self.winfo_rooty() + self.winfo_height() - win.winfo_height() - 48
            win.geometry(f"+{x}+{y}")
            for w in (frame, lbl, strip):
                w.bind("<Button-1>", lambda e: win.destroy())
            win.after(ms, lambda: win.winfo_exists() and win.destroy())
        except Exception:
            self.lbl_status.configure(text=msg)

    # ================= RELOJ, CUENTA REGRESIVA Y AUTO-ACTUALIZACIÓN =================

    def _tick_clock(self, reschedule: bool = True):
        # El corte real es el envío del ticket único de hoy; la hora habitual solo sirve de recordatorio
        now = datetime.now()
        habitual = parse_hora_corte(self.ent_hora_corte.get())
        habitual_dt = datetime.combine(now.date(), habitual)
        enviado = ((self.lote_info or {}).get('cortes') or {}).get(now.date())
        neutro = (COL["track"], COL["text_2"])
        if now.weekday() >= 5:
            text, (bg, fg) = f"{now:%H:%M}  ·  Fin de semana: entra como rezagada el lunes", neutro
        elif enviado:
            text, (bg, fg) = f"{now:%H:%M}  ·  Ticket único de hoy enviado a las {enviado:%H:%M}: lo nuevo va al siguiente", neutro
        elif now < habitual_dt:
            mins = int((habitual_dt - now).total_seconds() // 60)
            h, m = divmod(mins, 60)
            restante = f"{h} h {m:02d} min" if h else f"{m} min"
            text = f"{now:%H:%M}  ·  Ticket único de hoy pendiente  ·  habitual {habitual:%H:%M} (en {restante})"
            if mins > 60:
                bg, fg = COL["success_soft"], "#047857"
            elif mins > 15:
                bg, fg = COL["warning_soft"], COL["warning"]
            else:
                bg, fg = COL["danger_soft"], COL["danger"]
            if mins <= 15 and self._aviso_corte_dia != now.date():
                self._aviso_corte_dia = now.date()
                self._aviso_corte(mins)
        else:
            text = f"{now:%H:%M}  ·  Ticket único de hoy pendiente: lo que llegue aún entra hoy"
            bg, fg = COL["warning_soft"], COL["warning"]
        self.lbl_clock.configure(text=text, fg_color=bg, text_color=fg)
        if reschedule:
            self.after(20000, self._tick_clock)

    def _aviso_corte(self, mins: int):
        active_rows, _ = self._get_active_and_filtered_rows()
        pendientes = sum(1 for r in active_rows if r['Row_ID'] not in self.validated_row_ids)
        self.bell()
        self._toast(
            f"Faltan {mins} min para la hora habitual del ticket único.\n"
            + (f"Hay {pendientes} de {len(active_rows)} renovaciones sin validar." if self.all_rows
               else "Aún no se han extraído las renovaciones de hoy (F5)."),
            color="#EF4444", ms=12000
        )

    def _on_auto_toggle(self):
        self._save_config()
        self._schedule_auto_refresh()

    def _schedule_auto_refresh(self):
        if self._auto_job:
            self.after_cancel(self._auto_job)
            self._auto_job = None
        if not self.var_auto.get():
            self.lbl_auto_next.configure(text="Desactivada")
            return
        try:
            mins = max(1, int(self.cmb_auto_min.get()))
        except ValueError:
            mins = 10
        proximo = datetime.now() + timedelta(minutes=mins)
        self._auto_job = self.after(mins * 60000, self._auto_refresh_tick)
        self.lbl_auto_next.configure(text=f"Próximo escaneo a las {proximo:%H:%M}. Conserva validaciones y descartes.")

    def _auto_refresh_tick(self):
        self._auto_job = None
        if not self.is_extracting:
            self._start_extraction(auto=True)
        self._schedule_auto_refresh()

    # ================= EXTRACCIÓN =================

    def _start_extraction(self, auto: bool = False):
        if self.is_extracting:
            return

        modo = self.var_modo.get()
        agente_filter = self.ent_agente.get().strip()
        if modo == "AGENTE" and not agente_filter:
            if not auto:
                messagebox.showwarning(
                    "Falta gafete o nombre",
                    "En el modo 'Por compañero', escribe el gafete (ej. 2609, 7353) o el nombre del agente."
                )
            return

        self._save_config()
        sql_cfg = self._sql_params()
        if sql_cfg and (host_invalido(sql_cfg["servidor"])
                        or not (sql_cfg["servidor"] and sql_cfg["usuario"] and sql_cfg["password"])):
            if not auto:
                messagebox.showwarning(
                    "Falta la conexión a la base",
                    (host_invalido(sql_cfg["servidor"]) + "\n\n" if host_invalido(sql_cfg["servidor"]) else "")
                    + "Para validar en CR_RENOVACIONES escribe host, puerto, servicio, usuario y contraseña en "
                    "«Validaciones automáticas».\n\n"
                    "Esta vez se extrae sin validar en la base."
                )
            sql_cfg = None

        account = self.cmb_account.get().strip()
        if account in ("Sin conexión a Outlook", "Cargando cuentas..."):
            account = ""

        self._aplicar_inhabiles()
        if auto and self._ultima_busqueda:
            # La actualización automática repite lo último que se buscó, no lo que se está preparando en el panel
            revision, d_start, d_end, era_hoy = self._ultima_busqueda
            hoy = datetime.now().date()
            if era_hoy and not revision and d_end.date() != hoy:
                d_start, d_end = datetime.combine(hoy, datetime.min.time()), datetime.combine(hoy, datetime.max.time())
                self._ultima_busqueda = (revision, d_start, d_end, True)
        else:
            revision = self._es_revision()
            if not revision and self.cal_start.get_date() != self.cal_end.get_date():
                self.cal_start.set_date(self.cal_end.get_date())
            d_start = datetime.combine(self.cal_start.get_date(), datetime.min.time())
            d_end = datetime.combine(self.cal_end.get_date(), datetime.max.time())
            if d_start > d_end:
                d_start, d_end = d_end, d_start
            self._ultima_busqueda = (revision, d_start, d_end, d_end.date() == datetime.now().date())

        signature = (account, d_start.date(), d_end.date(), modo, agente_filter, revision,
                     bool(self.var_aplicar_corte.get()))
        if signature != self._prev_scan_signature:
            self._prev_keys = None
        self._prev_scan_signature = signature

        self.is_extracting = True
        self.seg_tipo.configure(state="disabled")  # que no se cambie de modo a media búsqueda
        self.lbl_extract_hint.configure(text=HINT_EXTRAER, text_color=COL["muted"])
        self.btn_extract.configure(state="disabled", text="Analizando renovaciones…")
        self.progress_bar.set(0.05)

        corte_opts = {
            "aplicar_corte": revision or bool(self.var_aplicar_corte.get()),
            "hora_corte": self.config_data.get("hora_corte", HORA_CORTE_DEFAULT),
            "asunto_concentrado": self.config_data.get("asunto_concentrado", ASUNTO_CONCENTRADO_DEFAULT),
            "ultimo_lote": self.config_data.get("ultimo_lote"),
            "revision": revision,
            "dias_atras": int(self.config_data.get("dias_atras_lote") or DIAS_ATRAS_LOTE),
            "leer_imagenes": bool(self.var_ocr.get())
        }
        self.ent_hora_corte.delete(0, "end")
        self.ent_hora_corte.insert(0, corte_opts["hora_corte"])

        threading.Thread(
            target=self._run_extraction_thread,
            args=(account, d_start, d_end, modo, agente_filter, corte_opts, auto, sql_cfg),
            daemon=True
        ).start()

    def _run_extraction_thread(self, account, d_start, d_end, modo, agente_filter, corte_opts, auto, sql_cfg=None):
        def progress_cb(msg, frac=None):
            self.after(0, lambda: self._update_progress(msg, frac))

        try:
            rows = self.engine.extract_renewals(
                account_name=account,
                start_dt=d_start,
                end_dt=d_end,
                mode=modo,
                agent_filter=agente_filter,
                progress_callback=progress_cb,
                **corte_opts
            )
            lote_info = self.engine.lote_info
            validaciones = {'ocr': dict(self.engine.ocr_info), 'sql': None}
            if sql_cfg and rows:
                progress_cb(f"🗄️ Validando {len(rows)} renovaciones en la base de datos...", 0.98)
                validaciones['sql'] = self._validar_sql(rows, lote_info, sql_cfg)
            # Calendario de emisión: lo generado antes de la emisión o del cierre de mes ya no se puede cancelar
            actualizar_calendario()
            validaciones['emision'] = marcar_bloqueos(rows)
            self.after(0, lambda: self._on_extraction_done(rows, None, lote_info, auto, validaciones))
        except Exception as e:
            err = str(e)
            self.after(0, lambda: self._on_extraction_done([], err, None, auto))

    def _update_progress(self, msg, frac):
        self.lbl_status.configure(text=plain_text(msg))
        if frac is not None:
            self.progress_bar.set(max(0.05, min(1.0, float(frac))))

    def _on_extraction_done(self, rows, error_msg, lote_info=None, auto=False, validaciones=None):
        self.is_extracting = False
        self.seg_tipo.configure(state="normal")
        self.btn_extract.configure(state="normal", text="Extraer y validar")
        self.progress_bar.set(1.0)

        if error_msg:
            self.lbl_status.configure(text=f"Error: {error_msg}")
            if not auto:
                messagebox.showerror("Error al extraer de Outlook", error_msg)
            return

        self.all_rows = rows
        self.lote_info = lote_info
        self.validaciones = validaciones or {}
        self._tick_clock(reschedule=False)
        self._key_by_rid = {r['Row_ID']: row_key(r) for r in rows}

        # Recuperar lo que el usuario ya había marcado (en escaneos anteriores o antes de cerrar la app)
        self.validated_row_ids = {rid for rid, k in self._key_by_rid.items() if self.estado.has('validados', k)}
        self.validated_row_ids.update(r['Row_ID'] for r in rows if r.get('Lote_Estado') == LOTE_EN_CONCENTRADO)
        self.discarded_row_ids = {rid for rid, k in self._key_by_rid.items() if self.estado.has('descartados', k)}
        self.forced_included_row_ids = {rid for rid, k in self._key_by_rid.items() if self.estado.has('incluidos', k)}

        current_keys = set(self._key_by_rid.values())
        self.new_keys = (current_keys - self._prev_keys) if self._prev_keys is not None else set()
        self._prev_keys = current_keys

        revision = bool(lote_info and lote_info.get('revision'))
        if not auto:
            self.seg_tipo.set(TIPOS_REVISION[1] if revision else TIPOS_REVISION[0])
            if revision:
                self.cmb_quick_filter.set(
                    FILTRO_NO_ENVIADAS if lote_info.get('no_enviadas') else
                    FILTRO_ERRORES if lote_info.get('errores_envio') else
                    FILTRO_SIN_COMPROBAR if lote_info.get('sin_comprobar') else FILTRO_TODAS)
            elif self.cmb_quick_filter.get() in FILTROS_REVISION:
                self.cmb_quick_filter.set(FILTRO_TODAS)
        self._refresh_views()
        self._actualizar_encabezado()

        active_rows, _ = self._get_active_and_filtered_rows()
        msg = f"Actualizado a las {datetime.now():%H:%M}  ·  {len(active_rows)} renovaciones en el lote"
        multi_tickets = {r['Numero_Ticket'] for r in active_rows if r.get('Es_Multiple')}
        if multi_tickets:
            msg += f"  ·  {len(multi_tickets)} ticket(s) con varias renovaciones"
        if lote_info:
            conteos = lote_info['conteos']
            if conteos.get(LOTE_REZAGADA):
                msg += f"  ·  {conteos[LOTE_REZAGADA]} rezagada(s)"
            if lote_info.get('no_enviadas'):
                msg += f"  ·  {lote_info['no_enviadas']} no enviada(s) en su día"
            if conteos.get(LOTE_SIGUIENTE):
                msg += f"  ·  {conteos[LOTE_SIGUIENTE]} para el siguiente lote"
        cerradas = {r['Numero_Ticket'] for r in active_rows if r.get('Cerrado_iTop')}
        if cerradas:
            msg += f"  ·  {len(cerradas)} ticket(s) ya cerrados en iTop"
        self.lbl_status.configure(text=msg)
        self._aprender_texto_correo(lote_info)

        if self.new_keys:
            nuevas = sum(1 for r in active_rows if self._key_by_rid.get(r['Row_ID']) in self.new_keys)
            if nuevas:
                self.bell()
                self._toast(f"{nuevas} renovación(es) nueva(s) desde el último escaneo.\nUsa el filtro «{FILTRO_NUEVAS}» para verlas.")

    def _aprender_texto_correo(self, lote_info):
        """Guarda el texto (saludo y 'FAVOR DE BLOQUEAR…') del último concentrado original para repetirlo en el correo de hoy."""
        concs = sorted((lote_info or {}).get('concentrados') or [], key=lambda c: c.get('recibido') or datetime.min, reverse=True)
        for conc in concs:
            if conc.get('intro') and not re.match(r'\s*(re|rv|fw|fwd)\s*:', conc.get('asunto', ''), re.I):
                if conc['intro'] != self.config_data.get("texto_correo_aprendido"):
                    self.config_data["texto_correo_aprendido"] = conc['intro']
                    self._save_config()
                return

    def _texto_correo(self) -> str:
        return (self.config_data.get("texto_correo") or self.config_data.get("texto_correo_aprendido")
                or INTRO_TICKET_UNICO_DEFAULT)

    # ================= CORTE / ÚLTIMA CONSIDERADA =================

    def _texto_marca_guardada(self) -> str:
        marca = self.config_data.get("ultimo_lote") or {}
        if not marca.get("hasta"):
            return ("Aún no hay un lote registrado en este equipo. Al copiar, exportar o crear el correo del lote "
                    "se guarda la última renovación considerada.")
        try:
            hasta = datetime.fromisoformat(marca["hasta"]).strftime("%d/%m/%Y %H:%M")
        except ValueError:
            hasta = marca["hasta"]
        return (
            f"Último lote registrado en este equipo: hasta {hasta}  ·  última considerada {marca.get('ticket', '-')}, "
            f"cliente {marca.get('cliente', '-')}, crédito {marca.get('credito', '-')}"
        )

    def _actualizar_encabezado(self):
        """Título 'Lote del dd/mm/aaaa', resumen del corte y aviso si no se encontró el concentrado anterior."""
        info = self.lote_info
        if not info:
            self.lbl_page_title.configure(text="Lote del día")
            self.lbl_lote_info.configure(text=self._texto_marca_guardada())
            avisos, rojo, filtro = self._avisos_validaciones()
            if avisos:
                self._mostrar_aviso("\n".join(avisos), rojo, filtro)
            else:
                self.notice.grid_remove()
            return

        hc = info['hora_corte'].strftime('%H:%M')
        c = info['conteos']
        revision = info.get('revision')
        if revision:
            ini, fin = info['fecha_inicio'], info['fecha_fin']
            self.lbl_page_title.configure(
                text=f"Revisión de envíos del {fin:%d/%m/%Y}" if ini == fin else f"Revisión de envíos del {ini:%d/%m} al {fin:%d/%m/%Y}"
            )
        else:
            self.lbl_page_title.configure(text=f"Lote del {info['corte_fin']:%d/%m/%Y}")
        lineas = []
        conc = info.get('concentrado_previo')
        if info['metodo'] == 'concentrado' and conc:
            lineas.append(
                f"Concentrado anterior: «{conc['asunto']}», recibido {conc['recibido']:%d/%m %H:%M}"
                f" ({conc['remitente'] or conc['carpeta']})"
            )
        if revision:
            semana = info.get('semana') or {}
            n_tu = len({f['asunto'] for f in semana.get('filas') or []})
            lineas.append(
                f"{info.get('en_periodo', 0)} renovaciones debían salir en el periodo  ·  enviadas: {info.get('enviadas', 0)} "
                f"(tarde: {info.get('tarde', 0)})  ·  NO enviadas: {info.get('no_enviadas', 0)}  ·  "
                f"datos equivocados: {info.get('errores_envio', 0)}  ·  sin comprobar: {info.get('sin_comprobar', 0)}"
            )
            lineas.append(
                f"Concentrado: {len(semana.get('filas') or [])} renglones de {n_tu} ticket(s) único(s) "
                f"(pestaña «Tickets únicos»)"
                + (f"  ·  {len(semana['sin_tabla'])} sin tabla legible" if semana.get('sin_tabla') else "")
            )
        else:
            ult = info.get('ultima_considerada')
            if ult:
                lineas.append(
                    f"Última considerada: {ult['ticket']} del {ult['fecha']:%d/%m %H:%M}  ·  "
                    f"cliente {ult['cliente']}, crédito {ult['credito']}"
                )
        dia_txt = f"del {info['lote_actual']:%d/%m}" if info.get('pasado') else "de hoy"
        envio_hoy = (f"Ticket único {dia_txt} enviado a las {info['corte_hoy']:%H:%M}" if info.get('corte_hoy')
                     else (f"No se encontró el ticket único {dia_txt}" if info.get('pasado')
                           else "Ticket único de hoy: pendiente"))
        if info.get('pasado') and not revision:
            en_el = c.get(LOTE_EN_CONCENTRADO, 0)
            lineas.append(f"Lote como debía salir el {info['lote_actual']:%d/%m}: {en_el} venían en su ticket único, "
                          f"{info.get('faltan_hoy', 0)} faltaron"
                          + (f" ({sum(1 for r in self.all_rows if r.get('Enviado_Despues'))} se mandaron después)"
                             if any(r.get('Enviado_Despues') for r in self.all_rows) else ""))
        if revision:
            dias_sin = [f"{d['dia']:%d/%m}" for d in info.get('por_dia') or [] if d['habil'] and not d['envios']]
            lineas.append(f"{envio_hoy}  ·  "
                          + (f"Días sin ticket único encontrado: {', '.join(dias_sin)}" if dias_sin
                             else "Todos los días tienen su ticket único")
                          + "  ·  detalle en la pestaña «Por día»")
        else:
            lineas.append(
                f"Rezagadas sumadas al lote: {c.get(LOTE_REZAGADA, 0)}  ·  Ya enviadas: {c.get(LOTE_ENVIADA, 0)}  ·  "
                f"{envio_hoy}  ·  Para el siguiente lote: {c.get(LOTE_SIGUIENTE, 0)}"
            )
        self.lbl_lote_info.configure(text="\n".join(lineas))

        avisos, rojo, filtro = [], False, None
        n_no = info.get('no_enviadas', 0)
        cerradas = info.get('cerradas_sin_envio', 0)
        n_err = info.get('errores_envio', 0)
        if revision:
            if n_no:
                rojo, filtro = True, FILTRO_NO_ENVIADAS
                texto = f"{n_no} renovación(es) NO se enviaron en el ticket único del día que les tocaba"
                detalle = []
                if n_no - cerradas:
                    detalle.append(f"{n_no - cerradas} con el incidente abierto se suman al lote {dia_txt}")
                if cerradas:
                    detalle.append(f"{cerradas} ya tienen el incidente cerrado en iTop (no entran al lote; «Restaurar» las incluye)")
                avisos.append(f"{texto}: {'; '.join(detalle)}.  Clic aquí para verlas.")
            if n_err:
                rojo = True
                avisos.append(f"{n_err} renovación(es) salieron en el ticket único con datos equivocados (cliente, crédito o "
                              f"folio de otro incidente): revisa «{FILTRO_ERRORES}».")
                filtro = filtro or FILTRO_ERRORES
        else:
            # Lote del día: en rojo solo lo que cambia el lote de hoy; lo de días anteriores, en una línea
            if n_no - cerradas:
                rojo, filtro = True, FILTRO_NO_ENVIADAS
                avisos.append(f"{n_no - cerradas} renovación(es) de días anteriores NO se enviaron y siguen abiertas: "
                              f"se suman al lote {dia_txt}.  Clic aquí para verlas.")
            err_hoy = sum(1 for r in self.all_rows if r.get('Error_Envio') and r.get('Dia_Lote') == info['lote_actual'])
            if err_hoy:
                rojo = True
                avisos.append(f"{err_hoy} renovación(es) salieron en el ticket único {dia_txt} con datos equivocados: "
                              f"revisa «{FILTRO_ERRORES}».")
                filtro = filtro or FILTRO_ERRORES
            antes = [t for t in (f"{cerradas} cerrada(s) en iTop sin enviar" if cerradas else "",
                                 f"{n_err - err_hoy} con datos equivocados" if n_err - err_hoy else "",
                                 f"{info['duplicadas']} posible(s) duplicada(s)" if info.get('duplicadas') else "") if t]
            if antes:
                avisos.append(f"De días anteriores: {', '.join(antes)}. No cambian el lote de hoy; el detalle está en "
                              f"«Revisión semanal».")
        if info.get('faltan_hoy'):
            avisos.append(f"{info['faltan_hoy']} renovación(es) {dia_txt} no vienen en el ticket único que salió a las "
                          f"{info['corte_hoy']:%H:%M}: agrégalas o mándalas en el siguiente.")
            filtro = filtro or FILTRO_ALERTAS
        if info.get('avisos_fecha'):
            avisos.append("Fecha corregida: " + "  ·  ".join(info['avisos_fecha']))
        if revision and info.get('duplicadas'):
            avisos.append(f"{info['duplicadas']} renovación(es) ya habían salido antes en otro ticket único "
                          f"(posible solicitud duplicada): están en «{FILTRO_ENVIADAS}».")
        if info.get('sin_comprobar'):
            avisos.append(f"{info['sin_comprobar']} renovación(es) son de días sin ticket único encontrado: revisa «{FILTRO_SIN_COMPROBAR}».")
            filtro = filtro or FILTRO_SIN_COMPROBAR
        if not revision and not (info['metodo'] == 'concentrado' and conc):
            aviso = f"No se encontró el correo concentrado del {info['dia_anterior']:%d/%m}. "
            if info['metodo'] == 'marca':
                aviso += f"Se usó la marca local ({info['marca']:%d/%m %H:%M}) para decidir qué faltó enviar."
            else:
                aviso += f"Se usó la hora habitual ({hc}); revisa las rezagadas antes de enviar el lote."
            avisos.append(aviso)
        extra, rojo_extra, filtro_extra = self._avisos_validaciones()
        avisos += extra
        rojo, filtro = rojo or rojo_extra, filtro or filtro_extra
        if avisos:
            self._mostrar_aviso("\n".join(avisos), rojo, filtro)
        else:
            self.notice.grid_remove()

    def _avisos_validaciones(self) -> tuple:
        """Avisos de la base de datos y de las capturas de pantalla: (avisos, rojo, filtro)."""
        avisos, rojo, filtro = [], False, None
        revision = bool(self.lote_info and self.lote_info.get('revision'))

        def cuenta(r):
            """En la revisión, lo del periodo; en el lote del día, solo lo que va en el lote (no lo ya enviado)."""
            if r.get('Row_ID') in self.discarded_row_ids:
                return False
            if revision:
                return r.get('En_Periodo', True)
            return not self.lote_info or r.get('Lote_Estado') in (LOTE_HOY, LOTE_REZAGADA)

        del_lote = [r for r in self.all_rows if cuenta(r)]
        bloqueadas = [r for r in del_lote if r.get('Bloqueo_Emision') and r.get('Lote_Estado') in (LOTE_HOY, LOTE_REZAGADA)]
        if bloqueadas:
            rojo, filtro = True, FILTRO_EMISION
            avisos.append(f"{len(bloqueadas)} renovación(es) ya NO se pueden cancelar: se generaron antes de la emisión "
                          f"o del cierre de mes (calendario de cédulas). Revisa «{FILTRO_EMISION}».")
        res = self.validaciones.get('sql')
        if res and not res.get('ok'):
            avisos.append(f"No se pudo validar en la base de datos: {res.get('error')}")
        elif res:
            malas = [r for r in del_lote if r.get('SQL_Estado') in SQL_PROBLEMAS]
            if malas:
                rojo, filtro = True, filtro or FILTRO_SQL
                no_procede = sum(1 for r in malas if r.get('SQL_Estado') == SQL_NO_PROCEDE)
                avisos.append(f"{len(malas)} renovación(es) no cuadran con la base"
                              + (f"; {no_procede} de otro día que no generó la APP con plazo menor: "
                                 f"«{LEYENDA_NO_PROCEDE}»" if no_procede else "")
                              + f": revisa «{FILTRO_SQL}».")
            n_app = sum(1 for r in del_lote if r.get('SQL_Estado') == SQL_APP_PLAZO)
            if n_app:
                avisos.append(f"{n_app} renovación(es) son de otro día pero las generó la APP con plazo menor: "
                              f"sí van (columna «Base de datos»).")
            n_ced = sum(1 for r in del_lote if r.get('SQL_Cedula'))
            if n_ced:
                avisos.append(f"{n_ced} renovación(es) ya emitieron (están en CR_CEDULAS_DETALLE): favor de validar.")
                filtro = filtro or FILTRO_ALERTAS
            n_sug = sum(1 for r in del_lote if r.get('SQL_Estado') == SQL_SUGERENCIA)
            if n_sug:
                avisos.append(f"{n_sug} renovación(es) a las que les falta un dato tienen sugerencia de la base "
                              f"(columna «Base de datos»).")
            if res.get('error_detalle'):
                avisos.append(f"La segunda consulta (APP / plazo / cédula) falló: {res['error_detalle']}")
            if revision and res.get('enviadas_con_problema'):
                rojo = True
                avisos.append(f"{res['enviadas_con_problema']} renglón(es) de los tickets únicos no cuadran con la base: "
                              f"pestaña «Tickets únicos».")
        ocr = self.validaciones.get('ocr') or {}
        de_imagen = sum(1 for r in del_lote if r.get('Dato_Imagen'))
        if de_imagen:
            avisos.append(f"{de_imagen} renovación(es) se completaron leyendo una captura de pantalla: "
                          f"revísalas antes de enviarlas (vista «{FILTRO_ALERTAS}»).")
            filtro = filtro or FILTRO_ALERTAS
        tickets_lote = {r.get('Numero_Ticket') for r in del_lote}
        sin_leer = [t for t in ocr.get('sin_leer') or [] if t in tickets_lote]
        if sin_leer:
            avisos.append("No se encontró cliente o crédito en la imagen de: " + ", ".join(sin_leer[:8])
                          + ("…" if len(sin_leer) > 8 else ""))
        if ocr.get('error'):
            avisos.append(f"No se pudieron leer las capturas de pantalla: {ocr['error']}")
        return avisos, rojo, filtro

    def _guardar_marca_lote(self, rows: list):
        """Guarda la renovación más reciente del lote copiado o del correo creado como 'última considerada'."""
        if self.lote_info and (self.lote_info.get('revision') or self.lote_info.get('pasado')):
            return
        en_lote = [r for r in rows if r.get('Fecha_Llegada_DT') or r.get('Fecha_DT')]
        if not en_lote:
            return
        ultima = max(en_lote, key=lambda r: r.get('Fecha_Llegada_DT') or r['Fecha_DT'])
        self.config_data["ultimo_lote"] = {
            "hasta": (ultima.get('Fecha_Llegada_DT') or ultima['Fecha_DT']).isoformat(timespec="minutes"),
            "ticket": ultima['Numero_Ticket'],
            "cliente": ultima['Numero_Cliente'],
            "credito": ultima['Credito_Renovado_Hijo'],
            "guardado": datetime.now().isoformat(timespec="minutes")
        }
        self._save_config()
        if not self.lote_info:
            self._actualizar_encabezado()

    # ================= FILTRADO Y VISTAS =================

    def _fuera_de_lote(self, r: dict) -> bool:
        return r.get('Lote_Estado') in LOTE_EXCLUIDOS and r['Row_ID'] not in self.forced_included_row_ids

    @staticmethod
    def _tiene_alerta(r: dict) -> bool:
        return bool(
            r.get('Datos_Incompletos')
            or r.get('Ticket_Duplicado')
            or r.get('Requiere_Validacion_Agente')
            or r.get('Cerrado_iTop')
            or r.get('No_Enviada')
            or r.get('Falta_Hoy')
            or r.get('Error_Envio')
            or r.get('Posible_Duplicado')
            or r.get('SQL_Estado') in SQL_PROBLEMAS
            or r.get('SQL_Cedula')
            or r.get('Bloqueo_Emision')
            or r.get('Dato_Imagen')
            or "DISCREPANCIA" in r.get('Alerta_Renovacion', '')
        )

    def _es_nueva(self, r: dict) -> bool:
        return self._key_by_rid.get(r['Row_ID']) in self.new_keys

    def _get_active_and_filtered_rows(self):
        auto_excl = bool(self.var_auto_exclude_reassigned.get()) and (self.var_modo.get() == "AGENTE")
        q_filter = self.cmb_quick_filter.get()
        search_txt = self.ent_search.get().strip().lower()

        revision = bool(self.lote_info and self.lote_info.get('revision'))
        active_rows = []
        visible_rows = []
        for r in self.all_rows:
            rid = r['Row_ID']
            is_disc = (rid in self.discarded_row_ids) or (auto_excl and r.get('Reasignado_A_Otro'))
            is_val = rid in self.validated_row_ids
            lote = r.get('Lote_Estado')
            fuera = self._fuera_de_lote(r)

            fuera_periodo = revision and not r.get('En_Periodo', True)
            if not is_disc and not fuera and not fuera_periodo:
                active_rows.append(r)

            if q_filter == FILTRO_PAPELERA:
                if not is_disc:
                    continue
            elif q_filter in FILTROS_REVISION:
                if revision and not r.get('En_Periodo', True):
                    continue  # De la semana anterior: solo se escaneó para sumar rezagadas
                if q_filter == FILTRO_ERRORES and not r.get('Error_Envio'):
                    continue
                if q_filter == FILTRO_SQL and r.get('SQL_Estado') not in SQL_PROBLEMAS:
                    continue
                if q_filter == FILTRO_ENVIADAS and lote != LOTE_ENVIADA:
                    continue
                if q_filter == FILTRO_SIGUIENTE and lote != LOTE_SIGUIENTE:
                    continue
                if q_filter == FILTRO_NO_ENVIADAS and not r.get('No_Enviada'):
                    continue
                if q_filter == FILTRO_TARDE and not r.get('Dias_Retraso'):
                    continue
                if q_filter == FILTRO_SIN_COMPROBAR and lote != LOTE_SIN_COMPROBAR:
                    continue
            else:
                if is_disc or fuera or fuera_periodo:
                    continue
                if q_filter == FILTRO_REZAGADAS and lote != LOTE_REZAGADA:
                    continue
                if q_filter == FILTRO_EMISION and not r.get('Bloqueo_Emision'):
                    continue
                if q_filter == FILTRO_MULTIPLES and not r.get('Es_Multiple'):
                    continue
                if q_filter == FILTRO_ALERTAS and not self._tiene_alerta(r):
                    continue
                if q_filter == FILTRO_PENDIENTES and is_val:
                    continue
                if q_filter == FILTRO_VALIDADAS and not is_val:
                    continue
                if q_filter == FILTRO_NUEVAS and not self._es_nueva(r):
                    continue
                if q_filter == FILTRO_CERRADAS and not r.get('Cerrado_iTop'):
                    continue

            if search_txt:
                haystack = (
                    f"{r.get('Numero_Ticket', '')} {r.get('Numero_Cliente', '')} "
                    f"{r.get('Credito_Renovado_Hijo', '')} {r.get('Credito_Origen_Padre', '')} "
                    f"{r.get('Gafete', '')} {r.get('Nombre_Asignado', '')} "
                    f"{r.get('Motivo', '')} {r.get('Resumen_iTop', '')} {r.get('Lote_Etiqueta', '')}"
                ).lower()
                if search_txt not in haystack:
                    continue

            visible_rows.append(r)

        return active_rows, visible_rows

    def _row_tag(self, r: dict, zebra_idx: int) -> str:
        rid = r['Row_ID']
        if rid in self.discarded_row_ids:
            return "descartado"
        if r.get('No_Enviada') or r.get('Error_Envio'):
            return "no_enviada"
        if r.get('SQL_Estado') in SQL_PROBLEMAS:
            # El cliente no tiene esa renovación: rojo aunque ya esté validada; de otra fecha: ámbar
            return "multiple" if r['SQL_Estado'] == SQL_OTRA_FECHA else "no_enviada"
        if r.get('Bloqueo_Emision') and r.get('Lote_Estado') in (LOTE_HOY, LOTE_REZAGADA):
            return "no_enviada"  # ya pasó la emisión o el cierre de mes: no se puede cancelar
        if r.get('Falta_Hoy'):
            return "rezagada"
        if self._fuera_de_lote(r):
            return "fuera_lote"
        if rid in self.validated_row_ids:
            return "validado"
        if r.get('Cerrado_iTop'):
            return "cerrado"
        if r.get('Lote_Estado') == LOTE_REZAGADA or r.get('Falta_Hoy'):
            return "rezagada"
        if r.get('Datos_Incompletos'):
            return "incompleto"
        if r.get('Ticket_Duplicado'):
            return "duplicado"
        if r.get('Requiere_Validacion_Agente'):
            return "reasignado"
        if r.get('Es_Multiple') or r.get('Nivel_Alerta') == "multiple":
            return "multiple"
        return "normal_alt" if zebra_idx % 2 else "normal"

    def _refresh_views(self):
        self._search_job = None
        active_rows, visible_rows = self._get_active_and_filtered_rows()

        # 1. KPIs sobre el lote activo
        unique_tickets = {r['Numero_Ticket'] for r in active_rows}
        multi_tickets = {r['Numero_Ticket'] for r in active_rows if r.get('Es_Multiple')}
        alert_rows = [r for r in active_rows if self._tiene_alerta(r)]
        rezagadas = [r for r in active_rows if r.get('Lote_Estado') == LOTE_REZAGADA]
        val_count = sum(1 for r in active_rows if r['Row_ID'] in self.validated_row_ids)

        self.lbl_kpi_tickets.configure(text=str(len(unique_tickets)))
        self.lbl_kpi_renovaciones.configure(text=str(len(active_rows)))
        self.lbl_kpi_multiples.configure(text=str(len(multi_tickets)))
        self.lbl_kpi_alertas.configure(
            text=str(len(alert_rows)), text_color=COL["danger"] if alert_rows else COL["text"]
        )
        self.lbl_kpi_rezagadas.configure(text=str(len(rezagadas)))
        self.lbl_kpi_validados.configure(text=f"{val_count} / {len(active_rows)}")
        self.kpi_progress.set(val_count / len(active_rows) if active_rows else 0)
        self.lbl_count.configure(text=f"{len(visible_rows)} de {len(self.all_rows)} renovaciones" if self.all_rows else "")

        # 2. Pestaña Desglose
        prev_sel = set(self.tree_desglose.selection())
        self.tree_desglose.delete(*self.tree_desglose.get_children())
        self._celdas_completas.clear()
        for i, r in enumerate(visible_rows):
            rid = r['Row_ID']
            estado = "✓  Validada" if rid in self.validated_row_ids else "Pendiente"
            if self._es_nueva(r):
                estado += "  ·  Nueva"
            lote = plain_text(r.get('Lote_Etiqueta', ''))
            self._celdas_completas[(rid, "lote")] = lote
            self._celdas_completas[(rid, "cierre")] = resumen_cierre(r.get('Cierre_iTop'))
            self._celdas_completas[(rid, "sql")] = r.get('SQL_Detalle', '')
            self.tree_desglose.insert(
                "", "end", iid=rid,
                values=(
                    estado,
                    r['Numero_Ticket'],
                    f"{r['Indice_En_Ticket']} de {r['Total_En_Ticket']}",
                    r['Numero_Cliente'],
                    r['Credito_Renovado_Hijo'],
                    r['Credito_Origen_Padre'],
                    texto_sql(r.get('SQL_Estado', '')),
                    r.get('Lote_Resumen') or lote_corto(lote),
                    resumen_cierre(r.get('Cierre_iTop'), corto=True),
                    r['Gafete'],
                    r['Nombre_Asignado'],
                    plain_text(r['Alerta_Renovacion']),
                    plain_text(r['Alerta_Asignacion']),
                    r['Motivo'],
                    self._fecha_llegada_str(r)
                ),
                tags=(self._row_tag(r, i),)
            )
        keep = [rid for rid in prev_sel if self.tree_desglose.exists(rid)]
        if keep:
            self.tree_desglose.selection_set(keep)
        self._reapply_sort(self.tree_desglose)

        if visible_rows:
            self.lbl_empty.place_forget()
        else:
            self.lbl_empty.configure(
                text="Aún no hay renovaciones. Presiona «Extraer y validar» (F5) para leer Outlook." if not self.all_rows
                else "No hay renovaciones con este filtro o búsqueda."
            )
            self.lbl_empty.place(relx=0.5, rely=0.45, anchor="center")
            self.lbl_empty.lift()

        # 3. Pestaña Por ticket
        self.tree_tickets.delete(*self.tree_tickets.get_children())
        tickets_grouped = {}
        for r in active_rows:
            g = tickets_grouped.setdefault(r['Numero_Ticket'], {
                'clientes': [], 'creditos_hijos': [], 'creditos_padres': [],
                'gafete': r['Gafete'], 'agente': r['Nombre_Asignado'], 'resumen': r['Resumen_iTop'],
                'alerta_asig': plain_text(r['Alerta_Asignacion']), 'fecha': self._fecha_llegada_str(r), 'has_missing': False,
                'cierre': resumen_cierre(r.get('Cierre_iTop'), corto=True)
            })
            self._celdas_completas[(f"TKT_{r['Numero_Ticket']}", "cierre")] = resumen_cierre(r.get('Cierre_iTop'))
            if r['Numero_Cliente'] not in g['clientes']:
                g['clientes'].append(r['Numero_Cliente'])
            g['creditos_hijos'].append(r['Credito_Renovado_Hijo'])
            padre = r['Credito_Origen_Padre']
            if padre and padre != "-" and padre not in g['creditos_padres']:
                g['creditos_padres'].append(padre)
            if r.get('Datos_Incompletos'):
                g['has_missing'] = True

        for i, (tnum, g) in enumerate(tickets_grouped.items()):
            cant = len(g['creditos_hijos'])
            tag = "incompleto" if g['has_missing'] else (
                "cerrado" if g['cierre'] else ("multiple" if cant > 1 else ("normal_alt" if i % 2 else "normal"))
            )
            self.tree_tickets.insert(
                "", "end", iid=f"TKT_{tnum}",
                values=(
                    tnum,
                    str(cant),
                    ", ".join(g['clientes']),
                    ", ".join(g['creditos_hijos']),
                    ", ".join(g['creditos_padres']) if g['creditos_padres'] else "-",
                    g['gafete'], g['agente'], g['resumen'], g['alerta_asig'], g['cierre'], g['fecha']
                ),
                tags=(tag,)
            )
        self._reapply_sort(self.tree_tickets)

        # 4. Pestaña Por compañero
        self.tree_agentes.delete(*self.tree_agentes.get_children())
        agentes = {}
        for r in active_rows:
            ag = agentes.setdefault((r['Gafete'], r['Nombre_Asignado']), {
                'tickets': set(), 'multi': set(), 'renovaciones': 0, 'alertas': 0, 'rezagadas': 0, 'validados': 0
            })
            ag['tickets'].add(r['Numero_Ticket'])
            if r.get('Es_Multiple'):
                ag['multi'].add(r['Numero_Ticket'])
            ag['renovaciones'] += 1
            ag['alertas'] += int(self._tiene_alerta(r))
            ag['rezagadas'] += int(r.get('Lote_Estado') == LOTE_REZAGADA)
            ag['validados'] += int(r['Row_ID'] in self.validated_row_ids)

        for i, ((gaf, nom), ag) in enumerate(sorted(agentes.items(), key=lambda x: x[1]['renovaciones'], reverse=True)):
            pendientes = ag['renovaciones'] - ag['validados']
            self.tree_agentes.insert(
                "", "end",
                values=(
                    gaf, nom, len(ag['tickets']), ag['renovaciones'], len(ag['multi']), ag['alertas'],
                    ag['rezagadas'], f"{ag['validados']} / {ag['renovaciones']}", pendientes
                ),
                tags=("validado" if pendientes == 0 else ("normal_alt" if i % 2 else "normal"),)
            )
        self._reapply_sort(self.tree_agentes)

        # Contadores en las pestañas
        self.tabs.set_label("desglose", f"Desglose  {len(visible_rows)}")
        self.tabs.set_label("tickets", f"Por ticket  {len(tickets_grouped)}")
        self.tabs.set_label("agentes", f"Por compañero  {len(agentes)}")
        self.tabs.set_label("enviado", f"Tickets únicos  {self._llenar_enviado(self.ent_search.get().strip().lower())}")
        n_dias = self._llenar_dias()
        self.tabs.set_label("dias", f"Por día  {n_dias}" if n_dias else "Por día")

    # ================= ORDENAMIENTO =================

    def _on_heading_click(self, tree, col):
        prev_col, prev_rev = self._sort_state.get(str(tree), (None, False))
        reverse = (not prev_rev) if prev_col == col else False
        self._sort_state[str(tree)] = (col, reverse)
        self._apply_sort(tree, col, reverse)

    def _reapply_sort(self, tree):
        state = self._sort_state.get(str(tree))
        if state:
            self._apply_sort(tree, *state)

    def _apply_sort(self, tree, col, reverse):
        items = [(tree.set(k, col), k) for k in tree.get_children("")]
        items.sort(key=lambda x: _sort_key(x[0]), reverse=reverse)
        for index, (_val, k) in enumerate(items):
            tree.move(k, "", index)
        for c in tree["columns"]:
            base = self._headers.get((str(tree), c), c)
            arrow = ("  ▼" if reverse else "  ▲") if c == col else ""
            tree.heading(c, text=base + arrow)

    # ================= DETALLE =================

    def _on_select_desglose_row(self, _event=None):
        sel = self.tree_desglose.selection()
        if not sel:
            return
        row_obj = next((r for r in self.all_rows if r['Row_ID'] == sel[0]), None)
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
                f"{tnum}   ·   Cliente {r['Numero_Cliente']}   ·   Crédito a cancelar {r['Credito_Renovado_Hijo']}"
                f"   ·   Origen {r['Credito_Origen_Padre']}   ·   {r['Gafete']} {r['Nombre_Asignado']}"
            )
        )

        box = self.txt_detail
        box.configure(state="normal")
        box.delete("1.0", "end")
        if r.get('No_Enviada') or r.get('Dias_Retraso') or r.get('Falta_Hoy') or r.get('Error_Envio'):
            etiqueta = plain_text(r.get('Lote_Etiqueta', ''))
            box.insert("end", f"{etiqueta}\n", "error" if r.get('No_Enviada') or r.get('Error_Envio') else "warn")
            if r.get('Error_Envio') and r['Error_Envio'] not in etiqueta:
                box.insert("end", f"DATOS EQUIVOCADOS en el ticket único: {r['Error_Envio']}\n", "error")
            box.insert("end", "\n")
        if r.get('Bloqueo_Emision'):
            box.insert("end", f"Emisión: {r['Bloqueo_Emision']}\n\n", "error")
        if r.get('SQL_Detalle') or r.get('Dato_Imagen') or r.get('SQL_Cedula'):
            if r.get('SQL_Cedula'):
                box.insert("end", f"Cédula: {r['SQL_Cedula']}\n", "warn")
            if r.get('SQL_Detalle'):
                box.insert("end", f"Base de datos: {r['SQL_Detalle']}\n",
                           "error" if r.get('SQL_Estado') in SQL_PROBLEMAS else "ok")
            if r.get('Dato_Imagen'):
                box.insert("end", f"Imagen: {r['Dato_Imagen']}. Compáralo con la imagen del correo antes de enviarlo.\n", "warn")
            box.insert("end", "\n")
        cierre = r.get('Cierre_iTop')
        if cierre:
            box.insert("end", f"Cerrado en iTop: {cierre.get('estado') or 'Cerrado'}", "cierre")
            if cierre.get('fecha_dt'):
                box.insert("end", f" el {cierre['fecha_dt']:%d/%m/%Y %H:%M}", "cierre")
            box.insert("end", "\n", "cierre")
            if cierre.get('cerro'):
                fuente = "" if cierre.get('cerro_fuente') == 'correo' else "  (agente asignado al cerrar)"
                box.insert("end", f"Lo cerró: {cierre['cerro']}{fuente}\n", "titulo")
            box.insert("end", f"Motivo: {cierre.get('motivo') or 'El aviso no trae el texto de la solución'}\n", "titulo")
            if cierre.get('codigo'):
                box.insert("end", f"Código de solución: {cierre['codigo']}\n", "muted")
            box.insert("end", "\n")
        box.insert("end", f"Ticket {tnum}  ·  {len(siblings)} renovación(es)\n", "titulo")
        box.insert("end", f"{r.get('Resumen_iTop', '-')}\n\n", "muted")
        for s in siblings:
            ok = s['Row_ID'] in self.validated_row_ids
            tag = "ok" if ok else ("warn" if s.get('Datos_Incompletos') or s.get('Es_Multiple') else "titulo")
            box.insert(
                "end",
                f"{'✓' if ok else '○'}  {s['Indice_En_Ticket']} de {s['Total_En_Ticket']}  ·  Cliente {s['Numero_Cliente']}"
                f"  ·  Cancelar {s['Credito_Renovado_Hijo']}  ·  Origen {s['Credito_Origen_Padre']}\n",
                tag
            )
            box.insert("end", f"     {plain_text(s['Alerta_Renovacion'])}\n", "muted")
        box.insert("end", f"\nLote: {plain_text(r.get('Lote_Etiqueta', '')) or '-'}\n", "titulo")
        box.insert("end", f"Llegó: {self._fecha_llegada_str(r)}", "muted")
        if r.get('Dia_Lote'):
            box.insert("end", f"  ·  debía ir en el ticket único del {r['Dia_Lote']:%d/%m/%Y}", "muted")
        box.insert("end", "\n", "muted")
        box.insert("end", f"\nAsignación: {plain_text(r.get('Alerta_Asignacion', '-'))}\n", "titulo")
        box.insert("end", f"{plain_text(r.get('Historial_Asignaciones', '-'))}\n", "muted")
        box.configure(state="disabled")

        # Correo original con los números extraídos resaltados
        body = r.get('Descripcion_Completa', '') or ''
        self.txt_body.configure(state="normal")
        self.txt_body.delete("1.0", "end")
        self.txt_body.insert("1.0", body)
        numeros = set()
        for s in siblings:
            for campo in ('Numero_Cliente', 'Credito_Renovado_Hijo', 'Credito_Origen_Padre'):
                numeros.update(n.lstrip('0') for n in re.findall(r'\d{3,10}', str(s.get(campo, ''))))
        for n in numeros:
            if not n:
                continue
            for m in re.finditer(rf'(?<!\d)0*{re.escape(n)}(?!\d)', body):
                self.txt_body.tag_add("hl", f"1.0+{m.start()}c", f"1.0+{m.end()}c")
        self.txt_body.configure(state="disabled")

    # ================= ACCIONES SOBRE FILAS =================

    def _selected_rows(self) -> list:
        sel = set(self.tree_desglose.selection())
        return [r for r in self.all_rows if r['Row_ID'] in sel]

    def _keys_for(self, rids) -> list:
        return [self._key_by_rid[rid] for rid in rids if rid in self._key_by_rid]

    def _toggle_selected_validation(self):
        sel = self.tree_desglose.selection()
        if not sel:
            return
        to_on = [rid for rid in sel if rid not in self.validated_row_ids]
        to_off = [rid for rid in sel if rid in self.validated_row_ids]
        self.validated_row_ids.update(to_on)
        self.validated_row_ids.difference_update(to_off)
        self.estado.set('validados', self._keys_for(to_on), True)
        self.estado.set('validados', self._keys_for(to_off), False)
        self._refresh_views()
        r = next((x for x in self.all_rows if x['Row_ID'] == sel[0]), None)
        if r:
            self._show_row_detail(r)

    def _validate_all_visible(self):
        _, visible_rows = self._get_active_and_filtered_rows()
        if not visible_rows:
            return
        rids = [r['Row_ID'] for r in visible_rows]
        all_already = all(rid in self.validated_row_ids for rid in rids)
        if all_already:
            self.validated_row_ids.difference_update(rids)
        else:
            self.validated_row_ids.update(rids)
        self.estado.set('validados', self._keys_for(rids), not all_already)
        self._refresh_views()

    def _discard_selected_rows(self):
        sel = self.tree_desglose.selection()
        if not sel:
            messagebox.showinfo("Selecciona una fila", "Selecciona una o más filas del desglose para descartarlas.")
            return
        self.discarded_row_ids.update(sel)
        self.estado.set('descartados', self._keys_for(sel), True)
        self._refresh_views()

    def _restore_selected_rows(self):
        sel = self.tree_desglose.selection()
        if sel:
            self.discarded_row_ids.difference_update(sel)
            self.estado.set('descartados', self._keys_for(sel), False)
            incluir = [
                rid for rid in sel
                if next((r for r in self.all_rows if r['Row_ID'] == rid), {}).get('Lote_Estado') in LOTE_EXCLUIDOS
            ]
            self.forced_included_row_ids.update(incluir)
            self.estado.set('incluidos', self._keys_for(incluir), True)
        else:
            if not self.discarded_row_ids:
                return
            if not messagebox.askyesno("Restaurar", "¿Restaurar TODAS las filas descartadas?"):
                return
            self.estado.set('descartados', self._keys_for(self.discarded_row_ids), False)
            self.discarded_row_ids.clear()
        self._refresh_views()

    def _popup_menu(self, event, tree, menu):
        if tree.identify_region(event.x, event.y) == "heading":
            self._menu_columnas(tree, event.x_root, event.y_root)
            return
        iid = tree.identify_row(event.y)
        if iid and iid not in tree.selection():
            tree.selection_set(iid)
        if not tree.selection():
            return
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    # ---------- Columnas visibles ----------

    def _clave_columnas(self, tree) -> str:
        return "tickets" if tree is self.tree_tickets else "desglose"

    def _aplicar_columnas(self, tree):
        ocultas = set(self.config_data.get("columnas_ocultas", {}).get(self._clave_columnas(tree), []))
        tree["displaycolumns"] = [c for c in tree["columns"] if c not in ocultas or c in COLUMNAS_FIJAS]

    def _abrir_menu_columnas(self):
        tree = self.tree_tickets if self.tabs.get() == self.tabs_names["tickets"] else self.tree_desglose
        b = self.btn_columnas
        self._menu_columnas(tree, b.winfo_rootx(), b.winfo_rooty() + b.winfo_height())

    def _menu_columnas(self, tree, x, y):
        clave = self._clave_columnas(tree)
        ocultas = set(self.config_data.get("columnas_ocultas", {}).get(clave, []))
        m = tk.Menu(self, **self._menu_opts)
        m.add_command(label="Columnas de la tabla", state="disabled")
        self._vars_columnas = []
        for col in tree["columns"]:
            if col in COLUMNAS_FIJAS:
                continue
            var = tk.BooleanVar(value=col not in ocultas)
            self._vars_columnas.append(var)
            m.add_checkbutton(
                label=self._headers[(str(tree), col)], variable=var,
                command=lambda c=col, v=var: self._mostrar_columna(tree, c, v.get())
            )
        m.add_separator()
        m.add_command(label="Mostrar en pantalla", state="disabled")
        vista = self.config_data.get("vista", {})
        for zona, texto in VISTA_ZONAS:
            var = tk.BooleanVar(value=vista.get(zona, True))
            self._vars_columnas.append(var)
            m.add_checkbutton(label=texto, variable=var, command=lambda z=zona, v=var: self._mostrar_zona(z, v.get()))
        m.add_separator()
        m.add_command(label="Restablecer vista", command=lambda: self._restablecer_columnas(tree))
        try:
            m.tk_popup(x, y)
        finally:
            m.grab_release()

    def _mostrar_columna(self, tree, col, visible: bool):
        clave = self._clave_columnas(tree)
        ocultas = self.config_data.setdefault("columnas_ocultas", {}).setdefault(clave, [])
        if visible and col in ocultas:
            ocultas.remove(col)
        elif not visible and col not in ocultas:
            ocultas.append(col)
        self._aplicar_columnas(tree)
        self._save_config()

    def _restablecer_columnas(self, tree):
        clave = self._clave_columnas(tree)
        self.config_data.setdefault("columnas_ocultas", {})[clave] = list(COLUMNAS_OCULTAS_DEFAULT[clave])
        self.config_data["vista"] = dict(VISTA_DEFAULT)
        self._aplicar_columnas(tree)
        self._aplicar_vista()
        self._save_config()

    def _mostrar_zona(self, zona: str, visible: bool):
        self.config_data.setdefault("vista", dict(VISTA_DEFAULT))[zona] = visible
        self._aplicar_vista()
        self._save_config()

    def _aplicar_vista(self):
        """Muestra u oculta indicadores, panel de detalle y configuración lateral según «Vista»."""
        vista = self.config_data.get("vista", {})
        if vista.get("indicadores", True):
            self.kpi_bar.grid()
            self.paned.grid_configure(pady=0)
        else:
            self.kpi_bar.grid_remove()
            self.paned.grid_configure(pady=(14, 0))
        if vista.get("configuracion", True):
            self.sidebar_config.grid()
        else:
            self.sidebar_config.grid_remove()   # queda solo «Extraer y validar»
        en_paned = str(self.detail_frame) in [str(p) for p in self.paned.panes()]
        if vista.get("detalle", True) and not en_paned:
            self.paned.add(self.detail_frame, weight=2)
        elif not vista.get("detalle", True) and en_paned:
            self.paned.forget(self.detail_frame)

    def _to_clipboard(self, text: str, aviso: str):
        self.clipboard_clear()
        self.clipboard_append(text)
        self.update()
        self.lbl_status.configure(text=aviso)

    def _copy_field(self, campos: tuple):
        rows = self._selected_rows()
        if not rows:
            return
        lines = ["\t".join(str(r.get(c, '')) for c in campos) for r in rows]
        self._to_clipboard("\n".join(lines), f"Copiado al portapapeles: {len(rows)} valor(es).")

    def _copy_selected_rows(self):
        rows = self._selected_rows()
        if rows:
            self._to_clipboard(build_lote_tsv(rows, self.validated_row_ids), f"{len(rows)} fila(s) copiadas al portapapeles.")
        return "break"

    def _copy_selected_ticket(self):
        tnum = self._selected_ticket()
        if tnum:
            self._to_clipboard(tnum, f"Ticket {tnum} copiado.")

    # ================= iTop =================

    @staticmethod
    def _fecha_llegada_str(r: dict) -> str:
        llegada = r.get('Fecha_Llegada_DT')
        return llegada.strftime('%d/%m/%Y %H:%M') if llegada else r.get('Fecha', '')

    def _itop_url(self, tnum: str, buscar: bool = False) -> str:
        if buscar:
            return build_itop_search_url(tnum)
        return build_itop_url(tnum, self.config_data.get("itop_url_template", ""))

    def _selected_ticket(self) -> str:
        """Ticket de la fila seleccionada en la pestaña visible."""
        if self.tabs.get() == self.tabs_names["tickets"]:
            sel_t = self.tree_tickets.selection()
            if sel_t:
                return sel_t[0].replace("TKT_", "")
        sel_d = self.tree_desglose.selection()
        if sel_d:
            row_obj = next((r for r in self.all_rows if r['Row_ID'] == sel_d[0]), None)
            if row_obj:
                return row_obj['Numero_Ticket']
        sel_t = self.tree_tickets.selection()
        if sel_t:
            return sel_t[0].replace("TKT_", "")
        return ""

    def _open_selected_in_itop(self, buscar: bool = False):
        tnum = self._selected_ticket()
        if not tnum:
            messagebox.showinfo("Abrir en iTop", "Selecciona primero una renovación o un ticket en la tabla.")
            return

        url = self._itop_url(tnum, buscar=buscar)
        opened = False
        try:
            opened = webbrowser.open_new_tab(url)
        except Exception:
            opened = False
        if not opened:
            try:
                os.startfile(url)
                opened = True
            except Exception:
                opened = False
        if not opened:
            self._to_clipboard(url, "La liga de iTop se copió al portapapeles.")
            messagebox.showwarning(
                "No se pudo abrir el navegador",
                f"No se pudo abrir el navegador automáticamente.\nLa liga se copió al portapapeles:\n\n{url}"
            )
            return
        self.lbl_status.configure(text=f"Abriendo {tnum} en iTop: {url}")

    # ================= LOTE: COPIAR, CORREO Y EXCEL =================

    def _lote_activo_o_aviso(self) -> list:
        if self.lote_info and self.lote_info.get('revision'):
            messagebox.showwarning(
                "Es una revisión", "Estás en «Revisión semanal»: el ticket único se arma desde «Lote del día» (F5).")
            return []
        active_rows, _ = self._get_active_and_filtered_rows()
        if not active_rows:
            messagebox.showwarning("Sin datos", "No hay renovaciones en el lote. Primero extrae (F5).")
        return active_rows

    def _copy_batch_to_clipboard(self):
        active_rows = self._lote_activo_o_aviso()
        if not active_rows:
            return
        tsv_text = build_ticket_unico_tsv(active_rows)
        html = build_ticket_unico_html(active_rows, url_fn=self._itop_url)
        como_tabla = set_clipboard_html(html, tsv_text)
        if not como_tabla:
            self.clipboard_clear()
            self.clipboard_append(tsv_text)
            self.update()
        self._guardar_marca_lote(active_rows)

        rezagadas = sum(1 for r in active_rows if r.get('Lote_Estado') == LOTE_REZAGADA)
        messagebox.showinfo(
            "Lote copiado al portapapeles",
            f"Se copiaron {len(active_rows)} renovaciones"
            + (f" (incluye {rezagadas} rezagada(s) del día anterior)" if rezagadas else "") + ".\n\n"
            + ("Pégalas con Ctrl+V: en Outlook quedan como tabla con formato y en Excel como celdas.\n\n" if como_tabla
               else "Pégalas con Ctrl+V en el correo o en Excel.\n\n")
            + self._texto_marca_guardada()
        )

    def _crear_correo_lote(self):
        active_rows = self._lote_activo_o_aviso()
        if not active_rows:
            return
        self._save_config()
        pendientes = sum(1 for r in active_rows if r['Row_ID'] not in self.validated_row_ids)
        cerradas = sorted({r['Numero_Ticket'] for r in active_rows if r.get('Cerrado_iTop')})
        avisos = []
        if pendientes:
            avisos.append(f"{pendientes} de {len(active_rows)} renovaciones aún no están validadas.")
        sql_mal = [r['Numero_Ticket'] for r in active_rows if r.get('SQL_Estado') in SQL_PROBLEMAS]
        if sql_mal:
            avisos.append(f"{len(sql_mal)} renovación(es) no cuadran con la base de datos: {', '.join(sql_mal[:8])}"
                          + ("…" if len(sql_mal) > 8 else "") + f". Revisa «{FILTRO_SQL}».")
        bloqueadas = sorted({r['Numero_Ticket'] for r in active_rows if r.get('Bloqueo_Emision')})
        if bloqueadas:
            avisos.append(f"{len(bloqueadas)} ticket(s) ya NO se pueden cancelar (pasó la emisión o el cierre de mes): "
                          f"{', '.join(bloqueadas[:8])}" + ("…" if len(bloqueadas) > 8 else "")
                          + f". Revisa «{FILTRO_EMISION}».")
        con_cedula = sorted({r['Numero_Ticket'] for r in active_rows if r.get('SQL_Cedula')})
        if con_cedula:
            avisos.append(f"{len(con_cedula)} ticket(s) ya emitieron (están en CR_CEDULAS_DETALLE): "
                          f"{', '.join(con_cedula[:8])}. Favor de validar.")
        de_imagen = [r['Numero_Ticket'] for r in active_rows if r.get('Dato_Imagen')]
        if de_imagen:
            avisos.append(f"{len(de_imagen)} renovación(es) tienen datos leídos de una imagen: {', '.join(de_imagen[:8])}.")
        if cerradas:
            avisos.append(f"{len(cerradas)} ticket(s) ya están cerrados en iTop: {', '.join(cerradas[:8])}"
                          + ("…" if len(cerradas) > 8 else "") + ". Revisa el motivo antes de enviarlos (vista «Cerradas en iTop»).")
        if avisos and not messagebox.askyesno("Antes de crear el correo", "\n\n".join(avisos) + "\n\n¿Crear el correo de todas formas?"):
            return

        # La fecha del lote que se extrajo (no la que se esté eligiendo en el panel sin buscar)
        fecha = (self.lote_info or {}).get('lote_actual') or self.cal_end.get_date()
        plantilla = self.config_data.get("asunto_borrador") or ASUNTO_BORRADOR_DEFAULT
        try:
            asunto = plantilla.format(fecha=fecha.strftime('%d-%m-%Y'))
        except (KeyError, IndexError, ValueError):
            asunto = plantilla
        html = build_ticket_unico_html(active_rows, intro=self._texto_correo(), url_fn=self._itop_url)
        try:
            self.engine.create_draft_email(
                asunto, html, self.config_data.get("correo_para", ""), self.config_data.get("correo_cc", "")
            )
        except Exception as e:
            messagebox.showerror("No se pudo crear el correo", f"Outlook no respondió:\n{e}")
            return
        self._guardar_marca_lote(active_rows)
        self.lbl_status.configure(
            text=f"Borrador «{asunto}» abierto en Outlook con {len(active_rows)} renovaciones. Revísalo y envíalo desde Outlook."
        )

    def _export_excel(self):
        active_rows, _ = self._get_active_and_filtered_rows()
        if not active_rows and not self.all_rows:
            messagebox.showwarning("Sin datos", "Primero realiza una extracción de renovaciones antes de exportar.")
            return

        now_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        modo_str = "Concentrado_Diario" if self.var_modo.get() == "CONCENTRADO" else f"Agente_{self.ent_agente.get().strip()}"
        filepath = filedialog.asksaveasfilename(
            title="Guardar reporte de validación de renovaciones",
            defaultextension=".xlsx",
            initialfile=f"Validador_Renovaciones_{modo_str}_{now_str}.xlsx",
            filetypes=[("Archivo de Excel", "*.xlsx")]
        )
        if not filepath:
            return

        try:
            self._write_excel_report(filepath, active_rows)
            self.lbl_status.configure(text=f"Reporte de Excel guardado en: {filepath}")
            if messagebox.askyesno("Reporte generado", f"El archivo se guardó correctamente:\n{filepath}\n\n¿Deseas abrirlo ahora?"):
                os.startfile(filepath)
        except Exception as e:
            messagebox.showerror("Error al exportar Excel", f"No se pudo guardar el archivo Excel:\n{e}")

    def _write_excel_report(self, filepath: str, active_rows: list):
        # 1. Desglose (1 fila = 1 crédito renovado a cancelar)
        desglose_data = []
        for r in active_rows:
            tnum = r['Numero_Ticket']
            desglose_data.append({
                "Validado": "SÍ" if r['Row_ID'] in self.validated_row_ids else "PENDIENTE",
                "Fecha / Hora": self._fecha_llegada_str(r),
                "Ticket iTop": tnum,
                "Desglose en Ticket": f"{r['Indice_En_Ticket']} de {r['Total_En_Ticket']}",
                "No. Cliente (Cuenta)": r['Numero_Cliente'],
                "Crédito Renovado (A Cancelar)": r['Credito_Renovado_Hijo'],
                "Crédito Origen (Padre)": r['Credito_Origen_Padre'],
                "Base de datos (SQL)": r.get('SQL_Detalle', ''),
                "Cédula": r.get('SQL_Cedula', ''),
                "Emisión": r.get('Bloqueo_Emision', ''),
                "Leído de imagen": r.get('Dato_Imagen', ''),
                "Lote / Corte": plain_text(r.get('Lote_Etiqueta', '')),
                "Gafete": r['Gafete'],
                "Agente Responsable": r['Nombre_Asignado'],
                "Estado / Alerta de Renovación": plain_text(r['Alerta_Renovacion']),
                "Estado Asignación": plain_text(r['Alerta_Asignacion']),
                "Historial Asignaciones": plain_text(r.get('Historial_Asignaciones', '')),
                "Cierre iTop - Estado": (r.get('Cierre_iTop') or {}).get('estado', ''),
                "Cierre iTop - Fecha": (
                    r['Cierre_iTop']['fecha_dt'].strftime('%d/%m/%Y %H:%M') if (r.get('Cierre_iTop') or {}).get('fecha_dt') else ''
                ),
                "Cierre iTop - Lo cerró": (r.get('Cierre_iTop') or {}).get('cerro', ''),
                "Cierre iTop - Motivo": (r.get('Cierre_iTop') or {}).get('motivo', ''),
                "Motivo / Observación": r['Motivo'],
                "Sucursal / Localidad": r.get('Sucursal', ''),
                "Resumen iTop": r.get('Resumen_iTop', ''),
                "Link Portal iTop": self._itop_url(tnum)
            })

        # 2. Resumen por ticket
        tickets_map = {}
        for r in active_rows:
            tm = tickets_map.setdefault(r['Numero_Ticket'], {
                "Ticket iTop": r['Numero_Ticket'],
                "Fecha / Hora": self._fecha_llegada_str(r),
                "Cantidad de Renovaciones": 0,
                "Clientes": [],
                "Créditos Renovados a Cancelar": [],
                "Créditos Padre": [],
                "Gafete": r['Gafete'],
                "Agente Responsable": r['Nombre_Asignado'],
                "Estado Asignación": plain_text(r['Alerta_Asignacion']),
                "Resumen iTop": r.get('Resumen_iTop', '')
            })
            tm["Cantidad de Renovaciones"] += 1
            if r['Numero_Cliente'] not in tm["Clientes"]:
                tm["Clientes"].append(r['Numero_Cliente'])
            tm["Créditos Renovados a Cancelar"].append(r['Credito_Renovado_Hijo'])
            if r['Credito_Origen_Padre'] != "-" and r['Credito_Origen_Padre'] not in tm["Créditos Padre"]:
                tm["Créditos Padre"].append(r['Credito_Origen_Padre'])

        tickets_data = [
            {
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
            }
            for tm in tickets_map.values()
        ]

        # 3. Resumen por agente
        agentes_map = {}
        for r in active_rows:
            ag = agentes_map.setdefault((r['Gafete'], r['Nombre_Asignado']), {
                "tickets": set(), "multi": set(), "renovaciones": 0, "alertas": 0, "validados": 0
            })
            ag["tickets"].add(r['Numero_Ticket'])
            if r.get('Es_Multiple'):
                ag["multi"].add(r['Numero_Ticket'])
            ag["renovaciones"] += 1
            ag["alertas"] += int(self._tiene_alerta(r))
            ag["validados"] += int(r['Row_ID'] in self.validated_row_ids)

        agentes_data = [
            {
                "Gafete": gaf,
                "Nombre del Agente": nom,
                "Tickets de Renovación": len(v["tickets"]),
                "Total Créditos Renovados a Cancelar": v["renovaciones"],
                "Tickets con Múltiples Renovaciones": len(v["multi"]),
                "Filas con Alerta": v["alertas"],
                "Validadas": v["validados"],
                "Pendientes": v["renovaciones"] - v["validados"]
            }
            for (gaf, nom), v in sorted(agentes_map.items(), key=lambda x: x[1]["renovaciones"], reverse=True)
        ]

        # 4. Descartados y fuera del lote (ya enviadas / después del corte)
        auto_excl = bool(self.var_auto_exclude_reassigned.get()) and (self.var_modo.get() == "AGENTE")
        descartados_data = [
            {
                "Fecha / Hora": self._fecha_llegada_str(r),
                "Ticket iTop": r['Numero_Ticket'],
                "No. Cliente": r['Numero_Cliente'],
                "Crédito Renovado": r['Credito_Renovado_Hijo'],
                "Gafete": r['Gafete'],
                "Agente": r['Nombre_Asignado'],
                "Motivo Descarte / Estado": plain_text(
                    r.get('Lote_Etiqueta', '') if self._fuera_de_lote(r) and r['Row_ID'] not in self.discarded_row_ids
                    else r['Alerta_Asignacion']
                )
            }
            for r in self.all_rows
            if (r['Row_ID'] in self.discarded_row_ids) or (auto_excl and r.get('Reasignado_A_Otro')) or self._fuera_de_lote(r)
        ]

        # 5. Revisión de envíos: cada renovación contra los tickets únicos (en qué día debía ir y si salió)
        def resultado(r):
            estado = r.get('Lote_Estado')
            if r.get('No_Enviada'):
                return "NO ENVIADA - incidente cerrado" if estado == LOTE_CERRADA_SIN_ENVIO else "NO ENVIADA - va en el lote de hoy"
            if r.get('Dias_Retraso'):
                return "Enviada tarde"
            return {
                LOTE_ENVIADA: "Enviada", LOTE_EN_CONCENTRADO: "Enviada", LOTE_SIN_COMPROBAR: "Sin comprobar",
                LOTE_SIGUIENTE: "Lote siguiente", LOTE_REZAGADA: "Rezagada (va hoy)",
            }.get(estado, "Lote de hoy")

        revision_data = [
            {
                "Resultado": resultado(r),
                "Ticket iTop": r['Numero_Ticket'],
                "Llegó": self._fecha_llegada_str(r),
                "Debía ir en el ticket único del": r['Dia_Lote'].strftime('%d/%m/%Y') if r.get('Dia_Lote') else '',
                "No. Cliente": r['Numero_Cliente'],
                "Crédito a cancelar": r['Credito_Renovado_Hijo'],
                "Enviada en": r.get('Enviado_En', ''),
                "Fecha de envío": r['Enviado_Fecha'].strftime('%d/%m/%Y') if r.get('Enviado_Fecha') else '',
                "Días de retraso": r.get('Dias_Retraso') or '',
                "Detalle": plain_text(r.get('Lote_Etiqueta', '')),
                "Datos equivocados en el ticket único": r.get('Error_Envio', ''),
                "Base de datos (SQL)": r.get('SQL_Detalle', ''),
                "Emisión": r.get('Bloqueo_Emision', ''),
                "Cierre iTop": resumen_cierre(r.get('Cierre_iTop')),
                "Gafete": r['Gafete'],
                "Agente": r['Nombre_Asignado'],
                "Descartada": "SÍ" if r['Row_ID'] in self.discarded_row_ids else "",
            }
            for r in sorted(self.all_rows, key=lambda x: (not x.get('No_Enviada'), x.get('Fecha_Llegada_DT') or datetime.min))
            if not (self.lote_info and self.lote_info.get('revision')) or r.get('En_Periodo', True)
        ]
        enviado_data = [
            {
                "Ticket único del": f['fecha'].strftime('%d/%m/%Y'),
                "Asunto": f['asunto'],
                "Incidente enviado": f['ticket'],
                "No. Cliente": f['cliente'],
                "Crédito": f['credito'],
                "Incidente en Outlook": f['incidente'],
                "Revisión": self._revision_enviado(f)[0],
                "Base de datos (SQL)": f.get('sql_detalle', ''),
            }
            for f in ((self.lote_info or {}).get('semana') or {}).get('filas') or []
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
            if revision_data and self.lote_info:
                pd.DataFrame(revision_data).to_excel(writer, sheet_name="5_Revision_de_Envios", index=False)
            if enviado_data:
                pd.DataFrame(enviado_data).to_excel(writer, sheet_name="6_Enviado_en_Tickets_Unicos", index=False)

            from openpyxl.styles import PatternFill, Font, Alignment
            wb = writer.book
            hdr_fill = PatternFill(start_color="0B3A6B", end_color="0B3A6B", fill_type="solid")
            hdr_font = Font(name="Segoe UI", color="FFFFFF", bold=True)
            link_font = Font(name="Segoe UI", color="0B5CAD", underline="single")
            fill_mult = PatternFill(start_color="FFFBEB", end_color="FFFBEB", fill_type="solid")
            fill_inc = PatternFill(start_color="FEF2F2", end_color="FEF2F2", fill_type="solid")
            fill_rez = PatternFill(start_color="EEF2FF", end_color="EEF2FF", fill_type="solid")
            fill_cerr = PatternFill(start_color="F5F3FF", end_color="F5F3FF", fill_type="solid")
            fill_no_env = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
            fill_tarde = PatternFill(start_color="FFFBEB", end_color="FFFBEB", fill_type="solid")

            if "6_Enviado_en_Tickets_Unicos" in wb.sheetnames:
                ws6 = wb["6_Enviado_en_Tickets_Unicos"]
                col_rev = [c.value for c in ws6[1]].index("Revisión")
                for row in ws6.iter_rows(min_row=2):
                    rev = str(row[col_rev].value or "")
                    fill = fill_no_env if rev.startswith("DATOS") else (fill_tarde if rev.startswith("No se encontró") else None)
                    if fill:
                        for cell in row:
                            cell.fill = fill

            if "5_Revision_de_Envios" in wb.sheetnames:
                ws5 = wb["5_Revision_de_Envios"]
                for row in ws5.iter_rows(min_row=2):
                    res = str(row[0].value or "")
                    fill = fill_no_env if res.startswith("NO ENVIADA") else (fill_tarde if res == "Enviada tarde" else None)
                    if fill:
                        for cell in row:
                            cell.fill = fill

            for ws in wb.worksheets:
                for cell in ws[1]:
                    cell.fill = hdr_fill
                    cell.font = hdr_font
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                ws.row_dimensions[1].height = 30
                for col in ws.columns:
                    max_len = max((len(str(c.value or "")) for c in col), default=12)
                    ws.column_dimensions[col[0].column_letter].width = min(max(max_len + 4, 12), 48)
                ws.freeze_panes = "A2"
                if ws.max_row > 1:
                    ws.auto_filter.ref = ws.dimensions

            ws1 = wb["1_Desglose_Renovaciones"]
            if desglose_data:
                ticket_col = next(
                    (c.column for c in ws1[1] if c.value == "Ticket iTop"), 3
                )
                for row_idx, r_orig in enumerate(active_rows, start=2):
                    t_cell = ws1.cell(row=row_idx, column=ticket_col)
                    tnum = str(t_cell.value or "")
                    if tnum.upper().startswith("I-"):
                        t_cell.hyperlink = self._itop_url(tnum)
                        t_cell.font = link_font

                    row_fill = None
                    if r_orig.get("No_Enviada"):
                        row_fill = fill_no_env
                    elif r_orig.get("Datos_Incompletos"):
                        row_fill = fill_inc
                    elif r_orig.get("Cerrado_iTop"):
                        row_fill = fill_cerr
                    elif r_orig.get("Lote_Estado") == LOTE_REZAGADA:
                        row_fill = fill_rez
                    elif r_orig.get("Es_Multiple"):
                        row_fill = fill_mult
                    if row_fill:
                        for c_idx in range(1, ws1.max_column + 1):
                            ws1.cell(row=row_idx, column=c_idx).fill = row_fill


if __name__ == "__main__":
    app = FonacotValidadorRenovacionesApp()
    app.mainloop()
