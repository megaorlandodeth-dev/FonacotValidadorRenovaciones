import os
import re
import sys
import json
import colorsys
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
    SQL_APP_PLAZO,
    SQL_MESA_CREDITO,
    SQL_PADRE_CORREGIDO,
)
import reglas_lote as reglas
from reglas_lote import aplicar_reglas, va_en_lote, veredicto, ICONOS
from calendario_cedulas import actualizar_calendario, marcar_bloqueos, cortes_alrededor, cortes_entre, URL_CALENDARIO
from core_renovaciones import (
    build_itop_url,
    build_itop_search_url,
    guardar_json,
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
    llaves_fila,
    aplicar_correccion,
    aplicar_correcciones,
    quitar_correccion,
    build_ticket_unico_tsv,
    build_ticket_unico_html,
    set_clipboard_html,
    plain_text,
    lote_corto,
    pl,
    INTRO_TICKET_UNICO_DEFAULT
)
import equipo_compartido as equipo
import registro


APP_VERSION = "3.3.7"
# Trabajo en equipo por carpeta compartida (lo que marcan los analistas le llega al validador): apagado por ahora
EQUIPO_ACTIVO = False

# Perfiles: el usuario (analista) solo ve su lote del día por su gafete; el validador ve todo y arma el ticket único
PERFIL_USUARIO, PERFIL_VALIDADOR = "usuario", "validador"
PERFILES = {PERFIL_USUARIO: "Usuario", PERFIL_VALIDADOR: "Validador"}

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

# Modo oscuro: la misma paleta con la luminosidad invertida (se elige en Configuración y se aplica al abrir)
_TEMA = {"oscuro": False}


def _luminancia(r: float, g: float, b: float) -> float:
    """Luminancia relativa (WCAG) de un color con canales 0..1."""
    def lineal(c):
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    return 0.2126 * lineal(r) + 0.7152 * lineal(g) + 0.0722 * lineal(b)


def _oscurecer(color: str) -> str:
    """#RRGGBB claro -> su equivalente oscuro (fondos claros a oscuros, textos oscuros a claros, mismo tono)."""
    r, g, b = (int(color[i:i + 2], 16) / 255 for i in (1, 3, 5))
    h, lum, s = colorsys.rgb_to_hls(r, g, b)
    lum = 0.09 + (1 - lum) * 0.84
    if lum < 0.25:
        s *= 0.55  # fondos: tinte apagado, sin colores chillones
    s = min(s, 1)
    # Textos y acentos: el rojo, el violeta o el índigo invertidos quedaban apagados sobre el fondo oscuro.
    # Se aclaran hasta tener contraste de al menos 5:1 con el fondo (los grises tenues, 3:1).
    piso = 0.3 if lum >= 0.45 else 0.15 if lum >= 0.38 else 0
    while lum < 0.92 and _luminancia(*colorsys.hls_to_rgb(h, lum, s)) < piso:
        lum += 0.02
    r, g, b = colorsys.hls_to_rgb(h, lum, s)
    return "#{:02X}{:02X}{:02X}".format(*(round(v * 255) for v in (r, g, b)))


def C(color: str) -> str:
    """Color fijo del código según el tema (claro tal cual; oscuro, invertido)."""
    return _oscurecer(color) if _TEMA["oscuro"] else color

FILTRO_TODAS = "Todas"
FILTRO_PENDIENTES = "Pendientes"
FILTRO_ALERTAS = "Revisar"
FILTRO_MULTIPLES = "Múltiples"
FILTRO_REZAGADAS = "Rezagadas"
FILTRO_NUEVAS = "Nuevas"
FILTRO_LISTAS = "Listas para enviar"
FILTRO_VALIDADAS = "Validadas"
FILTRO_ENVIADAS = "Ya enviadas"
FILTRO_INFORMATIVAS = "No se mandan"  # las del lote: de su ticket único o de después
FILTRO_INFORMATIVAS_TODAS = "No se mandan (todas)"  # de todos los días
FILTRO_SIGUIENTE = "Para el siguiente lote"
FILTRO_PAPELERA = "Descartadas"
FILTRO_CERRADAS = "Cerradas en iTop"
FILTRO_NO_ENVIADAS = "No enviadas en su día"
FILTRO_TARDE = "Enviadas tarde"
FILTRO_SIN_COMPROBAR = "Sin comprobar"
FILTRO_ERRORES = "Datos equivocados en el ticket único"
FILTRO_SQL = "No cuadran con la base (SQL)"
FILTRO_EQUIPO = "Marcadas por el equipo"
FILTRO_DUDAS = "Dudas del equipo"
FILTRO_AVISO = "Las del aviso"  # las renovaciones del aviso al que se le dio clic
OTRAS_VISTAS = "Otras vistas"

FILTROS_PRINCIPALES = [FILTRO_TODAS, FILTRO_PENDIENTES, FILTRO_ALERTAS, FILTRO_MULTIPLES, FILTRO_REZAGADAS, FILTRO_NUEVAS]
FILTROS_USUARIO = [FILTRO_TODAS, FILTRO_PENDIENTES, FILTRO_ALERTAS, FILTRO_ENVIADAS, FILTRO_INFORMATIVAS]
# Filtros cuyo número ya está en un indicador (En el lote, Avance, Revisar, Ya enviadas, No se mandan): no se repite
FILTROS_SIN_CONTEO = {FILTRO_TODAS, FILTRO_PENDIENTES, FILTRO_ALERTAS, FILTRO_ENVIADAS, FILTRO_INFORMATIVAS}
FILTROS_SECUNDARIOS = ([FILTRO_EQUIPO, FILTRO_DUDAS] if EQUIPO_ACTIVO else []) + [FILTRO_LISTAS, FILTRO_INFORMATIVAS, FILTRO_INFORMATIVAS_TODAS, FILTRO_NO_ENVIADAS, FILTRO_ERRORES, FILTRO_SQL, FILTRO_TARDE,
                       FILTRO_SIN_COMPROBAR, FILTRO_CERRADAS, FILTRO_VALIDADAS, FILTRO_ENVIADAS, FILTRO_SIGUIENTE,
                       FILTRO_PAPELERA]
# Vistas que muestran filas aunque estén fuera del lote de hoy (revisión de envíos e informativas)
FILTROS_REVISION = (FILTRO_ENVIADAS, FILTRO_SIGUIENTE, FILTRO_NO_ENVIADAS, FILTRO_ERRORES, FILTRO_SQL, FILTRO_TARDE,
                    FILTRO_SIN_COMPROBAR, FILTRO_INFORMATIVAS, FILTRO_INFORMATIVAS_TODAS, FILTRO_CERRADAS)

# Colores del veredicto de cada renovación (banner del desglose): (fondo, borde, texto, ícono)
VEREDICTO_ESTILO = {
    reglas.LISTA: ("#ECFDF5", "#A7F3D0", "#166534", "✔"),
    reglas.REVISAR: ("#FFFBEB", "#FDE68A", "#92400E", "⚠"),
    reglas.AUTORIZACION: ("#FFFBEB", "#FDE68A", "#92400E", "⚠"),
    reglas.CORREGIR: ("#FEF2F2", "#FECACA", "#991B1B", "✖"),
    reglas.NO_SE_MANDA: ("#F1F5F9", "#CBD5E1", "#334155", "⊘"),
    reglas.ENVIADA: ("#EAF2FB", "#BFDBFE", "#0B5CAD", "✉"),
    reglas.SIGUIENTE: ("#F1F5F9", "#CBD5E1", "#334155", "➜"),
    reglas.DESCARTADA: ("#F8FAFC", "#E2E8F0", "#64748B", "–"),
    reglas.SIN_COMPROBAR: ("#FFFBEB", "#FDE68A", "#92400E", "?"),
}
NIVELES_REVISAR = (reglas.REVISAR, reglas.CORREGIR, reglas.AUTORIZACION)


def _tema_guardado() -> str:
    """El tema elegido («claro» u «oscuro»), leído antes de crear la ventana."""
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f).get("tema", "claro")
    except (OSError, ValueError, AttributeError):
        return "claro"


def _aplicar_paleta(oscuro: bool):
    """Pasa la paleta al modo oscuro (una sola vez, antes de crear widgets)."""
    if not oscuro or _TEMA["oscuro"]:
        return
    _TEMA["oscuro"] = True
    for k, v in COL.items():
        COL[k] = _oscurecer(v)
    for k, (fondo, borde, texto, icono) in VEREDICTO_ESTILO.items():
        VEREDICTO_ESTILO[k] = (_oscurecer(fondo), _oscurecer(borde), _oscurecer(texto), icono)
HINT_EXTRAER = "Lee Outlook y arma el lote del día  ·  F5"
TIPOS_REVISION = ("Lote del día", "Revisión semanal")

MODOS = {"CONCENTRADO": "Por ticket único", "AGENTE": "Por compañero"}
DIAS_SEMANA = ("lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo")

# Mismo asunto que el correo que ya se envía: "Cancelación de renovaciones – ticket único 22-10-2026"
ASUNTO_BORRADOR_DEFAULT = "Cancelación de renovaciones – ticket único {fecha}"
ASUNTO_BORRADOR_ANTERIOR = "TICKET UNICO {fecha}"

# Columnas que se ocultan de inicio para no saturar la tabla (se activan en «Columnas» o con clic derecho en el encabezado)
COLUMNAS_OCULTAS_DEFAULT = {
    # Parte, gafete, agente, alerta y llegada están en el desglose del ticket; se muestran en «Columnas»
    "desglose": ["parte", "credito_padre", "gafete", "agente", "alerta_reno", "alerta_asig", "motivo", "fecha"],
    "tickets": ["creditos_padres", "alerta_asig"],
}
COLUMNAS_FIJAS = ("validado", "ticket")
VISTA_DEFAULT = {"indicadores": True, "detalle": True, "configuracion": True}
VISTA_ZONAS = (
    ("indicadores", "Indicadores (tarjetas)"),
    ("detalle", "Desglose del ticket y correo"),
    ("configuracion", "Panel izquierdo (periodo y configuración)  ·  Ctrl+B"),
)
POSICIONES_DETALLE = (("derecha", "Desglose a la derecha de la tabla"), ("abajo", "Desglose debajo de la tabla"))

ATAJOS = [
    ("F5", "Extraer y validar renovaciones"),
    ("Ctrl + K", "Buscar cualquier acción escribiendo (filtros, copiar, correo, tema…)"),
    ("Ctrl + F", "Ir al buscador (Esc lo limpia)"),
    ("Espacio / doble clic", "Marcar o desmarcar como validada (y pasa a la siguiente pendiente)"),
    ("Ctrl + Z", "Deshacer lo último (validar, descartar, restaurar…)"),
    ("N", "Ir a la siguiente pendiente"),
    ("E", "Corregir el cliente o el crédito a cancelar (por ejemplo, con la renovación que dice la base)"),
    ("J  /  K", "Renovación siguiente / anterior"),
    ("?", "Ver estos atajos"),
    ("Enter", "Abrir el ticket seleccionado en iTop"),
    ("Supr", "Descartar las filas seleccionadas"),
    ("Ctrl + C", "Copiar las filas seleccionadas"),
    ("Ctrl + Shift + C", "Copiar el lote (formato del ticket único, sin saludo)"),
    ("Ctrl + M", "Crear el correo 'Ticket Único' en Outlook"),
    ("Ctrl + E", "Exportar a Excel"),
    ("Clic derecho", "Más acciones: copiar cliente o crédito, buscar en iTop…"),
    ("Ctrl + B  /  ☰", "Ocultar o mostrar el panel izquierdo"),
    ("Vista / clic derecho en el encabezado", "Elegir columnas, dónde va el desglose y qué zonas se ven"),
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
            self.tip, text=texto, justify="left", wraplength=520, font=(FONT, 10), bg=C("#FFFFFF"), fg=COL["text"],
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
    ctk.set_appearance_mode("Dark" if _TEMA["oscuro"] else "Light")
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
        self.principales = list(FILTROS_PRINCIPALES)
        self._etiquetas = {f: f for f in self.principales}  # filtro -> texto del botón (con su conteo)
        seg.configure(command=self._desde_seg)
        menu.configure(command=self._desde_menu)
        self.set(FILTRO_TODAS)

    def configurar(self, principales: list, secundarios: list):
        """Cambia los filtros (cada perfil tiene los suyos)."""
        self.principales = list(principales)
        self._etiquetas = {f: f for f in self.principales}
        self.seg.configure(values=self.principales)
        self.menu.configure(values=[OTRAS_VISTAS] + [f for f in secundarios if f not in self.principales])
        self.set(self._valor if self._valor in self.principales + list(secundarios) else FILTRO_TODAS)

    def contar(self, conteos: dict):
        """Pone cuántas renovaciones hay en los filtros que no tienen indicador: «Rezagadas 1»."""
        etiquetas = {f: f"{f}  {conteos[f]}" if conteos.get(f) is not None and f not in FILTROS_SIN_CONTEO else f
                     for f in self.principales}
        if etiquetas != self._etiquetas:
            self._etiquetas = etiquetas
            self.seg.configure(values=[etiquetas[f] for f in self.principales])
            self.set(self._valor)

    def get(self) -> str:
        return self._valor

    def set(self, valor: str):
        self._valor = valor
        if valor in self.principales:
            self.seg.set(self._etiquetas.get(valor, valor))
            self.menu.set(OTRAS_VISTAS)
            self.menu.configure(fg_color=COL["track"], button_color=COL["track"], text_color=COL["text_2"])
        else:
            self.seg.set("")
            self.menu.set(valor)
            self.menu.configure(fg_color=COL["primary_soft"], button_color=COL["primary_soft"], text_color=COL["primary"])

    def _desde_seg(self, etiqueta):
        self.set(next((f for f, e in self._etiquetas.items() if e == etiqueta), etiqueta))
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

    def mostrar(self, nombre: str, visible: bool):
        """Muestra u oculta la pestaña (el perfil Usuario no ve las de revisión)."""
        box = self._tabs[nombre][0].master
        if visible:
            box.grid()
        else:
            box.grid_remove()
            if self._actual == nombre:
                self.set(next(iter(self._tabs)))


class FonacotValidadorRenovacionesApp(ctk.CTk):
    def __init__(self):
        _aplicar_paleta(_tema_guardado() == "oscuro")
        _aplicar_tema()
        super().__init__()
        self.report_callback_exception = self._error_en_boton

        self.title("Validador de Renovaciones · FONACOT")
        self._fit_to_screen()
        self.minsize(1000, 600)
        self.configure(fg_color=COL["bg"])
        self._angosta = False
        self._panel_auto_oculto = False
        self.bind("<Configure>", self._panel_segun_ancho, add="+")

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
        self._bloqueadas_avisadas = 0
        self._sort_state = {}
        self._headers = {}
        self._celdas_completas = {}   # {(iid, columna): texto completo} para el tooltip de celdas resumidas
        self._vars_columnas = []
        self._secciones = {}          # {panel: [(título, marco, línea, etiqueta, primera)]} para ocultarlas por perfil
        self._veredictos = {}         # {Row_ID: veredicto} del último refresco (tabla, indicadores y desglose)
        self._deshacer = []           # [(descripción, función que lo revierte)] lo último arriba
        self._barras = {}             # {color: imagen} de la barra de estado de cada fila
        # Trabajo en equipo: lo que publicaron los analistas en la carpeta compartida (lo lee el validador)
        self._equipo = {"analistas": [], "marcas": {}, "errores": []}
        self._equipo_vistas = set()   # llaves ya avisadas (para avisar solo lo nuevo)
        self._equipo_leyendo = False
        self._publicar_job = None
        self._equipo_estado = ""

        self._build_ui()
        self._build_context_menus()
        self._bind_shortcuts()

        self.after(50, self._maximize)
        self.after(300, self._init_accounts)
        self.after(500, self._tick_clock)
        self.after(800, self._schedule_auto_refresh)
        self.after(2500, self._equipo_tick)
        if self.config_data.get("perfil") in PERFILES:
            self.after(1500, lambda: self._recordatorio_emision(auto=True))
        else:
            self.after(600, self._elegir_perfil)

    # ================= CONFIGURACIÓN =================

    def _load_config(self) -> dict:
        default_cfg = {
            # "usuario" (analista: su lote del día por gafete) o "validador" (todo). Vacío = se pregunta al abrir.
            "perfil": "",
            # True: no se puede cambiar de perfil desde la app (lo fija quien instala)
            "perfil_fijo": False,
            # Al validar con Espacio se pasa sola a la siguiente renovación pendiente
            "avanzar_al_validar": True,
            # Carpeta compartida del equipo: los analistas publican lo que marcan y el validador lo ve (vacía = solo local)
            "carpeta_compartida": "",
            # "claro" u "oscuro" (se aplica al abrir) y letra grande en la tabla
            "tema": "claro",
            "letra_grande": False,
            # Pedir el motivo al descartar o meter a mano una renovación (queda en la bitácora)
            "pedir_motivo": True,
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
            # Recordatorio del calendario de emisión: se abre solo cuando faltan estos días (o menos) para un corte
            "aviso_emision_dias": 5,
            "recordatorio_emision_visto": "",
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
            # Ya se pasó a la vista compacta de columnas (v3.2): se hace una sola vez sobre lo guardado
            "columnas_compactas": False,
            # Zonas de la pantalla que se pueden ocultar desde «Vista» para dar más espacio a la tabla
            "vista": dict(VISTA_DEFAULT),
            # Desglose del ticket a la derecha o debajo de la tabla, y cuánto ocupa (fracción de la división)
            "detalle_posicion": "derecha",
            "detalle_division": 0.0,
            # Lo plegable: la configuración del panel izquierdo y el detalle del encabezado (los avisos empiezan
            # cerrados en cada sesión)
            "config_abierta": False,
            "info_expandida": False,
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
                if not default_cfg.get("columnas_compactas"):
                    ocultas = default_cfg.setdefault("columnas_ocultas", {}).setdefault("desglose", [])
                    ocultas.extend(c for c in COLUMNAS_OCULTAS_DEFAULT["desglose"] if c not in ocultas)
                    default_cfg["columnas_compactas"] = True
            except Exception:
                pass
        return default_cfg

    def _save_config(self):
        try:
            if self._es_usuario() and self.ent_agente.get().strip():
                self.config_data["gafete_usuario"] = self.ent_agente.get().strip()
            self.config_data.update({
                "account": self.cmb_account.get().strip(),
                # El usuario siempre busca por su gafete: no cambia el modo que el validador dejó elegido
                "modo_operacion": (self.config_data.get("modo_operacion", "CONCENTRADO") if self._es_usuario()
                                   else self.var_modo.get()),
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
                "auto_refresh_min": self.cmb_auto_min.get().strip() or "10",
                "carpeta_compartida": self.ent_carpeta.get().strip(),
            })
            guardar_json(CONFIG_FILE, self.config_data, indent=2, ensure_ascii=False)
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

        # La página va en un lienzo: si la ventana es más baja de lo que necesita (laptops a 125-150 %), sale una barra
        # a la derecha para bajar al desglose y al correo en lugar de cortarlos (ver _ajustar_pagina)
        self._lienzo = tk.Canvas(body, highlightthickness=0, bd=0, bg=COL["bg"], yscrollincrement=20)
        self._lienzo.grid(row=0, column=2, sticky="nsew")
        self._barra_pagina = ctk.CTkScrollbar(body, orientation="vertical", command=self._lienzo.yview)
        self._lienzo.configure(yscrollcommand=self._barra_pagina.set)
        pagina = ctk.CTkFrame(self._lienzo, fg_color="transparent", corner_radius=0)
        pagina.grid_columnconfigure(0, weight=1)
        pagina.grid_rowconfigure(0, weight=1)
        self._pagina_id = self._lienzo.create_window(0, 0, window=pagina, anchor="nw")
        self._lienzo.bind("<Configure>", lambda e: self._ajustar_pagina(), add="+")
        self.bind_all("<MouseWheel>", self._rueda_pagina, add="+")
        self._ajuste_pendiente = None

        main = self._main = ctk.CTkFrame(pagina, fg_color="transparent", corner_radius=0)
        main.grid(row=0, column=0, sticky="nsew", padx=24, pady=(18, 14))
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(3, weight=1)

        self._build_page_header(main)
        self._build_notice(main)
        self._build_kpis(main)
        self._build_work_area(main)
        self._build_status_bar()
        # Lo de arriba de la tabla cambia de alto (avisos, indicadores o bandeja del usuario): se vuelve a medir
        for w in (main.grid_slaves(row=0)[0], self.notice, self.kpi_bar, self.hero):
            for evento in ("<Configure>", "<Map>", "<Unmap>"):
                w.bind(evento, lambda e: self._pedir_ajuste_pagina(), add="+")

        self._on_mode_change()
        self._aplicar_perfil()

    # Alto mínimo (lógico) de la zona de tabla y desglose: con el desglose a la derecha, encabezado, veredicto,
    # comprobaciones y al menos 5 renglones del correo; abajo, además unas filas de la tabla
    ALTO_TRABAJO = {"derecha": 440, "abajo": 620, "sin_detalle": 300}

    def _pedir_ajuste_pagina(self):
        if self._ajuste_pendiente is None:
            self._ajuste_pendiente = self.after_idle(self._ajustar_pagina)

    def _ajustar_pagina(self):
        """La página llena el alto de la ventana; si no cabe lo mínimo, se hace más alta y sale la barra para bajar."""
        self._ajuste_pendiente = None
        lienzo = self._lienzo
        ancho, alto = lienzo.winfo_width(), lienzo.winfo_height()
        if ancho <= 1 or alto <= 1:
            return
        s = self._scaling()
        arriba = sum(w.winfo_reqheight() + int(20 * s) for w in self._main.grid_slaves()
                     if int(w.grid_info().get("row", 3)) < 3)
        vista = self.config_data.get("vista", {})
        tipo = ("sin_detalle" if not vista.get("detalle", True) else
                "derecha" if self._detalle_a_la_derecha() else "abajo")
        necesita = arriba + int((self.ALTO_TRABAJO[tipo] + 32) * s)
        alto_pagina = necesita if necesita > alto + int(6 * s) else alto  # sin barra por 1 o 2 píxeles
        lienzo.itemconfigure(self._pagina_id, width=ancho, height=alto_pagina)
        lienzo.configure(scrollregion=(0, 0, ancho, alto_pagina))
        if alto_pagina > alto:
            self._barra_pagina.grid(row=0, column=3, sticky="ns", padx=(0, 2))
        else:
            self._barra_pagina.grid_remove()
            lienzo.yview_moveto(0)

    def _rueda_pagina(self, e):
        """La rueda sobre la página (fuera de tablas y textos, que tienen su propia barra) la sube o la baja."""
        if not self._barra_pagina.winfo_ismapped():
            return
        w = e.widget
        if not isinstance(w, tk.Misc) or not str(w).startswith(str(self._lienzo)):
            return
        while w is not None and w is not self._lienzo:
            if isinstance(w, (ttk.Treeview, tk.Text, tk.Listbox)):
                return
            w = w.master
        self._lienzo.yview_scroll(-3 if e.delta > 0 else 3, "units")

    def _hline(self, parent) -> ctk.CTkFrame:
        return ctk.CTkFrame(parent, height=1, corner_radius=0, fg_color=COL["border"])

    def _btn(self, parent, text: str, command, kind: str = "secondary", **kw) -> ctk.CTkButton:
        estilos = {
            "primary": dict(fg_color=COL["primary"], hover_color=COL["primary_h"], text_color=C("#FFFFFF"), border_width=0),
            "secondary": dict(
                fg_color=COL["surface"], hover_color=COL["track"], text_color=COL["text"],
                border_width=1, border_color=COL["border_strong"]
            ),
            "danger": dict(
                fg_color=COL["surface"], hover_color=COL["danger_soft"], text_color=COL["danger"],
                border_width=1, border_color=C("#FECACA")
            ),
            "ghost": dict(fg_color="transparent", hover_color=COL["track"], text_color=COL["text_2"], border_width=0),
            "link": dict(fg_color="transparent", hover_color=COL["primary_soft"], text_color=COL["primary"], border_width=0),
        }
        opts = dict(height=32, width=10, corner_radius=6, font=F(12, kind == "primary"))
        opts.update(estilos[kind])
        opts.update(kw)
        return ctk.CTkButton(parent, text=text, command=command, **opts)

    def _build_topbar(self):
        bar = self._topbar = ctk.CTkFrame(self, fg_color=COL["surface"], corner_radius=0)
        bar.grid(row=0, column=0, sticky="ew")
        bar.grid_columnconfigure(1, weight=1)

        brand = self._topbar_brand = ctk.CTkFrame(bar, fg_color="transparent")
        brand.grid(row=0, column=0, padx=(10, 0), pady=11, sticky="w")
        self.btn_panel = ctk.CTkButton(
            brand, text="☰", width=36, height=34, corner_radius=8, font=F(17), fg_color="transparent",
            hover_color=COL["track"], text_color=COL["text_2"], command=self._toggle_sidebar
        )
        self.btn_panel.pack(side="left", padx=(0, 8))
        ctk.CTkLabel(
            brand, text="F", width=34, height=34, corner_radius=8,
            fg_color=COL["primary"], text_color=C("#FFFFFF"), font=F(17, True)
        ).pack(side="left")
        names = ctk.CTkFrame(brand, fg_color="transparent")
        names.pack(side="left", padx=(12, 0))
        ctk.CTkLabel(
            names, text="Validador de Renovaciones", font=F(15, True), text_color=COL["text"], height=20
        ).pack(anchor="w")
        self._top_subtitulo = ctk.CTkLabel(
            names, text="FONACOT  ·  Cancelación y reversa de renovaciones",
            font=F(11), text_color=COL["muted"], height=16
        )
        self._top_subtitulo.pack(anchor="w")

        right = self._topbar_right = ctk.CTkFrame(bar, fg_color="transparent")
        right.grid(row=0, column=2, padx=20, sticky="e")
        # Perfil (Usuario con su gafete / Validador): clic = cambiar de perfil o de gafete
        self.btn_perfil = ctk.CTkButton(
            right, text="", height=30, width=10, corner_radius=15, font=F(12, True), fg_color=COL["primary_soft"],
            hover_color=COL["select"], text_color=COL["primary"], command=self._menu_perfil
        )
        self.btn_perfil.pack(side="left", padx=(0, 10), ipadx=6)
        # Reloj y estado del ticket único de hoy
        self.lbl_corte = ctk.CTkLabel(
            right, text="", font=F(12, True), height=28, corner_radius=14, fg_color=COL["track"],
            text_color=COL["text_2"], cursor="hand2"
        )
        self.lbl_corte.pack(side="left", padx=(0, 10), ipadx=12)
        self.lbl_corte.bind("<Button-1>", lambda e: self._recordatorio_emision())
        self.lbl_clock = ctk.CTkLabel(
            right, text="--:--", font=F(12, True), height=28, corner_radius=14,
            fg_color=COL["track"], text_color=COL["text_2"]
        )
        self.lbl_clock.pack(side="left", padx=(0, 16), ipadx=12)
        self._top_sep = ctk.CTkFrame(right, width=1, height=28, fg_color=COL["border"], corner_radius=0)
        self._top_sep.pack(side="left", padx=(0, 16))
        self._top_buzon = ctk.CTkLabel(right, text="Buzón de Outlook", font=F(12), text_color=COL["muted"])
        self._top_buzon.pack(side="left", padx=(0, 8))
        self.cmb_account = ctk.CTkComboBox(
            right, values=["Cargando cuentas..."], width=260, height=32, font=F(12), dropdown_font=F(12)
        )
        self.cmb_account.pack(side="left")

        self._hline(self).grid(row=0, column=0, sticky="sew")
        # En pantallas angostas (laptop) la barra se compacta en vez de cortar la cuenta de Outlook a la derecha
        self._reloj = ("--:--", "--:--")
        self._top_nivel, self._top_necesita = 0, {}
        bar.bind("<Configure>", lambda e: self._ajustar_topbar(), add="+")

    def _compactar_topbar(self, nivel: int):
        """0: completa. 1: sin «Buzón de Outlook». 2: reloj corto. 3: cuenta más angosta y sin el subtítulo."""
        self._top_nivel = nivel
        if nivel >= 1:
            self._top_sep.pack_forget()
            self._top_buzon.pack_forget()
        elif not self._top_buzon.winfo_ismapped():
            self._top_sep.pack(side="left", padx=(0, 16), before=self.cmb_account)
            self._top_buzon.pack(side="left", padx=(0, 8), before=self.cmb_account)
        self.lbl_clock.configure(text=self._reloj[1] if nivel >= 2 else self._reloj[0])
        self.cmb_account.configure(width=190 if nivel >= 3 else 260)
        if nivel >= 3:
            self._top_subtitulo.pack_forget()
        elif not self._top_subtitulo.winfo_ismapped():
            self._top_subtitulo.pack(anchor="w")

    def _ajustar_topbar(self):
        """Compacta la barra de arriba un paso a la vez mientras no quepa; al agrandar la ventana, la regresa."""
        ancho = self._topbar.winfo_width()
        if ancho <= 1:
            return
        necesita = (self._topbar_brand.winfo_reqwidth() + self._topbar_right.winfo_reqwidth()
                    + int(50 * self._scaling()))
        nivel = self._top_nivel
        if necesita > ancho and nivel < 3:
            self._top_necesita[nivel] = necesita
            self._compactar_topbar(nivel + 1)
            self.after(30, self._ajustar_topbar)
        elif nivel > 0 and ancho >= self._top_necesita.get(nivel - 1, 10 ** 6):
            self._compactar_topbar(nivel - 1)
            self.after(30, self._ajustar_topbar)

    def _section(self, parent, titulo: str, first: bool = False) -> ctk.CTkFrame:
        # Cada sección en su propio marco: el perfil Usuario oculta las que no usa (_mostrar_secciones)
        box = ctk.CTkFrame(parent, fg_color="transparent")
        box.pack(fill="x")
        linea = self._hline(box)
        if not first:
            linea.pack(fill="x", padx=12, pady=(18, 0))
        lbl = ctk.CTkLabel(
            box, text=titulo.upper(), font=F(11, True), text_color=COL["muted"], anchor="w", height=16
        )
        lbl.pack(fill="x", padx=12, pady=(16 if not first else 4, 8))
        inner = ctk.CTkFrame(box, fg_color="transparent")
        inner.pack(fill="x", padx=12)
        inner.grid_columnconfigure(0, weight=1)
        self._secciones.setdefault(str(parent), []).append((titulo, box, linea, lbl, first))
        return inner

    def _mostrar_secciones(self, ocultas: set):
        """Vuelve a acomodar las secciones de cada panel en su orden, sin las `ocultas` (por título)."""
        for lista in self._secciones.values():
            for _, box, *_ in lista:
                box.pack_forget()
            primera = True
            for titulo, box, linea, lbl, first in lista:
                if titulo in ocultas:
                    continue
                antes = {"before": self._sep_config} if box.master is self.sidebar_config else {}
                box.pack(fill="x", **antes)
                linea.pack_forget()
                if not primera:
                    linea.pack(fill="x", padx=12, pady=(18, 0), before=lbl)
                lbl.pack_configure(pady=(4 if primera else 16, 8))
                primera = False

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
        wrap = self.sidebar_wrap = ctk.CTkFrame(parent, fg_color=COL["surface"], corner_radius=0)
        wrap.grid(row=0, column=0, sticky="ns")
        wrap.grid_rowconfigure(1, weight=1)
        wrap.grid_columnconfigure(0, weight=1)
        self.sidebar_sep = ctk.CTkFrame(parent, width=1, corner_radius=0, fg_color=COL["border"])
        self.sidebar_sep.grid(row=0, column=1, sticky="ns")

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
        self.lbl_agente = sec.grid_slaves(row=1, column=0)[0]
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

        # Configuración (se toca poco): plegada para no saturar el panel; «Configuración» la abre
        self._sep_config = self._hline(side)
        self._sep_config.pack(fill="x", padx=12, pady=(18, 0))
        self.btn_config = self._btn(side, "", lambda: self._toggle_config(guardar=True), "ghost", height=34,
                                    font=F(12, True), anchor="w")
        self.btn_config.pack(fill="x", padx=6, pady=(8, 0))
        adv = self.config_avanzada = ctk.CTkFrame(side, fg_color="transparent")

        # Trabajo en equipo: carpeta compartida donde cada analista publica lo que marca y el validador lo ve
        sec = self._section(adv, "Trabajo en equipo", first=True)
        self.ent_carpeta = self._field(sec, 0, "Carpeta compartida del equipo", self.config_data.get("carpeta_compartida", ""),
                                       r"\\servidor\renovaciones")
        botones = ctk.CTkFrame(sec, fg_color="transparent")
        botones.grid(row=2, column=0, columnspan=2, sticky="w", pady=(6, 0))
        self._btn(botones, "Elegir…", self._elegir_carpeta, height=28, font=F(12)).pack(side="left", ipadx=4)
        self._btn(botones, "Probar", self._probar_carpeta, height=28, font=F(12)).pack(side="left", padx=(6, 0), ipadx=4)
        self.lbl_equipo = ctk.CTkLabel(
            sec, text="Los analistas publican lo que marcan; el validador lo ve en su lote.", font=F(11),
            text_color=COL["muted"], justify="left", anchor="w", wraplength=230
        )
        self.lbl_equipo.grid(row=3, column=0, columnspan=2, sticky="w", pady=(6, 0))

        # Apariencia
        sec = self._section(adv, "Apariencia")
        self.var_oscuro = tk.BooleanVar(value=self.config_data.get("tema") == "oscuro")
        ctk.CTkSwitch(
            sec, text="Modo oscuro (al reabrir)", variable=self.var_oscuro, onvalue=True, offvalue=False,
            command=self._cambiar_tema, font=F(12), switch_width=36, switch_height=18
        ).grid(row=0, column=0, columnspan=2, sticky="w")
        self.var_letra = tk.BooleanVar(value=bool(self.config_data.get("letra_grande")))
        ctk.CTkSwitch(
            sec, text="Letra grande en la tabla", variable=self.var_letra, onvalue=True, offvalue=False,
            command=self._cambiar_letra, font=F(12), switch_width=36, switch_height=18
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(8, 0))

        # Corte
        sec = self._section(adv, "Envío del ticket único")
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
        sec = self._section(adv, "Validaciones automáticas")
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
        sec = self._section(adv, "Correo del lote")
        self.ent_asunto_borrador = self._field(
            sec, 0, "Asunto", self.config_data.get("asunto_borrador", ASUNTO_BORRADOR_DEFAULT), ASUNTO_BORRADOR_DEFAULT
        )
        self.ent_correo_para = self._field(sec, 2, "Para", self.config_data.get("correo_para", ""), "correo@fonacot.gob.mx")
        self.ent_correo_cc = self._field(sec, 4, "CC", self.config_data.get("correo_cc", ""), "Opcional")
        ctk.CTkLabel(
            sec, text="{fecha} se cambia por dd-mm-aaaa. El texto del correo se toma del último ticket único.",
            font=F(11), text_color=COL["muted"], justify="left", anchor="w", wraplength=230
        ).grid(row=6, column=0, columnspan=2, sticky="w", pady=(6, 0))

        # Auto-actualizar
        sec = self._section(adv, "Actualización automática")
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
        self._toggle_config(self.config_data.get("config_abierta", False))

    def _toggle_config(self, abrir: bool = None, guardar: bool = False):
        """Abre o pliega la configuración del panel izquierdo (lo que se toca poco)."""
        abrir = (not self.config_data.get("config_abierta", False)) if abrir is None else abrir
        self.config_data["config_abierta"] = abrir
        if abrir:
            self.config_avanzada.pack(fill="x", after=self.btn_config)
        else:
            self.config_avanzada.pack_forget()
        self.btn_config.configure(text=f"⚙   Configuración   {'▾' if abrir else '▸'}")
        if guardar:
            self._save_config()

    def _panel_segun_ancho(self, e):
        """Por debajo de 1200 de ancho (lógico) el panel izquierdo se oculta solo; al agrandar, regresa."""
        if e.widget is not self:
            return
        angosta = e.width / self._scaling() < 1200
        if angosta != self._angosta:
            self._angosta = self._panel_auto_oculto = angosta
            if hasattr(self, "sidebar_wrap"):
                self._aplicar_vista()

    def _toggle_sidebar(self):
        vista = self.config_data.setdefault("vista", dict(VISTA_DEFAULT))
        if self._panel_auto_oculto:
            # Se ocultó solo por el ancho: Ctrl+B lo muestra de todos modos
            self._panel_auto_oculto = False
            vista["configuracion"] = True
            self._aplicar_vista()
            return
        vista["configuracion"] = not vista.get("configuracion", True)
        self._aplicar_vista()
        self._save_config()

    # ================= PERFILES =================

    def _es_usuario(self) -> bool:
        return self.config_data.get("perfil") == PERFIL_USUARIO

    def _aplicar_perfil(self):
        """
        Usuario: su gafete y su periodo (sin revisión semanal, correo del ticket único ni descartes).
        Validador: todo.
        """
        usuario = self._es_usuario()
        gafete = self.ent_agente.get().strip()
        self.btn_perfil.configure(text=f"Usuario  ·  gafete {gafete or '—'}  ▾" if usuario else "Validador  ▾")

        # El usuario elige el día de su lote en «Periodo de recepción» (la revisión semanal y el correo son del validador)
        ocultas = {"¿Qué quieres revisar?", "Correo del lote"} if usuario else set()
        if not EQUIPO_ACTIVO:
            ocultas.add("Trabajo en equipo")
        self._mostrar_secciones(ocultas)
        for titulo, _box, _linea, lbl, _first in self._secciones.get(str(self.sidebar_config), []):
            if titulo == "Modo de validación":
                lbl.configure(text="TU GAFETE" if usuario else titulo.upper())
        self.lbl_agente.configure(text="Tu gafete (lo que iTop te asigna)" if usuario else "Gafete o nombre del compañero")
        if usuario:
            self.seg_modo.grid_remove()
            self.var_modo.set("AGENTE")
            self.seg_tipo.set(TIPOS_REVISION[0])
            if self.cal_start.get_date() != self.cal_end.get_date():
                # Venía de una semana (validador): su lote es de un día
                hoy = datetime.now().date()
                self.cal_start.set_date(hoy)
                self.cal_end.set_date(hoy)
        else:
            self.seg_modo.grid()
            self.var_modo.set(self.config_data.get("modo_operacion", "CONCENTRADO"))
            self.seg_modo.set(MODOS.get(self.var_modo.get(), MODOS["CONCENTRADO"]))
        self._on_mode_change()
        self.lbl_extract_hint.configure(text="Lee tus correos de iTop y arma tu lote de hoy  ·  F5" if usuario
                                        else HINT_EXTRAER, text_color=COL["muted"])

        for nombre in ("agentes", "enviado", "dias"):
            self.tabs.mostrar(nombre, not usuario)

        # Acciones del encabezado: el usuario copia sus renovaciones y se las pasa a quien manda el ticket único
        for b in (self.btn_copiar, self.btn_excel, self.btn_correo):
            b.pack_forget()
        if usuario:
            self.btn_copiar.configure(text="Copiar mis renovaciones", fg_color=COL["primary"], hover_color=COL["primary_h"],
                                      text_color=C("#FFFFFF"), border_width=0, font=F(12, True))
            self.btn_copiar.pack(side="left", ipadx=10)
        else:
            self.btn_copiar.configure(text="Copiar lote", fg_color=COL["surface"], hover_color=COL["track"],
                                      text_color=COL["text"], border_width=1, font=F(12))
            self.btn_copiar.pack(side="left", padx=(0, 8), ipadx=8)
            self.btn_excel.pack(side="left", padx=(0, 8), ipadx=8)
            self.btn_correo.pack(side="left", ipadx=8)

        self.cmb_quick_filter.configurar(FILTROS_USUARIO if usuario else FILTROS_PRINCIPALES,
                                         [FILTRO_INFORMATIVAS_TODAS] if usuario else FILTROS_SECUNDARIOS)
        if usuario:
            self.menu_vistas.pack_forget()
        else:
            self.menu_vistas.pack(side="left", padx=(10, 0))
        for b in (self.btn_validar_visibles, self.btn_equipo, self.btn_descartar, self.btn_restaurar):
            b.pack_forget()
        if not usuario:
            self.btn_validar_visibles.pack(side="left", padx=(8, 0), ipadx=6)
            if self._carpeta_equipo():
                self.btn_equipo.pack(side="left", padx=(8, 0), ipadx=6)
            self.btn_descartar.pack(side="left", padx=(8, 4), ipadx=6)
            self.btn_restaurar.pack(side="left", ipadx=6)
        self._estilo_tabla()
        self._aplicar_vista()
        self._refresh_views()
        self._actualizar_encabezado()

    def _menu_perfil(self):
        m = tk.Menu(self, **self._menu_opts)
        actual = PERFILES.get(self.config_data.get("perfil"), "Validador")
        m.add_command(label=f"Perfil: {actual}", state="disabled")
        if self._es_usuario():
            m.add_command(label="Cambiar mi gafete…", command=self._cambiar_gafete)
        if not self.config_data.get("perfil_fijo"):
            otro = PERFIL_VALIDADOR if self._es_usuario() else PERFIL_USUARIO
            m.add_command(label=f"Cambiar a {PERFILES[otro]}", command=lambda: self._cambiar_perfil(otro))
        m.add_separator()
        m.add_command(label="Atajos de teclado", command=self._show_shortcuts)
        b = self.btn_perfil
        try:
            m.tk_popup(b.winfo_rootx(), b.winfo_rooty() + b.winfo_height() + 2)
        finally:
            m.grab_release()

    def _pedir_gafete(self, titulo: str = "Tu gafete"):
        dlg = ctk.CTkInputDialog(title=titulo, text="Escribe tu gafete (el número con el que iTop te asigna):")
        return re.sub(r"\s+", "", dlg.get_input() or "")

    def _poner_gafete(self, valor: str):
        """Escribe el gafete aunque el campo esté deshabilitado (modo «Por ticket único»: si no, no se cambia)."""
        estado = self.ent_agente.cget("state")
        self.ent_agente.configure(state="normal")
        self.ent_agente.delete(0, "end")
        self.ent_agente.insert(0, valor)
        self.ent_agente.configure(state=estado)

    def _cambiar_gafete(self):
        valor = self._pedir_gafete()
        if valor:
            self.config_data["gafete_usuario"] = valor
            self._poner_gafete(valor)
            self._save_config()
            self._aplicar_perfil()

    def _cambiar_perfil(self, perfil: str, gafete: str = None) -> bool:
        """
        Cada perfil guarda su gafete: el del usuario es el suyo; el del validador, el compañero que estaba revisando
        (así el usuario no hereda el gafete que el validador tenía puesto).
        """
        previo = self.config_data.get("perfil")
        if self.is_extracting and perfil != previo:
            # Lo que se está buscando es del perfil anterior: al terminar se mostraría con el perfil nuevo
            messagebox.showinfo("Cambiar de perfil", "Espera a que termine la búsqueda para cambiar de perfil.")
            return False
        if previo == PERFIL_USUARIO and self.ent_agente.get().strip():
            self.config_data["gafete_usuario"] = self.ent_agente.get().strip()
        if perfil == PERFIL_USUARIO:
            gafete = gafete or self.config_data.get("gafete_usuario") or self._pedir_gafete()
            if not gafete:
                return False
        if previo == PERFIL_VALIDADOR:
            self.config_data["agente_validador"] = self.ent_agente.get().strip()
        if perfil == PERFIL_VALIDADOR:
            gafete = gafete or self.config_data.get("agente_validador") or ""
        else:
            self.config_data["gafete_usuario"] = gafete
        self.config_data["perfil"] = perfil
        if perfil == PERFIL_VALIDADOR and previo == PERFIL_USUARIO:
            # El usuario busca siempre «Por compañero» (su gafete): el validador regresa al modo que tenía (si no,
            # _save_config guardaba «Por compañero» como su modo y ya no volvía a «Por ticket único»)
            self.var_modo.set(self.config_data.get("modo_operacion", "CONCENTRADO"))
        if gafete or perfil == PERFIL_VALIDADOR:
            self._poner_gafete(gafete)
        if perfil != previo:
            # La actualización automática repetía la última búsqueda del otro perfil (al usuario le hacía la revisión
            # semanal del validador): cada perfil empieza con la suya
            self._ultima_busqueda = None
            self._prev_scan_signature = None
        if perfil != previo and previo in PERFILES and (self.all_rows or self.lote_info):
            # El lote del validador (o una revisión) no es del usuario, ni el del usuario es todo el lote del validador:
            # se limpia hasta que se extraiga el suyo
            self.all_rows, self.lote_info = [], None
            self.validaciones = {}
            self._limpiar_detalle()
            self.lbl_status.configure(text="Presiona «Extraer y validar» (F5) para ver tus renovaciones."
                                      if perfil == PERFIL_USUARIO else
                                      "Presiona «Extraer y validar» (F5) para ver el lote completo.")
        self._save_config()
        self._aplicar_perfil()
        return True

    def _elegir_perfil(self):
        """Primera vez: ¿usuario (analista) o validador? El usuario escribe su gafete."""
        win = ctk.CTkToplevel(self)
        win.title("Bienvenido · Validador de Renovaciones")
        win.configure(fg_color=COL["bg"])
        win.resizable(False, False)
        win.transient(self)
        ctk.CTkLabel(win, text="¿Cómo vas a usar el validador?", font=F(22, True), text_color=COL["text"]).pack(
            anchor="w", padx=28, pady=(26, 2))
        ctk.CTkLabel(win, text="Elige tu perfil. Lo puedes cambiar después con el botón de perfil (arriba a la derecha).",
                     font=F(12), text_color=COL["muted"]).pack(anchor="w", padx=28)
        cards = ctk.CTkFrame(win, fg_color="transparent")
        cards.pack(fill="both", padx=22, pady=(18, 24))
        elegido = {"listo": False}

        def entrar(perfil, gafete=None):
            if perfil == PERFIL_USUARIO and not gafete:
                ent.configure(border_color=COL["danger"])
                ent.focus_set()
                return
            if not self._cambiar_perfil(perfil, gafete):
                return
            elegido["listo"] = True
            win.destroy()
            self.after(400, lambda: self._recordatorio_emision(auto=True))

        def tarjeta(col, titulo, puntos, color):
            c = ctk.CTkFrame(cards, fg_color=COL["surface"], corner_radius=12, border_width=1, border_color=COL["border"])
            c.grid(row=0, column=col, padx=6, sticky="nsew")
            ctk.CTkFrame(c, height=4, corner_radius=2, fg_color=color).pack(fill="x", padx=18, pady=(16, 0))
            ctk.CTkLabel(c, text=titulo, font=F(17, True), text_color=COL["text"]).pack(anchor="w", padx=18, pady=(12, 6))
            for p in puntos:
                ctk.CTkLabel(c, text=f"•  {p}", font=F(12), text_color=COL["text_2"], justify="left", anchor="w",
                             wraplength=290).pack(anchor="w", padx=18, pady=1)
            return c

        cu = tarjeta(0, "Usuario (analista)", ["Tus renovaciones de hoy, por tu gafete.",
                                              "Qué va, qué revisar y qué ya no se manda.",
                                              "Copias tus renovaciones para quien manda el ticket único."], COL["primary"])
        ctk.CTkLabel(cu, text="Tu gafete", font=F(12), text_color=COL["text_2"]).pack(anchor="w", padx=18, pady=(14, 3))
        ent = ctk.CTkEntry(cu, height=34, font=F(13), placeholder_text="Ej. 2609")
        ent.pack(fill="x", padx=18)
        # Solo el gafete que ya se dio como usuario (no el del compañero que tenía puesto el validador)
        if self.config_data.get("gafete_usuario"):
            ent.insert(0, self.config_data["gafete_usuario"])
        ent.bind("<Return>", lambda e: entrar(PERFIL_USUARIO, ent.get().strip()))
        self._btn(cu, "Entrar como usuario", lambda: entrar(PERFIL_USUARIO, ent.get().strip()), "primary",
                  height=38).pack(fill="x", padx=18, pady=(12, 18))
        cv = tarjeta(1, "Validador", ["El lote de todos o de un compañero, y la revisión semanal.",
                                      "Valida renovación por renovación con el desglose.",
                                      "Arma y manda el correo del ticket único."], C("#0F766E"))
        ctk.CTkFrame(cv, fg_color="transparent", height=1).pack(expand=True, fill="both", padx=18)
        self._btn(cv, "Entrar como validador", lambda: entrar(PERFIL_VALIDADOR), height=38).pack(
            fill="x", side="bottom", padx=18, pady=(12, 18))
        cards.grid_columnconfigure((0, 1), weight=1, uniform="perfil")

        def cerrar():
            # Sin elegir: como usuario; el gafete y el perfil se cambian en el botón de perfil
            win.destroy()
            if not elegido["listo"]:
                self.config_data["perfil"] = PERFIL_USUARIO
                self._save_config()
                self._aplicar_perfil()
        win.protocol("WM_DELETE_WINDOW", cerrar)
        win.after(150, lambda: (win.lift(), win.focus_force(), ent.focus_set()))
        try:
            win.grab_set()
        except tk.TclError:
            pass

    def _build_page_header(self, parent):
        hdr = ctk.CTkFrame(parent, fg_color="transparent")
        hdr.grid(row=0, column=0, sticky="ew")
        hdr.grid_columnconfigure(0, weight=1)

        self.lbl_page_title = ctk.CTkLabel(hdr, text="Lote del día", font=F(20, True), text_color=COL["text"], anchor="w")
        self.lbl_page_title.grid(row=0, column=0, sticky="w")
        info = ctk.CTkFrame(hdr, fg_color="transparent")
        info.grid(row=1, column=0, sticky="ew")
        # «Más detalle» va pegado al texto; lo que sobra queda en la columna 2
        info.grid_columnconfigure(2, weight=1)
        self.lbl_lote_info = ctk.CTkLabel(
            info, text=self._texto_marca_guardada(), font=F(12), text_color=COL["muted"],
            justify="left", anchor="w", wraplength=800
        )
        self.lbl_lote_info.grid(row=0, column=0, sticky="w")
        self.btn_info_mas = self._btn(info, "", self._toggle_info, "link", height=22, font=F(11))
        self.btn_info_mas.grid(row=0, column=1, sticky="nw", padx=(8, 0))
        self.btn_info_mas.grid_remove()
        self._info_lineas = ([self._texto_marca_guardada()], [])

        actions = ctk.CTkFrame(hdr, fg_color="transparent")
        actions.grid(row=0, column=1, rowspan=2, sticky="ne", padx=(16, 0))
        # Con el panel izquierdo oculto, «Extraer y validar» se pasa aquí
        self.btn_extract_hdr = self._btn(actions, "Extraer y validar", self._start_extraction)
        self._primer_accion = self.btn_copiar = self._btn(actions, "Copiar lote", self._copy_batch_to_clipboard)
        self.btn_copiar.pack(side="left", padx=(0, 8), ipadx=8)
        self.btn_excel = self._btn(actions, "Exportar a Excel", self._export_excel)
        self.btn_excel.pack(side="left", padx=(0, 8), ipadx=8)
        self.btn_correo = self._btn(actions, "Crear correo Ticket Único", self._crear_correo_lote, "primary")
        self.btn_correo.pack(side="left", ipadx=8)

        # e.width llega en píxeles reales; wraplength de CTk va en unidades lógicas. Se mide el espacio del texto
        # (no todo el encabezado) y se deja lugar para «Más detalle».
        info.bind(
            "<Configure>",
            lambda e: self.lbl_lote_info.configure(wraplength=max(300, int(e.width / self._scaling()) - 130))
        )

    def _pintar_info(self):
        """Una línea con lo principal; el resto, con «Más detalle»."""
        principal, extra = self._info_lineas
        abierta = self.config_data.get("info_expandida", False)
        texto = "\n".join(t for t in principal + (extra if abierta else []) if t)
        self.lbl_lote_info.configure(text=texto)
        if texto:
            self.lbl_lote_info.grid()
        else:
            self.lbl_lote_info.grid_remove()
        if extra:
            self.btn_info_mas.configure(text="Menos detalle ▴" if abierta else "Más detalle ▾")
            self.btn_info_mas.grid()
        else:
            self.btn_info_mas.grid_remove()

    def _toggle_info(self):
        self.config_data["info_expandida"] = not self.config_data.get("info_expandida", False)
        self._pintar_info()

    def _build_notice(self, parent):
        """
        Avisos bajo el título, uno por renglón: rojo si algo cambia el lote (no se envió, datos equivocados),
        ámbar para lo demás. Clic en un aviso = ver solo esas renovaciones, con el detalle de la primera.
        """
        self.notice = ctk.CTkFrame(parent, fg_color=COL["warning_soft"], corner_radius=8, border_width=1, border_color=C("#FDE68A"))
        self.notice.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        self.notice_bar = ctk.CTkFrame(self.notice, width=3, height=1, corner_radius=2, fg_color=C("#F59E0B"))
        self.notice_bar.pack(side="left", fill="y", padx=(10, 0), pady=8)
        self.btn_notice_mas = ctk.CTkButton(
            self.notice, text="", width=10, height=24, corner_radius=6, font=F(11, True), fg_color="transparent",
            hover_color=C("#FDE68A"), text_color=C("#92400E"), command=self._toggle_avisos
        )
        self.notice_lista = ctk.CTkFrame(self.notice, fg_color="transparent")
        self.notice_lista.pack(side="left", fill="x", expand=True, padx=12, pady=4)
        self._notice_ancho = 900
        self.notice.bind("<Configure>", self._ajustar_avisos)
        self._notice_labels = []
        self._avisos = []
        self._avisos_abiertos = False
        self._ids_aviso = set()
        self.notice.grid_remove()

    def _ajustar_avisos(self, e=None):
        if e is not None:
            self._notice_ancho = max(300, int(e.width / self._scaling()) - 190)
        for lbl in self._notice_labels:
            lbl.configure(wraplength=self._notice_ancho)

    def _mostrar_aviso(self, avisos, rojo: bool = False):
        """
        Cada aviso es un texto o (texto, destino[, urgente]); destino = lista de Row_ID, un filtro, «pestaña:x» o None;
        urgente = cambia el lote de hoy (va en rojo; lo demás, en ámbar). Se ve el primero (el más importante); los
        demás con «+N avisos».
        """
        avisos = [avisos] if isinstance(avisos, (str, tuple)) else list(avisos)
        avisos = [a if isinstance(a, tuple) else (a, None) for a in avisos]
        self._avisos = [(a[0], a[1], a[2] if len(a) > 2 else False) for a in avisos]
        if rojo:
            self.notice.configure(fg_color=COL["danger_soft"], border_color=C("#FECACA"))
            self.notice_bar.configure(fg_color=C("#DC2626"))
            self.btn_notice_mas.configure(text_color=C("#991B1B"), hover_color=C("#FECACA"))
        else:
            self.notice.configure(fg_color=COL["warning_soft"], border_color=C("#FDE68A"))
            self.notice_bar.configure(fg_color=C("#F59E0B"))
            self.btn_notice_mas.configure(text_color=C("#92400E"), hover_color=C("#FDE68A"))
        self._pintar_avisos()
        self.notice.grid()

    def _pintar_avisos(self):
        avisos = self._avisos
        abierto = self._avisos_abiertos
        for lbl in self._notice_labels:
            lbl.destroy()
        self._notice_labels = []
        varios = len(avisos) > 1
        for texto, destino, urgente in (avisos if abierto else avisos[:1]):
            lbl = ctk.CTkLabel(
                self.notice_lista, text=("•  " if varios and abierto else "") + texto + ("   ›" if destino else ""),
                font=F(12), text_color=C("#991B1B") if urgente else C("#92400E"), justify="left", anchor="w",
                wraplength=self._notice_ancho
            )
            lbl.pack(fill="x", pady=1)
            if destino:
                lbl.configure(cursor="hand2")
                lbl.bind("<Button-1>", lambda e, d=destino: self._ir_a_aviso(d))
                lbl.bind("<Enter>", lambda e, w=lbl: w.configure(font=F(12, True)))
                lbl.bind("<Leave>", lambda e, w=lbl: w.configure(font=F(12)))
            self._notice_labels.append(lbl)
        if varios:
            self.btn_notice_mas.configure(text="Ocultar ▴" if abierto else f"+{pl(len(avisos) - 1, '# [aviso|avisos] más')} ▾")
            self.btn_notice_mas.pack(side="right", padx=(0, 10), pady=6, before=self.notice_lista)
        else:
            self.btn_notice_mas.pack_forget()

    def _toggle_avisos(self):
        # Solo en esta sesión: al abrir la app los avisos empiezan cerrados (se ve el más importante)
        self._avisos_abiertos = not self._avisos_abiertos
        self._pintar_avisos()

    def _ir_a_aviso(self, destino):
        """Clic en un aviso: el desglose muestra solo esas renovaciones y abajo el detalle de la primera."""
        if isinstance(destino, str) and destino.startswith("pestaña:"):
            self.tabs.set(self.tabs_names[destino.split(":", 1)[1]])
            return
        if isinstance(destino, str):
            self._apply_filter(destino)
        else:
            self._ids_aviso = set(destino)
            self.ent_search.delete(0, "end")  # que la búsqueda no esconda ninguna
            self._apply_filter(FILTRO_AVISO)
        if not self.config_data.get("vista", {}).get("detalle", True):
            self._mostrar_zona("detalle", True)
        hijos = self.tree_desglose.get_children()
        if hijos:
            self.tree_desglose.selection_set(hijos[0])
            self.tree_desglose.focus(hijos[0])
            self.tree_desglose.see(hijos[0])
            self._on_select_desglose_row()
            self.after(150, self._division_para_pocas)

    def _division_para_pocas(self):
        """Detalle abajo y pocas filas (≤ 8): la tabla toma solo su alto. No se guarda: al mover la división, gana esa."""
        if self.paned is None or len(self.paned.panes()) < 2 or self._detalle_a_la_derecha():
            return
        hijos = self.tree_desglose.get_children()
        caja = self.tree_desglose.bbox(hijos[0]) if hijos else None
        if not caja or len(hijos) > 8:
            return
        total = self.paned.winfo_height()
        arriba = self.tree_desglose.winfo_rooty() - self.paned.winfo_rooty()
        alto = arriba + caja[1] + caja[3] * len(hijos) + int(30 * self._scaling())
        self.paned.sashpos(0, max(int(total * 0.25), min(alto, int(total * 0.55))))

    def _build_kpis(self, parent):
        self.kpi_bar = kpi_bar = ctk.CTkFrame(parent, fg_color="transparent")
        kpi_bar.grid(row=2, column=0, sticky="ew", pady=(10, 10))
        for i in range(6):
            kpi_bar.grid_columnconfigure(i, weight=1, uniform="kpi")

        # Lo que importa del día, por veredicto: qué va en el ticket único, qué falta revisar y qué ya no se manda
        self.lbl_kpi_renovaciones = self._kpi(kpi_bar, 0, "En el lote", "Van en el ticket único de hoy", C("#0B5CAD"), FILTRO_TODAS)
        self.lbl_kpi_listas = self._kpi(kpi_bar, 1, "Listas", "Todo cuadra: listas para enviar", C("#16A34A"), FILTRO_LISTAS)
        self.lbl_kpi_alertas = self._kpi(kpi_bar, 2, "Revisar", "Corrige o revisa antes de enviar", C("#DC2626"), FILTRO_ALERTAS)
        self.lbl_kpi_enviadas = self._kpi(kpi_bar, 3, "Ya enviadas", "Ya salieron en un ticket único", C("#0F766E"), FILTRO_ENVIADAS)
        self.lbl_kpi_info = self._kpi(kpi_bar, 4, "No se mandan", "Emitidas, resueltas o no proceden", C("#64748B"),
                                      FILTRO_INFORMATIVAS)
        self.lbl_kpi_validados = self._kpi(kpi_bar, 5, "Avance", "Clic para ver pendientes", C("#4F46E5"), FILTRO_PENDIENTES,
                                           progreso=True)

        # Usuario: en vez de indicadores, una bandeja con lo que le falta y, al terminar, el botón para copiar
        self.hero = hero = ctk.CTkFrame(parent, fg_color=COL["surface"], corner_radius=10, border_width=1,
                                        border_color=COL["border"])
        hero.grid(row=2, column=0, sticky="ew", pady=(10, 10))
        hero.grid_remove()
        hero.grid_columnconfigure(0, weight=1)
        self.lbl_hero = ctk.CTkLabel(hero, text="", font=F(20, True), text_color=COL["text"], anchor="w")
        self.lbl_hero.grid(row=0, column=0, sticky="w", padx=18, pady=(12, 0))
        self.lbl_hero_sub = ctk.CTkLabel(hero, text="", font=F(12), text_color=COL["muted"], anchor="w", justify="left")
        self.lbl_hero_sub.grid(row=1, column=0, sticky="w", padx=18)
        self.hero_progress = ctk.CTkProgressBar(hero, height=8, corner_radius=4, progress_color=COL["success"])
        self.hero_progress.grid(row=2, column=0, sticky="ew", padx=18, pady=(8, 14))
        self.hero_progress.set(0)
        self.btn_hero = self._btn(hero, "Copiar mis renovaciones", self._copy_batch_to_clipboard, "primary",
                                  height=42, font=F(14, True))
        self.btn_hero_siguiente = self._btn(hero, "Siguiente pendiente  (N)", self._siguiente_pendiente, height=38,
                                            font=F(13))

    def _pintar_hero(self, active_rows: list, val_count: int):
        """Bandeja del usuario: «Te faltan X de Y» y, cuando termina, «Listo» con el botón de copiar."""
        total, faltan = len(active_rows), len(active_rows) - val_count
        revisar = sum(1 for r in active_rows if r['Row_ID'] not in self.validated_row_ids and self._tiene_alerta(r))
        dudas = sum(1 for r in active_rows if self.estado.duda(self._key_by_rid.get(r['Row_ID'], '')))
        self.btn_hero.grid_forget()
        self.btn_hero_siguiente.grid_forget()
        if not self.all_rows:
            titulo, sub, color = "Extrae tu lote de hoy", "Presiona «Extraer y validar» (F5) para leer tus correos de iTop.", \
                COL["text"]
        elif not total:
            titulo, sub, color = "No tienes renovaciones para el ticket único de hoy", \
                "Lo que llegó ya salió, es del siguiente o no se manda.", COL["text"]
        elif faltan:
            titulo, color = f"Te faltan {faltan} de {total}", COL["text"]
            sub = f"{val_count} validadas  ·  {revisar} de las que faltan tienen algo que revisar" + (
                f"  ·  {pl(dudas, '# [duda|dudas]')} al validador" if dudas else "")
            self.btn_hero_siguiente.grid(row=0, column=1, rowspan=3, padx=18)
        else:
            titulo, color = f"✔ Listo: validaste tus {total} renovaciones", COL["success"]
            sub = ("El validador ya las ve en la carpeta compartida. También puedes copiarlas y pasárselas."
                   if self._carpeta_equipo() else "Cópialas y pásalas a quien manda el ticket único.")
            self.btn_hero.grid(row=0, column=1, rowspan=3, padx=18)
        self.lbl_hero.configure(text=titulo, text_color=color)
        self.lbl_hero_sub.configure(text=sub)
        self.hero_progress.set(val_count / total if total else 0)

    def _kpi(self, parent, col, titulo, subtitulo, acento, filtro, progreso=False):
        # Una sola línea por tarjeta (punto, título y número): ocupa poco alto; con el mouse encima, el título dice qué cuenta
        card = ctk.CTkFrame(parent, fg_color=COL["surface"], corner_radius=8, border_width=1, border_color=COL["border"])
        card.grid(row=0, column=col, padx=(0, 0 if col == 5 else 10), sticky="nsew")
        card.grid_columnconfigure(2, weight=1)
        dot = ctk.CTkFrame(card, width=8, height=8, corner_radius=4, fg_color=acento)
        dot.grid(row=0, column=0, padx=(12, 0), pady=(9, 9 if not progreso else 3))
        lbl_t = ctk.CTkLabel(card, text=titulo.upper(), font=F(10, True), text_color=COL["muted"], height=14)
        lbl_t.grid(row=0, column=1, padx=(8, 0), sticky="w")
        lbl_v = ctk.CTkLabel(card, text="0", font=F(17, True), text_color=COL["text"], anchor="e", height=24)
        lbl_v.grid(row=0, column=2, padx=(8, 12), sticky="e")
        widgets = [card, dot, lbl_t, lbl_v]
        if progreso:
            self.kpi_progress = ctk.CTkProgressBar(card, height=3, corner_radius=2, progress_color=acento)
            self.kpi_progress.grid(row=1, column=0, columnspan=3, sticky="ew", padx=12, pady=(0, 6))
            self.kpi_progress.set(0)
            widgets.append(self.kpi_progress)

        def subtitulo_visible(on):
            lbl_t.configure(text=subtitulo if on else titulo.upper(), font=F(10, not on))

        def hover(on, _e=None):
            if not on:
                w = self.winfo_containing(*self.winfo_pointerxy())
                if w is not None and str(w).startswith(str(card)):
                    return
            card.configure(border_color=COL["primary"] if on else COL["border"])
            subtitulo_visible(on)

        card.bind("<Enter>", lambda e: hover(True))
        for w in widgets:
            w.bind("<Button-1>", lambda e, f=filtro: self._apply_filter(f))
            w.bind("<Leave>", lambda e: hover(False))
        return lbl_v

    def _build_work_area(self, parent):
        # La tabla y el desglose del ticket van en un divisor que se arma según «Vista» (a la derecha o abajo)
        self._work_parent, self.paned = parent, None
        top = ctk.CTkFrame(parent, fg_color=COL["surface"], corner_radius=10, border_width=1, border_color=COL["border"])
        bottom = ctk.CTkFrame(parent, fg_color=COL["surface"], corner_radius=10, border_width=1, border_color=COL["border"])
        self.table_card, self.detail_frame = top, bottom

        top.grid_columnconfigure(0, weight=1)
        top.grid_rowconfigure(3, weight=1)

        # Encabezado: pestañas a la izquierda, contador y buscador a la derecha
        head = ctk.CTkFrame(top, fg_color="transparent")
        head.grid(row=0, column=0, sticky="ew", padx=(12, 16), pady=(6, 0))
        tabs_box = ctk.CTkFrame(head, fg_color="transparent")
        buscar = ctk.CTkFrame(head, fg_color="transparent")
        self._dos_renglones(head, tabs_box, buscar)

        self.lbl_count = ctk.CTkLabel(buscar, text="", font=F(12), text_color=COL["muted"])
        self.lbl_count.pack(side="left", padx=(0, 12))
        # En «Las del aviso»: cómo regresar a todas
        self.btn_salir_aviso = self._btn(buscar, "✕ Ver todas", lambda: self._apply_filter(FILTRO_TODAS), "link",
                                         height=22, font=F(11))
        search = ctk.CTkFrame(buscar, fg_color=COL["surface"], corner_radius=6, border_width=1, border_color=COL["border_strong"])
        search.pack(side="left", pady=6)
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
        # «Vista» es configuración, no filtro: va junto al buscador (en la barra de filtros se cortaba a 125 %)
        self.btn_columnas = self._btn(buscar, "Vista  ▾", self._abrir_menu_columnas)
        self.btn_columnas.pack(side="left", padx=(8, 0), ipadx=6)

        self._hline(top).grid(row=1, column=0, sticky="ew", padx=1)

        # Barra de filtros y acciones sobre filas
        bar = ctk.CTkFrame(top, fg_color="transparent")
        bar.grid(row=2, column=0, sticky="ew", padx=16, pady=8)
        filtros = ctk.CTkFrame(bar, fg_color="transparent")
        acciones = ctk.CTkFrame(bar, fg_color="transparent")
        self._dos_renglones(bar, filtros, acciones)
        self._btn(acciones, "✔  Validar", self._toggle_selected_validation, "primary").pack(side="left", ipadx=10)
        self._btn(acciones, "Siguiente pendiente  ➜", self._siguiente_pendiente).pack(side="left", padx=(8, 0), ipadx=6)
        self.btn_validar_visibles = self._btn(acciones, "Validar visibles", self._validate_all_visible)
        self.btn_validar_visibles.pack(side="left", padx=(8, 0), ipadx=6)
        self.btn_equipo = self._btn(acciones, "Validar lo del equipo", self._aceptar_equipo)
        self.btn_equipo.pack(side="left", padx=(8, 0), ipadx=6)
        self.btn_descartar = self._btn(acciones, "Descartar", self._discard_selected_rows, "danger")
        self.btn_descartar.pack(side="left", padx=(8, 4), ipadx=6)
        self.btn_restaurar = self._btn(acciones, "Restaurar", self._restore_selected_rows, "ghost")
        self.btn_restaurar.pack(side="left", ipadx=6)

        seg = ctk.CTkSegmentedButton(filtros, values=FILTROS_PRINCIPALES, font=F(12), height=32)
        seg.pack(side="left")
        menu = ctk.CTkOptionMenu(
            filtros, values=[OTRAS_VISTAS] + FILTROS_SECUNDARIOS, width=175, height=32, font=F(12),
            dropdown_font=F(12), dynamic_resizing=False
        )
        menu.pack(side="left", padx=(10, 0))
        self.menu_vistas = menu
        self.cmb_quick_filter = SelectorFiltro(seg, menu, self._refresh_views)

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
        self._armar_paneles()

    @staticmethod
    def _dos_renglones(cont, izq, der):
        """`izq` a la izquierda y `der` a la derecha; si no caben juntos (tabla angosta), `der` baja a otro renglón."""
        cont.grid_columnconfigure(1, weight=1)
        izq.grid(row=0, column=0, sticky="sw")
        der.grid(row=0, column=1, sticky="e")

        def acomodar(_e=None):
            ancho = cont.winfo_width()
            if ancho <= 1:
                return
            juntos = izq.winfo_reqwidth() + der.winfo_reqwidth() + 24 <= ancho
            fila = int(der.grid_info().get("row", 0))
            if juntos and fila != 0:
                der.grid(row=0, column=1, columnspan=1, sticky="e", pady=0)
            elif not juntos and fila != 1:
                der.grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))
        cont.bind("<Configure>", acomodar, add="+")

    def _detalle_a_la_derecha(self) -> bool:
        return self.config_data.get("detalle_posicion", "derecha") != "abajo"

    def _armar_paneles(self):
        """Divisor tabla | desglose del ticket: lado a lado (a la derecha) o uno sobre otro (abajo)."""
        derecha = self._detalle_a_la_derecha()
        if self.paned is not None:
            for p in self.paned.panes():
                self.paned.forget(p)
            self.paned.destroy()
        self.paned = ttk.PanedWindow(self._work_parent, orient=tk.HORIZONTAL if derecha else tk.VERTICAL)
        self.paned.grid(row=3, column=0, sticky="nsew")
        self.paned.add(self.table_card, weight=3)
        self.table_card.lift()
        self._colocar_detalle(derecha)
        self.paned.bind("<ButtonRelease-1>", lambda e: self._guardar_division(), add="+")
        self._aplicar_vista()
        self.after(350, self._ajustar_division)

    def _ajustar_division(self, intentos: int = 10):
        """Pone el divisor donde lo dejó el usuario (o 58 % tabla / 42 % desglose)."""
        if self.paned is None or len(self.paned.panes()) < 2:
            return
        derecha = self._detalle_a_la_derecha()
        total = self.paned.winfo_width() if derecha else self.paned.winfo_height()
        if total < 300:
            if intentos:
                self.after(200, lambda: self._ajustar_division(intentos - 1))
            return
        frac = self.config_data.get("detalle_division") or 0.0
        if not 0.25 <= frac <= 0.85:
            frac = 0.58 if derecha else 0.55
        self.paned.sashpos(0, int(total * frac))

    def _guardar_division(self):
        if self.paned is None or len(self.paned.panes()) < 2:
            return
        total = self.paned.winfo_width() if self._detalle_a_la_derecha() else self.paned.winfo_height()
        if total > 300:
            self.config_data["detalle_division"] = round(self.paned.sashpos(0) / total, 3)

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
            borderwidth=0, relief="flat"
        )
        self._estilo_tabla()
        # Sin marco ni anillo de foco alrededor de la tabla
        style.layout("Renovaciones.Treeview", [("Treeview.treearea", {"sticky": "nswe"})])
        # Columna #0 (barra de estado del desglose): solo la imagen, sin sangría ni triángulo de árbol que la tapen
        style.layout("Renovaciones.Treeview.Item", [("Treeitem.padding", {"sticky": "nswe", "children": [
            ("Treeitem.image", {"side": "left", "sticky": ""}),
            ("Treeitem.focus", {"side": "left", "sticky": "", "children": [("Treeitem.text", {"side": "left", "sticky": ""})]}),
        ]})])
        style.configure("Renovaciones.Treeview", indent=0)
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

    def _estilo_tabla(self):
        """Alto de renglón y letra de las tablas: más grandes con «Letra grande» o en el perfil de usuario."""
        s = self._scaling()
        grande = bool(self.config_data.get("letra_grande")) or self._es_usuario()
        ttk.Style().configure("Renovaciones.Treeview", rowheight=int((40 if grande else 32) * s),
                              font=(FONT, 11 if grande else 10))
        ttk.Style().configure("Renovaciones.Treeview.Heading", font=(FONT_SB, 10 if grande else 9))

    def _barra(self, color: str) -> tk.PhotoImage:
        """Barra de color al inicio de la fila (el estado se lee por color, ícono y texto: no solo por color)."""
        if color not in self._barras:
            alto = int(26 * self._scaling())
            img = tk.PhotoImage(width=6, height=alto)
            img.put(COL["surface"], to=(0, 0, 6, alto))
            img.put(color, to=(0, 2, 4, alto - 2))
            self._barras[color] = img
        return self._barras[color]

    def _make_tree(self, parent, headers: dict, selectmode="extended", barra: bool = False) -> ttk.Treeview:
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)
        tree = ttk.Treeview(
            parent, columns=tuple(headers.keys()), show="tree headings" if barra else "headings",
            style="Renovaciones.Treeview", selectmode=selectmode
        )
        if barra:
            tree.column("#0", width=int(14 * self._scaling()), minwidth=10, stretch=False)
            tree.heading("#0", text="")
        for col, (text, width, anchor) in headers.items():
            self._headers[(str(tree), col)] = text
            tree.heading(col, text=text, anchor=anchor, command=lambda c=col, t=tree: self._on_heading_click(t, c))
            # Los anchos de ttk no siguen el zoom de Windows (la fuente sí): se escalan para que el texto quepa
            tree.column(col, width=int(width * self._scaling()), minwidth=50, anchor=anchor, stretch=False)

        # Colores por estado, suaves (los de alerta ganan a la franja alterna porque cada fila lleva un solo tag)
        tree.tag_configure("normal", background=C("#FFFFFF"), foreground=COL["text"])
        tree.tag_configure("normal_alt", background=C("#FAFBFC"), foreground=COL["text"])
        tree.tag_configure("multiple", background=C("#FFFBEB"), foreground=C("#92400E"))      # Ámbar: varias renovaciones en 1 ticket
        tree.tag_configure("incompleto", background=C("#FEF2F2"), foreground=C("#B91C1C"))    # Rojo: falta cliente o crédito
        tree.tag_configure("duplicado", background=C("#FDF2F8"), foreground=C("#9D174D"))     # Rosa: crédito repetido en otro ticket
        tree.tag_configure("reasignado", background=C("#FFF7ED"), foreground=C("#9A3412"))    # Naranja: reasignado
        tree.tag_configure("rezagada", background=C("#EEF2FF"), foreground=C("#3730A3"))      # Índigo: rezagada del día anterior
        tree.tag_configure("fuera_lote", background=C("#FFFFFF"), foreground=COL["faint"])  # Gris: ya enviada / después del corte
        tree.tag_configure("validado", background=C("#F0FDF4"), foreground=C("#166534"))      # Verde: validada
        tree.tag_configure("cerrado", background=C("#F5F3FF"), foreground=C("#5B21B6"))       # Violeta: ya cerrado en iTop
        tree.tag_configure("descartado", background=COL["subtle"], foreground=COL["faint"])  # Gris: descartada
        tree.tag_configure("no_enviada", background=C("#FEE2E2"), foreground=C("#991B1B"))    # Rojo fuerte: no se envió en su día
        tree.tag_configure("informativa", background=C("#F8FAFC"), foreground=C("#64748B"))   # Gris: ya no se manda (informativa)
        tree.tag_configure("lista", background=C("#FFFFFF"), foreground=COL["text"])        # Lista para enviar (sin validar)
        # Desglose: filas sin pintar; el estado va en la barra de color y en el texto del veredicto
        tree.tag_configure("fila", background=COL["surface"], foreground=COL["text"])
        tree.tag_configure("fila_alt", background=COL["subtle"], foreground=COL["text"])
        tree.tag_configure("apagada", background=COL["surface"], foreground=COL["faint"])

        vsb = ctk.CTkScrollbar(parent, orientation="vertical", command=tree.yview, width=12)
        hsb = ctk.CTkScrollbar(parent, orientation="horizontal", command=tree.xview, height=12)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        return tree

    def _build_tab_desglose(self):
        self.tree_desglose = self._make_tree(self.tab_desglose, {
            "validado": ("Veredicto", 130, "w"),
            **({"equipo": ("Equipo", 150, "w")} if EQUIPO_ACTIVO else {}),
            "ticket": ("Ticket iTop", 105, "center"),
            "parte": ("Parte", 64, "center"),
            "cliente": ("No. cliente", 115, "center"),
            "credito_hijo": ("Crédito a cancelar", 135, "center"),
            "credito_padre": ("Crédito origen", 120, "center"),
            "sql": ("Base de datos", 130, "w"),
            "lote": ("Lote / envío", 270, "w"),
            "cierre": ("Cierre en iTop", 170, "w"),
            "gafete": ("Gafete", 70, "center"),
            "agente": ("Agente responsable", 190, "w"),
            "alerta_reno": ("Alerta de renovación", 290, "w"),
            "alerta_asig": ("Asignación", 240, "w"),
            "motivo": ("Motivo", 230, "w"),
            "fecha": ("Llegada", 135, "center"),
        }, barra=True)
        self.tree_desglose.bind("<<TreeviewSelect>>", self._on_select_desglose_row)
        teclas = [("j", lambda: self._mover_seleccion(1)), ("k", lambda: self._mover_seleccion(-1))]
        if EQUIPO_ACTIVO:
            teclas.append(("d", self._poner_duda))
        for tecla, accion in teclas:
            self.tree_desglose.bind(f"<{tecla}>", lambda e, a=accion: a())
            self.tree_desglose.bind(f"<{tecla.upper()}>", lambda e, a=accion: a())
        self.tree_desglose.bind("<Double-1>", lambda e: self._toggle_selected_validation())
        self.tree_desglose.bind("<space>", lambda e: self._toggle_selected_validation())
        self.tree_desglose.bind("<Return>", lambda e: self._open_selected_in_itop())
        self.tree_desglose.bind("<Delete>", lambda e: self._discard_selected_rows())
        self.tree_desglose.bind("<Control-c>", lambda e: self._copy_selected_rows())
        self.tree_desglose.bind("<n>", lambda e: self._siguiente_pendiente())
        self.tree_desglose.bind("<N>", lambda e: self._siguiente_pendiente())
        self.tree_desglose.bind("<e>", lambda e: self._corregir_dato())
        self.tree_desglose.bind("<E>", lambda e: self._corregir_dato())
        # Estado vacío sobre la tabla
        self.lbl_empty = ctk.CTkLabel(self.tab_desglose, text="", font=F(13), text_color=COL["muted"], fg_color=COL["surface"])

    def _build_tab_tickets(self):
        self.tree_tickets = self._make_tree(self.tab_tickets, {
            "ticket": ("Ticket iTop", 105, "center"),
            "cant_renovaciones": ("Renovaciones", 110, "center"),
            "clientes": ("Clientes", 170, "w"),
            "creditos_hijos": ("Créditos a cancelar", 220, "w"),
            "creditos_padres": ("Créditos origen", 150, "w"),
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
            **({"equipo_marcadas": ("Marcó el analista", 150, "center"),
                "equipo_senal": ("Última señal", 150, "center")} if EQUIPO_ACTIVO else {}),
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
        dias_de_corte = (self.lote_info or {}).get('dias_de_corte') or {}
        for d in dias:
            envios = "  ·  ".join(f"{rec:%H:%M} {rem} ({n} renglones)" for rec, rem, n in d['envios'])
            if not envios:
                envios = ("Sin guardia: lo que llegó va en el ticket único del lunes" if d['dia'].weekday() == 5
                          and not d['habil'] else "Día inhábil" if not d['habil'] else
                          "No se encontró ticket único" if d['debian'] else "Sin ticket único")
            if d['dia'] in dias_de_corte:
                # Día de emisión o de cierre: puede haber más de un ticket único (cuenta el último)
                envios = f"📅 {dias_de_corte[d['dia']][0].upper()}{dias_de_corte[d['dia']][1:]}  ·  {envios}"
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

    def _limpiar_detalle(self):
        """Panel de detalle sin renovación (al vaciar la tabla no debe quedar la de antes)."""
        self.lbl_detail_header.configure(text="Selecciona una renovación para ver el desglose del ticket.")
        self.banner.configure(fg_color=COL["subtle"], border_color=COL["border"])
        self.lbl_ver_icono.configure(text="", fg_color=COL["track"], text_color=COL["muted"])
        self.lbl_ver_titulo.configure(text="Sin renovación seleccionada", text_color=COL["text_2"])
        self.lbl_ver_accion.configure(text="Elige una renovación de la tabla.", text_color=COL["text_2"])
        for box in (self.txt_detail, self.txt_body):
            box.configure(state="normal")
            box.delete("1.0", "end")
            box.configure(state="disabled")
        self._fila_detalle = None
        self.lbl_status.configure(text="Listo. Presiona «Extraer y validar» (F5).")

    def _build_detail_panel(self, parent):
        """Desglose del ticket (lo principal) y, debajo o a un lado, el correo original. Se acomoda en _colocar_detalle."""
        hdr = self.detail_hdr = ctk.CTkFrame(parent, fg_color="transparent")
        hdr.grid_columnconfigure(0, weight=1)
        self.lbl_detail_header = ctk.CTkLabel(
            hdr, text="Selecciona una renovación para ver el desglose del ticket.",
            font=F(14, True), text_color=COL["text"], anchor="w", justify="left", wraplength=500
        )
        self.lbl_detail_header.grid(row=0, column=0, sticky="ew")
        botones = self._detail_botones = ctk.CTkFrame(hdr, fg_color="transparent")
        botones.grid(row=0, column=1, sticky="ne", padx=(8, 0))
        self._btn(botones, "Abrir en iTop", self._open_selected_in_itop, height=30).pack(side="left", ipadx=6)
        self._btn(botones, "Corregir dato", self._corregir_dato, height=30).pack(side="left", padx=(8, 0), ipadx=6)
        self._btn(
            botones, "Copiar cliente y crédito", lambda: self._copy_field(('Numero_Cliente', 'Credito_Renovado_Hijo')),
            height=30
        ).pack(side="left", padx=(8, 0), ipadx=6)
        hdr.bind("<Configure>", self._ajustar_encabezado_detalle)

        # Veredicto de la renovación: qué es y qué hacer, en grande y con color (lo primero que se lee)
        self.banner = ctk.CTkFrame(parent, corner_radius=10, border_width=1, fg_color=COL["subtle"], border_color=COL["border"])
        self.banner.grid_columnconfigure(1, weight=1)
        self.lbl_ver_icono = ctk.CTkLabel(self.banner, text="", width=40, height=40, corner_radius=20, font=F(18, True),
                                          fg_color=COL["track"], text_color=COL["muted"])
        self.lbl_ver_icono.grid(row=0, column=0, rowspan=2, padx=(14, 12), pady=12, sticky="n")
        self.lbl_ver_titulo = ctk.CTkLabel(self.banner, text="Sin renovación seleccionada", font=F(16, True),
                                           text_color=COL["text_2"], anchor="w", justify="left")
        self.lbl_ver_titulo.grid(row=0, column=1, sticky="ew", padx=(0, 14), pady=(12, 0))
        self.lbl_ver_accion = ctk.CTkLabel(self.banner, text="Elige una renovación de la tabla.", font=F(13),
                                           text_color=COL["text_2"], anchor="w", justify="left", wraplength=420)
        self.lbl_ver_accion.grid(row=1, column=1, sticky="ew", padx=(0, 14), pady=(2, 12))
        self.banner.bind("<Configure>", lambda e: self.lbl_ver_accion.configure(
            wraplength=max(200, int(e.width / self._scaling()) - 90)))

        self.lbl_det_titulo = ctk.CTkLabel(parent, text="DESGLOSE DEL TICKET  ·  comprobaciones", font=F(10, True),
                                           text_color=COL["primary"], anchor="w", height=14)
        # Título del correo con «Abrir en ventana»: en pantallas chicas el correo se lee completo aparte
        self.lbl_body_titulo = ctk.CTkFrame(parent, fg_color="transparent")
        ctk.CTkLabel(self.lbl_body_titulo, text="CORREO ORIGINAL  ·  números detectados resaltados", font=F(10, True),
                     text_color=COL["muted"], anchor="w", height=14).pack(side="left")
        self._btn(self.lbl_body_titulo, "Abrir en ventana ↗", self._ventana_correo, "link", height=18,
                  font=F(10)).pack(side="left", padx=(8, 0))

        self.txt_detail = ctk.CTkTextbox(parent, font=F(13), fg_color=COL["subtle"], wrap="word")
        self.txt_detail.tag_config("titulo", foreground=COL["text"])
        self.txt_detail.tag_config("ok", foreground=COL["success"])
        self.txt_detail.tag_config("warn", foreground=COL["warning"])
        self.txt_detail.tag_config("muted", foreground=COL["muted"])
        self.txt_detail.tag_config("cierre", foreground=C("#5B21B6"))
        self.txt_detail.tag_config("error", foreground=C("#B91C1C"))
        # Encabezados de sección y renglones con sangría (en el Text de tk: CTkTextbox no deja poner fuente a un tag)
        interno = self.txt_detail._textbox
        interno.tag_configure("seccion", font=(FONT_SB, 9), foreground=COL["muted"], spacing1=10, spacing3=4)
        interno.tag_configure("check", lmargin1=4, lmargin2=26, spacing1=3, spacing3=3)
        interno.tag_configure("dato", font=(FONT_SB, 11), foreground=COL["text"])
        s = self._scaling()
        interno.tag_configure("tabla", tabs=(int(95 * s), int(235 * s), int(470 * s)), lmargin1=4, lmargin2=4,
                              spacing1=2, spacing3=2)
        interno.bind("<Configure>", lambda e: interno.tag_configure(
            "tabla", tabs=tuple(max(int(m * s), int(e.width * f)) for m, f in ((70, 0.20), (150, 0.46), (250, 0.73)))),
            add="+")
        interno.tag_configure("tabla_enc", font=(FONT_SB, 9), foreground=COL["muted"])
        interno.tag_configure("mas", font=(FONT_SB, 10), foreground=COL["primary"], lmargin1=4, spacing1=3, spacing3=3)
        interno.tag_bind("mas", "<Button-1>", lambda e: self._toggle_checks())
        interno.tag_bind("mas", "<Enter>", lambda e: interno.configure(cursor="hand2"))
        interno.tag_bind("mas", "<Leave>", lambda e: interno.configure(cursor=""))
        self._checks_abiertos = False
        self._fila_detalle = None
        self._detalle_derecha = True

        self.txt_detail.insert("1.0", "Aquí sale el desglose del ticket: el veredicto de cada renovación, qué hacer y "
                                      "las comprobaciones (correo, base de datos, calendario, ticket único e iTop).")
        self.txt_detail.configure(state="disabled")

        self.txt_body = ctk.CTkTextbox(parent, font=F(12), wrap="word")
        # Resaltado: el amarillo invertido quedaba verde olivo con texto naranja (no se leía); en oscuro, su propio par
        fondo, texto = ("#854D0E", "#FEF3C7") if _TEMA["oscuro"] else ("#FEF08A", "#713F12")
        self.txt_body.tag_config("hl", background=fondo, foreground=texto)
        self.txt_body._textbox.bind("<Double-1>", lambda e: self._ventana_correo(), add="+")
        self._win_correo = None
        # Al cambiar el alto del panel, el correo se vuelve a repartir con las comprobaciones
        parent.bind("<Configure>", lambda e: self._ajustar_alto_correo(), add="+")

    def _colocar_detalle(self, derecha: bool):
        """A la derecha: desglose arriba (más alto) y correo abajo. Abajo: desglose a la izquierda (más ancho) y correo."""
        p = self.detail_frame
        for w in (self.detail_hdr, self.banner, self.lbl_det_titulo, self.lbl_body_titulo, self.txt_detail, self.txt_body):
            w.grid_forget()
        for i in range(7):
            p.grid_rowconfigure(i, weight=0)
            p.grid_columnconfigure(i, weight=0)
        if derecha:
            # El desglose se queda con el alto que sobra; el correo, solo el que necesita (ver _ajustar_alto_correo)
            p.grid_columnconfigure(0, weight=1)
            p.grid_rowconfigure(3, weight=1, minsize=int(80 * self._scaling()))
            p.grid_rowconfigure(5, weight=0)
            self.detail_hdr.grid(row=0, column=0, padx=16, pady=(12, 8), sticky="ew")
            self.banner.grid(row=1, column=0, padx=16, pady=(0, 10), sticky="ew")
            self.lbl_det_titulo.grid(row=2, column=0, padx=16, pady=(0, 6), sticky="w")
            self.txt_detail.grid(row=3, column=0, padx=16, pady=(0, 10), sticky="nsew")
            self.lbl_body_titulo.grid(row=4, column=0, padx=16, pady=(0, 6), sticky="w")
            self.txt_body.grid(row=5, column=0, padx=16, pady=(0, 16), sticky="nsew")
        else:
            p.grid_columnconfigure(0, weight=3)
            p.grid_columnconfigure(1, weight=2)
            p.grid_rowconfigure(3, weight=1, minsize=int(160 * self._scaling()))
            self.detail_hdr.grid(row=0, column=0, columnspan=2, padx=16, pady=(12, 6), sticky="ew")
            self.banner.grid(row=1, column=0, columnspan=2, padx=16, pady=(0, 8), sticky="ew")
            self.lbl_det_titulo.grid(row=2, column=0, padx=(16, 6), pady=(0, 6), sticky="w")
            self.lbl_body_titulo.grid(row=2, column=1, padx=(6, 16), pady=(0, 6), sticky="w")
            self.txt_detail.grid(row=3, column=0, padx=(16, 6), pady=(0, 16), sticky="nsew")
            self.txt_body.grid(row=3, column=1, padx=(6, 16), pady=(0, 16), sticky="nsew")
        self._detalle_derecha = derecha
        self._ajustar_alto_correo()

    def _ajustar_encabezado_detalle(self, e):
        """El folio, cliente y crédito usan todo el ancho que dejan los botones (sin cortarse)."""
        botones, s = self._detail_botones, self._scaling()
        ancho_botones = botones.winfo_reqwidth()
        if ancho_botones < 20:  # todavía no se dibujan
            ancho_botones = int(330 * s)
        # Panel angosto (laptop): los botones bajan a otro renglón y el folio usa todo el ancho
        abajo = e.width - ancho_botones < int(280 * s)
        if abajo != (int(botones.grid_info().get("row", 0)) == 1):
            if abajo:
                botones.grid(row=1, column=0, columnspan=2, sticky="w", padx=0, pady=(6, 0))
            else:
                botones.grid(row=0, column=1, columnspan=1, sticky="ne", padx=(8, 0), pady=0)
        libre = e.width if abajo else e.width - ancho_botones
        self.lbl_detail_header.configure(wraplength=max(220, int(libre / s) - 16))

    def _ajustar_alto_correo(self):
        """
        A la derecha, el correo original toma el alto de su texto (entre 4 y 14 renglones) y no le quita espacio a las
        comprobaciones; abajo, comparten el alto. Si el panel es bajo (laptop), las comprobaciones ceden su espacio
        hasta que el correo tenga al menos 5 renglones, para que nunca desaparezca.
        """
        if not getattr(self, "_detalle_derecha", True):
            alto = 200
        else:
            texto = self.txt_body.get("1.0", "end").rstrip("\n")
            renglones = sum(1 + len(linea) // 70 for linea in texto.split("\n")) if texto else 3
            alto = max(4, min(14, renglones)) * 18 + 16
            s = self._scaling()
            panel = self.detail_frame.winfo_height() / s
            if panel > 100:
                fijo = sum(w.winfo_reqheight() for w in (self.detail_hdr, self.banner, self.lbl_det_titulo,
                                                         self.lbl_body_titulo)) / s + 72
                alto = min(alto, max(5 * 18 + 16, int(panel - fijo - 220)))
        if int(self.txt_body.cget("height")) != alto:
            self.txt_body.configure(height=alto)

    def _ventana_correo(self, enfocar: bool = True):
        """El correo original completo en su propia ventana, con los números resaltados (para pantallas chicas)."""
        texto = self.txt_body.get("1.0", "end").rstrip("\n")
        if self._fila_detalle is None or not texto.strip():
            self.lbl_status.configure(text="Selecciona una renovación para ver su correo original.")
            return
        win = self._win_correo
        if win is None or not win.winfo_exists():
            win = self._win_correo = ctk.CTkToplevel(self)
            win.configure(fg_color=COL["surface"])
            win.transient(self)
            caja = win.caja = ctk.CTkTextbox(win, font=F(13), wrap="word")
            caja.tag_config("hl", background=self.txt_body._textbox.tag_cget("hl", "background"),
                            foreground=self.txt_body._textbox.tag_cget("hl", "foreground"))
            botones = ctk.CTkFrame(win, fg_color="transparent")
            botones.pack(side="bottom", fill="x", padx=16, pady=(0, 14))
            caja.pack(fill="both", expand=True, padx=16, pady=(16, 10))
            self._btn(botones, "Cerrar", win.destroy, "primary").pack(side="right", ipadx=12)
            self._btn(botones, "Copiar texto", lambda: self._to_clipboard(
                caja.get("1.0", "end").rstrip("\n"), "Correo original copiado al portapapeles.")).pack(side="right",
                                                                                                    padx=(0, 8))
            win.bind("<Escape>", lambda e: win.destroy())
            # Hasta 760 x 640, sin pasarse del 70 % de la pantalla (geometry de CTkToplevel va en unidades lógicas)
            s = self._scaling()
            ancho = int(min(760, self.winfo_screenwidth() / s * 0.7))
            alto = int(min(640, self.winfo_screenheight() / s * 0.7))
            x = int(self.winfo_rootx() / s + max(0, (self.winfo_width() / s - ancho) / 2))
            y = int(self.winfo_rooty() / s + max(0, (self.winfo_height() / s - alto) / 3))
            win.geometry(f"{ancho}x{alto}+{x}+{y}")
        ticket = self._fila_detalle.get('Numero_Ticket', '')
        win.title(f"Correo original · {ticket}" if ticket else "Correo original")
        caja = win.caja
        caja.configure(state="normal")
        caja.delete("1.0", "end")
        caja.insert("1.0", texto)
        rangos = self.txt_body._textbox.tag_ranges("hl")
        for inicio, fin in zip(rangos[::2], rangos[1::2]):
            caja.tag_add("hl", inicio, fin)
        caja.configure(state="disabled")
        if enfocar:
            win.deiconify()
            win.lift()
            win.after(100, lambda: win.winfo_exists() and win.focus_force())

    def _toggle_checks(self):
        self._checks_abiertos = not self._checks_abiertos
        if self._fila_detalle is not None:
            self._show_row_detail(self._fila_detalle)

    def _build_context_menus(self):
        self._menu_opts = opts = dict(
            tearoff=0, font=(FONT, 10), bg=COL["surface"], fg=COL["text"], activebackground=COL["primary_soft"],
            activeforeground=COL["text"], relief="flat", borderwidth=1
        )
        m = tk.Menu(self, **opts)
        m.add_command(label="Marcar / desmarcar validada", accelerator="Espacio", command=self._toggle_selected_validation)
        m.add_command(label="Abrir en iTop", accelerator="Enter", command=self._open_selected_in_itop)
        m.add_command(label="Buscar el folio en iTop", command=lambda: self._open_selected_in_itop(buscar=True))
        m.add_command(label="Corregir cliente o crédito…", accelerator="E", command=self._corregir_dato)
        if EQUIPO_ACTIVO:
            m.add_command(label="Pasar duda al validador…", accelerator="D", command=self._poner_duda)
        m.add_separator()
        m.add_command(label="Copiar No. cliente", command=lambda: self._copy_field(('Numero_Cliente',)))
        m.add_command(label="Copiar crédito a cancelar", command=lambda: self._copy_field(('Credito_Renovado_Hijo',)))
        m.add_command(label="Copiar ticket", command=lambda: self._copy_field(('Numero_Ticket',)))
        m.add_command(label="Copiar filas", accelerator="Ctrl+C", command=self._copy_selected_rows)
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
        self.bind("<Control-b>", lambda e: self._toggle_sidebar())
        self.bind("<Control-k>", lambda e: self._paleta_comandos())
        self.bind("<Control-z>", lambda e: None if self._escribiendo(e) else self._deshacer_ultimo())
        self.bind("<question>", lambda e: None if self._escribiendo(e) else self._show_shortcuts())

    @staticmethod
    def _escribiendo(e) -> bool:
        """El foco está en un campo de texto (ahí las teclas son para escribir)."""
        return isinstance(e.widget, (tk.Entry, tk.Text))

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
        self.lbl_sql_estado.configure(text=texto, text_color=COL["success"] if ok else C("#B91C1C"))
        if ok:
            self._save_config()
        else:
            registro.log().warning("Probar conexión a la base: %s", texto)

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
            registro.log().warning("Validar en la base: %s", e)
            return {'ok': False, 'error': str(e)}
        except Exception as e:
            registro.error("Validar en la base")
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
        """
        Pone de lunes a sábado de la semana (la actual llega hasta hoy): el sábado también llegan incidentes y, si hay
        guardia, sale ticket único. Se busca con «Extraer y validar».
        """
        hoy = datetime.now().date()
        lunes = hoy - timedelta(days=hoy.weekday() + 7 * semanas_atras)
        sabado = lunes + timedelta(days=5)
        self.cal_start.set_date(lunes)
        self.cal_end.set_date(min(sabado, hoy) if semanas_atras == 0 else sabado)
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
        ctk.CTkLabel(win, text="Atajos de teclado", font=F(16, True), text_color=COL["text"]).pack(
            anchor="w", padx=24, pady=(20, 12))
        caja, lista, ajustar = self._cuerpo_con_barra(win)
        for i, (tecla, accion) in enumerate(ATAJOS):
            ctk.CTkLabel(
                lista, text=tecla, font=F(11, True), fg_color=COL["track"], corner_radius=4,
                text_color=COL["text_2"], height=24
            ).grid(row=i, column=0, sticky="w", padx=(24, 16), pady=3, ipadx=8)
            ctk.CTkLabel(lista, text=accion, font=F(12), text_color=COL["text"]).grid(row=i, column=1, sticky="w",
                                                                                   padx=(0, 24))
        botones = ctk.CTkFrame(win, fg_color="transparent")
        botones.pack(side="bottom", fill="x", padx=24, pady=(16, 20))
        caja.pack(fill="both", expand=True)
        self._btn(botones, "Cerrar", win.destroy, "primary").pack(side="right", ipadx=12)
        ajustar()
        win.after(150, win.focus_force)

    def _cuerpo_con_barra(self, win):
        """
        Cuerpo de una ventana de tamaño fijo que en una laptop puede no caber en la pantalla: si su contenido es más
        alto que lo que hay (menos el título y los botones), se desplaza con una barra y la rueda; si cabe, se ve igual.
        Devuelve (caja para acomodar en `win`, interior donde van los widgets, ajustar() para llamar al terminar).
        """
        caja = ctk.CTkFrame(win, fg_color="transparent", corner_radius=0)
        lienzo = tk.Canvas(caja, highlightthickness=0, bd=0, bg=win.cget("bg"), yscrollincrement=20)
        barra = ctk.CTkScrollbar(caja, orientation="vertical", command=lienzo.yview)
        lienzo.configure(yscrollcommand=barra.set)
        lienzo.pack(side="left", fill="both", expand=True)
        interior = ctk.CTkFrame(lienzo, fg_color="transparent", corner_radius=0)
        lienzo.create_window(0, 0, window=interior, anchor="nw")

        def ajustar():
            win.update_idletasks()
            ancho, alto = interior.winfo_reqwidth(), interior.winfo_reqheight()
            # Lo que no es el cuerpo (título, botones, barra de la ventana) y la barra de tareas de Windows
            resto = win.winfo_reqheight() - caja.winfo_reqheight() + int(110 * self._scaling())
            visible = max(int(150 * self._scaling()), min(alto, self.winfo_screenheight() - resto))
            lienzo.configure(width=ancho, height=visible, scrollregion=(0, 0, ancho, alto))
            if alto > visible:
                barra.pack(side="right", fill="y")
                win.bind("<MouseWheel>", lambda e: lienzo.yview_scroll(-3 if e.delta > 0 else 3, "units"), add="+")
        return caja, interior, ajustar

    def _toast(self, msg: str, color: str = None, ms: int = 7000, accion: tuple = None):
        """
        Aviso flotante en la esquina inferior derecha que se cierra solo (o con clic). `accion`: (texto, función) para
        un botón, p. ej. («Deshacer», …). Solo hay un aviso a la vez: el nuevo reemplaza al anterior.
        """
        try:
            previo = getattr(self, "_toast_win", None)
            if previo is not None and previo.winfo_exists():
                previo.destroy()
            # Siempre oscuro (en el tema oscuro, un poco más claro que el fondo)
            fondo, texto = ("#1F2937", "#F9FAFB") if not _TEMA["oscuro"] else ("#334155", "#F9FAFB")
            win = self._toast_win = ctk.CTkToplevel(self)
            win.overrideredirect(True)
            win.attributes("-topmost", True)
            win.configure(fg_color=fondo)
            frame = ctk.CTkFrame(win, fg_color=fondo, corner_radius=0)
            frame.pack(fill="both", expand=True)
            strip = ctk.CTkFrame(frame, width=4, height=4, corner_radius=2, fg_color=color or COL["primary"])
            strip.pack(side="left", fill="y", padx=(12, 0), pady=12)
            lbl = ctk.CTkLabel(
                frame, text=msg, text_color=texto, font=F(12), wraplength=360, justify="left", anchor="w"
            )
            lbl.pack(side="left", padx=(12, 12 if accion else 18), pady=12)
            if accion:
                def hacer():
                    win.destroy()
                    accion[1]()
                ctk.CTkButton(frame, text=accion[0], command=hacer, height=28, width=10, corner_radius=6, font=F(12, True),
                              fg_color="transparent", hover_color="#475569", text_color="#93C5FD").pack(
                    side="left", padx=(0, 12), ipadx=6)
            win.update_idletasks()
            # Del tamaño que pide su contenido (CTkToplevel nace de 200x200 y escala lo que se le pase a geometry)
            ancho, alto = frame.winfo_reqwidth(), frame.winfo_reqheight()
            x = self.winfo_rootx() + self.winfo_width() - ancho - 24
            y = self.winfo_rooty() + self.winfo_height() - alto - 48
            tk.Toplevel.wm_geometry(win, f"{ancho}x{alto}+{x}+{y}")
            for w in (frame, lbl, strip):
                w.bind("<Button-1>", lambda e: win.destroy())
            win.after(ms, lambda: win.winfo_exists() and win.destroy())
        except Exception:
            self.lbl_status.configure(text=msg)

    # ================= DESHACER =================

    def _con_deshacer(self, descripcion: str, revertir):
        """Guarda cómo revertir lo que se acaba de hacer y lo avisa con el botón «Deshacer» (Ctrl+Z)."""
        self._deshacer.append((descripcion, revertir))
        del self._deshacer[:-20]
        self._toast(f"{descripcion}.", COL["success"], ms=6000, accion=("Deshacer  Ctrl+Z", self._deshacer_ultimo))

    def _deshacer_ultimo(self):
        if not self._deshacer:
            self._toast("No hay nada que deshacer.", COL["muted"], ms=2500)
            return
        descripcion, revertir = self._deshacer.pop()
        revertir()
        self._refrescar_tras_accion()
        self._mostrar_seleccion()
        self._publicar_equipo()
        self._toast(f"Se deshizo: {descripcion[0].lower()}{descripcion[1:]}.", COL["muted"], ms=3500)

    def _mostrar_seleccion(self):
        sel = self.tree_desglose.selection()
        r = next((x for x in self.all_rows if sel and x['Row_ID'] == sel[0]), None)
        if r:
            self._show_row_detail(r)

    def _estado_de(self, rids) -> tuple:
        """Copia de las marcas de esas filas (para poder deshacer)."""
        rids = set(rids)
        return (rids & self.validated_row_ids, rids & self.discarded_row_ids, rids & self.forced_included_row_ids)

    def _restaurar_estado(self, rids, copia):
        """Regresa las marcas de esas filas a como estaban (en pantalla y en disco)."""
        validadas, descartadas, incluidas = copia
        rids = set(rids)
        for conjunto, categoria, antes in ((self.validated_row_ids, 'validados', validadas),
                                           (self.discarded_row_ids, 'descartados', descartadas),
                                           (self.forced_included_row_ids, 'incluidos', incluidas)):
            quitar, poner = (conjunto & rids) - antes, antes - conjunto
            conjunto.difference_update(quitar)
            conjunto.update(poner)
            self.estado.set(categoria, self._keys_for(quitar), False)
            self.estado.set(categoria, self._keys_for(poner), True)
        self.estado.registrar(self._keys_for(rids), "Deshecho")

    # ================= PALETA DE COMANDOS (Ctrl+K) =================

    def _comandos(self) -> list:
        """[(texto, función)] de todo lo que se puede hacer escribiendo en Ctrl+K."""
        usuario = self._es_usuario()
        cmds = [
            ("Extraer y validar (F5)", self._start_extraction),
            ("Copiar mis renovaciones" if usuario else "Copiar lote", self._copy_batch_to_clipboard),
            ("Siguiente pendiente (N)", self._siguiente_pendiente),
            ("Validar la seleccionada (Espacio)", self._toggle_selected_validation),
            ("Deshacer lo último (Ctrl+Z)", self._deshacer_ultimo),
            ("Abrir en iTop (Enter)", self._open_selected_in_itop),
            ("Corregir cliente o crédito a cancelar (E)", self._corregir_dato),
            ("Ver el calendario de emisión", self._recordatorio_emision),
            ("Ver atajos de teclado (?)", self._show_shortcuts),
            ("Cambiar a modo oscuro" if not _TEMA["oscuro"] else "Cambiar a modo claro",
             lambda: (self.var_oscuro.set(not _TEMA["oscuro"]), self._cambiar_tema())),
            ("Letra grande en la tabla: " + ("quitar" if self.config_data.get("letra_grande") else "poner"),
             lambda: (self.var_letra.set(not self.config_data.get("letra_grande")), self._cambiar_letra())),
            ("Ocultar o mostrar el panel izquierdo (Ctrl+B)", self._toggle_sidebar),
        ]
        if usuario and EQUIPO_ACTIVO:
            cmds.insert(4, ("Pasar una duda al validador (D)", self._poner_duda))
        elif not usuario:
            cmds += [("Crear correo Ticket Único (Ctrl+M)", self._crear_correo_lote),
                     ("Exportar a Excel (Ctrl+E)", self._export_excel),
                     ("Validar visibles", self._validate_all_visible),
                     ("Revisión semanal: esta semana", lambda: self._revisar_semana(0))]
        filtros = self.cmb_quick_filter.principales + ([FILTRO_INFORMATIVAS_TODAS] if usuario else FILTROS_SECUNDARIOS)
        cmds += [(f"Ver: {f}", lambda f=f: self._apply_filter(f)) for f in dict.fromkeys(filtros)]
        return cmds

    def _paleta_comandos(self):
        win = ctk.CTkToplevel(self)
        win.overrideredirect(True)
        win.attributes("-topmost", True)
        win.configure(fg_color=COL["surface"])
        marco = ctk.CTkFrame(win, fg_color=COL["surface"], corner_radius=10, border_width=1, border_color=COL["border_strong"])
        marco.pack(fill="both", expand=True)
        ent = ctk.CTkEntry(marco, width=520, height=38, font=F(14), placeholder_text="¿Qué quieres hacer? (Esc cierra)")
        ent.pack(fill="x", padx=12, pady=(12, 6))
        lista = tk.Listbox(marco, height=10, font=(FONT, 11), activestyle="none", borderwidth=0, highlightthickness=0,
                           bg=COL["surface"], fg=COL["text"], selectbackground=COL["primary_soft"],
                           selectforeground=COL["text"])
        lista.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        comandos = self._comandos()
        visibles = []

        def filtrar(_e=None):
            palabras = ent.get().lower().split()
            visibles[:] = [c for c in comandos if all(p in c[0].lower() for p in palabras)]
            lista.delete(0, "end")
            for texto, _ in visibles:
                lista.insert("end", f"  {texto}")
            if visibles:
                lista.selection_set(0)

        def mover(paso):
            if not visibles:
                return "break"
            i = (lista.curselection() or (0,))[0]
            i = max(0, min(len(visibles) - 1, i + paso))
            lista.selection_clear(0, "end")
            lista.selection_set(i)
            lista.see(i)
            return "break"

        def ejecutar(_e=None):
            sel = lista.curselection()
            if visibles and sel:
                accion = visibles[sel[0]][1]
                win.destroy()
                self.after(50, accion)

        ent.bind("<KeyRelease>", lambda e: e.keysym in ("Up", "Down", "Return") or filtrar())
        ent.bind("<Down>", lambda e: mover(1))
        ent.bind("<Up>", lambda e: mover(-1))
        ent.bind("<Return>", ejecutar)
        lista.bind("<Double-1>", ejecutar)
        for w in (win, ent, lista):
            w.bind("<Escape>", lambda e: win.destroy())
        win.bind("<FocusOut>", lambda e: win.after(150, lambda: win.winfo_exists() and win.focus_get() is None
                                                   and win.destroy()))
        filtrar()
        win.update_idletasks()
        x = self.winfo_rootx() + (self.winfo_width() - win.winfo_width()) // 2
        y = self.winfo_rooty() + int(self.winfo_height() * 0.18)
        win.geometry(f"+{x}+{y}")
        win.after(60, lambda: (win.focus_force(), ent.focus_set()))

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
            corto = f"{now:%H:%M}  ·  Fin de semana"
        elif enviado:
            text, (bg, fg) = f"{now:%H:%M}  ·  Ticket único de hoy enviado a las {enviado:%H:%M}: lo nuevo va al siguiente", neutro
            corto = f"{now:%H:%M}  ·  Ticket único enviado {enviado:%H:%M}"
        elif now < habitual_dt:
            mins = int((habitual_dt - now).total_seconds() // 60)
            h, m = divmod(mins, 60)
            restante = f"{h} h {m:02d} min" if h else f"{m} min"
            text = f"{now:%H:%M}  ·  Ticket único de hoy pendiente  ·  habitual {habitual:%H:%M} (en {restante})"
            corto = f"{now:%H:%M}  ·  Ticket único en {restante}"
            if mins > 60:
                bg, fg = COL["success_soft"], C("#047857")
            elif mins > 15:
                bg, fg = COL["warning_soft"], COL["warning"]
            else:
                bg, fg = COL["danger_soft"], COL["danger"]
            if mins <= 15 and self._aviso_corte_dia != now.date():
                self._aviso_corte_dia = now.date()
                self._aviso_corte(mins)
        else:
            text = f"{now:%H:%M}  ·  Ticket único de hoy pendiente: lo que llegue aún entra hoy"
            corto = f"{now:%H:%M}  ·  Ticket único pendiente"
            bg, fg = COL["warning_soft"], COL["warning"]
        self._reloj = (text, corto)
        self.lbl_clock.configure(text=corto if self._top_nivel >= 2 else text, fg_color=bg, text_color=fg)
        self.after_idle(self._ajustar_topbar)
        self._actualizar_chip_corte()
        if reschedule:
            self.after(20000, self._tick_clock)

    # ================= RECORDATORIO DEL CALENDARIO DE EMISIÓN =================

    @staticmethod
    def _cuando(fecha, hoy) -> str:
        dias = (fecha - hoy).days
        return "HOY" if dias == 0 else "mañana" if dias == 1 else f"en {dias} días"

    def _actualizar_chip_corte(self):
        hoy = datetime.now().date()
        _, proximos = cortes_alrededor(hoy, 1)
        if not proximos:
            self.lbl_corte.pack_forget()
            return
        fecha, texto = proximos[0]
        dias = (fecha - hoy).days
        corto = "Emisión" if "emisión" in texto else "Cierre de mes"
        bg, fg = ((COL["danger_soft"], COL["danger"]) if dias <= 1 else
                  (COL["warning_soft"], COL["warning"]) if dias <= int(self.config_data.get("aviso_emision_dias") or 5)
                  else (COL["track"], COL["text_2"]))
        self.lbl_corte.configure(text=f"📅 {corto} {fecha:%d/%m} ({self._cuando(fecha, hoy)})", fg_color=bg, text_color=fg)

    def _recordatorio_emision(self, auto: bool = False):
        """
        Ventana con el próximo corte del calendario de emisión (generación de la emisión o cierre de mes) y lo que
        implica. Sola: una vez al día cuando faltan pocos días para el corte o acaba de pasar uno, y cuando el lote
        trae renovaciones que ya no se pueden cancelar. Siempre: con clic en el aviso de la barra de arriba.
        """
        hoy = datetime.now().date()
        ultimo, proximos = cortes_alrededor(hoy, 2)
        if not proximos:
            return
        # Las que iban al lote y ya se emitieron (salieron del lote: solo informativas)
        bloqueadas = [r for r in self.all_rows if r.get('Informativa_Tipo') == reglas.INFO_EMISION
                      and r.get('Row_ID') not in self.discarded_row_ids]
        if auto:
            dias_aviso = int(self.config_data.get("aviso_emision_dias") or 5)
            toca = ((proximos[0][0] - hoy).days <= dias_aviso or (ultimo and (hoy - ultimo[0]).days <= 2))
            visto = self.config_data.get("recordatorio_emision_visto") == hoy.isoformat()
            nuevas = len(bloqueadas) > self._bloqueadas_avisadas
            if not ((toca and not visto) or nuevas):
                return
        self._bloqueadas_avisadas = len(bloqueadas)
        self.config_data["recordatorio_emision_visto"] = hoy.isoformat()
        self._save_config()

        fecha, texto = proximos[0]
        if auto:
            # Sin ventana que tape el trabajo: un aviso con el botón para ver el calendario
            msg = f"📅 Próximo corte: {texto} el {fecha:%d/%m} ({self._cuando(fecha, hoy)})."
            if bloqueadas:
                msg += f"\n{pl(len(bloqueadas), '# [renovación ya pasó|renovaciones ya pasaron] la emisión: solo [informativa|informativas], no se [manda|mandan]')}."
            self._toast(msg, COL["warning"], ms=15000, accion=("Ver calendario", self._recordatorio_emision))
            return
        win = ctk.CTkToplevel(self)
        win.title("Recordatorio: calendario de emisión")
        win.resizable(False, False)
        win.transient(self)
        win.configure(fg_color=COL["surface"])
        urgente = (fecha - hoy).days <= 1 or bloqueadas
        ctk.CTkFrame(win, height=6, corner_radius=0, fg_color=C("#DC2626") if urgente else C("#F59E0B")).pack(fill="x")
        caja, cuerpo, ajustar = self._cuerpo_con_barra(win)

        def renglon(t, fuente=F(12), color=COL["text"], arriba=6):
            ctk.CTkLabel(cuerpo, text=t, font=fuente, text_color=color, justify="left", anchor="w",
                         wraplength=520).pack(fill="x", pady=(arriba, 0))

        renglon("📅  Calendario de emisión", F(17, True), arriba=0)
        renglon(f"Próximo corte: {texto}, {DIAS_SEMANA[fecha.weekday()]} {fecha:%d/%m/%Y} "
                f"({self._cuando(fecha, hoy)}).", F(14, True), C("#B91C1C") if urgente else COL["warning"], 12)
        renglon(f"Lo dado de alta hasta el {fecha:%d/%m} se cancela solo si sale en un ticket único de ese día o antes. "
                f"Ese día se puede mandar más de un ticket único: lo que llegue después del primero se manda en otro "
                f"el mismo día. Después de la emisión ya no se puede cancelar (sale del lote y queda como informativa); "
                f"después del cierre de mes solo con autorización.")
        if len(proximos) > 1:
            f2, t2 = proximos[1]
            renglon(f"Después: {t2}, {DIAS_SEMANA[f2.weekday()]} {f2:%d/%m/%Y} ({self._cuando(f2, hoy)}).",
                    color=COL["text_2"], arriba=10)
        if ultimo and (hoy - ultimo[0]).days <= 7:
            renglon(f"Corte reciente: {ultimo[1]}, el {ultimo[0]:%d/%m}. Lo dado de alta hasta ese día que no salió en "
                    f"un ticket único de ese día ya no se puede cancelar.", color=COL["text_2"], arriba=10)
        if bloqueadas:
            tickets = sorted({r['Numero_Ticket'] for r in bloqueadas})
            renglon(f"⚠️  {pl(len(bloqueadas), '# [renovación ya se emitió y salió|renovaciones ya se emitieron y salieron] del lote (solo [informativa|informativas])')}: "
                    f"{', '.join(tickets[:8])}{'…' if len(tickets) > 8 else ''}.", F(13, True), C("#B91C1C"), 14)
        elif not self.var_sql.get():
            renglon("Sin la base de datos se usa el día en que llegó el incidente; con «Validar en la base de datos» se "
                    "usa la fecha de alta de cada renovación (más exacto).", color=COL["muted"], arriba=14)

        # El calendario: cortes del último mes y de los próximos meses
        ctk.CTkLabel(cuerpo, text="CALENDARIO DE CORTES", font=F(10, True), text_color=COL["muted"], anchor="w",
                     height=14).pack(fill="x", pady=(16, 6))
        tabla = ctk.CTkFrame(cuerpo, fg_color=COL["subtle"], corner_radius=8, border_width=1, border_color=COL["border"])
        tabla.pack(fill="x")
        tabla.grid_columnconfigure(1, weight=1)
        for col, enc in enumerate(("Fecha", "Corte", "Cuándo")):
            ctk.CTkLabel(tabla, text=enc, font=F(11, True), text_color=COL["muted"], anchor="w", height=18).grid(
                row=0, column=col, sticky="w", padx=(12, 8), pady=(8, 4))
        siguiente = proximos[0][0]
        for i, (f, t) in enumerate(sorted(cortes_entre(hoy - timedelta(days=35), hoy + timedelta(days=100)).items()), 1):
            pasado = f < hoy
            color = COL["faint"] if pasado else (C("#B91C1C") if f == siguiente and urgente else
                                                 COL["warning"] if f == siguiente else COL["text"])
            fuente = F(12, f == siguiente)
            cuando = "ya pasó" if pasado else self._cuando(f, hoy)
            for col, valor in enumerate((f"{DIAS_SEMANA[f.weekday()][:3]} {f:%d/%m/%Y}", t[0].upper() + t[1:], cuando)):
                ctk.CTkLabel(tabla, text=valor, font=fuente, text_color=color, anchor="w", height=20).grid(
                    row=i, column=col, sticky="w", padx=(12, 8), pady=(0, 3))
        ctk.CTkFrame(tabla, height=6, fg_color="transparent").grid(row=99, column=0)

        botones = ctk.CTkFrame(win, fg_color="transparent")
        botones.pack(side="bottom", fill="x", padx=28, pady=(14, 20))
        caja.pack(fill="both", expand=True, padx=28, pady=(18, 8))
        self._btn(botones, "Entendido", win.destroy, "primary").pack(side="right", ipadx=12)
        self._btn(botones, "Página del calendario (FONACOT)", lambda: webbrowser.open(URL_CALENDARIO), "link").pack(
            side="left")
        if bloqueadas:
            def ver():
                win.destroy()
                self._apply_filter(FILTRO_INFORMATIVAS_TODAS)
            self._btn(botones, "Ver las que ya no se mandan", ver, "secondary").pack(side="right", padx=(0, 8))
        ajustar()
        win.update_idletasks()
        x = self.winfo_rootx() + (self.winfo_width() - win.winfo_width()) // 2
        y = self.winfo_rooty() + (self.winfo_height() - win.winfo_height()) // 3
        win.geometry(f"+{max(0, x)}+{max(0, y)}")
        win.attributes("-topmost", True)
        win.after(200, lambda: (win.winfo_exists() and win.focus_force()))
        if urgente:
            self.bell()

    def _aviso_corte(self, mins: int):
        active_rows, _ = self._get_active_and_filtered_rows()
        pendientes = sum(1 for r in active_rows if r['Row_ID'] not in self.validated_row_ids)
        self.bell()
        self._toast(
            f"Faltan {mins} min para la hora habitual del ticket único.\n"
            + (f"Hay {pendientes} de {len(active_rows)} renovaciones sin validar." if self.all_rows
               else "Aún no se han extraído las renovaciones de hoy (F5)."),
            color=C("#EF4444"), ms=12000
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

        if self._es_usuario():
            # El usuario ve su lote de un día (el que eligió en «Periodo de recepción»), por su gafete
            self.var_modo.set("AGENTE")
            self.seg_tipo.set(TIPOS_REVISION[0])
        modo = self.var_modo.get()
        agente_filter = self.ent_agente.get().strip()
        if modo == "AGENTE" and not agente_filter:
            if not auto:
                if self._es_usuario():
                    self._cambiar_gafete()
                    return
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
                    "«Configuración» → «Validaciones automáticas» (panel izquierdo).\n\n"
                    "Esta vez se extrae sin validar en la base."
                )
            sql_cfg = None

        account = self.cmb_account.get().strip()
        if account in ("Sin conexión a Outlook", "Cargando cuentas..."):
            account = ""

        self._aplicar_inhabiles()
        if auto and self._ultima_busqueda and self._es_usuario() and self._ultima_busqueda[0]:
            self._ultima_busqueda = None  # El usuario no hace la revisión semanal (era del validador)
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
        self.btn_extract_hdr.configure(state="disabled", text="Analizando…")
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
            # Lo corregido a mano (antes de validar en la base: así la base confirma el número bueno)
            aplicar_correcciones(rows, dict(self.estado.data.get('correcciones') or {}))
            validaciones = {'ocr': dict(self.engine.ocr_info), 'sql': None,
                            'errores_lectura': list(getattr(self.engine, 'errores_lectura', []))}
            if sql_cfg and rows:
                progress_cb(f"🗄️ Validando {len(rows)} renovaciones en la base de datos...", 0.98)
                validaciones['sql'] = self._validar_sql(rows, lote_info, sql_cfg)
            try:
                # Calendario de emisión: lo generado antes de la emisión o del cierre de mes ya no se puede cancelar
                actualizar_calendario()
                validaciones['emision'] = marcar_bloqueos(rows, (lote_info or {}).get('lote_actual'))
                # Lo emitido, lo resuelto en iTop y lo que no procede ya no va en el lote (solo informativo)
                validaciones['reglas'] = aplicar_reglas(rows)
                reglas.recontar_envios(lote_info, rows)  # lo ya resuelto en iTop no cuenta como «NO enviada»
            except Exception as e:
                # Las renovaciones se muestran igual, con un aviso rojo: sin reglas no se sabe qué ya no se manda
                registro.error("Fallaron las reglas del lote")
                validaciones['reglas_error'] = str(e) or type(e).__name__
            self.after(0, lambda: self._on_extraction_done(rows, None, lote_info, auto, validaciones))
        except Exception as e:
            registro.error("Falló la extracción")
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
        self.btn_extract_hdr.configure(state="normal", text="Extraer y validar")
        self.progress_bar.set(1.0)

        if error_msg:
            self.lbl_status.configure(text=f"Error: {error_msg}")
            if not auto:
                messagebox.showerror("Error al extraer de Outlook", error_msg)
            return

        # Una corrección guardada mientras se extraía: se pone ya; la base la confirma en la siguiente actualización
        for r in aplicar_correcciones(rows, self.estado.data.get('correcciones') or {}):
            if r.get('SQL_Estado'):
                r['SQL_Estado'], r['SQL_Detalle'], r['SQL_Opciones'], r['SQL_Fecha_Alta'] = None, '', [], None
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
        if self.cmb_quick_filter.get() == FILTRO_AVISO:
            self.cmb_quick_filter.set(FILTRO_TODAS)  # los avisos cambian con cada extracción
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
        # Que el desglose del ticket no quede vacío: se muestra la primera renovación de la tabla
        hijos = self.tree_desglose.get_children()
        if hijos and not self.tree_desglose.selection():
            self.tree_desglose.selection_set(hijos[0])
            self.tree_desglose.see(hijos[0])
            self._on_select_desglose_row()

        active_rows, _ = self._get_active_and_filtered_rows()
        # Barra de estado: cuándo se actualizó; los números están en los indicadores, los filtros y los avisos
        msg = f"Actualizado a las {datetime.now():%H:%M}"
        if lote_info and lote_info.get('no_enviadas'):
            msg += f"  ·  {pl(lote_info['no_enviadas'], '# no [salió|salieron]')} en su día"
        self.lbl_status.configure(text=msg)
        self._aprender_texto_correo(lote_info)
        # Equipo: el analista publica su avance; el validador lee lo que ya marcaron
        self._publicar_equipo(ya=True)
        self._leer_equipo()
        if not (lote_info and lote_info.get('revision')):
            self.after(300, lambda: self._recordatorio_emision(auto=True))

        if self.new_keys:
            nuevas = sum(1 for r in active_rows if self._key_by_rid.get(r['Row_ID']) in self.new_keys)
            if nuevas:
                self.bell()
                self._toast(f"{pl(nuevas, '# [renovación nueva|renovaciones nuevas]')} desde el último escaneo.\nUsa el filtro «{FILTRO_NUEVAS}» para verlas.")

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
        usuario = self._es_usuario()
        if not info:
            self.lbl_page_title.configure(text="Mis renovaciones de hoy" if usuario else "Lote del día")
            self._info_lineas = ([self._texto_marca_guardada()], [])
            self._pintar_info()
            avisos, rojo = self._avisos_validaciones()
            if avisos:
                self._mostrar_aviso(avisos, rojo)
            else:
                self.notice.grid_remove()
            return

        hc = info['hora_corte'].strftime('%H:%M')
        c = info['conteos']
        # Rezagadas que siguen en el lote (las que ya pasaron la emisión o se resolvieron solo son informativas)
        rezagadas = sum(1 for r in self.all_rows if r.get('Lote_Estado') == LOTE_REZAGADA and va_en_lote(r))
        revision = info.get('revision')
        if revision:
            ini, fin = info['fecha_inicio'], info['fecha_fin']
            self.lbl_page_title.configure(
                text=f"Revisión de envíos del {fin:%d/%m/%Y}" if ini == fin else f"Revisión de envíos del {ini:%d/%m} al {fin:%d/%m/%Y}"
            )
        elif usuario:
            self.lbl_page_title.configure(text=f"Mis renovaciones  ·  {info['lote_actual']:%d/%m/%Y}")
        else:
            self.lbl_page_title.configure(text=f"Lote del {info['corte_fin']:%d/%m/%Y}")
        lineas = []
        conc = info.get('concentrado_previo')
        if info['metodo'] == 'concentrado' and conc:
            lineas.append(
                f"Ticket único anterior: «{conc['asunto']}», recibido {conc['recibido']:%d/%m %H:%M}"
                f" ({conc['remitente'] or conc['carpeta']})"
            )
        principal = ""
        if revision:
            semana = info.get('semana') or {}
            n_tu = len({f['asunto'] for f in semana.get('filas') or []})
            principal = (
                f"{info.get('en_periodo', 0)} renovaciones debían salir en el periodo  ·  enviadas: {info.get('enviadas', 0)} "
                f"(tarde: {info.get('tarde', 0)})  ·  NO enviadas: {info.get('no_enviadas', 0)}  ·  "
                f"datos equivocados: {info.get('errores_envio', 0)}  ·  sin comprobar: {info.get('sin_comprobar', 0)}"
            )
            lineas.append(
                f"Tickets únicos: {len(semana.get('filas') or [])} renglones de {pl(n_tu, '# [ticket único|tickets únicos]')} "
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
        envios = info.get('envios_hoy') or []
        envio_hoy = (f"Tickets únicos {dia_txt}: {', '.join(f'{e:%H:%M}' for e in envios)}" if len(envios) > 1
                     else f"Ticket único {dia_txt} enviado a las {info['corte_hoy']:%H:%M}" if info.get('corte_hoy')
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
            # Lo que ya está en un indicador (ya enviadas, en el lote) o en la chapa del reloj (pendiente) no se repite
            principal = "  ·  ".join(t for t in (
                f"Gafete {self.ent_agente.get().strip()}" if usuario else "",
                envio_hoy if info.get('corte_hoy') or info.get('pasado') else "",
                f"Rezagadas sumadas al lote: {rezagadas}" if rezagadas else "",
                f"Para el siguiente ticket único: {c[LOTE_SIGUIENTE]}" if c.get(LOTE_SIGUIENTE) else "",
            ) if t)
        self._info_lineas = ([principal], lineas)
        self._pintar_info()

        # Cada aviso: (texto, destino). Destino = las renovaciones de ese aviso (clic = verlas), un filtro o nada
        def ids(rows):
            return [r['Row_ID'] for r in rows]

        avisos, rojo = [], False
        aviso_corte = []
        if info.get('corte_de_hoy') and not revision:
            # Día de emisión o de cierre: lo que llegue después del ticket único se manda hoy en otro
            aviso_corte.append((f"📅 Hoy es {info['corte_de_hoy']}: lo que llegue después del ticket único se manda HOY "
                                f"en otro ticket único (mañana ya no se puede cancelar: ya se emitió).", None))
        base = [r for r in self.all_rows if not revision or r.get('En_Periodo', True)]
        if revision:
            # «NO enviadas» = siguen abiertas (o se cerraron por error): lo ya resuelto en iTop no se manda y no es falta
            no_env = [r for r in base if r.get('No_Enviada')]
            cerradas = [r for r in base if r.get('Cerrada_Sin_Envio')]
            errores = [r for r in base if r.get('Error_Envio')]
            if no_env:
                rojo = True
                avisos.append((f"{pl(len(no_env), '# [renovación NO salió|renovaciones NO salieron] en el ticket único del día que [le|les] tocaba')} "
                               f"y {pl(len(no_env), '[sigue pendiente: se suma|siguen pendientes: se suman]')} al lote {dia_txt}.", ids(no_env), True))
            if errores:
                rojo = True
                avisos.append((f"{pl(len(errores), '# [renovación salió|renovaciones salieron]')} en el ticket único con datos equivocados "
                               f"(cliente, crédito o folio de otro incidente).", ids(errores), True))
            # Lo que no salió en su día pero ya no se manda: no cuenta como NO enviada
            emitidas = [r for r in base if r.get('No_Enviada_Emitida')]
            no_van = [t for t in (f"{len(emitidas)} ya se emitieron (ya no se pueden cancelar)" if emitidas else "",
                                  f"{len(cerradas)} tienen el incidente resuelto en iTop" if cerradas else "") if t]
            if no_van:
                avisos.append((f"No salieron en su día y no se mandan: {' · '.join(no_van)}.", ids(emitidas + cerradas)))
            duplicadas = [r for r in base if r.get('Posible_Duplicado')]
            if duplicadas:
                avisos.append((f"{pl(len(duplicadas), '# [renovación ya había salido|renovaciones ya habían salido]')} antes en otro ticket único "
                               f"(posible solicitud duplicada).", ids(duplicadas)))
        else:
            # Lote del día: en rojo solo lo que cambia el lote de hoy; lo de días anteriores, en una línea
            # Solo las que todavía van (lo ya emitido o resuelto no se suma: queda en «No se mandan»)
            suman = [r for r in self.all_rows if r.get('No_Enviada') and not r.get('Cerrado_Por_Error')
                     and not self._fuera_de_lote(r) and r['Row_ID'] not in self.discarded_row_ids]
            if suman:
                rojo = True
                avisos.append((f"{pl(len(suman), '# [renovación|renovaciones] de días anteriores NO [salió y sigue abierta|salieron y siguen abiertas]')}: "
                               f"{pl(len(suman), 'se [suma|suman]')} al lote {dia_txt}.", ids(suman), True))
            err_hoy = [r for r in self.all_rows if r.get('Error_Envio') and r.get('Dia_Lote') == info['lote_actual']]
            if err_hoy:
                rojo = True
                avisos.append((f"{pl(len(err_hoy), '# [renovación salió|renovaciones salieron]')} en el ticket único {dia_txt} con datos "
                               f"equivocados.", ids(err_hoy), True))
            # Las resueltas en iTop ya van en «No se mandan» (no se cuentan dos veces)
            n_err = info.get('errores_envio', 0) - len(err_hoy)
            antes = [t for t in (f"{n_err} con datos equivocados" if n_err else "",
                                 f"{pl(info['duplicadas'], '# [posible duplicada|posibles duplicadas]')}" if info.get('duplicadas') else "") if t]
            if antes and not usuario:
                avisos.append((f"De días anteriores: {', '.join(antes)}. No cambian el lote de hoy; el detalle está en "
                               f"«Revisión semanal».", None))
        faltan = [r for r in self.all_rows if r.get('Falta_Hoy')]
        if faltan:
            avisos.append((f"{pl(len(faltan), '# [renovación|renovaciones]')} {dia_txt} no {pl(len(faltan), '[viene|vienen]')} en el ticket único que salió a las "
                           f"{info['corte_hoy']:%H:%M}: agrégalas o mándalas en el siguiente.", ids(faltan)))
        if info.get('avisos_fecha'):
            avisos.append(("Fecha corregida: " + "  ·  ".join(info['avisos_fecha']), None))
        sin_comprobar = [r for r in base if r.get('Lote_Estado') == LOTE_SIN_COMPROBAR]
        if sin_comprobar:
            avisos.append((f"{pl(len(sin_comprobar), '# [renovación es|renovaciones son]')} de días sin ticket único encontrado.",
                           ids(sin_comprobar)))
        if not revision and not usuario and not (info['metodo'] == 'concentrado' and conc):
            aviso = f"No se encontró el ticket único del {info['dia_anterior']:%d/%m}. "
            if info['metodo'] == 'marca':
                aviso += f"Se usó la marca local ({info['marca']:%d/%m %H:%M}) para decidir qué faltó enviar."
            else:
                aviso += f"Se usó la hora habitual ({hc}); revisa las rezagadas antes de enviar el lote."
            avisos.append((aviso, None))
        extra, rojo_extra = self._avisos_validaciones()
        # Arriba lo más importante (lo que va en rojo); el aviso del día de corte, enseguida
        primero, despues = (extra, avisos) if rojo_extra and not rojo else (avisos, extra)
        avisos = (primero[:1] + aviso_corte + primero[1:] + despues if rojo or rojo_extra
                  else aviso_corte + primero + despues)
        if avisos:
            self._mostrar_aviso(avisos, rojo or rojo_extra)
        else:
            self.notice.grid_remove()

    def _avisos_validaciones(self) -> tuple:
        """Avisos de la base de datos y de las capturas de pantalla: ([(texto, destino)], rojo)."""
        avisos, rojo = [], False
        revision = bool(self.lote_info and self.lote_info.get('revision'))

        def ids(rows):
            return [r['Row_ID'] for r in rows]

        if self.validaciones.get('reglas_error'):
            rojo = True
            avisos.append((f"No se aplicaron las reglas del lote (calendario, iTop, cierre por error): "
                           f"{self.validaciones['reglas_error']}. No envíes sin revisar cada una; el detalle está en "
                           f"{registro.ARCHIVO}.", None, True))
        if self.validaciones.get('errores_lectura'):
            rojo = True
            tickets = self.validaciones['errores_lectura']
            avisos.append((f"No se pudo leer el correo de {pl(len(tickets), '# [incidente|incidentes]')}: {', '.join(tickets[:8])}"
                           + ("…" if len(tickets) > 8 else "") + f". No están en la tabla: revísalos en iTop "
                           f"(detalle en {registro.ARCHIVO}).", None, True))

        def cuenta(r):
            """En la revisión, lo del periodo; en el lote del día, solo lo que va en el lote (no lo ya enviado)."""
            if r.get('Row_ID') in self.discarded_row_ids:
                return False
            if revision:
                return r.get('En_Periodo', True)
            return not self.lote_info or r.get('Lote_Estado') in (LOTE_HOY, LOTE_REZAGADA)

        del_lote = [r for r in self.all_rows if cuenta(r)]
        # Lo que todavía hay que mandar (en la revisión, lo ya enviado o resuelto no se vuelve a validar)
        pendientes = [r for r in del_lote if not self._fuera_de_lote(r)]
        if not revision:
            # Lo que iba al lote pero ya no se manda (informativo) y lo que necesita algo antes de enviarse
            # Solo las que de verdad van hoy (no las descartadas ni las que ya se emitieron)
            grupos = reglas.grupos_del_dia(self.all_rows, self.discarded_row_ids, self.forced_included_row_ids,
                                           desde=self._desde_lote())
            mal_cerradas = grupos['cerradas_error']
            if mal_cerradas:
                rojo = True
                avisos.append((f"{pl(len(mal_cerradas), '# [incidente se cerró por error y no se mandó|incidentes se cerraron por error y no se mandaron]')}: "
                               + ", ".join(sorted({r['Numero_Ticket'] for r in mal_cerradas})[:6])
                               + pl(len(mal_cerradas), ". [Va hoy; avisa a quien lo cerró|Van hoy; avisa a quien los cerró]."), ids(mal_cerradas), True))
            posibles = [r for r in grupos['posibles_cierre'] if not r.get('Cierre_Sin_Cancelar')]
            if posibles:
                avisos.append((f"{pl(len(posibles), '# [incidente quizá se cerró|incidentes quizá se cerraron]')} por error (falta TZ_LOCK): "
                               f"{pl(len(posibles), '[confírmalo|confírmalos]')} antes de enviar el lote.", ids(posibles)))
            sin_cancelar = [r for r in grupos['posibles_cierre'] if r.get('Cierre_Sin_Cancelar')]
            if sin_cancelar:
                rojo = True
                avisos.append((f"{pl(len(sin_cancelar), '# [incidente se resolvió|incidentes se resolvieron]')} en iTop, pero "
                               f"{pl(len(sin_cancelar), '[la renovación no está cancelada|las renovaciones no están canceladas]')} "
                               f"en la base y sí procede: "
                               + ", ".join(sorted({r['Numero_Ticket'] for r in sin_cancelar})[:6])
                               + ". Confirma con quien lo cerró; «Restaurar» la manda hoy.", ids(sin_cancelar), True))
            informativas = grupos['no_se_mandan']
            if informativas:
                avisos.append((f"No se mandan: {reglas.texto_no_se_mandan(grupos['por_tipo'])}.", ids(informativas)))
            sin_cal = [r for r in pendientes if r.get('Sin_Calendario')]
            if sin_cal:
                meses = ", ".join(sorted({r['Sin_Calendario'] for r in sin_cal}))
                avisos.append((f"El calendario no tiene la fecha de la emisión de {meses}: no se sabe si "
                               f"{pl(len(sin_cal), '# [renovación ya se emitió|renovaciones ya se emitieron]')}. Revisa el calendario de FONACOT "
                               f"antes de enviar.", ids(sin_cal)))
            autorizacion = [r for r in pendientes if r.get('Requiere_Autorizacion')]
            if autorizacion:
                avisos.append((f"{pl(len(autorizacion), '# [renovación pasó|renovaciones pasaron] el cierre de mes: solo se [manda|mandan]')} con "
                               f"autorización (veredicto «Autorización»).", ids(autorizacion)))
        res = self.validaciones.get('sql')
        if res and not res.get('ok'):
            avisos.append((f"No se pudo validar en la base de datos: {res.get('error')}", None))
        elif res:
            malas = [r for r in (del_lote if revision else pendientes) if r.get('SQL_Estado') in SQL_PROBLEMAS]
            corregidos = [r for r in pendientes if r.get('SQL_Estado') == SQL_PADRE_CORREGIDO]
            if corregidos:
                avisos.append((f"{pl(len(corregidos), '# [renovación traía|renovaciones traían]')} el crédito PADRE: ya se puso su renovación "
                               f"(hijo) como crédito a cancelar. Valídalas: "
                               + ", ".join(f"{r['Numero_Ticket']} ({r.get('Credito_Padre_Enviado')} → "
                                           f"{r['Credito_Renovado_Hijo']})" for r in corregidos[:5])
                               + ("…" if len(corregidos) > 5 else "") + ".", ids(corregidos)))
            if malas:
                rojo = True
                avisos.append((f"{pl(len(malas), '# [renovación no cuadra|renovaciones no cuadran]')} con la base.", ids(malas), True))
            sin_renov = [r for r in pendientes if r.get('SQL_Clientes_Sin_Renov')]
            if sin_renov:
                avisos.append(("Incidentes con varios clientes donde alguno no cuenta con renovaciones: "
                               + ", ".join(sorted({r['Numero_Ticket'] for r in sin_renov})[:6]) + ".", ids(sin_renov)))
            if res.get('error_padre'):
                avisos.append((f"No se pudo revisar si el crédito es el padre: {res['error_padre']}", None))
            if res.get('error_tz_lock'):
                avisos.append((f"No se pudo leer TZ_LOCK (para saber si un incidente se cerró por error): "
                               f"{res['error_tz_lock']}", None))
            mesa = [r for r in pendientes if r.get('SQL_Estado') == SQL_MESA_CREDITO]
            if mesa:
                avisos.append((f"{pl(len(mesa), '# [renovación es de otro día pero la manda|renovaciones son de otro día pero las manda]')} la Mesa de "
                               f"Crédito: sí {pl(len(mesa), '[va|van]')} (no les tocó la emisión ni el cierre de mes).", ids(mesa)))
            app_plazo = [r for r in pendientes if r.get('SQL_Estado') == SQL_APP_PLAZO]
            if app_plazo:
                avisos.append((f"{pl(len(app_plazo), '# [renovación es de otro día pero la generó|renovaciones son de otro día pero las generó]')} la APP con plazo "
                               f"menor: sí {pl(len(app_plazo), '[va|van]')} (columna «Base de datos»).", ids(app_plazo)))
            cedula = [r for r in pendientes if r.get('SQL_Cedula')] if revision else []
            if cedula:
                avisos.append((f"{pl(len(cedula), '# [renovación ya se emitió (está|renovaciones ya se emitieron (están]')} en CR_CEDULAS_DETALLE): favor de "
                               f"validar.", ids(cedula)))
            sugerencias = [r for r in del_lote if r.get('SQL_Estado') == SQL_SUGERENCIA]
            if sugerencias:
                avisos.append((f"{pl(len(sugerencias), '# [renovación a la que le falta un dato tiene|renovaciones a las que les falta un dato tienen]')} sugerencia de la "
                               f"base (columna «Base de datos»).", ids(sugerencias)))
            if res.get('error_detalle'):
                avisos.append((f"La segunda consulta (APP / plazo / cédula) falló: {res['error_detalle']}", None))
            if revision and res.get('enviadas_con_problema'):
                rojo = True
                avisos.append((f"{pl(res['enviadas_con_problema'], '# [renglón|renglones] de los tickets únicos no [cuadra|cuadran]')} con la "
                               f"base: pestaña «Tickets únicos».", "pestaña:enviado", True))
        ocr = self.validaciones.get('ocr') or {}
        de_imagen = [r for r in del_lote if r.get('Dato_Imagen')]
        if de_imagen:
            avisos.append((f"{pl(len(de_imagen), '# [renovación se completó|renovaciones se completaron]')} leyendo una captura de pantalla: "
                           f"revísalas antes de enviarlas.", ids(de_imagen)))
        tickets_lote = {r.get('Numero_Ticket') for r in del_lote}
        sin_leer = [t for t in ocr.get('sin_leer') or [] if t in tickets_lote]
        if sin_leer:
            avisos.append(("No se encontró cliente o crédito en la imagen de: " + ", ".join(sin_leer[:8])
                           + ("…" if len(sin_leer) > 8 else ""),
                           ids(r for r in del_lote if r.get('Numero_Ticket') in sin_leer)))
        if ocr.get('error'):
            avisos.append((f"No se pudieron leer las capturas de pantalla: {ocr['error']}", None))
        return avisos, rojo

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

    # ================= TRABAJO EN EQUIPO (carpeta compartida) =================

    def _carpeta_equipo(self) -> str:
        return (self.config_data.get("carpeta_compartida") or "").strip() if EQUIPO_ACTIVO else ""

    def _elegir_carpeta(self):
        carpeta = filedialog.askdirectory(title="Carpeta compartida del equipo", initialdir=self._carpeta_equipo() or None)
        if carpeta:
            self.ent_carpeta.delete(0, "end")
            self.ent_carpeta.insert(0, os.path.normpath(carpeta))
            self._probar_carpeta()

    def _probar_carpeta(self):
        self._save_config()
        error = equipo.carpeta_valida(self._carpeta_equipo())
        self._estado_equipo(error or ("Lista: tus marcas se publican ahí." if self._es_usuario()
                                      else "Lista: aquí verás lo que marca el equipo."), bool(error))
        if not error:
            if not self._es_usuario() and not self.btn_equipo.winfo_manager():
                self.btn_equipo.pack(side="left", padx=(8, 0), ipadx=6, after=self.btn_validar_visibles)
            self._publicar_equipo(ya=True)
            self._leer_equipo()

    def _estado_equipo(self, texto: str, error: bool = False):
        self._equipo_estado = texto
        self.lbl_equipo.configure(text=texto, text_color=COL["danger"] if error else COL["muted"])

    def _dia_equipo(self):
        dia = (self.lote_info or {}).get('lote_actual') or datetime.now().date()
        return dia.date() if isinstance(dia, datetime) else dia

    def _publicar_equipo(self, ya: bool = False):
        """Usuario: publica en la carpeta compartida lo que validó y sus dudas (un segundo después del último cambio)."""
        if not self._es_usuario() or not self._carpeta_equipo() or not self.all_rows:
            return
        if self._publicar_job:
            self.after_cancel(self._publicar_job)
        self._publicar_job = self.after(50 if ya else 900, self._publicar_ahora)

    def _publicar_ahora(self):
        self._publicar_job = None
        carpeta, gafete = self._carpeta_equipo(), self.ent_agente.get().strip()
        if not carpeta or not gafete:
            return
        validados = self.estado.data.get('validados', {})
        marcas, en_lote = {}, []
        for r in self.all_rows:
            llaves = llaves_fila(r)
            k = llaves[0]
            fuera = self._fuera_de_lote(r) or r['Row_ID'] in self.discarded_row_ids
            if not fuera:
                en_lote.append(r)
            nota = self.estado.duda(k)
            if nota:
                hora = (self.estado.data['dudas'].get(k) or {}).get('hora')
                marcas[k] = equipo.marca(r, equipo.DUDA, datetime.fromisoformat(hora) if hora else None, nota, llaves)
            elif r['Row_ID'] in self.validated_row_ids and not fuera:
                hora = validados.get(k)
                marcas[k] = equipo.marca(r, equipo.VALIDADA, datetime.fromisoformat(hora) if hora else None, "", llaves)
        avance = {"total": len(en_lote), "validadas": sum(1 for r in en_lote if r['Row_ID'] in self.validated_row_ids),
                  "revisar": sum(1 for r in en_lote if r['Row_ID'] not in self.validated_row_ids and self._tiene_alerta(r)),
                  "dudas": sum(1 for m in marcas.values() if m["estado"] == equipo.DUDA)}
        nombre = next((r['Nombre_Asignado'] for r in self.all_rows if str(r.get('Gafete')) == gafete), "")
        dia = self._dia_equipo()

        def publicar():
            error = equipo.publicar(carpeta, gafete, nombre, dia, marcas, avance)
            self.after(0, lambda: self._estado_equipo(
                error or f"Publicado a las {datetime.now():%H:%M}: {pl(avance['validadas'], '# [validada|validadas]')}"
                         + (f", {pl(avance['dudas'], '# [duda|dudas]')}" if avance['dudas'] else "") + " para el validador.", bool(error)))
        threading.Thread(target=publicar, daemon=True).start()

    def _equipo_tick(self):
        self.after(15000, self._equipo_tick)
        self._leer_equipo()

    def _leer_equipo(self):
        """Validador: lee lo que publicó el equipo (en segundo plano, para no trabar la pantalla si la red tarda)."""
        carpeta = self._carpeta_equipo()
        if self._es_usuario() or not carpeta or self._equipo_leyendo:
            return
        self._equipo_leyendo = True
        dia = self._dia_equipo()

        def leer():
            res = equipo.leer_equipo(carpeta, dia)
            self.after(0, lambda: self._equipo_leido(res))
        threading.Thread(target=leer, daemon=True).start()

    def _equipo_leido(self, res: dict):
        self._equipo_leyendo = False
        firma = lambda d: {(k, m.get("estado"), m.get("hora")) for k, m in d["marcas"].items()}  # noqa: E731
        cambio = firma(res) != firma(self._equipo)
        nuevas = [m for k, m in res["marcas"].items() if (k, m.get("estado"), m.get("hora")) not in self._equipo_vistas]
        self._equipo_vistas |= firma(res)
        self._equipo = res
        if res["errores"]:
            self._estado_equipo(res["errores"][0], True)
        else:
            n = len(res["analistas"])
            self._estado_equipo(f"{pl(n, '# [analista|analistas]')} publicando hoy · {pl(len(res['marcas']), '# [marca|marcas]')}. "
                                f"Revisado a las {datetime.now():%H:%M}." if n else
                                "Conectado. Ningún analista ha publicado hoy.")
        if nuevas and self.all_rows:
            por_quien = {}
            for m in nuevas:
                quien = f"{m.get('nombre') or 'Gafete'} ({m.get('gafete')})"
                v, d = por_quien.get(quien, (0, 0))
                por_quien[quien] = (v + (m.get("estado") == equipo.VALIDADA), d + (m.get("estado") == equipo.DUDA))
            partes = [f"{q} validó {v}" + (f" y dejó {pl(d, '# [duda|dudas]')}" if d else "") if v else f"{q} dejó {pl(d, '# [duda|dudas]')}"
                      for q, (v, d) in por_quien.items()]
            self._toast("Equipo: " + "; ".join(partes) + ".", COL["primary"], ms=9000,
                        accion=("Ver", lambda: self._apply_filter(FILTRO_DUDAS if any(
                            m.get("estado") == equipo.DUDA for m in nuevas) else FILTRO_EQUIPO)))
        if cambio and self.all_rows:
            self._refrescar_tras_accion()
            self._mostrar_seleccion()

    def _marca_equipo(self, r: dict):
        """Lo que el equipo publicó de esta renovación (solo lo ve el validador)."""
        if self._es_usuario() or not self._equipo["marcas"]:
            return None
        return equipo.marca_de_fila(llaves_fila(r), self._equipo["marcas"])

    def _texto_equipo(self, r: dict) -> str:
        """Columna «Equipo»: quién la marcó (validador) o la duda que dejaste (usuario)."""
        if not EQUIPO_ACTIVO:
            return ""
        if self._es_usuario():
            nota = self.estado.duda(self._key_by_rid.get(r['Row_ID'], ''))
            return f"? Duda: {nota}" if nota else ""
        m = self._marca_equipo(r)
        if not m:
            return ""
        hora = str(m.get("hora", ""))[11:16]
        quien = (m.get("nombre") or "").split(" ")[0] or m.get("gafete")
        return f"? {quien}: {m.get('nota')}" if m.get("estado") == equipo.DUDA else f"✓ {quien} {hora}"

    def _aceptar_equipo(self):
        """Validador: valida de una vez lo que el equipo ya marcó (menos lo que tiene algo que revisar)."""
        active_rows, _ = self._get_active_and_filtered_rows()
        marcadas = [(r, m) for r in active_rows if r['Row_ID'] not in self.validated_row_ids
                    for m in [self._marca_equipo(r)] if m and m.get("estado") == equipo.VALIDADA]
        if not marcadas:
            messagebox.showinfo("Equipo", "No hay renovaciones marcadas por el equipo pendientes de validar.")
            return
        aceptar = [(r, m) for r, m in marcadas if not self._tiene_alerta(r)]
        omitidas = [r for r, m in marcadas if self._tiene_alerta(r)]
        texto = pl(len(aceptar), "Se [validará|validarán] # [renovación|renovaciones] que el equipo ya marcó.")
        if omitidas:
            texto += (f"\n\nSe dejan fuera {len(omitidas)} con algo que revisar (revísalas tú): "
                      + ", ".join(r['Numero_Ticket'] for r in omitidas[:8]) + ("…" if len(omitidas) > 8 else ""))
        if not aceptar or not messagebox.askyesno("Validar lo del equipo", texto + "\n\n¿Continuar?"):
            if not aceptar:
                messagebox.showinfo("Validar lo del equipo", texto)
            return
        rids = [r['Row_ID'] for r, _ in aceptar]
        copia = self._estado_de(rids)
        self.validated_row_ids.update(rids)
        self.estado.set('validados', self._keys_for(rids), True)
        for r, m in aceptar:
            self.estado.registrar([self._key_by_rid.get(r['Row_ID'])],
                                  f"Validada (la marcó {m.get('nombre') or m.get('gafete')} a las {str(m.get('hora'))[11:16]})",
                                  guardar=False)
        self.estado.save()
        self._refrescar_tras_accion()
        self._con_deshacer(f"Validaste {len(rids)} que marcó el equipo", lambda: self._restaurar_estado(rids, copia))

    # ---- Corregir el cliente o el crédito a cancelar ----

    def _corregir_dato(self):
        """
        Corrige el cliente o el crédito a cancelar de la renovación seleccionada (por ejemplo, con la renovación de ese
        día que encontró la base): así al copiar y en el ticket único ya va el número bueno. Se guarda en este equipo y
        se vuelve a poner en cada actualización; lo del correo se sigue viendo en el desglose.
        """
        rows = self._selected_rows()
        if len(rows) != 1:
            self._toast("Selecciona una sola renovación para corregirla.", COL["muted"], ms=4000)
            return
        r = rows[0]
        antes = r.get('Correccion_Manual') or {}
        win = ctk.CTkToplevel(self)
        win.title(f"Corregir dato · {r.get('Numero_Ticket', '')}")
        win.configure(fg_color=COL["bg"])
        win.resizable(False, False)
        win.transient(self)
        ctk.CTkLabel(win, text="Corregir cliente o crédito a cancelar", font=F(17, True), text_color=COL["text"]).pack(
            anchor="w", padx=22, pady=(18, 2))
        correo = (f"El correo trae: cliente {antes.get('cliente', r.get('Numero_Cliente'))} · crédito "
                  f"{antes.get('credito', r.get('Credito_Renovado_Hijo'))}")
        ctk.CTkLabel(win, text=correo, font=F(12), text_color=COL["text_2"], anchor="w").pack(anchor="w", padx=22)
        detalle = re.sub(r"\s+", " ", str(r.get('SQL_Detalle') or '')).strip()
        if detalle:
            ctk.CTkLabel(win, text=f"Base de datos: {detalle}", font=F(11), text_color=COL["muted"], justify="left",
                         anchor="w", wraplength=440).pack(anchor="w", padx=22, pady=(6, 0))

        campos = ctk.CTkFrame(win, fg_color="transparent")
        campos.pack(fill="x", padx=22, pady=(14, 0))
        ctk.CTkLabel(campos, text="Cliente", font=F(12), text_color=COL["text_2"]).grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(campos, text="Crédito a cancelar", font=F(12), text_color=COL["text_2"]).grid(
            row=0, column=1, sticky="w", padx=(12, 0))
        ent_cli = ctk.CTkEntry(campos, width=200, height=34, font=F(13))
        ent_cred = ctk.CTkEntry(campos, width=200, height=34, font=F(13))
        ent_cli.grid(row=1, column=0)
        ent_cred.grid(row=1, column=1, padx=(12, 0))
        for ent, valor in ((ent_cli, r.get('Numero_Cliente')), (ent_cred, r.get('Credito_Renovado_Hijo'))):
            if valor and not str(valor).upper().startswith('SIN'):
                ent.insert(0, str(valor))

        def poner(cli, cred):
            for ent, valor in ((ent_cli, cli), (ent_cred, cred)):
                ent.delete(0, "end")
                ent.insert(0, valor)

        opciones = r.get('SQL_Opciones') or []
        if opciones:
            ctk.CTkLabel(win, text="La base tiene:", font=F(12), text_color=COL["text_2"]).pack(
                anchor="w", padx=22, pady=(12, 2))
            for cli, cred, fecha in opciones:
                texto = f"Usar {cred}  ·  cliente {cli}" + (f"  ·  alta {fecha:%d/%m/%Y}" if fecha else "")
                self._btn(win, texto, lambda c=cli, n=cred: poner(c, n), "link", height=28, anchor="w").pack(
                    anchor="w", padx=16)
        lbl_error = ctk.CTkLabel(win, text="", font=F(11), text_color=COL["danger"])
        lbl_error.pack(anchor="w", padx=22)

        def guardar(_e=None):
            cli, cred = re.sub(r"\D", "", ent_cli.get()), re.sub(r"\D", "", ent_cred.get())
            if not re.fullmatch(r"\d{5,10}", cli) or not re.fullmatch(r"\d{3,8}", cred):
                lbl_error.configure(text="Escribe el cliente (5 a 10 dígitos) y el crédito a cancelar (3 a 8 dígitos).")
                return
            win.destroy()
            self._guardar_correccion(r, cli, cred)

        def quitar():
            win.destroy()
            self._guardar_correccion(r, "", "")

        botones = ctk.CTkFrame(win, fg_color="transparent")
        botones.pack(fill="x", padx=22, pady=(8, 18))
        self._btn(botones, "Guardar", guardar, "primary").pack(side="right", ipadx=12)
        self._btn(botones, "Cancelar", win.destroy, "ghost").pack(side="right", padx=(0, 8))
        if antes:
            self._btn(botones, "Volver a lo del correo", quitar, "danger").pack(side="left")
        for ent in (ent_cli, ent_cred):
            ent.bind("<Return>", guardar)
        win.bind("<Escape>", lambda e: win.destroy())
        win.after(150, lambda: (win.lift(), win.focus_force(), ent_cred.focus_set(), ent_cred.select_range(0, "end")))
        try:
            win.grab_set()
        except tk.TclError:
            pass

    def _guardar_correccion(self, r: dict, cliente: str, credito: str):
        """Guarda la corrección (vacía = regresa lo del correo), la pone en pantalla y, con la base activa, actualiza
        para que la base confirme el número nuevo."""
        k = self._key_by_rid.get(r['Row_ID']) or row_key(r)
        previo = f"cliente {r.get('Numero_Cliente')} · crédito {r.get('Credito_Renovado_Hijo')}"
        quitar_correccion(r)
        if cliente and aplicar_correccion(r, cliente, credito, datetime.now().isoformat(timespec='seconds')):
            self.estado.corregir(k, cliente, credito)
            accion, msg = "Corrigió el dato", (f"Corregido: cliente {cliente} · crédito {credito}. Al copiar y en el "
                                               f"ticket único ya va así.")
        else:
            self.estado.corregir(k)
            accion, msg = "Quitó la corrección", "Se regresó a lo que trae el correo."
        self.estado.registrar([k], accion, f"{previo} → cliente {r.get('Numero_Cliente')} · crédito "
                                           f"{r.get('Credito_Renovado_Hijo')}")
        sql = bool(self.var_sql.get())
        if sql:
            # Lo que la base dijo era del número anterior: se vuelve a validar al actualizar
            r['SQL_Estado'], r['SQL_Detalle'], r['SQL_Opciones'], r['SQL_Fecha_Alta'] = None, '', [], None
        self._refrescar_tras_accion()
        self._show_row_detail(r)
        if sql and not self.is_extracting and self._ultima_busqueda:
            msg += " Actualizando para confirmarlo en la base…"
            self.after(200, lambda: self._start_extraction(auto=True))
        self._toast(msg, COL["success"], ms=6000)

    def _poner_duda(self):
        """Usuario: deja una duda al validador sobre la renovación seleccionada (le llega por la carpeta compartida)."""
        if not self._es_usuario():
            self._toast("Las dudas las dejan los analistas; aquí las ves en la columna «Equipo».", COL["muted"], ms=4000)
            return
        sel = self.tree_desglose.selection()
        r = next((x for x in self.all_rows if sel and x['Row_ID'] == sel[0]), None)
        if not r:
            return
        k = self._key_by_rid.get(r['Row_ID'])
        previa = self.estado.duda(k)
        dlg = ctk.CTkInputDialog(title="Duda para el validador",
                                 text=f"{r['Numero_Ticket']} · cliente {r['Numero_Cliente']} · crédito "
                                      f"{r['Credito_Renovado_Hijo']}\n\n¿Qué duda tienes? (vacío = quitar la duda)"
                                      + (f"\n\nAhora: «{previa}»" if previa else ""))
        nota = dlg.get_input()
        if nota is None:
            return
        nota = nota.strip()
        self.estado.poner_duda(k, nota)
        self.estado.registrar([k], "Duda para el validador" if nota else "Quitó la duda", nota)
        self._refrescar_tras_accion()
        self._show_row_detail(r)
        self._publicar_equipo()
        if not self._carpeta_equipo():
            self._toast("Duda guardada. Para que le llegue al validador configura la carpeta compartida "
                        "(Configuración → Trabajo en equipo).", COL["warning"], ms=8000)
        else:
            self._con_deshacer("Duda enviada al validador" if nota else "Duda quitada",
                               lambda: (self.estado.poner_duda(k, previa), self._publicar_equipo()))

    def _mover_seleccion(self, paso: int):
        tree = self.tree_desglose
        filas = list(tree.get_children(""))
        if not filas:
            return
        actual = (tree.selection() or [None])[0]
        i = filas.index(actual) + paso if actual in filas else 0
        i = max(0, min(len(filas) - 1, i))
        tree.selection_set(filas[i])
        tree.focus(filas[i])
        tree.see(filas[i])

    def _cambiar_tema(self):
        self.config_data["tema"] = "oscuro" if self.var_oscuro.get() else "claro"
        self._save_config()
        self._toast("El modo " + ("oscuro" if self.var_oscuro.get() else "claro") + " se aplica al volver a abrir la app.",
                    COL["primary"], ms=5000)

    def _cambiar_letra(self):
        self.config_data["letra_grande"] = bool(self.var_letra.get())
        self._save_config()
        self._estilo_tabla()

    # ================= FILTRADO Y VISTAS =================

    def _desde_lote(self):
        """Día del lote que se ve (None en la revisión semanal): «No se mandan» solo trae las de ese día o de después."""
        if not self.lote_info or self.lote_info.get('revision'):
            return None
        return self.lote_info.get('lote_actual')

    def _fuera_de_lote(self, r: dict) -> bool:
        """
        Fuera del ticket único de hoy: ya enviada, del siguiente, informativa (emitida, resuelta, no procede)… Lo que
        ya está cancelado en la base queda fuera aunque se haya metido a mano antes.
        """
        if r.get('Informativa_Tipo') == reglas.INFO_CANCELADA:
            return True
        return not va_en_lote(r) and r['Row_ID'] not in self.forced_included_row_ids

    def _veredicto(self, r: dict) -> dict:
        rid = r['Row_ID']
        if rid not in self._veredictos:
            self._veredictos[rid] = veredicto(
                r, en_lote=not self._fuera_de_lote(r), validada=rid in self.validated_row_ids,
                descartada=rid in self.discarded_row_ids, sql_activo=bool(self.var_sql.get()))
        return self._veredictos[rid]

    def _tiene_alerta(self, r: dict) -> bool:
        """Hay que corregir, revisar o pedir autorización antes de mandarla."""
        return self._veredicto(r)['nivel'] in NIVELES_REVISAR

    def _es_nueva(self, r: dict) -> bool:
        return self._key_by_rid.get(r['Row_ID']) in self.new_keys

    def _get_active_and_filtered_rows(self):
        auto_excl = bool(self.var_auto_exclude_reassigned.get()) and (self.var_modo.get() == "AGENTE")
        q_filter = self.cmb_quick_filter.get()
        search_txt = self.ent_search.get().strip().lower()
        self._veredictos = {}  # se recalculan: validar, descartar o restaurar los cambian

        revision = bool(self.lote_info and self.lote_info.get('revision'))
        active_rows = []
        visible_rows = []
        self._conteo_filtros = {f: 0 for f in self.cmb_quick_filter.principales}
        for r in self.all_rows:
            rid = r['Row_ID']
            is_disc = (rid in self.discarded_row_ids) or (auto_excl and r.get('Reasignado_A_Otro'))
            fuera = self._fuera_de_lote(r)
            fuera_periodo = revision and not r.get('En_Periodo', True)
            if not is_disc and not fuera and not fuera_periodo:
                active_rows.append(r)

            if search_txt:
                haystack = (
                    f"{r.get('Numero_Ticket', '')} {r.get('Numero_Cliente', '')} "
                    f"{r.get('Credito_Renovado_Hijo', '')} {r.get('Credito_Origen_Padre', '')} "
                    f"{r.get('Gafete', '')} {r.get('Nombre_Asignado', '')} "
                    f"{r.get('Motivo', '')} {r.get('Resumen_iTop', '')} {r.get('Lote_Etiqueta', '')}"
                ).lower()
                if search_txt not in haystack:
                    continue
            # Conteo de cada filtro rápido (para los botones: «Revisar 3»)
            for f in self._conteo_filtros:
                if self._pasa_filtro(r, f, is_disc, fuera, fuera_periodo, revision):
                    self._conteo_filtros[f] += 1
            if self._pasa_filtro(r, q_filter, is_disc, fuera, fuera_periodo, revision):
                visible_rows.append(r)

        return active_rows, visible_rows

    def _pasa_filtro(self, r: dict, q_filter: str, is_disc, fuera, fuera_periodo, revision) -> bool:
        rid, lote = r['Row_ID'], r.get('Lote_Estado')
        if q_filter == FILTRO_AVISO:
            return rid in self._ids_aviso
        if q_filter == FILTRO_PAPELERA:
            return bool(is_disc)
        if q_filter in FILTROS_REVISION:
            if revision and not r.get('En_Periodo', True):
                return False  # De la semana anterior: solo se escaneó para sumar rezagadas
            return {
                FILTRO_ERRORES: lambda: bool(r.get('Error_Envio')),
                FILTRO_SQL: lambda: r.get('SQL_Estado') in SQL_PROBLEMAS,
                FILTRO_ENVIADAS: lambda: lote in (LOTE_ENVIADA, LOTE_EN_CONCENTRADO),
                FILTRO_INFORMATIVAS: lambda: bool(r.get('Informativa') and fuera
                                                  and reglas.es_del_dia(r, self._desde_lote())),
                FILTRO_INFORMATIVAS_TODAS: lambda: bool(r.get('Informativa') and fuera),
                FILTRO_CERRADAS: lambda: bool(r.get('Cerrado_iTop')),
                FILTRO_SIGUIENTE: lambda: lote == LOTE_SIGUIENTE,
                FILTRO_NO_ENVIADAS: lambda: bool(r.get('No_Enviada')),
                FILTRO_TARDE: lambda: bool(r.get('Dias_Retraso')),
                FILTRO_SIN_COMPROBAR: lambda: lote == LOTE_SIN_COMPROBAR,
            }.get(q_filter, lambda: True)()
        if q_filter == FILTRO_ALERTAS and fuera and r.get('Posible_Cierre_Error') and not is_disc and not fuera_periodo:
            return True  # ¿Se cerró por error? No va sola, pero hay que revisarla
        if is_disc or fuera or fuera_periodo:
            return False
        is_val = rid in self.validated_row_ids
        return {
            FILTRO_REZAGADAS: lambda: lote == LOTE_REZAGADA,
            FILTRO_LISTAS: lambda: self._veredicto(r)['nivel'] == reglas.LISTA,
            FILTRO_MULTIPLES: lambda: bool(r.get('Es_Multiple')),
            FILTRO_ALERTAS: lambda: self._tiene_alerta(r),
            FILTRO_PENDIENTES: lambda: not is_val,
            FILTRO_VALIDADAS: lambda: is_val,
            FILTRO_NUEVAS: lambda: self._es_nueva(r),
            FILTRO_EQUIPO: lambda: bool(self._marca_equipo(r)),
            FILTRO_DUDAS: lambda: (self._marca_equipo(r) or {}).get("estado") == equipo.DUDA,
        }.get(q_filter, lambda: True)()

    def _chip(self, r: dict) -> tuple:
        """Desglose: (color de la barra, tag de la fila). La fila no se pinta; el estado va en la barra y el texto."""
        rid, nivel = r['Row_ID'], self._veredicto(r)['nivel']
        apagada = (rid in self.discarded_row_ids or self._fuera_de_lote(r)
                   or nivel in (reglas.NO_SE_MANDA, reglas.SIGUIENTE, reglas.DESCARTADA))
        if rid in self.validated_row_ids and not apagada:
            color = COL["success"]
        else:
            color = {reglas.CORREGIR: COL["danger"], reglas.REVISAR: COL["warning"],
                     reglas.AUTORIZACION: COL["warning"], reglas.SIN_COMPROBAR: COL["warning"],
                     reglas.LISTA: VEREDICTO_ESTILO[reglas.LISTA][1], reglas.ENVIADA: COL["primary"]}.get(nivel, COL["faint"])
        return color, "apagada" if apagada else "fila"

    def _error_en_boton(self, tipo, valor, rastro):
        """Un error al dar clic o con una tecla: queda en validador.log y se avisa abajo (la app sigue funcionando)."""
        registro.log().error("Error en la interfaz", exc_info=(tipo, valor, rastro))
        try:
            self.lbl_status.configure(text=f"Algo falló ({tipo.__name__}: {valor}). El detalle quedó en {registro.ARCHIVO}.")
        except Exception:
            pass

    def _refrescar_tras_accion(self):
        """Después de validar, descartar, restaurar o deshacer: la tabla y también los avisos (al buscar, solo la tabla)."""
        self._refresh_views()
        self._actualizar_encabezado()

    def _refresh_views(self):
        self._search_job = None
        active_rows, visible_rows = self._get_active_and_filtered_rows()

        # 1. KPIs por veredicto
        revision = bool(self.lote_info and self.lote_info.get('revision'))
        grupos = reglas.grupos_del_dia(self.all_rows, self.discarded_row_ids, self.forced_included_row_ids,
                                       solo_periodo=revision, desde=self._desde_lote())
        # ¿Se cerró por error? también se revisa
        alert_rows = [r for r in active_rows if self._tiene_alerta(r)] + grupos['posibles_cierre']
        listas = [r for r in active_rows if self._veredicto(r)['nivel'] == reglas.LISTA]
        val_count = sum(1 for r in active_rows if r['Row_ID'] in self.validated_row_ids)
        vivas = [r for r in self.all_rows if r['Row_ID'] not in self.discarded_row_ids
                 and (not revision or r.get('En_Periodo', True))]
        enviadas = [r for r in vivas if r.get('Lote_Estado') in (LOTE_ENVIADA, LOTE_EN_CONCENTRADO)]
        informativas = grupos['no_se_mandan']

        self.lbl_kpi_renovaciones.configure(text=str(len(active_rows)))
        self.lbl_kpi_listas.configure(text=str(len(listas)))
        self.lbl_kpi_alertas.configure(
            text=str(len(alert_rows)), text_color=COL["danger"] if alert_rows else COL["text"]
        )
        self.lbl_kpi_enviadas.configure(text=str(len(enviadas)))
        self.lbl_kpi_info.configure(text=str(len(informativas)))
        self.lbl_kpi_validados.configure(text=f"{val_count} / {len(active_rows)}")
        self.kpi_progress.set(val_count / len(active_rows) if active_rows else 0)
        if self._es_usuario():
            self._pintar_hero(active_rows, val_count)
        en_aviso = self.cmb_quick_filter.get() == FILTRO_AVISO
        if not self.all_rows:
            texto = ""
        elif en_aviso:
            texto = f"{len(visible_rows)} del aviso  ·  {len(active_rows)} en el lote"
        else:
            texto = f"{len(visible_rows)} de {len(self.all_rows)} renovaciones"
        self.lbl_count.configure(text=texto)
        if en_aviso:
            self.btn_salir_aviso.pack(side="left", padx=(0, 12), before=self.lbl_count)
        else:
            self.btn_salir_aviso.pack_forget()

        # 2. Pestaña Desglose
        prev_sel = set(self.tree_desglose.selection())
        self.tree_desglose.delete(*self.tree_desglose.get_children())
        self._celdas_completas.clear()
        for i, r in enumerate(visible_rows):
            rid = r['Row_ID']
            ver = self._veredicto(r)
            estado = ("✓  Validada" if rid in self.validated_row_ids and not self._fuera_de_lote(r)
                      else f"{VEREDICTO_ESTILO[ver['nivel']][3]}  {ver['corto']}")
            if self._es_nueva(r):
                estado += "  ·  Nueva"
            lote = plain_text(r.get('Lote_Etiqueta', ''))
            self._celdas_completas[(rid, "validado")] = f"{ver['titulo']}: {ver['accion']}"
            self._celdas_completas[(rid, "lote")] = (f"{lote}\n{plain_text(r['Informativa'])}" if r.get('Informativa')
                                                     else lote)
            self._celdas_completas[(rid, "cierre")] = resumen_cierre(r.get('Cierre_iTop'))
            self._celdas_completas[(rid, "sql")] = r.get('SQL_Detalle', '')
            texto_eq = self._texto_equipo(r)
            if texto_eq:
                self._celdas_completas[(rid, "equipo")] = texto_eq
            color, tag = self._chip(r)
            if tag == "fila" and i % 2:
                tag = "fila_alt"
            self.tree_desglose.insert(
                "", "end", iid=rid, image=self._barra(color),
                values=(estado,) + ((texto_eq,) if EQUIPO_ACTIVO else ()) + (
                    # La parte va junto al ticket solo si hay varias (la columna «Parte» empieza oculta)
                    r['Numero_Ticket'] + (f" ({r['Indice_En_Ticket']}/{r['Total_En_Ticket']})"
                                          if (r.get('Total_En_Ticket') or 1) > 1 else ""),
                    f"{r['Indice_En_Ticket']} de {r['Total_En_Ticket']}",
                    r['Numero_Cliente'],
                    r['Credito_Renovado_Hijo'],
                    r['Credito_Origen_Padre'],
                    texto_sql(r.get('SQL_Estado', ''))
                    + (f" → {r['SQL_Renovacion_Hija']}" if r.get('SQL_Renovacion_Hija') else ""),
                    r.get('Lote_Resumen') or lote_corto(lote),
                    resumen_cierre(r.get('Cierre_iTop'), corto=True),
                    r['Gafete'],
                    r['Nombre_Asignado'],
                    plain_text(r['Alerta_Renovacion']),
                    plain_text(r['Alerta_Asignacion']),
                    r['Motivo'],
                    self._fecha_llegada_str(r)
                ),
                tags=(tag,)
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

        publicados = {str(a.get('gafete')): a for a in self._equipo['analistas']}
        for i, ((gaf, nom), ag) in enumerate(sorted(agentes.items(), key=lambda x: x[1]['renovaciones'], reverse=True)):
            pendientes = ag['renovaciones'] - ag['validados']
            pub = publicados.get(str(gaf))
            if pub:
                av = pub.get('avance') or {}
                marco = f"{av.get('validadas', 0)} de {av.get('total', 0)}" + (
                    f" · {pl(av['dudas'], '# [duda|dudas]')}" if av.get('dudas') else "")
                senal = "sin señal" if pub.get('sin_senal') else str(pub.get('actualizado', ''))[11:16]
            else:
                marco, senal = ("—", "no publica") if self._carpeta_equipo() and not self._es_usuario() else ("", "")
            self.tree_agentes.insert(
                "", "end",
                values=(
                    gaf, nom, len(ag['tickets']), ag['renovaciones'], len(ag['multi']), ag['alertas'],
                    ag['rezagadas'], f"{ag['validados']} / {ag['renovaciones']}", pendientes
                ) + ((marco, senal) if EQUIPO_ACTIVO else ()),
                tags=("validado" if pendientes == 0 else ("normal_alt" if i % 2 else "normal"),)
            )
        self._reapply_sort(self.tree_agentes)

        # Contadores en las pestañas y en los filtros rápidos
        self.cmb_quick_filter.contar(self._conteo_filtros)
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
                + (f"   ·   (traía el padre {r['Credito_Padre_Enviado']})" if r.get('Credito_Padre_Enviado') else "")
            )
        )

        # Veredicto: qué es y qué hacer (con el color del estado)
        ver = self._veredicto(r)
        fondo, borde, texto, icono = VEREDICTO_ESTILO[ver['nivel']]
        self.banner.configure(fg_color=fondo, border_color=borde)
        self.lbl_ver_icono.configure(text=icono, fg_color=borde, text_color=texto)
        titulo = ver['titulo'] + ("  ·  validada" if r['Row_ID'] in self.validated_row_ids
                                  and ver['nivel'] in (reglas.LISTA,) + NIVELES_REVISAR else "")
        self.lbl_ver_titulo.configure(text=titulo, text_color=texto)
        self.lbl_ver_accion.configure(text=ver['accion'], text_color=COL["text_2"])

        box = self.txt_detail
        box.configure(state="normal")
        box.delete("1.0", "end")
        estilo = {reglas.OK: "ok", reglas.AVISO: "warn", reglas.MAL: "error", reglas.INFO: "muted"}

        box.insert("end", "COMPROBACIONES\n", "seccion")
        if self._fila_detalle is None or self._fila_detalle.get('Row_ID') != r.get('Row_ID'):
            self._checks_abiertos = False  # otra renovación: empieza plegado
        self._fila_detalle = r
        # Primero la que decide el veredicto y lo que hay que atender; lo demás («Otras comprobaciones»), plegado
        checks = ver['checks']
        otras = len(checks) - (ver.get('principales') or len(checks))
        for estado, texto in (checks if self._checks_abiertos or not otras else checks[:len(checks) - otras]):
            box.insert("end", f"{ICONOS[estado]}  ", (estilo[estado], "check"))
            fuente, _, resto = texto.partition(": ")
            box.insert("end", f"{fuente}: " if resto else "", ("titulo", "check"))
            box.insert("end", f"{resto or texto}\n", ("check",))
        if otras:
            box.insert("end", ("▾ Ocultar las otras comprobaciones" if self._checks_abiertos
                               else f"▸ Otras comprobaciones ({otras})") + "\n", ("mas",))
        if r.get('Imagen_Info') and not r.get('Dato_Imagen'):
            box.insert("end", f"{ICONOS[reglas.INFO]}  Imagen: {r['Imagen_Info']}.\n", ("muted", "check"))

        # Lo que dice cada fuente, lado a lado
        box.insert("end", "COMPARACIÓN\n", "seccion")
        filas = reglas.comparacion(r, en_lote=not self._fuera_de_lote(r))
        # Si no se manda, la columna del ticket único solo repetiría «no se manda»
        con_ticket = ver['nivel'] not in (reglas.NO_SE_MANDA, reglas.DESCARTADA)
        if filas and all(f['estado'] == reglas.OK for f in filas):
            fuentes = "el incidente, la base de datos y el ticket único" if con_ticket else "el incidente y la base de datos"
            box.insert("end", f"{ICONOS[reglas.OK]}  Cliente, crédito y fecha coinciden en {fuentes}.\n", ("ok", "check"))
        else:
            box.insert("end", "\tIncidente\tBase de datos" + ("\tTicket único" if con_ticket else "") + "\n",
                       ("tabla", "tabla_enc"))
            for fila in filas:
                box.insert("end", f"{ICONOS[fila['estado']]} {fila['dato']}\t", ("tabla", estilo[fila['estado']]))
                box.insert("end", f"{fila['incidente']}\t{fila['base']}" + (f"\t{fila['ticket']}" if con_ticket else "")
                           + "\n", ("tabla", "titulo"))

        # Lo que dijo el equipo (validador) o la duda que dejaste (usuario)
        texto_eq = self._texto_equipo(r)
        if texto_eq:
            box.insert("end", "EQUIPO\n", "seccion")
            m = self._marca_equipo(r)
            if m and m.get("estado") == equipo.VALIDADA:
                box.insert("end", f"{ICONOS[reglas.OK]}  La validó {m.get('nombre') or ''} ({m.get('gafete')}) "
                                  f"a las {str(m.get('hora'))[11:16]}.\n", ("ok", "check"))
            elif m:
                box.insert("end", f"{ICONOS[reglas.AVISO]}  Duda de {m.get('nombre') or ''} ({m.get('gafete')}): "
                                  f"{m.get('nota')}\n", ("warn", "check"))
            else:
                box.insert("end", f"{ICONOS[reglas.AVISO]}  Le dejaste esta duda al validador: "
                                  f"{texto_eq[len('? Duda: '):]}\n", ("warn", "check"))

        box.insert("end", "DATOS\n", "seccion")
        llegada = self._fecha_llegada_str(r)
        box.insert("end", "Cliente  ", "muted")
        box.insert("end", f"{r['Numero_Cliente']}", "dato")
        box.insert("end", "     Crédito a cancelar  ", "muted")
        box.insert("end", f"{r['Credito_Renovado_Hijo']}", "dato")
        if str(r.get('Credito_Origen_Padre') or '-') not in ('-', ''):
            box.insert("end", "     Origen  ", "muted")
            box.insert("end", f"{r['Credito_Origen_Padre']}", "dato")
        box.insert("end", "\n")
        box.insert("end", f"Llegó el {llegada}"
                   + (f"  ·  le toca el ticket único del {r['Dia_Lote']:%d/%m/%Y}" if r.get('Dia_Lote') else "")
                   + f"\nAsignado a {r['Gafete']} {r['Nombre_Asignado']}  ·  {plain_text(r.get('Alerta_Asignacion', ''))}\n",
                   "muted")
        if r.get('Resumen_iTop'):
            box.insert("end", f"Resumen en iTop: {r['Resumen_iTop']}\n", "muted")

        if len(siblings) > 1:
            box.insert("end", f"LAS {len(siblings)} RENOVACIONES DEL INCIDENTE {tnum}\n", "seccion")
            for s in siblings:
                vs = self._veredicto(s)
                ok = s['Row_ID'] in self.validated_row_ids
                tag = "ok" if ok or vs['nivel'] == reglas.LISTA else (
                    "error" if vs['nivel'] == reglas.CORREGIR else "warn" if vs['nivel'] in NIVELES_REVISAR else "muted")
                marca = "▶" if s is r else ("✓" if ok else "○")
                box.insert("end", f"{marca}  {s['Indice_En_Ticket']} de {s['Total_En_Ticket']}  ·  Cliente "
                                  f"{s['Numero_Cliente']}  ·  Crédito {s['Credito_Renovado_Hijo']}  —  {vs['corto']}\n",
                           (tag, "check"))

        cierre = r.get('Cierre_iTop')
        if cierre:
            box.insert("end", "CIERRE EN iTop\n", "seccion")
            fuente = "" if cierre.get('cerro_fuente') == 'correo' else " (agente asignado al cerrar)"
            box.insert("end", f"{cierre.get('estado') or 'Cerrado'}"
                       + (f" el {cierre['fecha_dt']:%d/%m/%Y %H:%M}" if cierre.get('fecha_dt') else "")
                       + (f"  ·  lo cerró {cierre['cerro']}{fuente}" if cierre.get('cerro') else "") + "\n", "cierre")
            box.insert("end", f"Motivo: {cierre.get('motivo') or 'el aviso no trae el texto de la solución'}\n", "titulo")
        if r.get('Historial_Asignaciones'):
            box.insert("end", "HISTORIAL DE ASIGNACIÓN\n", "seccion")
            box.insert("end", f"{plain_text(r['Historial_Asignaciones'])}\n", "muted")
        historial = self.estado.historial(self._key_by_rid.get(r['Row_ID'], ''))
        if historial:
            box.insert("end", "BITÁCORA\n", "seccion")
            for hora, accion, detalle in reversed(historial[-8:]):
                box.insert("end", f"{str(hora)[8:10]}/{str(hora)[5:7]} {str(hora)[11:16]}  ", "muted")
                box.insert("end", accion, "titulo")
                box.insert("end", f" — {detalle}\n" if detalle else "\n", "muted")
        box.configure(state="disabled")

        # Correo original con los números extraídos resaltados
        body = r.get('Descripcion_Completa', '') or ''
        self.txt_body.configure(state="normal")
        self.txt_body.delete("1.0", "end")
        self.txt_body.insert("1.0", body)
        numeros = set()
        for s in siblings:
            for campo in ('Numero_Cliente', 'Credito_Renovado_Hijo', 'Credito_Origen_Padre', 'Credito_Padre_Enviado'):
                numeros.update(n.lstrip('0') for n in re.findall(r'\d{3,10}', str(s.get(campo, ''))))
            for valor in [(s.get('Correccion_Manual') or {}).get(c) for c in ('cliente', 'credito')]:  # lo que traía el correo (se corrigió a mano)
                numeros.update(n.lstrip('0') for n in re.findall(r'\d{3,10}', str(valor or '')))
        for n in numeros:
            if not n:
                continue
            for m in re.finditer(rf'(?<!\d)0*{re.escape(n)}(?!\d)', body):
                self.txt_body.tag_add("hl", f"1.0+{m.start()}c", f"1.0+{m.end()}c")
        self.txt_body.configure(state="disabled")
        self._ajustar_alto_correo()
        if self._win_correo is not None and self._win_correo.winfo_exists():
            self._ventana_correo(enfocar=False)  # la ventana abierta sigue a la renovación seleccionada

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
        # Lo que no se manda (ya enviada, emitida, resuelta, descartada) no se valida; quitar una validación sí se puede
        por_id = {r['Row_ID']: r for r in self.all_rows}
        sel = tuple(rid for rid in sel if rid in self.validated_row_ids or (
            rid in por_id and not self._fuera_de_lote(por_id[rid]) and rid not in self.discarded_row_ids))
        if not sel:
            self._toast("No se manda: no hay nada que validar.")
            return
        to_on = [rid for rid in sel if rid not in self.validated_row_ids]
        to_off = [rid for rid in sel if rid in self.validated_row_ids]
        copia = self._estado_de(sel)
        self.validated_row_ids.update(to_on)
        self.validated_row_ids.difference_update(to_off)
        self.estado.set('validados', self._keys_for(to_on), True)
        self.estado.set('validados', self._keys_for(to_off), False)
        self.estado.registrar(self._keys_for(to_on), "Validada", guardar=False)
        self.estado.registrar(self._keys_for(to_off), "Quitó la validación")
        self._refrescar_tras_accion()
        self._publicar_equipo()
        if len(sel) > 1 or to_off:
            self._con_deshacer(f"Validaste {len(to_on)}" if to_on and not to_off else
                               f"Quitaste la validación a {len(to_off)}" if to_off and not to_on else
                               f"Cambiaste {len(sel)} validaciones", lambda: self._restaurar_estado(sel, copia))
        else:  # Una a una no se avisa (sería ruido), pero Ctrl+Z la deshace
            self._deshacer.append(("Validaste 1", lambda: self._restaurar_estado(sel, copia)))
            del self._deshacer[:-20]
        if len(sel) == 1 and to_on and self.config_data.get('avanzar_al_validar', True):
            if self._siguiente_pendiente(desde=sel[0]):
                return
        r = next((x for x in self.all_rows if x['Row_ID'] == sel[0]), None)
        if r:
            self._show_row_detail(r)

    def _siguiente_pendiente(self, desde: str = None) -> bool:
        """Selecciona la siguiente renovación visible que va en el lote y falta validar (da la vuelta al final)."""
        tree = self.tree_desglose
        filas = list(tree.get_children(""))
        if not filas:
            return False
        por_id = {r['Row_ID']: r for r in self.all_rows}
        actual = desde or (tree.selection() or [None])[0]
        inicio = filas.index(actual) + 1 if actual in filas else 0
        for rid in filas[inicio:] + filas[:inicio]:
            r = por_id.get(rid)
            if r and rid != actual and rid not in self.validated_row_ids and not self._fuera_de_lote(r) \
                    and rid not in self.discarded_row_ids:
                tree.selection_set(rid)
                tree.focus(rid)
                tree.see(rid)
                self._show_row_detail(r)
                return True
        # En esta vista ya no quedan; se dice cuántas faltan en todo el lote (con filtro o búsqueda puede haber más)
        faltan = sum(1 for r in self.all_rows if r['Row_ID'] not in self.validated_row_ids
                     and r['Row_ID'] not in self.discarded_row_ids and not self._fuera_de_lote(r))
        filtrada = self.cmb_quick_filter.get() != FILTRO_TODAS or bool(self.ent_search.get().strip())
        self.lbl_status.configure(
            text="Ya validaste todo el lote. ✔" if not faltan
            else f"No quedan pendientes en esta vista; en todo el lote faltan {faltan} por validar."
            if filtrada else f"Faltan {faltan} por validar.")
        return False

    def _validate_all_visible(self):
        """Valida lo visible, con vista previa: lo que hay que revisar o no va en el lote se queda fuera."""
        _, visible_rows = self._get_active_and_filtered_rows()
        if not visible_rows:
            return
        en_lote = [r for r in visible_rows if not self._fuera_de_lote(r) and r['Row_ID'] not in self.discarded_row_ids]
        pendientes = [r for r in en_lote if r['Row_ID'] not in self.validated_row_ids]
        if en_lote and not pendientes:
            rids = [r['Row_ID'] for r in en_lote]
            if not messagebox.askyesno("Quitar validación", f"Las {len(rids)} visibles ya están validadas.\n\n"
                                                            "¿Quitarles la validación?"):
                return
            copia = self._estado_de(rids)
            self.validated_row_ids.difference_update(rids)
            self.estado.set('validados', self._keys_for(rids), False)
            self.estado.registrar(self._keys_for(rids), "Quitó la validación", "en bloque")
            self._refrescar_tras_accion()
            self._publicar_equipo()
            self._con_deshacer(f"Quitaste la validación a {len(rids)}", lambda: self._restaurar_estado(rids, copia))
            return
        aceptar = [r for r in pendientes if not self._tiene_alerta(r)]
        revisar = [r for r in pendientes if self._tiene_alerta(r)]
        fuera = len(visible_rows) - len(en_lote)
        texto = pl(len(aceptar), "Se [validará|validarán] # [renovación visible|renovaciones visibles].")
        if revisar:
            texto += (f"\n\nSe dejan fuera {len(revisar)} que tienen algo que revisar (valídalas una por una): "
                      + ", ".join(r['Numero_Ticket'] for r in revisar[:8]) + ("…" if len(revisar) > 8 else ""))
        if fuera:
            texto += f"\n\n{fuera} no van en el ticket único de hoy y no se tocan."
        if not aceptar:
            messagebox.showinfo("Validar visibles", texto)
            return
        if not messagebox.askyesno("Validar visibles", texto + "\n\n¿Continuar?"):
            return
        rids = [r['Row_ID'] for r in aceptar]
        copia = self._estado_de(rids)
        self.validated_row_ids.update(rids)
        self.estado.set('validados', self._keys_for(rids), True)
        self.estado.registrar(self._keys_for(rids), "Validada", "en bloque")
        self._refrescar_tras_accion()
        self._publicar_equipo()
        self._con_deshacer(f"Validaste {len(rids)}", lambda: self._restaurar_estado(rids, copia))

    def _pedir_motivo(self, titulo: str, n: int):
        """El motivo para la bitácora ('' si no se pide; None si se canceló)."""
        if not self.config_data.get("pedir_motivo", True):
            return ""
        dlg = ctk.CTkInputDialog(title=titulo, text=f"{titulo} {pl(n, '# [renovación|renovaciones]')}.\n\n¿Por qué? (queda en la bitácora)")
        motivo = dlg.get_input()
        return None if motivo is None else motivo.strip()

    def _discard_selected_rows(self):
        sel = self.tree_desglose.selection()
        if not sel:
            messagebox.showinfo("Selecciona una fila", "Selecciona una o más filas del desglose para descartarlas.")
            return
        motivo = self._pedir_motivo("Descartar", len(sel))
        if motivo is None:
            return
        copia = self._estado_de(sel)
        self.discarded_row_ids.update(sel)
        self.estado.set('descartados', self._keys_for(sel), True)
        self.estado.registrar(self._keys_for(sel), "Descartada", motivo)
        self._refrescar_tras_accion()
        self._publicar_equipo()
        self._con_deshacer(f"Descartaste {len(sel)}", lambda: self._restaurar_estado(sel, copia))

    def _restore_selected_rows(self):
        sel = self.tree_desglose.selection()
        if sel:
            # Lo que quedó fuera del lote se puede meter a mano, salvo lo que ya se emitió o ya está cancelado en la
            # base (eso ya no se cancela)
            por_id = {r['Row_ID']: r for r in self.all_rows}
            incluir = [
                rid for rid in sel
                if rid in por_id and not va_en_lote(por_id[rid])
                and por_id[rid].get('Informativa_Tipo') not in (reglas.INFO_EMISION, reglas.INFO_CANCELADA)
            ]
            sel = [rid for rid in sel if rid in self.discarded_row_ids or rid in incluir]
            if not sel:
                self._toast("Nada que restaurar: lo seleccionado no está descartado y ya va en el lote, ya se emitió o "
                            "ya está cancelado en la base.")
                return
            motivo = self._pedir_motivo("Restaurar", len(sel))
            if motivo is None:
                return
            copia = self._estado_de(sel)
            self.discarded_row_ids.difference_update(sel)
            self.estado.set('descartados', self._keys_for(sel), False)
            self.estado.registrar(self._keys_for(sel), "Restaurada", motivo)
            self.forced_included_row_ids.update(incluir)
            self.estado.set('incluidos', self._keys_for(incluir), True)
            self._con_deshacer(f"Restauraste {len(sel)}", lambda: self._restaurar_estado(sel, copia))
        else:
            if not self.discarded_row_ids:
                return
            if not messagebox.askyesno("Restaurar", "¿Restaurar TODAS las filas descartadas?"):
                return
            todas = list(self.discarded_row_ids)
            copia = self._estado_de(todas)
            self.estado.set('descartados', self._keys_for(todas), False)
            self.estado.registrar(self._keys_for(todas), "Restaurada", "todas las descartadas")
            self.discarded_row_ids.clear()
            self._con_deshacer(f"Restauraste {len(todas)}", lambda: self._restaurar_estado(todas, copia))
        self._refrescar_tras_accion()
        self._publicar_equipo()

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
        visibles = [c for c in tree["columns"] if c not in ocultas or c in COLUMNAS_FIJAS]
        tree["displaycolumns"] = visibles
        # La última columna visible toma el espacio que sobra (no queda un hueco a la derecha)
        for c in tree["columns"]:
            tree.column(c, stretch=(c == visibles[-1]))

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
        self._var_posicion = tk.StringVar(value="derecha" if self._detalle_a_la_derecha() else "abajo")
        for valor, texto in POSICIONES_DETALLE:
            m.add_radiobutton(label=texto, value=valor, variable=self._var_posicion,
                              command=lambda: self._posicion_detalle(self._var_posicion.get()))
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
        self.config_data["detalle_posicion"], self.config_data["detalle_division"] = "derecha", 0.0
        self._aplicar_columnas(tree)
        self._armar_paneles()
        self._save_config()

    def _posicion_detalle(self, posicion: str):
        self.config_data["detalle_posicion"], self.config_data["detalle_division"] = posicion, 0.0
        self._armar_paneles()
        self._save_config()

    def _mostrar_zona(self, zona: str, visible: bool):
        self.config_data.setdefault("vista", dict(VISTA_DEFAULT))[zona] = visible
        self._aplicar_vista()
        self._save_config()

    def _aplicar_vista(self):
        """Muestra u oculta indicadores, panel de detalle y configuración lateral según «Vista»."""
        vista = self.config_data.get("vista", {})
        usuario = self._es_usuario()
        self.kpi_bar.grid_remove()
        self.hero.grid_remove()
        if vista.get("indicadores", True):
            (self.hero if usuario else self.kpi_bar).grid()
            self.paned.grid_configure(pady=0)
        else:
            self.paned.grid_configure(pady=(14, 0))
        if vista.get("configuracion", True) and not self._panel_auto_oculto:
            self.sidebar_wrap.grid()
            self.sidebar_sep.grid()
            self.btn_extract_hdr.pack_forget()
        else:
            # Panel izquierdo oculto: «Extraer y validar» pasa a los botones de arriba
            self.sidebar_wrap.grid_remove()
            self.sidebar_sep.grid_remove()
            self.btn_extract_hdr.pack(side="left", padx=(0, 8), ipadx=8, before=self._primer_accion)
        self.btn_panel.configure(fg_color=COL["track"] if not vista.get("configuracion", True) or self._panel_auto_oculto
                                 else "transparent")
        en_paned = str(self.detail_frame) in [str(p) for p in self.paned.panes()]
        if vista.get("detalle", True) and not en_paned:
            self.paned.add(self.detail_frame, weight=2)
            self.detail_frame.lift()
            self.after(200, self._ajustar_division)
        elif not vista.get("detalle", True) and en_paned:
            self.paned.forget(self.detail_frame)
        self._pedir_ajuste_pagina()

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
        self._to_clipboard("\n".join(lines), f"Copiado al portapapeles: {pl(len(rows), '# [valor|valores]')}.")

    def _copy_selected_rows(self):
        """Las filas seleccionadas con el formato del ticket único (INCIDENTE | FECHA | CLIENTE | CREDITO (HIJO))."""
        rows = self._selected_rows()
        if rows:
            tsv_text = build_ticket_unico_tsv(rows)
            if not set_clipboard_html(build_ticket_unico_html(rows, url_fn=self._itop_url), tsv_text):
                self._to_clipboard(tsv_text, "")
            self.lbl_status.configure(text=f"{pl(len(rows), '# [fila copiada|filas copiadas]')} con el formato del ticket único.")
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
        info = self.lote_info or {}
        if info.get('corte_de_hoy') and info.get('corte_hoy'):
            # Día de corte con el ticket único ya enviado: el siguiente solo lleva lo que no salió en el anterior
            active_rows = [r for r in active_rows if r.get('Lote_Estado') != LOTE_EN_CONCENTRADO]
            if not active_rows:
                messagebox.showinfo("Nada pendiente", f"Todo el lote ya salió en el ticket único de las "
                                                      f"{info['corte_hoy']:%H:%M}. Hoy es {info['corte_de_hoy']}: "
                                                      f"si llega algo más, se manda en otro ticket único hoy.")
                return []
        if not active_rows:
            messagebox.showwarning("Sin datos", "No hay renovaciones en el lote. Primero extrae (F5).")
        return active_rows

    def _copy_batch_to_clipboard(self):
        active_rows = self._lote_activo_o_aviso()
        if not active_rows:
            return
        # Solo la tabla con el formato del ticket único (sin saludo): cada quien copia las suyas y se las pasa a
        # quien manda el ticket único, que las junta en su correo
        tsv_text = build_ticket_unico_tsv(active_rows)
        html = build_ticket_unico_html(active_rows, url_fn=self._itop_url)
        como_tabla = set_clipboard_html(html, tsv_text)
        if not como_tabla:
            self.clipboard_clear()
            self.clipboard_append(tsv_text)
            self.update()
        if not self._es_usuario():
            self._guardar_marca_lote(active_rows)
        pendientes = sum(1 for r in active_rows if r['Row_ID'] not in self.validated_row_ids)
        revisar = sum(1 for r in active_rows if self._tiene_alerta(r))

        rezagadas = sum(1 for r in active_rows if r.get('Lote_Estado') == LOTE_REZAGADA)
        messagebox.showinfo(
            "Lote copiado al portapapeles",
            f"Se copiaron {len(active_rows)} renovaciones"
            + (f" (incluye {pl(rezagadas, '# [rezagada|rezagadas]')} del día anterior)" if rezagadas else "") + ".\n\n"
            + ("Van con el formato del ticket único, solo la tabla (sin saludo). Pégalas con Ctrl+V: en Outlook "
               "quedan como tabla y en Excel como celdas.\n\n" if como_tabla
               else "Pégalas con Ctrl+V en el correo o en Excel.\n\n")
            + (f"⚠ {revisar} todavía tienen algo que revisar (filtro «{FILTRO_ALERTAS}»).\n" if revisar else "")
            + (f"⚠ {pendientes} aún no están validadas.\n" if pendientes else "")
            + ("\nPásaselas a quien manda el ticket único." if self._es_usuario() else "\n" + self._texto_marca_guardada())
        )

    def _crear_correo_lote(self):
        active_rows = self._lote_activo_o_aviso()
        if not active_rows:
            return
        self._save_config()
        pendientes = sum(1 for r in active_rows if r['Row_ID'] not in self.validated_row_ids)
        cerradas = sorted({r['Numero_Ticket'] for r in active_rows if r.get('Cerrado_Por_Error')})
        avisos = []
        if pendientes:
            avisos.append(f"{pendientes} de {len(active_rows)} renovaciones aún no están validadas.")
        padres = [r for r in active_rows if r.get('Credito_Padre_Enviado')]
        if padres:
            avisos.append("Traían el crédito PADRE; en el correo va su renovación (valídalas): "
                          + ", ".join(f"{r['Numero_Ticket']} ({r['Credito_Padre_Enviado']}"
                                      f" → {r['Credito_Renovado_Hijo']})" for r in padres[:8]) + ".")
        sql_mal = [r['Numero_Ticket'] for r in active_rows if r.get('SQL_Estado') in SQL_PROBLEMAS]
        if sql_mal:
            avisos.append(f"{pl(len(sql_mal), '# [renovación no cuadra|renovaciones no cuadran]')} con la base de datos: {', '.join(sql_mal[:8])}"
                          + ("…" if len(sql_mal) > 8 else "") + f". Revisa «{FILTRO_SQL}».")
        autorizacion = sorted({r['Numero_Ticket'] for r in active_rows if r.get('Requiere_Autorizacion')})
        if autorizacion:
            avisos.append(f"{pl(len(autorizacion), '# [ticket llegó|tickets llegaron] después del cierre de mes y solo se [cancela|cancelan]')} con "
                          f"AUTORIZACIÓN: {', '.join(autorizacion[:8])}" + ("…" if len(autorizacion) > 8 else "") + ".")
        emitidas = sorted({r['Numero_Ticket'] for r in active_rows if r.get('Informativa_Tipo') == reglas.INFO_EMISION})
        if emitidas:
            avisos.append(f"{pl(len(emitidas), '# [ticket metido a mano ya pasó|tickets metidos a mano ya pasaron] la emisión y NO se [puede|pueden] cancelar')}: "
                          f"{', '.join(emitidas[:8])}.")
        de_imagen = [r['Numero_Ticket'] for r in active_rows if r.get('Dato_Imagen')]
        if de_imagen:
            avisos.append(f"{pl(len(de_imagen), '# [renovación tiene|renovaciones tienen]')} datos leídos de una imagen: {', '.join(de_imagen[:8])}.")
        if cerradas:
            avisos.append(f"{pl(len(cerradas), '# [ticket se cerró|tickets se cerraron] en iTop por error y sí [va|van]')} en el lote: "
                          f"{', '.join(cerradas[:8])}" + ("…" if len(cerradas) > 8 else "") + pl(len(cerradas), ". Avisa para que [lo|los] reabran."))
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
                return "NO SALIÓ - incidente cerrado" if estado == LOTE_CERRADA_SIN_ENVIO else "NO SALIÓ - va en el lote de hoy"
            if r.get('No_Enviada_Emitida'):
                return "No salió en su día - ya se emitió"
            if r.get('Cerrada_Sin_Envio') or str(r.get('Lote_Resumen') or '').startswith("No salió el"):
                return "No salió en su día - no se manda"
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
                    fill = fill_no_env if res.startswith("NO SALIÓ") else (fill_tarde if res == "Enviada tarde" else None)
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
    if "--autoprueba" in sys.argv:
        # Revisa que este equipo y el ejecutable tengan todo (Revisar_Equipo.bat); ver autoprueba.py
        import autoprueba
        sys.exit(autoprueba.main(APP_VERSION, con_ventana="--sin-ventana" not in sys.argv))
    app = FonacotValidadorRenovacionesApp()
    app.mainloop()
