#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Utilidades del lote diario (sin dependencias de Outlook ni de la interfaz):
- EstadoValidaciones: guarda en disco qué renovaciones se validaron / descartaron / incluyeron a mano,
  para que no se pierdan al volver a escanear (auto-actualización) ni al cerrar la aplicación.
- build_lote_tsv / build_lote_html: el lote en texto tabulado (Excel) y en tabla HTML (correo de Outlook).
- make_cf_html: empaqueta HTML en el formato "HTML Format" del portapapeles de Windows.
- insert_into_html_body: inserta la tabla al inicio del cuerpo de un correo conservando la firma.
"""

import os
import re
import json
from datetime import datetime, timedelta
from html import escape

from core_renovaciones import first_digits, LOTE_REZAGADA


_EMOJI_RE = re.compile('[\U0001F000-\U0001FAFF☀-➿⬀-⯿⌚-⏿️‍]')


def plain_text(value) -> str:
    """Quita emojis/íconos de los textos del motor para mostrarlos en tablas y reportes con aspecto sobrio."""
    s = str(value or '').replace('➔', '→')
    s = _EMOJI_RE.sub('', s)
    return re.sub(r'\s{2,}', ' ', s).strip(" |·-")


def lote_corto(etiqueta: str) -> str:
    """'Rezagada 22/10 14:53 (el ticket está en el concentrado pero NO este…)' -> 'Rezagada 22/10 14:53 · otros datos en el concentrado'."""
    m = re.match(r'(.*?)\s*\((.*)\)$', etiqueta or '', re.S)
    if not m:
        return etiqueta or ''
    base, motivo = m.groups()
    if motivo.startswith("el ticket está en el concentrado"):
        motivo = "otros datos en el concentrado"
    elif motivo == "no está en el concentrado":
        motivo = "no se envió"
    return f"{base}  ·  {motivo}"


def row_key(row: dict) -> str:
    """Llave estable de una renovación (no depende del orden ni del gafete): Ticket|Cliente|Crédito."""
    cli = first_digits(row.get('Numero_Cliente')) or 'SC'
    cred = first_digits(row.get('Credito_Renovado_Hijo')) or 'SR'
    return f"{row.get('Numero_Ticket', '')}|{cli}|{cred}|{row.get('Indice_En_Ticket', '') if cred == 'SR' else ''}"


class EstadoValidaciones:
    """Persistencia simple en JSON: {categoria: {llave: fecha_iso}}. Purga lo que tenga más de `dias_retencion` días."""

    CATEGORIAS = ('validados', 'descartados', 'incluidos')

    def __init__(self, path: str, dias_retencion: int = 30):
        self.path = path
        self.dias_retencion = dias_retencion
        self.data = {c: {} for c in self.CATEGORIAS}
        self.load()

    def load(self):
        try:
            if os.path.exists(self.path):
                with open(self.path, 'r', encoding='utf-8') as f:
                    saved = json.load(f)
                for c in self.CATEGORIAS:
                    if isinstance(saved.get(c), dict):
                        self.data[c] = saved[c]
        except Exception:
            pass
        self._purge()

    def save(self):
        try:
            with open(self.path, 'w', encoding='utf-8') as f:
                json.dump(self.data, f, indent=1, ensure_ascii=False)
        except Exception:
            pass

    def _purge(self):
        limite = (datetime.now() - timedelta(days=self.dias_retencion)).isoformat()
        for c in self.CATEGORIAS:
            self.data[c] = {k: v for k, v in self.data[c].items() if str(v) >= limite}

    def has(self, categoria: str, key: str) -> bool:
        return key in self.data[categoria]

    def set(self, categoria: str, keys, on: bool):
        now = datetime.now().isoformat(timespec='seconds')
        for k in keys:
            if on:
                self.data[categoria][k] = now
            else:
                self.data[categoria].pop(k, None)
        self.save()


# ======================= LOTE PARA CORREO / PORTAPAPELES =======================

LOTE_COLUMNAS = [
    ('#', None),
    ('Ticket iTop', 'Numero_Ticket'),
    ('No. Cliente', 'Numero_Cliente'),
    ('Crédito Renovado (Cancelar)', 'Credito_Renovado_Hijo'),
    ('Crédito Origen (Padre)', 'Credito_Origen_Padre'),
    ('Gafete', 'Gafete'),
    ('Agente', 'Nombre_Asignado'),
    ('Motivo', 'Motivo'),
    ('Llegada', '_llegada'),
    ('Observación', '_obs'),
    ('Validado', '_validado'),
]


def _valor(row: dict, campo: str, idx: int, validados: set) -> str:
    if campo is None:
        return str(idx)
    if campo == '_llegada':
        llegada = row.get('Fecha_Llegada_DT')
        return llegada.strftime('%d/%m/%Y %H:%M') if llegada else str(row.get('Fecha', ''))
    if campo == '_obs':
        if row.get('Lote_Estado') == LOTE_REZAGADA:
            return "Rezagada del día anterior"
        if row.get('Total_En_Ticket', 1) > 1:
            return f"{row.get('Indice_En_Ticket')} de {row.get('Total_En_Ticket')} en el ticket"
        return ""
    if campo == '_validado':
        return "SÍ" if row.get('Row_ID') in validados else "PENDIENTE"
    return str(row.get(campo, '') or '')


def build_lote_tsv(rows: list, validados: set) -> str:
    lines = ["\t".join(h for h, _ in LOTE_COLUMNAS)]
    for i, r in enumerate(rows, 1):
        lines.append("\t".join(_valor(r, c, i, validados).replace("\t", " ") for _, c in LOTE_COLUMNAS))
    return "\n".join(lines)


def build_lote_html(rows: list, validados: set, titulo: str = "") -> str:
    """Tabla con estilos en línea (Outlook ignora las hojas de estilo)."""
    th = ("background:#0F2942;color:#FFFFFF;font-family:Segoe UI,Arial;font-size:10pt;"
          "padding:5px 8px;border:1px solid #CBD5E1;text-align:center")
    td = "font-family:Segoe UI,Arial;font-size:10pt;padding:4px 8px;border:1px solid #CBD5E1"
    partes = []
    if titulo:
        partes.append(f'<p style="font-family:Segoe UI,Arial;font-size:11pt"><b>{escape(titulo)}</b></p>')
    partes.append('<table style="border-collapse:collapse" cellspacing="0" cellpadding="0">')
    partes.append("<tr>" + "".join(f'<th style="{th}">{escape(h)}</th>' for h, _ in LOTE_COLUMNAS) + "</tr>")
    for i, r in enumerate(rows, 1):
        if r.get('Lote_Estado') == LOTE_REZAGADA:
            bg = "#E0E7FF"
        elif r.get('Total_En_Ticket', 1) > 1:
            bg = "#FEF3C7"
        else:
            bg = "#FFFFFF" if i % 2 else "#F8FAFC"
        celdas = "".join(
            f'<td style="{td};background:{bg}">{escape(_valor(r, c, i, validados))}</td>' for _, c in LOTE_COLUMNAS
        )
        partes.append(f"<tr>{celdas}</tr>")
    partes.append("</table>")
    rezagadas = sum(1 for r in rows if r.get('Lote_Estado') == LOTE_REZAGADA)
    resumen = f"Total de renovaciones a cancelar: {len(rows)}"
    if rezagadas:
        resumen += f" (incluye {rezagadas} rezagada(s) del día anterior)"
    partes.append(f'<p style="font-family:Segoe UI,Arial;font-size:10pt">{escape(resumen)}</p>')
    return "\n".join(partes)


# ======================= FORMATO "TICKET ÚNICO" (el correo que se envía cada día) =======================

TICKET_UNICO_COLUMNAS = ('INCIDENTE', 'FECHA', 'CLIENTE', 'CREDITO (HIJO)')

INTRO_TICKET_UNICO_DEFAULT = (
    "Hola,\n\n"
    "Se envía “{total}” solicitudes de cancelaciones para su atención, favor de asignar a incidentes."
)


def _fila_ticket_unico(r: dict) -> tuple:
    llegada = r.get('Fecha_Llegada_DT') or r.get('Fecha_DT')
    fecha = llegada.strftime('%d/%m/%Y') if llegada else str(r.get('Fecha', ''))[:10]
    return (str(r.get('Numero_Ticket', '')), fecha,
            str(r.get('Numero_Cliente', '') or ''), str(r.get('Credito_Renovado_Hijo', '') or ''))


def build_ticket_unico_tsv(rows: list) -> str:
    lines = ["\t".join(TICKET_UNICO_COLUMNAS)]
    lines += ["\t".join(v.replace("\t", " ") for v in _fila_ticket_unico(r)) for r in rows]
    return "\n".join(lines)


def texto_intro(intro: str, total: int) -> str:
    try:
        return intro.format(total=total).strip()
    except (KeyError, IndexError, ValueError):
        return intro.strip()


def build_ticket_unico_texto(rows: list, intro: str = "") -> str:
    """Lo mismo que el correo del ticket único en texto plano (saludo y tabla separada por tabuladores)."""
    tabla = build_ticket_unico_tsv(rows)
    return f"{texto_intro(intro, len(rows))}\n\n{tabla}" if intro else tabla


def build_ticket_unico_html(rows: list, intro: str = "", url_fn=None) -> str:
    """
    Cuerpo del correo con el mismo formato del "ticket único" que ya se envía:
    saludo con el total, y tabla INCIDENTE (con liga a iTop) | FECHA | CLIENTE | CREDITO (HIJO).
    `intro` admite {total}; vacío = sin texto (solo la tabla, para pegar).
    """
    font = "font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#000000"
    celda = f"{font};border:1px solid #000000;padding:2px 10px;text-align:center"
    partes = []
    if intro:
        for parrafo in re.split(r'\n\s*\n', texto_intro(intro, len(rows))):
            partes.append(f'<p style="{font};margin:0 0 10pt 0">{escape(parrafo).replace(chr(10), "<br>")}</p>')
    partes.append('<table style="border-collapse:collapse" cellspacing="0" cellpadding="0">')
    partes.append("<tr>" + "".join(f'<td style="{celda};font-weight:bold">{escape(h)}</td>' for h in TICKET_UNICO_COLUMNAS) + "</tr>")
    for r in rows:
        ticket, fecha, cli, cred = (escape(v) for v in _fila_ticket_unico(r))
        url = url_fn(r.get('Numero_Ticket', '')) if url_fn else ''
        if url:
            ticket = f'<a href="{escape(url)}" style="color:#000000">{ticket}</a>'
        partes.append(f'<tr><td style="{celda}">{ticket}</td><td style="{celda}">{fecha}</td>'
                      f'<td style="{celda}">{cli}</td><td style="{celda}">{cred}</td></tr>')
    partes.append("</table><p></p>")
    return "\n".join(partes)


def make_cf_html(fragment: str) -> bytes:
    """Empaqueta un fragmento HTML con el encabezado de offsets que exige el formato 'HTML Format' de Windows."""
    header_tpl = (
        "Version:0.9\r\nStartHTML:{0:010d}\r\nEndHTML:{1:010d}\r\n"
        "StartFragment:{2:010d}\r\nEndFragment:{3:010d}\r\n"
    )
    prefix = "<html><body><!--StartFragment-->"
    suffix = "<!--EndFragment--></body></html>"
    header_len = len(header_tpl.format(0, 0, 0, 0).encode('utf-8'))
    frag_b = fragment.encode('utf-8')
    start_html = header_len
    start_frag = start_html + len(prefix.encode('utf-8'))
    end_frag = start_frag + len(frag_b)
    end_html = end_frag + len(suffix.encode('utf-8'))
    header = header_tpl.format(start_html, end_html, start_frag, end_frag)
    return header.encode('utf-8') + prefix.encode('utf-8') + frag_b + suffix.encode('utf-8')


def insert_into_html_body(full_html: str, fragment: str) -> str:
    """Inserta `fragment` justo después de <body ...> (antes de la firma). Si no hay <body>, lo antepone."""
    if full_html and re.search(r'<body[^>]*>', full_html, re.I):
        return re.sub(r'(<body[^>]*>)', lambda m: m.group(1) + fragment, full_html, count=1, flags=re.I)
    return fragment + (full_html or '')


def set_clipboard_html(fragment: str, text: str) -> bool:
    """Pone en el portapapeles la tabla HTML (se pega como tabla en Outlook/Excel) y el texto tabulado. False si falla."""
    try:
        import win32clipboard
        import win32con
    except ImportError:
        return False
    try:
        cf_html = win32clipboard.RegisterClipboardFormat("HTML Format")
        win32clipboard.OpenClipboard()
        try:
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(cf_html, make_cf_html(fragment))
            win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, text)
        finally:
            win32clipboard.CloseClipboard()
        return True
    except Exception:
        return False
