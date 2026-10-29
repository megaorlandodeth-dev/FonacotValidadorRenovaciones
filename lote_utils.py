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

from core_renovaciones import first_digits, guardar_json, LOTE_REZAGADA


_EMOJI_RE = re.compile('[\U0001F000-\U0001FAFF☀-➿⬀-⯿⌚-⏿️‍]')


def pl(n: int, texto: str) -> str:
    """
    Concordancia con el número: «# [renovación|renovaciones] ya se [emitió|emitieron]» -> «1 renovación ya se emitió» o
    «3 renovaciones ya se emitieron». # es el número; de cada [uno|varios] queda el que toca.
    """
    return re.sub(r"\[([^|\]]*)\|([^\]]*)\]", lambda m: m.group(1) if n == 1 else m.group(2), texto).replace("#", str(n))


def plain_text(value) -> str:
    """Quita emojis/íconos de los textos del motor para mostrarlos en tablas y reportes con aspecto sobrio."""
    s = str(value or '').replace('➔', '→')
    s = _EMOJI_RE.sub('', s)
    return re.sub(r'\s{2,}', ' ', s).strip(" |·-")


def lote_corto(etiqueta: str) -> str:
    """'Rezagada 22/10 14:53 (el incidente viene en el ticket único pero NO este…)' -> '… · otros datos en el ticket único'."""
    m = re.match(r'(.*?)\s*\((.*)\)$', etiqueta or '', re.S)
    if not m:
        return etiqueta or ''
    base, motivo = m.groups()
    if motivo.startswith("el incidente viene en el ticket único pero NO"):
        motivo = "otros datos en el ticket único"
    elif motivo == "no viene en el ticket único":
        motivo = "no salió"
    return f"{base}  ·  {motivo}"


def row_key(row: dict) -> str:
    """
    Llave estable de una renovación (no depende del orden ni del gafete): Ticket|Cliente|Crédito. Si se corrigió a mano,
    con lo que traía el correo (así no se pierde lo ya marcado ni la corrección en la siguiente actualización).
    """
    antes = row.get('Correccion_Manual') or {}
    cli = first_digits(antes.get('cliente', row.get('Numero_Cliente'))) or 'SC'
    cred = first_digits(antes.get('credito', row.get('Credito_Renovado_Hijo'))) or 'SR'
    return f"{row.get('Numero_Ticket', '')}|{cli}|{cred}|{row.get('Indice_En_Ticket', '') if cred == 'SR' else ''}"


def _sin_dato(valor) -> bool:
    return not first_digits(valor) or str(valor).upper().startswith('SIN')


def aplicar_correccion(row: dict, cliente: str, credito: str, hora: str = '') -> bool:
    """
    El analista corrigió el cliente o el crédito a cancelar (por ejemplo, con la renovación de ese día que dice la
    base): lo que traía el correo queda en 'Correccion_Manual'. False si ya estaba corregida o no cambia nada.
    """
    if row.get('Correccion_Manual'):
        return False
    cli_antes, cred_antes = row.get('Numero_Cliente', ''), row.get('Credito_Renovado_Hijo', '')
    cliente, credito = (cliente or cli_antes), (credito or cred_antes)
    if (first_digits(cliente), first_digits(credito)) == (first_digits(cli_antes), first_digits(cred_antes)):
        return False
    row['Correccion_Manual'] = {'cliente': cli_antes, 'credito': cred_antes, 'hora': hora,
                                'nivel': row.get('Nivel_Alerta'), 'alerta': row.get('Alerta_Renovacion')}
    row['Numero_Cliente'], row['Credito_Renovado_Hijo'] = cliente, credito
    row['Datos_Incompletos'] = _sin_dato(cliente) or _sin_dato(credito)
    if not row['Datos_Incompletos'] and row.get('Nivel_Alerta') == 'incompleto':
        row['Nivel_Alerta'] = 'multiple' if row.get('Es_Multiple') else 'normal'
        row['Alerta_Renovacion'] = "✏️ Dato corregido a mano"
    return True


def quitar_correccion(row: dict) -> bool:
    """Regresa lo que traía el correo. False si no estaba corregida."""
    antes = row.pop('Correccion_Manual', None)
    if not antes:
        return False
    row['Numero_Cliente'], row['Credito_Renovado_Hijo'] = antes['cliente'], antes['credito']
    row['Datos_Incompletos'] = _sin_dato(antes['cliente']) or _sin_dato(antes['credito'])
    if antes.get('nivel'):
        row['Nivel_Alerta'], row['Alerta_Renovacion'] = antes['nivel'], antes.get('alerta', '')
    return True


def aplicar_correcciones(rows: list, correcciones: dict) -> list:
    """Vuelve a poner lo corregido a mano ({llave: {'cliente', 'credito', 'hora'}}) en las renovaciones recién leídas.
    Regresa las que se corrigieron ahora (las que ya lo estaban no cuentan)."""
    hechas = []
    for r in rows:
        cor = correcciones.get(row_key(r)) if not r.get('Correccion_Manual') else None
        if cor and aplicar_correccion(r, cor.get('cliente', ''), cor.get('credito', ''), cor.get('hora', '')):
            hechas.append(r)
    return hechas


def llaves_fila(row: dict) -> list:
    """La llave de la renovación y, si el crédito padre se cambió por su renovación, la llave con el padre (así se
    reconoce en otro equipo que no validó en la base)."""
    llaves = [row_key(row)]
    if row.get('Credito_Padre_Enviado'):
        llaves.append(row_key(dict(row, Credito_Renovado_Hijo=row['Credito_Padre_Enviado'])))
    return llaves


def _fecha_de(valor) -> str:
    """Fecha ISO de un registro guardado: texto, {'hora': …} o lista de eventos [[hora, …], …]."""
    if isinstance(valor, dict):
        return str(valor.get('hora', ''))
    if isinstance(valor, list):
        return str(valor[-1][0]) if valor and valor[-1] else ''
    return str(valor)


class EstadoValidaciones:
    """
    Persistencia simple en JSON: {categoria: {llave: fecha_iso}}, más 'dudas' ({llave: {'hora', 'nota'}}) y la
    'bitacora' ({llave: [[hora, acción, detalle], …]}). Purga lo que tenga más de `dias_retencion` días.
    """

    CATEGORIAS = ('validados', 'descartados', 'incluidos', 'dudas', 'bitacora', 'correcciones')

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
            guardar_json(self.path, self.data, indent=1, ensure_ascii=False)
        except Exception:
            pass

    def _purge(self):
        limite = (datetime.now() - timedelta(days=self.dias_retencion)).isoformat()
        for c in self.CATEGORIAS:
            self.data[c] = {k: v for k, v in self.data[c].items() if _fecha_de(v) >= limite}

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

    def registrar(self, keys, accion: str, detalle: str = '', guardar: bool = True):
        """Agrega a la bitácora de cada renovación: quién hizo qué y cuándo (se queda en este equipo)."""
        now = datetime.now().isoformat(timespec='seconds')
        for k in keys:
            self.data['bitacora'].setdefault(k, []).append([now, accion, detalle])
            del self.data['bitacora'][k][:-30]  # las últimas 30 por renovación
        if guardar:
            self.save()

    def historial(self, key: str) -> list:
        return [tuple(e) for e in self.data['bitacora'].get(key, []) if len(e) == 3]

    def poner_duda(self, key: str, nota: str):
        """Duda para el validador (nota vacía = se quita)."""
        if nota:
            self.data['dudas'][key] = {'hora': datetime.now().isoformat(timespec='seconds'), 'nota': nota}
        else:
            self.data['dudas'].pop(key, None)
        self.save()

    def duda(self, key: str) -> str:
        return (self.data['dudas'].get(key) or {}).get('nota', '')

    def corregir(self, key: str, cliente: str = '', credito: str = ''):
        """Guarda lo corregido a mano de esa renovación (sin cliente ni crédito = se quita la corrección)."""
        if cliente or credito:
            self.data['correcciones'][key] = {'hora': datetime.now().isoformat(timespec='seconds'),
                                              'cliente': cliente, 'credito': credito}
        else:
            self.data['correcciones'].pop(key, None)
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
        resumen += f" (incluye {pl(rezagadas, '# [rezagada|rezagadas]')} del día anterior)"
    partes.append(f'<p style="font-family:Segoe UI,Arial;font-size:10pt">{escape(resumen)}</p>')
    return "\n".join(partes)


# ======================= FORMATO "TICKET ÚNICO" (el correo que se envía cada día) =======================

TICKET_UNICO_COLUMNAS = ('INCIDENTE', 'FECHA', 'CLIENTE', 'CREDITO (HIJO)')

INTRO_TICKET_UNICO_DEFAULT = (
    "Hola,\n\n"
    "Se envía “{total}” solicitudes de cancelaciones para su atención, favor de asignar a incidentes."
)


def _fila_ticket_unico(r: dict) -> tuple:
    # FECHA es la de alta de la renovación (FECHA_ALTA de la primera consulta a la base), no la del incidente; sin
    # base, la que trae el correo, y si tampoco, la de llegada
    alta = r.get('SQL_Fecha_Alta') or r.get('Fecha_Alta_Correo') or r.get('Fecha_Llegada_DT') or r.get('Fecha_DT')
    fecha = alta.strftime('%d/%m/%Y') if alta else str(r.get('Fecha', ''))[:10]
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
