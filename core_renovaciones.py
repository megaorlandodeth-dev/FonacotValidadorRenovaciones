#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Motor Analizador y Desglosador Especializado en Renovaciones de Crédito FONACOT (iTop / CREDERE).
Identifica tickets de Cancelación/Reversa de Renovación (incluso cuando vienen con títulos
como "Cancelación de crédito", "Apoyo", "Eventual", "Análisis | cancelación renovaciones", etc.)
y desglosa 1 Ticket -> N Filas de Renovación individuales (Cliente + Crédito Renovado/Hijo + Crédito Origen/Padre).
"""

import os
import re
import sys
import unicodedata
from datetime import datetime, date, time, timedelta
from urllib.parse import quote
from bs4 import BeautifulSoup


def make_soup(html: str) -> BeautifulSoup:
    """BeautifulSoup con lxml si está instalado; si no, el parser nativo de Python (antes fallaba en silencio)."""
    try:
        return BeautifulSoup(html, 'lxml')
    except Exception:
        return BeautifulSoup(html, 'html.parser')


# ======================= ABREVIATURAS DE UNA LETRA / CORTAS =======================
# Las sucursales escriben "C5864479", "T 4530500", "Cte. 3386153", "Cred. 61116", "R-91258", "C-74544"...
# Como la letra va pegada al número, los patrones con \b no los veían. Aquí se traducen a etiquetas completas.
_ABBR_CLIENTE = {'clte', 'cte', 'cli', 'clt', 'cta', 'trab', 'tr', 't'}
_ABBR_CREDITO = {'cred', 'crd', 'cr', 'cto', 'cdto', 'renov', 'reno', 'ren', 'r', 'pag'}
_ABBR_PADRE = {'orig', 'org', 'ant'}
_ABBR_AMBIGUO = {'c'}  # "C" puede ser Cliente o Crédito: se decide por la cantidad de dígitos

_ABBR_RE = re.compile(
    r'(?<![A-Za-zÁÉÍÓÚÑáéíóúñ\d])'
    r'(?P<lbl>clte|cte|cli|clt|cta|trab|tr|t|cred|crd|cdto|cto|cr|renov|reno|ren|r|pag|orig|org|ant|c)'
    r'\s*[\.:#\-/]?\s*(?P<num>\d{3,10})(?!\d)',
    re.I
)

# Los números de cliente FONACOT tienen 7 dígitos; los créditos/renovaciones, de 3 a 6
MIN_DIGITOS_CLIENTE_AMBIGUO = 7


def _expand_abbr_match(m: re.Match) -> str:
    lbl = m.group('lbl').lower()
    num = m.group('num')
    if num.startswith('0'):
        # Folios tipo R-012345 (UserRequest de iTop) u otros con ceros a la izquierda: no son clientes ni créditos
        return m.group(0)
    n = len(num)
    if lbl in _ABBR_CLIENTE:
        return f"Cliente: {num}" if 5 <= n <= 8 else m.group(0)
    if lbl in _ABBR_CREDITO:
        return f"Credito: {num}" if n <= 7 else m.group(0)
    if lbl in _ABBR_PADRE:
        return f"Credito anterior: {num}" if n <= 7 else m.group(0)
    if lbl in _ABBR_AMBIGUO:
        if MIN_DIGITOS_CLIENTE_AMBIGUO <= n <= 8:
            return f"Cliente: {num}"
        if n < MIN_DIGITOS_CLIENTE_AMBIGUO:
            return f"Credito: {num}"
    return m.group(0)


def expand_abbreviated_labels(text: str) -> str:
    """Convierte 'C5864479 C56572 Y 56574' -> 'Cliente: 5864479 Credito: 56572 Y 56574'."""
    if not text:
        return ""
    return _ABBR_RE.sub(_expand_abbr_match, text)


def strip_letter_prefixes(text: str) -> str:
    """Quita prefijos de letras pegados a números ('C57137' -> '57137', 'T-4530500' -> '4530500'). Para celdas de tablas."""
    if not text:
        return ""
    return re.sub(r'(?<![A-Za-zÁÉÍÓÚÑáéíóúñ])[A-Za-z]{1,5}\s*[\.:#\-]?\s*(?=\d{3,})', '', str(text))


def normalizar_correo(texto: str) -> str:
    """Espacios duros (&nbsp;), Ñ y acentos compuestos de dos caracteres y saltos \\r\\n como en un texto normal."""
    texto = unicodedata.normalize("NFC", str(texto or ""))
    texto = re.sub(r'[\u00a0\u2007\u202f\u200b\ufeff]', lambda m: '' if m.group(0) in '\u200b\ufeff' else ' ', texto)
    return texto.replace('\r\n', '\n').replace('\r', '\n')


def first_digits(value: str) -> str:
    """Primer bloque numérico de un valor ('5864479 (⚠️ En texto también: 123)' -> '5864479'). '' si no hay."""
    m = re.search(r'\d+', str(value or ''))
    return m.group(0).lstrip('0') if m else ''


# ======================= ENLACES A iTop =======================
ITOP_BASE_URL = "https://mservicio.fonacot.gob.mx/pages/UI.php"
# Prefijo de la referencia de iTop -> clase del objeto (I-xxxxxx es Incident, no UserRequest)
ITOP_CLASES_POR_PREFIJO = {'I': 'Incident', 'R': 'UserRequest', 'C': 'Change', 'P': 'Problem'}


def _split_ticket_ref(ticket: str) -> tuple:
    m = re.match(r'\s*([A-Za-z]+)\s*-\s*0*(\d+)', str(ticket or ''))
    if not m:
        digits = re.sub(r'\D', '', str(ticket or '')).lstrip('0')
        return 'I', digits
    return m.group(1).upper(), m.group(2)


def build_itop_url(ticket: str, template: str = "") -> str:
    """
    URL de detalle del ticket en iTop. `template` (opcional, desde la configuración) admite
    {id}, {ref} y {clase}, p. ej. 'https://otro-servidor/pages/UI.php?operation=details&class={clase}&id={id}'.
    """
    prefix, num = _split_ticket_ref(ticket)
    if not num:
        return ""
    clase = ITOP_CLASES_POR_PREFIJO.get(prefix, 'Ticket')
    ref = f"{prefix}-{num}"
    if template and template.strip():
        return template.strip().format(id=num, num_id=num, ref=quote(ref), clase=clase)
    return f"{ITOP_BASE_URL}?operation=details&class={clase}&id={num}"


def build_itop_search_url(ticket: str) -> str:
    """Búsqueda global de iTop por referencia (sirve aunque el id interno no coincida con el folio)."""
    _prefix, num = _split_ticket_ref(ticket)
    if not num:
        return ""
    return f"{ITOP_BASE_URL}?operation=full_text&text={quote(str(ticket).strip().upper())}"


def get_app_dir() -> str:
    """Devuelve la carpeta base donde corre el ejecutable o el script."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def normalize_text(text: str) -> str:
    """Convierte a minúsculas, elimina acentos y normaliza espacios."""
    if not text:
        return ""
    text = str(text).lower()
    text = ''.join(
        c for c in unicodedata.normalize('NFD', text)
        if unicodedata.category(c) != 'Mn'
    )
    text = re.sub(r'\s+', ' ', text)
    return text.strip()


class RenovacionTicketParser:
    """
    Parser de alta precisión para correos de iTop FONACOT enfocado en:
    1. Detectar si un correo oficial de asignación es una Solicitud de Cancelación / Reversa de Renovación.
    2. Excluir trámites que mencionan la palabra 'renovación' pero son correcciones de nombre/domicilio,
       ajuste de centavos o cambio de bandera eventual/indeterminado.
    3. Extraer y desglosar TODAS las renovaciones contenidas en el ticket:
       - Tablas HTML de CREDERE (Cuenta | Operación Origen | Renovado | Sucursal | Fecha)
       - Tablas de sucursal (CLIENTE | CREDITO | SOLICITUD | FECHA | GAFETE)
       - Un cliente con múltiples créditos renovados separados por 'y', ',', '-', '+' (ej. I-359504, I-359592)
       - Múltiples clientes con sus respectivos créditos dentro del mismo correo
       - Separación explícita entre Crédito Padre (Origen / Anterior / Original) y Crédito Hijo (Renovado / Nuevo / Pagaré a cancelar)
    """

    # Números que pertenecen a plantillas o direcciones institucionales de FONACOT y NUNCA son clientes ni créditos
    IGNORED_NUMBERS = {
        '111', '452', '85000', '06760', '2024', '2025', '2026', '2027', '2028', '2029',
        '1400', '1003', '301', '104'
    }

    @staticmethod
    def clean_person_name(raw_name: str) -> str:
        if not raw_name:
            return ""
        cleaned = re.split(
            r'\b(Se\s+te\s+asign|Se\s+le\s+informa|para\s+su\s+atenci|Incidente|Resumen|Descripci)\b',
            raw_name,
            flags=re.I
        )[0]
        cleaned = re.sub(r'[\r\n\t]+', ' ', cleaned)
        cleaned = re.sub(r'\s+', ' ', cleaned).strip(" .,:;-_")
        return cleaned[:65]

    @staticmethod
    def is_official_assignment_email(subject: str, plain_body: str = "", sender: str = "") -> tuple:
        """
        Verifica si el correo es una notificación oficial de asignación de iTop.
        Retorna: (es_asignacion_oficial: bool, gafete: str, nombre_agente: str, ticket_num: str)
        """
        subject = (subject or "").strip()
        if re.match(r'^(re|rv|fw|fwd)\s*:', subject, re.I):
            return False, "", "", ""
        if '|' in subject:
            return False, "", "", ""

        m_subj = re.search(
            r'El\s+Ticket\s+(I-\d+)\s+se\s+le\s+ha\s+asignado\s+al\s+agente:\s*(\d{3,5})\s+([^\r\n]+)',
            subject,
            re.I
        )
        if m_subj:
            ticket_num = m_subj.group(1).upper().strip()
            gafete = m_subj.group(2).strip()
            agente = RenovacionTicketParser.clean_person_name(m_subj.group(3))
            return True, gafete, agente, ticket_num

        if plain_body:
            top_body = plain_body[:350]
            m_body = re.search(
                r'Apreciable\s+(\d{3,5})\s+([^\r\n]+?)\s+Se\s+te\s+asign[óo]\s+el\s+reporte\s+(I-\d+)',
                top_body,
                re.I | re.S
            )
            if m_body:
                gafete = m_body.group(1).strip()
                agente = RenovacionTicketParser.clean_person_name(m_body.group(2))
                ticket_num = m_body.group(3).upper().strip()
                return True, gafete, agente, ticket_num

        return False, "", "", ""

    @staticmethod
    def parse_itop_base_fields(html_body: str, plain_body: str) -> dict:
        """Extrae los campos de la tabla principal de iTop y conserva las tablas HTML internas de la Descripción."""
        result = {
            'incidente': '',
            'resumen': '',
            'solicitante': '',
            'localidad': '',
            'prioridad': 'Media',
            'estado': 'Asignado',
            'categoria': '',
            'descripcion_texto': '',
            'inner_tables': []  # Lista de tablas internas dentro de Descripción (como filas de celdas)
        }

        if html_body:
            try:
                soup = make_soup(html_body)
                for tag in soup(['script', 'style']):
                    tag.decompose()

                for tr in soup.find_all('tr'):
                    tds = tr.find_all('td', recursive=False)
                    if not tds:
                        tds = tr.find_all('td')

                    if len(tds) == 2:
                        key_txt = normalize_text(tds[0].get_text(" ", strip=True)).rstrip(':')
                        val_txt = tds[1].get_text(" ", strip=True)

                        if 'incidente #' in key_txt or key_txt == 'incidente':
                            m = re.search(r'(I-\d+)', val_txt, re.I)
                            if m:
                                result['incidente'] = m.group(1).upper()
                        elif key_txt == 'resumen':
                            result['resumen'] = val_txt.strip()
                        elif key_txt == 'nombre':
                            result['solicitante'] = RenovacionTicketParser.clean_person_name(val_txt)
                        elif key_txt == 'localidad':
                            result['localidad'] = val_txt.strip()
                        elif key_txt == 'prioridad':
                            result['prioridad'] = val_txt.strip()
                        elif key_txt == 'estado':
                            result['estado'] = val_txt.strip()
                        elif 'categoria' in key_txt:
                            result['categoria'] = val_txt.strip()

                # Extraer Descripción e inspeccionar tablas anidadas dentro de <pre> o <div class="HTML">
                pre_tags = soup.find_all('pre')
                if pre_tags:
                    desc_container = pre_tags[-1]
                    result['descripcion_texto'] = desc_container.get_text("\n", strip=True)
                    for tbl in desc_container.find_all('table'):
                        rows_data = []
                        for r in tbl.find_all('tr'):
                            cells = [c.get_text(" ", strip=True) for c in r.find_all(['td', 'th'])]
                            cells = [c for c in cells if c != '']
                            if cells:
                                rows_data.append(cells)
                        if rows_data:
                            result['inner_tables'].append(rows_data)
                else:
                    full_txt = soup.get_text("\n", strip=True)
                    if 'Descripción' in full_txt:
                        result['descripcion_texto'] = full_txt.split('Descripción', 1)[-1].strip()
                    else:
                        result['descripcion_texto'] = full_txt
            except Exception:
                pass

        if not result['descripcion_texto'] and plain_body:
            result['descripcion_texto'] = plain_body

        if not result['resumen'] and plain_body:
            m_res = re.search(r'Resumen:\s*([^\r\n]+)', plain_body, re.I)
            if m_res:
                result['resumen'] = m_res.group(1).strip()

        return result

    @classmethod
    def classify_renewal_ticket(cls, resumen: str, descripcion: str, inner_tables: list) -> tuple:
        """
        Determina si el ticket es una CANCELACIÓN / REVERSA DE RENOVACIÓN.
        Retorna: (es_cancelacion_renovacion: bool, subtipo_renovacion: str, motivo_detectado: str)
        """
        norm_res = normalize_text(resumen)
        norm_desc = normalize_text(descripcion)
        combined = f"{norm_res} {norm_desc}"

        # 1. ¿Es tabla de CREDERE ("Análisis | cancelación renovaciones" o tabla con columnas Cuenta / Operación / Renovado)?
        if 'cancelacion renovaciones' in combined or 'seguimiento a la cancelacion en credere' in combined:
            return True, "Concentrado CREDERE (Tabla)", "Cancelación masiva en CREDERE"

        for tbl in inner_tables:
            flat_hdr = normalize_text(" ".join(" ".join(r) for r in tbl[:2]))
            if ('cuenta' in flat_hdr and 'renovado' in flat_hdr) or ('operacion' in flat_hdr and 'renovado' in flat_hdr):
                return True, "Concentrado CREDERE (Tabla)", "Cancelación en CREDERE"

        # 2. Exclusiones estrictas de tickets que mencionan "renovación" pero SON OTRO TRÁMITE
        #    (solo se excluyen si NO piden explícitamente cancelar/reversar la renovación)
        explicit_cancel_renewal = bool(re.search(
            r'(canc\w*\.?\s+(de\s+|la\s+|el\s+|las\s+|los\s+)?(credito\s+renovado|renov|reno\b|ren\b)|reversa\s+(de\s+|a\s+la\s+)?renovaci|numero\s+de\s+credito\s*[\(/]\s*renovaci|no\.?\s*de\s+renovaci|numero\s+de\s+renovaci)',
            combined
        ))

        if not explicit_cancel_renewal:
            # Corrección de datos / domicilio / nombre / DICE - DEBE DECIR
            if ('dice' in combined and 'debe decir' in combined) or 'correccion de domicilio' in combined or 'reimpresion de los documentos' in combined or 'correccion e impresion' in combined:
                return False, "Corrección de Datos / Impresión", ""
            # Adeudo por centavos para poder renovar
            if 'adeudo por centavos' in combined or 'diferencia de centavos' in combined or 'ajuste de centavos' in combined:
                return False, "Ajuste de Centavos", ""
            # Cambio de bandera eventual/indeterminado para poder renovar
            if 'banderas para tipo de contrato' in combined or 'cambio en tipo de bandera' in combined:
                return False, "Cambio de Bandera", ""
            # Error en sistema que no permite renovar o registro trabado en originación
            if 'no me permite renovar' in combined or 'registro asociado al trabajador en etapa' in combined:
                return False, "Error al Intentar Renovar", ""

        # 3. Verificación positiva de Cancelación / Reversa de Renovación
        has_renewal_word = bool(re.search(
            r'\b(renov\w*|reno|ren)\b|credito\s+padre|credito\s+hijo|credito\s+original|credito\s+anterior|antes\s+de\s+renovar|operacion\s*\(\s*origen\s*\)',
            combined
        ))
        has_cancel_word = bool(re.search(
            r'\b(canc\w*|reversa\w*|sin\s+efecto|baja)\b',
            combined
        ))
        has_capacity_reason = bool(re.search(
            r'(no\s+libero\s+(disponible|capacidad|credito)|no\s+dio\s+disponible|no\s+da\s+disponible|bajo\s+(su\s+|la\s+|el\s+)?(capacidad|nivel)|topado\s+por\s+buro|capacidad\s+(de\s+cuota\s+|de\s+credito\s+)?agotada|plazo\s+menor\s+en\s+la\s+app)',
            combined
        ))

        if explicit_cancel_renewal or (has_renewal_word and (has_cancel_word or has_capacity_reason)):
            motivo = cls._extract_motivo(descripcion, combined)
            subtipo = "Cancelación de Renovación"
            if 'reversa' in combined:
                subtipo = "Reversa de Renovación"
            elif 'cancelacion de credito' in norm_res or 'cancelar credito' in norm_res:
                subtipo = "Cancelación de Crédito (Renovación)"
            return True, subtipo, motivo

        return False, "", ""

    @staticmethod
    def _extract_motivo(raw_desc: str, norm_combined: str) -> str:
        """Extrae un resumen limpio del motivo de la cancelación."""
        clean_d = re.split(r'NOTA IMPORTANTE:|Requisitos:|El calendario de emisi', raw_desc or '', flags=re.I)[0]
        m = re.search(
            r'Motivo(?:\s+de\s+(?:la\s+)?cancelaci[óo]n)?\s*[:\.-]?\s*([^\r\n]+)',
            clean_d,
            re.I
        )
        if m:
            mot = m.group(1).strip(" .:-_")
            # Evitar textos genéricos de plantilla si están vacíos
            if mot and not mot.lower().startswith('intermitencia / error de cuenta / cambio'):
                return mot[:85]

        if 'buro' in norm_combined or '10%' in norm_combined or '15%' in norm_combined or 'disponible' in norm_combined or 'capacidad' in norm_combined:
            return "Sin disponible / Topado por Buró / Bajó capacidad"
        if 'app' in norm_combined or 'plazo menor' in norm_combined:
            return "Trámite en App Móvil / Plazo menor"
        if 'peticion del cliente' in norm_combined or 'no lo quiso' in norm_combined or 'no le convino' in norm_combined:
            return "A petición del trabajador"
        return "Cancelación de renovación solicitada"

    @classmethod
    def _clean_user_text_for_extraction(cls, text: str) -> str:
        """
        Elimina el texto fijo de las plantillas de iTop ("Envía un correo electrónico a: 111@fonacot.gob.mx",
        "Requisitos: Para los días de generación de la emisión... hasta las 14:00 hrs...", etc.)
        PERO conserva la sección de "Información requerida:" o datos del cliente que a veces viene al final.
        """
        if not text:
            return ""
        cleaned = text
        # Remover bloque de Requisitos hasta antes de "Información requerida" (si existe)
        cleaned = re.sub(
            r'Requisitos:\s*Para los d[íi]as de generaci[óo]n.*?https?://\S+',
            ' ',
            cleaned,
            flags=re.I | re.S
        )
        cleaned = re.sub(
            r'NOTA IMPORTANTE:\s*Para los d[íi]as de generaci[óo]n.*?https?://\S+',
            ' ',
            cleaned,
            flags=re.I | re.S
        )
        # Remover frases literales de plantilla sin números reales
        cleaned = re.sub(
            r'[“"]?Cancelaci[óo]n de (?:renovaci[óo]n|cr[ée]dito)\s*\+\s*n[úu]mero de cliente\s*\+\s*n[úu]mero de cr[ée]dito[”"]?',
            ' ',
            cleaned,
            flags=re.I
        )
        cleaned = re.sub(
            r'Intermitencia\s*/\s*Error de cuenta\s*/\s*Cambio de monto\s*/\s*Cliente ya no lo quiere,?\s*etc\.?',
            ' ',
            cleaned,
            flags=re.I
        )
        return cleaned

    @classmethod
    def _is_valid_number_token(cls, num_str: str, min_len: int = 3, max_len: int = 10) -> bool:
        if not num_str or not num_str.isdigit():
            return False
        if not (min_len <= len(num_str) <= max_len):
            return False
        if num_str in cls.IGNORED_NUMBERS:
            return False
        return True

    @classmethod
    def _extract_crede_table_records(cls, inner_tables: list, plain_desc: str) -> list:
        """
        Extrae filas de tablas tipo CREDERE:
        Cuenta | Operación (Origen) | Renovado | Sucursal | Fecha
        o tablas horizontales de sucursal:
        CLIENTE | CREDITO | SOLICITUD | FECHA | GAFETE
        Retorna lista de dicts: [{'cliente': ..., 'credito_padre': ..., 'credito_renovado': ..., 'sucursal': ..., 'fecha_op': ...}]
        """
        records = []

        # 1. Buscar en tablas HTML anidadas
        for raw_tbl in inner_tables:
            if len(raw_tbl) < 2:
                continue
            # Celdas con letra pegada ("C5864479", "T-4530500", "Cred. 57137") -> solo el número
            tbl = [[strip_letter_prefixes(c).strip() for c in row] for row in raw_tbl]

            # Revisar si es tabla CREDERE (Cuenta | Operación (Origen) | Renovado)
            header_joined = normalize_text(" ".join(" ".join(r) for r in tbl[:2]))
            if ('cuenta' in header_joined and 'renovado' in header_joined) or ('operacion' in header_joined and 'renovado' in header_joined):
                for row in tbl:
                    nums_in_row = [c.strip() for c in row if re.fullmatch(r'\d{3,10}', c.strip())]
                    date_in_row = next((c.strip() for c in row if re.search(r'\d{2}[/-]\d{2}[/-]\d{4}', c.strip())), "")
                    if len(nums_in_row) >= 3:
                        records.append({
                            'cliente': nums_in_row[0],
                            'credito_padre': nums_in_row[1],
                            'credito_renovado': nums_in_row[2],
                            'sucursal': nums_in_row[3] if len(nums_in_row) >= 4 else '',
                            'fecha_op': date_in_row,
                            'origen_extraccion': 'Tabla CREDERE'
                        })
                if records:
                    return records

            # Revisar si es tabla horizontal de sucursal: CLIENTE | CREDITO | [SOLICITUD] | [FECHA] | [GAFETE]
            # Debe tener al menos 2 columnas distintas (cli_idx != cred_idx) para no confundirse con tablas plantilla de 1 columna
            for r_idx, row in enumerate(tbl[:-1]):
                if len(row) < 2:
                    continue
                row_norm = [normalize_text(c) for c in row]
                cli_idx = next((i for i, h in enumerate(row_norm) if h in ('cliente', 'no. cliente', 'numero de cliente', 'trabajador', 'cuenta', 'id cliente')), None)
                cred_idx = next((i for i, h in enumerate(row_norm) if h in ('credito', 'no. credito', 'numero de credito', 'renovacion', 'no. renovacion', 'renovado')), None)
                padre_idx = next((i for i, h in enumerate(row_norm) if 'origen' in h or 'anterior' in h or 'original' in h or 'padre' in h), None)

                if cli_idx is not None and cred_idx is not None and cli_idx != cred_idx:
                    for data_row in tbl[r_idx + 1:]:
                        if len(data_row) > max(cli_idx, cred_idx):
                            c_val = re.findall(r'\b\d{4,10}\b', data_row[cli_idx])
                            r_val = re.findall(r'\b\d{3,8}\b', data_row[cred_idx])
                            p_val = re.findall(r'\b\d{3,8}\b', data_row[padre_idx]) if (padre_idx is not None and len(data_row) > padre_idx) else []
                            if c_val and r_val:
                                for rv in r_val:
                                    records.append({
                                        'cliente': c_val[0],
                                        'credito_padre': p_val[0] if p_val else '',
                                        'credito_renovado': rv,
                                        'sucursal': '',
                                        'fecha_op': '',
                                        'origen_extraccion': 'Tabla Sucursal'
                                    })
                    if records:
                        return records

        # 2. Respaldo en texto plano si la tabla CREDERE venía como texto ("Cuenta Operación (Origen) Renovado Sucursal Fecha ...")
        if plain_desc and re.search(r'Cuenta\s+Operaci[óo]n\s*\(\s*Origen\s*\)\s+Renovado', plain_desc, re.I):
            after_hdr = strip_letter_prefixes(re.split(r'Renovado\s+Sucursal\s+Fecha', plain_desc, flags=re.I)[-1])
            row_matches = re.findall(
                r'\b(\d{5,10})\s+(\d{3,8})\s+(\d{3,8})\s+(\d{2,4})\s+(\d{2}[/-]\d{2}[/-]\d{4})\b',
                after_hdr
            )
            for cli, pad, ren, suc, f_op in row_matches:
                records.append({
                    'cliente': cli,
                    'credito_padre': pad,
                    'credito_renovado': ren,
                    'sucursal': suc,
                    'fecha_op': f_op,
                    'origen_extraccion': 'Tabla CREDERE (Texto)'
                })

        return records

    # Palabras que no son parte de un nombre: si aparecen, el número es de otra cosa (crédito, folio, teléfono…)
    _NO_NOMBRE = re.compile(
        r'^(CR[EÉ]DITOS?|RENOVAD[OA]S?|RENOVACI[OÓ]N(ES)?|CLIENT[EA]|TRABAJADOR[A]?|SOLICITUD|FOLIO|GAFETE|TEL[EÉ]FONO|'
        r'TEL|CEL|NSS|CURP|RFC|SUC|SUCURSAL|ANALISTA|INCIDENTE|TICKET|PAGAR[EÉ]|OPERACI[OÓ]N|CUENTA|FECHA|NO|NUM|'
        r'N[UÚ]MERO|ID|CANCELAR|CANCELACI[OÓ]N|EMPRESA|CENTRO|TRABAJO|ORIGINAL|ANTERIOR|NUEVO|HIJO|PADRE)$', re.I)

    @classmethod
    def _clientes_junto_al_nombre(cls, texto: str) -> list:
        """
        Números de 6 a 8 dígitos pegados a un nombre en mayúsculas o con mayúscula inicial (al menos 2 palabras),
        antes o después: «LUIS RENE MENDEZ ZUÑIGA 6650919», «6650919 Luis Rene Mendez». Sin etiqueta «cliente».
        """
        palabra = r"[A-ZÁÉÍÓÚÑÜ][A-Za-zÁÉÍÓÚÑÜáéíóúñü]+"
        encontrados = []
        anterior = ''
        for linea in normalizar_correo(texto).splitlines():
            # «PEREZ DIAZ CARLOS ERNESTO» y en el renglón de abajo solo «3326312»
            if re.fullmatch(r'\s*\d{6,8}\s*', linea) and cls._es_renglon_de_nombre(anterior):
                encontrados.append(linea.strip())
            if linea.strip():
                anterior = linea
            for m in re.finditer(rf'((?:{palabra}[^\S\n]+){{1,5}}{palabra})[^\S\n,:\-]*[,:\-]?[^\S\n]*(\d{{6,8}})\b(?![/\-\.]\d)', linea):
                nombre, num = m.group(1), m.group(2)
                if any(cls._NO_NOMBRE.match(p) for p in nombre.split()):
                    continue
                encontrados.append(num)
            for m in re.finditer(rf'(?<![\d/\-\.])\b(\d{{6,8}})[^\S\n]*[,:\-]?[^\S\n]*((?:{palabra}[^\S\n]+){{1,5}}{palabra})', linea):
                num, nombre = m.group(1), m.group(2)
                antes = linea[:m.start()].split()
                if any(cls._NO_NOMBRE.match(p) for p in nombre.split()[:2]) or (
                        antes and cls._NO_NOMBRE.match(antes[-1].strip(':.#'))):
                    continue
                encontrados.append(num)

        def es_fecha(n):  # 26102026 / 20261026
            return len(n) == 8 and (re.fullmatch(r'(0[1-9]|[12]\d|3[01])(0[1-9]|1[0-2])20\d\d', n)
                                    or re.fullmatch(r'20\d\d(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])', n))
        return [n for i, n in enumerate(encontrados) if n not in encontrados[:i] and not es_fecha(n)]

    _PALABRA_NOMBRE = r"(?:[A-ZÁÉÍÓÚÑÜ][A-Za-zÁÉÍÓÚÑÜáéíóúñü]+|de|del|la|las|los)"
    # Palabras que pueden acompañar a los créditos en su renglón («C/33363», «CREDITO RENOVADO», «53290 Y 61780»)
    _ETIQUETAS_CREDITO = {'c', 'cr', 'cred', 'credito', 'creditos', 'ren', 'renov', 'renovacion', 'renovaciones',
                          'renovado', 'renovada', 'renovados', 'renovadas', 'no', 'num', 'numero', 'y', 'e'}

    @classmethod
    def _es_renglon_de_nombre(cls, linea: str, con_numero: bool = False):
        """El renglón es solo un nombre (2 a 6 palabras, con o sin el número al final). Devuelve el match o None."""
        p = cls._PALABRA_NOMBRE
        m = re.fullmatch(rf'\s*(?P<nom>[A-ZÁÉÍÓÚÑÜ][A-Za-zÁÉÍÓÚÑÜáéíóúñü]+(?:[^\S\n]+{p}){{1,5}})'
                         rf'(?:[^\S\n]*[,:\-]?[^\S\n]*(?P<num>\d{{6,8}}))?\s*', linea or '')
        if not m or (m.group('num') and not con_numero):
            return None
        palabras = m.group('nom').split()
        if any(cls._NO_NOMBRE.match(w) for w in palabras) or sum(1 for w in palabras if w[0].isupper()) < 2:
            return None
        return m

    @classmethod
    def _bloques_por_nombre(cls, texto: str) -> list:
        """
        Correos sin etiquetas, un trabajador por bloque:
            PEREZ DIAZ CARLOS ERNESTO            LUIS RENE MENDEZ ZUÑIGA 6650919
            3326312                              CREDITO RENOVADO
            C/33363                              . 53290 Y 61780
        [(cliente, [créditos]), …] solo si cada bloque trae su cliente y al menos un crédito; si no, [].
        """
        bloques, actual = [], None
        for linea in normalizar_correo(texto).splitlines():
            l = linea.strip()
            if not l:
                continue
            m = cls._es_renglon_de_nombre(l, con_numero=True)
            if m:
                actual = {'cliente': m.group('num') or '', 'creditos': []}
                bloques.append(actual)
                continue
            if actual is None:
                continue
            if re.fullmatch(r'\d{6,8}', l) and not actual['cliente'] and not actual['creditos']:
                actual['cliente'] = l
                continue
            tokens = re.findall(r'[^\W\d_]+|\d+', l)
            nums = [t for t in tokens if t.isdigit()]
            etiquetas = all(normalize_text(t) in cls._ETIQUETAS_CREDITO for t in tokens if not t.isdigit())
            if etiquetas and nums and actual['cliente'] and all(3 <= len(n) <= 7 for n in nums):
                actual['creditos'] += [n for n in nums if n not in actual['creditos'] and n != actual['cliente']]
                continue
            if etiquetas and not nums:
                continue  # «CREDITO RENOVADO» sin número: los créditos vienen abajo
            actual = None  # sucursal, analista, gafete…: termina el bloque
        bloques = [b for b in bloques if b['cliente'] or b['creditos']]
        if not bloques or any(not (b['cliente'] and b['creditos']) for b in bloques):
            return []
        return [(b['cliente'], b['creditos']) for b in bloques]

    @classmethod
    def extract_disaggregated_renewals(cls, resumen: str, descripcion: str, inner_tables: list) -> list:
        """
        Extrae y desglosa TODAS las renovaciones individuales de un ticket.
        Siempre retorna una lista de diccionarios (al menos 1 elemento):
        [
            {
                'cliente': str,
                'credito_renovado': str,   # Crédito Hijo / Nuevo / Pagaré a cancelar
                'credito_padre': str,      # Crédito Origen / Anterior / Original (si se menciona)
                'sucursal': str,
                'fecha_op': str,
                'origen_extraccion': str
            },
            ...
        ]
        """
        # PASO 1: ¿Contiene una tabla estructurada CREDERE o de Sucursal?
        table_records = cls._extract_crede_table_records(inner_tables, descripcion)
        if table_records:
            return table_records

        resumen, descripcion = normalizar_correo(resumen), normalizar_correo(descripcion)
        clean_desc = cls._clean_user_text_for_extraction(descripcion)
        # Traducir abreviaturas ("C5864479", "Cte. 3386153", "R 91258") a etiquetas completas antes de aplicar los patrones
        resumen_exp = expand_abbreviated_labels(resumen or '')
        full_context = f"{resumen_exp}\n{expand_abbreviated_labels(clean_desc or '')}"

        # Sin etiquetas: «NOMBRE DEL TRABAJADOR / número de cliente / créditos», uno o varios trabajadores
        bloques = cls._bloques_por_nombre(full_context)
        if bloques:
            total = sum(len(creds) for _, creds in bloques)
            return [{
                'cliente': cli,
                'credito_renovado': cred,
                'credito_padre': '',
                'sucursal': '',
                'fecha_op': '',
                'origen_extraccion': ('Múltiples Clientes y Créditos' if len(bloques) > 1 else
                                      'Múltiples Créditos en Ticket' if total > 1 else 'Individual')
            } for cli, creds in bloques for cred in creds]

        # PASO 2: ¿Vienen múltiples bloques de "Cliente ... Crédito ..." para DIFERENTES clientes en el mismo texto?
        # (Ej. cuando en un mismo correo enlistan Cliente A con Crédito X y abajo Cliente B con Crédito Y)
        clientes_found = []
        creditos_renovados_found = []
        creditos_padre_found = []

        # 2.A: Extraer explícitamente Créditos Padre (Origen / Anterior / Original / Antes de renovar)
        padre_patterns = [
            r'(?:cr[ée]dito\s+(?:anterior|original|de\s+origen|origen|padre)|antes\s+de\s+renovar|operaci[óo]n\s*\(\s*origen\s*\))\s*[:\.\-\(]*\s*([\d,\s yY\-]+)',
            r'cancelar\s+la\s+renovaci[óo]n\s+del\s+cr[ée]dito\s+(\d{3,7})\s+con\s+no\.?\s+de\s+pagar[ée]',
            r'\b(\d{4,7})\s+\d{2,3}\s+se\s+renov[óo]\s+al\s+\d{4,7}\b'
        ]
        for pat in padre_patterns:
            for m in re.finditer(pat, full_context, re.I):
                raw_list = m.group(1)
                for n in re.findall(r'\b\d{3,7}\b', raw_list):
                    if cls._is_valid_number_token(n, 3, 7) and n not in creditos_padre_found:
                        creditos_padre_found.append(n)

        # Caso especial: "No. Credito Renovado : 15744 No. Renovación Nueva: 45529" (donde 'Credito Renovado' es el padre y 'Renovación Nueva' es el hijo)
        m_old_new = re.search(
            r'No\.?\s*Cr[ée]dito\s+Renovado\s*:\s*(\d{3,7})\s+No\.?\s*Renovaci[óo]n\s+Nueva\s*:\s*(\d{3,7})',
            full_context,
            re.I
        )
        if m_old_new:
            p_num, r_num = m_old_new.group(1), m_old_new.group(2)
            if p_num not in creditos_padre_found:
                creditos_padre_found.append(p_num)
            if r_num not in creditos_renovados_found:
                creditos_renovados_found.append(r_num)

        # 2.B: Extraer Números de Cliente (Trabajador / Trabajadora / Credere / Fonacot / T-)
        cliente_patterns = [
            # Etiquetas explícitas: "Número de Cliente: 5864479", "No. trabajador: 6046732", "Credere: 6400729", "T-3854458"
            r'(?:n[úu]mero|num\.?|no\.?|#)\s*(?:de\s+)?(?:client[ea]|trabajador(?:a)?|fonacot|credere|cuenta)\s*[:\.\-#]*\s*(\d{5,10})\b',
            r'\b(?:client[ea]|trabajador(?:a)?|credere)\s*(?:no\.?|num\.?|#|:|\.)?\s*'
            r'(?P<nom>(?:(?-i:[A-ZÁÉÍÓÚÑ]{2,20})[^\S\n]+){0,4})[#:\.]?[^\S\n]*(?P<num>\d{5,10})\b',
            r'\bT\s*[\-\.]\s*(\d{5,10})\b',
            r'/\s*T\s+(\d{5,10})\b',
            r'[\+\-]\s*(\d{6,10})\s*[\+\-]\s*\d{4,7}\b',
            r'\b(?:cancelaci[óo]n\s+(?:de\s+)?renovaci(?:[óo]n|ones)|cancelaci[óo]n\s+(?:de\s+)?cr[ée]dito)\s*[\+\-:]?\s*(\d{6,10})\b',
            r'\bid\s+(\d{6,8})\s+cr[ée]dito\s+\d{4,7}\b'
        ]
        for pat in cliente_patterns:
            for m in re.finditer(pat, full_context, re.I):
                if m.groupdict().get('nom') and any(cls._NO_NOMBRE.match(w.strip('.:#')) for w in m.group('nom').split()):
                    continue
                c_num = (m.groupdict().get('num') or m.group(1)).strip()
                # Excluir NSS (11 dígitos) y números de solicitud (empiezan con 507 o 508 de 9 dígitos)
                if len(c_num) == 9 and c_num.startswith(('507', '508')):
                    continue
                if cls._is_valid_number_token(c_num, 5, 10) and c_num not in clientes_found and c_num not in creditos_padre_found:
                    clientes_found.append(c_num)

        # Sin etiqueta: el número del trabajador junto a su nombre ("LUIS RENE MENDEZ ZUÑIGA 6650919")
        for c_num in cls._clientes_junto_al_nombre(full_context):
            if c_num not in clientes_found and c_num not in creditos_padre_found:
                clientes_found.append(c_num)

        # Buscar en tabla estilo texto: "CLIENTE CREDITO ... 4951216 45730"
        for m_inline_tbl in re.finditer(
            r'CLIENTE\s+CR[ÉE]DITO(?:\s+SOLICITUD)?(?:\s+FECHA)?(?:\s+GAFETE)?\s+(\d{5,10})\s+(\d{3,7})\b',
            full_context,
            re.I
        ):
            c_num, r_num = m_inline_tbl.group(1), m_inline_tbl.group(2)
            if c_num not in clientes_found:
                clientes_found.append(c_num)
            if r_num not in creditos_renovados_found and r_num not in creditos_padre_found:
                creditos_renovados_found.append(r_num)

        # 2.C: Extraer TODOS los Números de Crédito Renovado (Hijo / A cancelar / Pagaré)
        # Soporta listas múltiples como:
        # "CREDITO 56572 Y 56574", "Número de Crédito (renovación): 61116,61117,61118",
        # "Número de renovación 48736 - 48735", "C- 74544 y C- 74545", "+ 2431191 + 20757 + 20758"
        multi_credit_label_patterns = [
            r'(?:n[úu]mero|num\.?|no\.?)\s+de\s+cr[ée]ditos?\s*(?:[\(/]\s*renovaci(?:[óo]n|ones)\s*\)?)?\s*[:\.\-]*\s*([C\-\d,\s yY\+]{3,45})',
            r'(?:n[úu]mero|num\.?|no\.?)\s+de\s+renovaci(?:[óo]n|ones)(?:\s+nueva)?\s*[:\.\-]*\s*([C\-\d,\s yY\+]{3,45})',
            r'(?:n[úu]mero|num\.?|no\.?)\s+de(?:l)?\s+(?:cr[ée]dito|pagar[ée])\s+a\s+cancelar\s*[:\.\-]*\s*([C\-\d,\s yY\+]{3,45})',
            r'(?:ya\s+renovado|cr[ée]dito\s+nuevo|cr[ée]dito\s+renovado|cr[ée]dito\s+hijo|pagar[ée])\s*[:\.\-\(]*\s*([C\-\d,\s yY]{3,35})',
            r'\b(?:cr[ée]ditos?|renovaci(?:[óo]n|ones)|ren)(?:\s*(?:n[úu]mero|num\.?|no\.?|n°)\s*|[^\S\n]*)(?:[:#\.]\s*)?([C\-\d]+(?:\s*(?:,|y|Y|\-|\+)\s*[C\-\d]+)+)',
            r'\b(?:cr[ée]ditos?|renovaci(?:[óo]n|ones)|ren)(?:\s*(?:n[úu]mero|num\.?|no\.?|n°)\s*|[^\S\n]*)(?:[:#\.]\s*)?[C\-]?[^\S\n]*(\d{3,7})\b',
            r'\bse\s+renov[óo]\s+al\s+(\d{4,7})\b',
        ]

        for pat in multi_credit_label_patterns:
            for m in re.finditer(pat, full_context, re.I):
                chunk = m.group(1)
                # Cortar el chunk si empieza otra etiqueta ("Motivo", "Gafete", "Cliente", "Trabajador", "Fecha", "Sucursal")
                chunk = re.split(
                    r'\b(motivo|gafete|client[ea]|trabajador|fecha|sucursal|estado|solicitud|nombre|antes|original|anterior)\b',
                    chunk,
                    flags=re.I
                )[0]
                for token in re.findall(r'\b\d{3,7}\b', chunk):
                    if not cls._is_valid_number_token(token, 3, 7):
                        continue
                    if token in clientes_found or token in creditos_padre_found:
                        continue
                    if token not in creditos_renovados_found:
                        creditos_renovados_found.append(token)

        # Último recurso, solo si no se encontró ningún crédito: la etiqueta cierra el renglón y abajo solo vienen los
        # créditos («renovaciones de crédito\n45375, 45377,45380,»). De 3 a 6 dígitos (un número de cliente solo no
        # cuenta) y no si el renglón es de otra cosa («Analista de Crédito\n2609», «capacidad de crédito\n15000»).
        if not creditos_renovados_found:
            for m in re.finditer(
                    r'^([^\n]*\b(?:cr[ée]ditos?|renovaci(?:[óo]n|ones)))[^\S\n]*[:#\.]?[^\S\n]*\n[^\S\n]*'
                    r'((?:\d{3,6}[^\S\n]*(?:,|y|Y|\-)[^\S\n]*)*\d{3,6})[^\S\n]*[,\.]?[^\S\n]*(?=\n|$)',
                    full_context, re.I | re.M):
                if re.search(r'\b(analista|ejecutiv[oa]s?|capacidad|folio|depto|departamento|[áa]rea|gerencia|gerente|'
                             r'jef[ea]|coordinaci[óo]n|direcci[óo]n|subdirecci[óo]n|sistema|monto|importe|saldo|'
                             r'l[íi]mite|sucursal|oficina)\b', m.group(1), re.I):
                    continue
                for token in re.findall(r'\b\d{3,6}\b', m.group(2)):
                    if (cls._is_valid_number_token(token, 3, 6) and token not in clientes_found
                            and token not in creditos_padre_found and token not in creditos_renovados_found):
                        creditos_renovados_found.append(token)

        # Patrón Cliente + Crédito con signos: "+ 7239770+50074", "-8220974-60117", "+ 2431191 + 20757 + 20758",
        # o con asterisco: "Cancelación de crédito 1184282*10430", "1184282 * 10430 * 10431"
        for pat_par in (
            r'[\+\-]\s*(\d{5,10})\s*[\+\-]\s*(\d{3,7})\b(?:\s*(?:[\+\-,]|y)\s*(\d{3,7})\b)?',
            r'\b(\d{5,10})\s*\*\s*(\d{3,7})\b(?:\s*(?:[\*,]|y)\s*(\d{3,7})\b)?',
        ):
            for m_plus in re.finditer(pat_par, full_context, re.I):
                c_cand = m_plus.group(1)
                r1 = m_plus.group(2)
                r2 = m_plus.group(3)
                if cls._is_valid_number_token(c_cand, 5, 10) and c_cand not in clientes_found:
                    clientes_found.append(c_cand)
                for rx in (r1, r2):
                    if rx and cls._is_valid_number_token(rx, 3, 7) and rx not in clientes_found and rx not in creditos_padre_found:
                        if rx not in creditos_renovados_found:
                            creditos_renovados_found.append(rx)

        # Patrón en Resumen tipo: "CANCELACION DE RENOVACION 5042466 57521" o "82377 y 82379 7113273"
        if not clientes_found or not creditos_renovados_found:
            res_nums = [
                n for n in re.findall(r'\b\d{4,10}\b', strip_letter_prefixes(re.sub(
                    r'\b[IR]-\d+|\b(?:agente|gafete|ticket|incidente)\s*[:#]?\s*\d+', ' ', resumen or '', flags=re.I)))
                if cls._is_valid_number_token(n, 4, 10) and not (len(n) == 9 and n.startswith(('507', '508')))
            ]
            for n in res_nums:
                if len(n) >= 6 and n not in creditos_renovados_found and n not in clientes_found:
                    clientes_found.append(n)
                elif len(n) <= 5 and n not in clientes_found and n not in creditos_padre_found and n not in creditos_renovados_found:
                    creditos_renovados_found.append(n)

        # Filtrar de creditos_renovados cualquier número que sea en realidad el Gafete del analista de sucursal
        gafete_sucursal_matches = re.findall(
            r'(?:gafete(?:\s+del\s+analista[^0-9]{0,25})?|usuario)\s*[:#]?\s*(\d{3,5})\b',
            full_context,
            re.I
        )
        gafetes_sucursal = set(gafete_sucursal_matches)
        if len(creditos_renovados_found) > 1 and gafetes_sucursal:
            creditos_renovados_found = [c for c in creditos_renovados_found if c not in gafetes_sucursal]

        # PASO 3: Emparejar Clientes, Créditos Renovados (Hijos) y Créditos Origen (Padres)
        disaggregated = []

        # Caso A: 1 Cliente (o sin cliente) y N Créditos Renovados (ej. I-359504 con 2 créditos, I-359592 con 3 créditos, I-361952 con 4 créditos)
        if len(clientes_found) <= 1 and len(creditos_renovados_found) >= 1:
            cli_str = clientes_found[0] if clientes_found else "SIN CLIENTE"
            for idx, r_cred in enumerate(creditos_renovados_found):
                # Si hay igual cantidad de créditos padre que renovados (ej. I-361952 tiene 4 originales y 4 a cancelar), emparejar 1 a 1
                if len(creditos_padre_found) == len(creditos_renovados_found):
                    pad_str = creditos_padre_found[idx]
                elif len(creditos_padre_found) == 1:
                    pad_str = creditos_padre_found[0]
                else:
                    pad_str = ", ".join(creditos_padre_found) if creditos_padre_found else ""

                disaggregated.append({
                    'cliente': cli_str,
                    'credito_renovado': r_cred,
                    'credito_padre': pad_str,
                    'sucursal': '',
                    'fecha_op': '',
                    'origen_extraccion': 'Múltiples Créditos en Ticket' if len(creditos_renovados_found) > 1 else 'Individual'
                })
            return disaggregated

        # Caso B: Múltiples Clientes diferentes en el mismo ticket (ej. 2 números de cliente con 2 o más cancelaciones)
        if len(clientes_found) > 1:
            # Todos los clientes del incidente van en cada renovación: si alguno no tiene renovaciones se avisa
            todos_clientes = list(dict.fromkeys(clientes_found))

            def posicion(num):
                m_pos = re.search(rf'(?<!\d){re.escape(num)}(?!\d)', full_context)
                return m_pos.start() if m_pos else len(full_context)
            if len(creditos_renovados_found) == len(clientes_found):
                clientes_found = sorted(clientes_found, key=posicion)
                creditos_renovados_found = sorted(creditos_renovados_found, key=posicion)
                for idx, cli_str in enumerate(clientes_found):
                    pad_str = creditos_padre_found[idx] if idx < len(creditos_padre_found) else ""
                    disaggregated.append({
                        'cliente': cli_str,
                        'credito_renovado': creditos_renovados_found[idx],
                        'credito_padre': pad_str,
                        'sucursal': '',
                        'fecha_op': '',
                        'origen_extraccion': 'Múltiples Clientes en Ticket',
                        'clientes_ticket': todos_clientes
                    })
                return disaggregated
            elif len(creditos_renovados_found) > len(clientes_found):
                # Asociar cada crédito al cliente más cercano en la misma línea o en el bloque inmediato anterior
                cli_positions = []
                for cli_str in clientes_found:
                    m_pos = re.search(rf'\b{re.escape(cli_str)}\b', full_context)
                    cli_positions.append((m_pos.start() if m_pos else 0, cli_str))
                cli_positions.sort(key=lambda x: x[0])

                for idx, r_cred in enumerate(creditos_renovados_found):
                    m_cpos = re.search(rf'\b{re.escape(r_cred)}\b', full_context)
                    cpos = m_cpos.start() if m_cpos else 0
                    # Buscar el cliente inmediato anterior (o el primero si el crédito apareció antes en el resumen)
                    preceding = [c_str for p_pos, c_str in cli_positions if p_pos <= cpos]
                    assigned_cli = preceding[-1] if preceding else cli_positions[0][1]
                    pad_str = creditos_padre_found[idx] if idx < len(creditos_padre_found) else ""
                    disaggregated.append({
                        'cliente': assigned_cli,
                        'credito_renovado': r_cred,
                        'credito_padre': pad_str,
                        'sucursal': '',
                        'fecha_op': '',
                        'origen_extraccion': 'Múltiples Clientes y Créditos',
                        'clientes_ticket': todos_clientes
                    })
                return disaggregated
            else:
                # Si hay múltiples números de cliente pero SOLO 1 crédito único, y ese mismo crédito se repite
                # junto a ambos clientes (ej. error de sucursal donde pusieron un cliente en Asunto y otro en Número de Cliente):
                if len(creditos_renovados_found) == 1:
                    single_cred = creditos_renovados_found[0]
                    occurrences = len(re.findall(rf'\b{re.escape(single_cred)}\b', full_context))
                    if occurrences >= 2:
                        pad_str = creditos_padre_found[0] if creditos_padre_found else ""
                        # Priorizar el último ("Número de Cliente:") pero mostrar ambos para que el validador detecte la discrepancia
                        cli_combined = f"{clientes_found[-1]} (⚠️ En texto también: {clientes_found[0]})"
                        return [{
                            'cliente': cli_combined,
                            'credito_renovado': single_cred,
                            'credito_padre': pad_str,
                            'sucursal': '',
                            'fecha_op': '',
                            'origen_extraccion': 'Discrepancia de Cliente en Ticket',
                            'discrepancia_cliente': True,
                            'clientes_ticket': todos_clientes
                        }]

                # Más clientes que créditos detectados: crear 1 fila por cada cliente para no omitir ninguno
                for idx, cli_str in enumerate(clientes_found):
                    r_cred = creditos_renovados_found[idx] if idx < len(creditos_renovados_found) else "SIN CRÉDITO"
                    pad_str = creditos_padre_found[idx] if idx < len(creditos_padre_found) else ""
                    disaggregated.append({
                        'cliente': cli_str,
                        'credito_renovado': r_cred,
                        'credito_padre': pad_str,
                        'sucursal': '',
                        'fecha_op': '',
                        'origen_extraccion': 'Múltiples Clientes en Ticket',
                        'clientes_ticket': todos_clientes
                    })
                return disaggregated

        # Caso C: No se encontró Crédito Renovado (datos incompletos en el correo, ej. solo pusieron el Cliente o solo texto)
        cli_fallback = clientes_found[0] if clientes_found else "SIN CLIENTE"
        pad_fallback = creditos_padre_found[0] if creditos_padre_found else ""
        return [{
            'cliente': cli_fallback,
            'credito_renovado': "SIN CRÉDITO",
            'credito_padre': pad_fallback,
            'sucursal': '',
            'fecha_op': '',
            'origen_extraccion': 'Datos Incompletos'
        }]

    @classmethod
    def parse_and_disaggregate_email(
        cls,
        subject: str,
        html_body: str,
        plain_body: str,
        received_dt: datetime,
        ticket_num: str,
        gafete: str,
        nombre_asignado: str,
        carpeta_origen: str = "Bandeja de entrada"
    ) -> list:
        """
        Analiza un correo de asignación de iTop. Si es de Cancelación/Reversa de Renovación,
        devuelve la lista de filas desglosadas (1 fila por cada renovación individual).
        Si NO es de Cancelación de Renovación, devuelve una lista vacía [].
        """
        base = cls.parse_itop_base_fields(html_body, plain_body)
        resumen = base['resumen'] or subject
        descripcion = base['descripcion_texto'] or plain_body or ""
        inner_tables = base['inner_tables']

        is_ren_cancel, subtipo, motivo = cls.classify_renewal_ticket(resumen, descripcion, inner_tables)
        if not is_ren_cancel:
            return []

        items = cls.extract_disaggregated_renewals(resumen, descripcion, inner_tables)
        total_in_ticket = len(items)

        rows = []
        for idx, item in enumerate(items, 1):
            cli = item.get('cliente', 'SIN CLIENTE')
            cred_ren = item.get('credito_renovado', 'SIN CRÉDITO')
            cred_pad = item.get('credito_padre', '')
            discrepancia_cli = bool(item.get('discrepancia_cliente', False))

            datos_incompletos = (cli == "SIN CLIENTE" or cred_ren == "SIN CRÉDITO")
            es_multiple = total_in_ticket > 1

            if datos_incompletos:
                alerta_reno = "🚨 FALTA DATO (Revisar Cliente/Crédito en correo)"
                nivel_alerta = "incompleto"
            elif discrepancia_cli:
                alerta_reno = "⚠️ DISCREPANCIA DE CLIENTE: El ticket trae 2 números de cliente distintos para este crédito"
                nivel_alerta = "multiple"
            elif es_multiple:
                alerta_reno = f"⚠️ MÚLTIPLE: Renovación {idx} de {total_in_ticket} en este ticket"
                nivel_alerta = "multiple"
            else:
                alerta_reno = "✅ Individual (1 renovación en ticket)"
                nivel_alerta = "normal"

            rows.append({
                # Incluye el gafete: en modo por compañero el mismo ticket puede venir de 2 agentes filtrados
                'Row_ID': f"{ticket_num}_{gafete}_{idx}_{first_digits(cli) or 'SC'}_{first_digits(cred_ren) or 'SR'}",
                'Fecha_DT': received_dt,
                'Fecha': received_dt.strftime('%d/%m/%Y %H:%M'),
                'Numero_Ticket': ticket_num,
                'Indice_En_Ticket': idx,
                'Total_En_Ticket': total_in_ticket,
                'Es_Multiple': es_multiple,
                'Datos_Incompletos': datos_incompletos,
                'Nivel_Alerta': nivel_alerta,
                'Alerta_Renovacion': alerta_reno,
                'Carpeta_Origen': carpeta_origen,
                'Gafete': gafete,
                'Nombre_Asignado': nombre_asignado,
                'Numero_Cliente': cli,
                'Credito_Renovado_Hijo': cred_ren,
                'Credito_Origen_Padre': cred_pad or "-",
                'Clientes_Ticket': list(item.get('clientes_ticket') or []),
                'Subtipo_Renovacion': subtipo,
                'Motivo': motivo,
                'Sucursal': item.get('sucursal', '') or base['localidad'][:35],
                'Fecha_Operacion': item.get('fecha_op', ''),
                'Solicitante': base['solicitante'],
                'Prioridad': base['prioridad'],
                'Estado_iTop': base['estado'],
                'Resumen_iTop': resumen,
                'Descripcion_Completa': descripcion[:3500],
                'Asunto_Correo': subject,
                'Validado_Check': False,
                'Requiere_Validacion_Agente': False,
                'Reasignado_A_Otro': False,
                'Alerta_Asignacion': "✅ Asignación directa",
                'Historial_Asignaciones': ""
            })

        return rows


# ======================= DATOS LEÍDOS DE IMÁGENES (capturas de pantalla) =======================

def completar_con_imagenes(rows: list, lecturas: list) -> list:
    """
    Completa el cliente o el crédito que el texto del incidente no trae con lo leído en sus imágenes (OCR).
    `rows`: renglones de un mismo correo. `lecturas`: [{'cliente', 'credito', 'imagen', 'dudosa'}], una por imagen.
    Si el incidente no trae nada y hay varias imágenes con distintos datos, queda una renovación por imagen.
    Lo completado se marca para revisarse: una imagen se puede leer con un dígito equivocado.
    """
    pares = []
    for lec in lecturas:
        par = (first_digits(lec.get('cliente') or ''), first_digits(lec.get('credito') or ''))
        if any(par) and par not in [p[:2] for p in pares]:
            pares.append((par[0], par[1], lec.get('imagen') or 'imagen', bool(lec.get('dudosa'))))
    if not pares or not any(r.get('Datos_Incompletos') for r in rows):
        return rows

    def llenar(r, par):
        pc, pr, imagen, dudosa = par
        leido = []
        if r['Numero_Cliente'] == "SIN CLIENTE" and pc:
            r['Numero_Cliente'] = pc
            leido.append(f"Cliente {pc}")
        if r['Credito_Renovado_Hijo'] == "SIN CRÉDITO" and pr:
            r['Credito_Renovado_Hijo'] = pr
            leido.append(f"Crédito {pr}")
        r['Datos_Incompletos'] = r['Numero_Cliente'] == "SIN CLIENTE" or r['Credito_Renovado_Hijo'] == "SIN CRÉDITO"
        r['Dato_Imagen'] = f"{' y '.join(leido)} leído de la imagen «{imagen}»" + (" (lectura dudosa)" if dudosa else "")
        if r['Datos_Incompletos']:
            r['Alerta_Renovacion'] = f"🚨 FALTA DATO · {r['Dato_Imagen']}: revísalo"
            r['Nivel_Alerta'] = "incompleto"
        else:
            r['Alerta_Renovacion'] = f"🖼️ {r['Dato_Imagen']}: revísalo antes de enviarlo"
            r['Nivel_Alerta'] = "multiple"

    usados, salida = set(), []
    for r in rows:
        if r.get('Datos_Incompletos'):
            cli = '' if r['Numero_Cliente'] == "SIN CLIENTE" else first_digits(r['Numero_Cliente'])
            cred = '' if r['Credito_Renovado_Hijo'] == "SIN CRÉDITO" else first_digits(r['Credito_Renovado_Hijo'])
            compatibles = [i for i, (pc, pr, _, _) in enumerate(pares) if i not in usados
                           and (not cli or not pc or pc == cli) and (not cred or not pr or pr == cred)]
            if compatibles:
                usados.add(compatibles[0])
                llenar(r, pares[compatibles[0]])
        salida.append(r)

    base = salida[0] if len(rows) == 1 and salida[0].get('Dato_Imagen') else None
    extra = [p for i, p in enumerate(pares) if i not in usados and p[0] and p[1]]
    if base and extra:
        for n, par in enumerate(extra, 2):
            copia = dict(base)
            copia.update({'Numero_Cliente': "SIN CLIENTE", 'Credito_Renovado_Hijo': "SIN CRÉDITO",
                          'Row_ID': f"{base['Row_ID']}_img{n}"})
            llenar(copia, par)
            salida.append(copia)
        for i, r in enumerate(salida, 1):
            r['Indice_En_Ticket'], r['Total_En_Ticket'], r['Es_Multiple'] = i, len(salida), True
    return salida


# ======================= CORTE DIARIO, CONCENTRADO "TICKET ÚNICO" Y REZAGADAS =======================
# El corte de cada día es la hora en que realmente salió su correo concentrado ("Ticket Único <fecha>"):
# lo que llegó después, o lo que no quedó en él, debe sumarse al lote siguiente. Mientras el ticket único
# de hoy no sale, todo lo que llega entra hoy. La hora habitual (17:50) solo se usa en días sin ticket único.

HORA_CORTE_DEFAULT = "17:50"
ASUNTO_CONCENTRADO_DEFAULT = "ticket unico"
# Lo que llegó poco antes de que saliera el ticket único y no viene en él se toma como "llegó mientras se armaba"
MARGEN_ENVIO = timedelta(minutes=30)

LOTE_HOY = 'HOY'                         # Llegó hoy antes de que saliera el ticket único de hoy
LOTE_REZAGADA = 'REZAGADA'               # Llegó antes de hoy y no se envió (incidente abierto) -> entra hoy
LOTE_ENVIADA = 'ENVIADA'                 # Llegó antes de hoy y ya se envió (aparece en un ticket único)
LOTE_EN_CONCENTRADO = 'EN_CONCENTRADO'   # Es de hoy y ya aparece en un concentrado enviado
LOTE_SIGUIENTE = 'SIGUIENTE_LOTE'        # Llegó hoy después del ticket único de hoy -> va al lote siguiente
LOTE_CERRADA_SIN_ENVIO = 'CERRADA_SIN_ENVIO'  # Debía ir en un ticket único que ya pasó, no viene y el incidente ya se cerró
LOTE_SIN_COMPROBAR = 'SIN_COMPROBAR'     # Día anterior sin ticket único encontrado para comprobar si se envió

_MESES_ES = {
    'enero': 1, 'febrero': 2, 'marzo': 3, 'abril': 4, 'mayo': 5, 'junio': 6, 'julio': 7,
    'agosto': 8, 'septiembre': 9, 'setiembre': 9, 'octubre': 10, 'noviembre': 11, 'diciembre': 12
}


def parse_hora_corte(text: str) -> time:
    """'17:50', '5:50 pm', '1750', '17.50' -> time(17, 50). Si no se entiende, usa 17:50."""
    t = normalize_text(text or '').replace('.', ':').replace(' ', '')
    m = re.fullmatch(r'(\d{1,2}):?(\d{2})(am|pm|a:m:|p:m:)?', t)
    if not m:
        return time(17, 50)
    hh, mm = int(m.group(1)), int(m.group(2))
    ampm = (m.group(3) or '').replace(':', '')
    if ampm == 'pm' and hh < 12:
        hh += 12
    elif ampm == 'am' and hh == 12:
        hh = 0
    if not (0 <= hh <= 23 and 0 <= mm <= 59):
        return time(17, 50)
    return time(hh, mm)


# Días inhábiles: los oficiales de la Ley Federal del Trabajo + los que se capturen en la configuración.
# Un día en que sí salió un ticket único cuenta como hábil aunque esté en la lista.
_INHABILES_EXTRA = set()
_USAR_OFICIALES = True
_HABILES_FORZADOS = set()


def _n_lunes(year: int, month: int, n: int) -> date:
    d = date(year, month, 1)
    d += timedelta(days=(7 - d.weekday()) % 7)
    return d + timedelta(weeks=n - 1)


def feriados_oficiales(year: int) -> set:
    """1 ene, 1er lunes de feb, 3er lunes de mar, 1 may, 16 sep, 3er lunes de nov, 25 dic."""
    return {date(year, 1, 1), _n_lunes(year, 2, 1), _n_lunes(year, 3, 3), date(year, 5, 1),
            date(year, 9, 16), _n_lunes(year, 11, 3), date(year, 12, 25)}


def parse_dias_inhabiles(texto: str) -> set:
    """'16/11/2026, 25-12-2026; 02/11' -> {date, ...} (sin año = año actual)."""
    dias = set()
    for tok in re.split(r'[,;\s]+', str(texto or '')):
        d = parse_fecha_asunto(tok) if tok else None
        if d:
            dias.add(d)
    return dias


def set_dias_inhabiles(dias=(), oficiales: bool = True):
    """Configura los días inhábiles que usan next/previous_business_day y el corte de cada día."""
    global _USAR_OFICIALES
    _INHABILES_EXTRA.clear()
    _INHABILES_EXTRA.update(dias or ())
    _USAR_OFICIALES = bool(oficiales)


def es_inhabil(d: date) -> bool:
    """Feriado oficial o día capturado como inhábil (no cuenta fines de semana)."""
    return d in _INHABILES_EXTRA or (_USAR_OFICIALES and d in feriados_oficiales(d.year))


def es_habil(d: date) -> bool:
    if d in _HABILES_FORZADOS:
        return True
    return d.weekday() < 5 and not es_inhabil(d)


def previous_business_day(d: date) -> date:
    """Día hábil anterior (lunes -> viernes; se salta fines de semana y días inhábiles)."""
    prev = d - timedelta(days=1)
    while not es_habil(prev):
        prev -= timedelta(days=1)
    return prev


def next_business_day(d: date) -> date:
    """Día hábil siguiente (viernes -> lunes; se salta fines de semana y días inhábiles)."""
    nxt = d + timedelta(days=1)
    while not es_habil(nxt):
        nxt += timedelta(days=1)
    return nxt


def dia_de_lote(llegada: datetime, corte=None) -> date:
    """
    Día hábil en cuyo ticket único debe ir una renovación: el mismo día si llegó a más tardar al corte de ese día
    (`corte`: datetime real de envío o una hora fija); sin corte (el ticket único aún no sale) es el mismo día.
    """
    d = llegada.date()
    if isinstance(corte, time):
        corte = datetime.combine(d, corte)
    if not es_habil(d) or (corte is not None and llegada > corte):
        d = next_business_day(d)
    return d


def es_respuesta(asunto: str) -> bool:
    """'RE: …', 'RV: …', 'FW: …': respuesta o reenvío del concentrado (no marca la hora de envío)."""
    return bool(re.match(r'\s*(re|rv|fw|fwd)\s*:', asunto or '', re.I))


def normalizar_fechas_concentrados(concentrados: list):
    """
    Decide de qué día es cada ticket único ('fecha_efectiva'): la fecha del asunto, salvo que sea un error evidente:
    - fechada en el futuro (después de cuando se envió) -> el día en que se envió;
    - fechada en un día que ya tiene su propio ticket único enviado ese mismo día, y el día del envío no tiene
      ninguno -> era el de ese día con la fecha mal escrita.
    Deja 'aviso_fecha' con la explicación cuando corrige.
    """
    originales = [c for c in concentrados if c.get('recibido') and not es_respuesta(c.get('asunto'))]
    mismo_dia = {c['fecha_asunto'] for c in originales if c.get('fecha_asunto') == c['recibido'].date()}
    for c in concentrados:
        c.pop('fecha_efectiva', None)
        c.pop('aviso_fecha', None)
        f, rec = c.get('fecha_asunto'), c.get('recibido')
        if not f or not rec or es_respuesta(c.get('asunto')):
            continue
        dia_envio = rec.date()
        if f > dia_envio:
            c['fecha_efectiva'] = dia_envio
            c['aviso_fecha'] = f"el asunto dice {f:%d/%m} pero se envió el {dia_envio:%d/%m}: se tomó como del {dia_envio:%d/%m}"
        elif f < dia_envio and f in mismo_dia and dia_envio not in mismo_dia and rec.time() >= time(6, 0):
            c['fecha_efectiva'] = dia_envio
            c['aviso_fecha'] = (f"el asunto dice {f:%d/%m} (que ya tiene su ticket único) pero se envió el "
                                f"{dia_envio:%d/%m}: se tomó como del {dia_envio:%d/%m}")


def cortes_reales(concentrados: list, dias_ultimo=()) -> dict:
    """
    {día: hora en que salió el PRIMER ticket único original de ese día}. Los que salen después (complementos)
    cuentan como evidencia de envío pero no mueven el corte. Ignora respuestas y reenvíos.
    En los `dias_ultimo` (días de emisión o de cierre de mes) cuenta el ÚLTIMO del día: ese día se manda otro ticket
    único con lo que llega después del primero, porque al día siguiente ya no se puede cancelar.
    """
    cortes, tardios = {}, {}
    for c in concentrados or []:
        dia, rec = _fecha_concentrado(c), c.get('recibido')
        if not dia or not rec or es_respuesta(c.get('asunto')) or rec.date() < dia:
            continue
        destino = cortes if rec.date() == dia else tardios  # tardío: enviado pasada la medianoche o días después
        destino[dia] = (max if destino is cortes and dia in dias_ultimo else min)(destino.get(dia, rec), rec)
    for dia, rec in tardios.items():
        cortes.setdefault(dia, rec)
    return cortes


def parse_fecha_asunto(subject: str, ref_year: int = None):
    """Fecha escrita en el asunto del concentrado: '22/10/2026', '22-10-26', '22.10', '22 de octubre'. None si no hay."""
    s = normalize_text(subject or '')
    ref_year = ref_year or datetime.now().year
    m = re.search(r'\b(\d{1,2})[/\-\.](\d{1,2})(?:[/\-\.](\d{2,4}))?\b', s)
    if m:
        dd, mm = int(m.group(1)), int(m.group(2))
        yy = int(m.group(3)) if m.group(3) else ref_year
        if yy < 100:
            yy += 2000
        try:
            return date(yy, mm, dd)
        except ValueError:
            pass
    m = re.search(r'\b(\d{1,2})\s+(?:de\s+)?([a-z]+)(?:\s+(?:de\s+|del\s+)?(\d{4}))?', s)
    if m and m.group(2) in _MESES_ES:
        try:
            return date(int(m.group(3) or ref_year), _MESES_ES[m.group(2)], int(m.group(1)))
        except ValueError:
            pass
    return None


def extract_concentrado_tokens(text: str) -> tuple:
    """
    Del cuerpo/adjuntos del concentrado obtiene:
    - tickets: set de folios numéricos ('I-361515' -> '361515')
    - numeros: set de todos los números de 3 a 10 dígitos (clientes y créditos, con o sin letra pegada)
    """
    text = str(text or '')
    tickets = {t.lstrip('0') for t in re.findall(r'\bI\s*-\s*(\d{3,10})\b', text, re.I)}
    numeros = {n.lstrip('0') for n in re.findall(r'(?<!\d)\d{3,10}(?!\d)', strip_letter_prefixes(text))}
    return tickets, numeros


def ticket_key(ticket: str) -> str:
    """Forma canónica de un folio para comparar: 'i-0362386' -> 'I-362386'."""
    prefix, num = _split_ticket_ref(ticket)
    return f"{prefix}-{num}" if num else ''


def extract_concentrado_rows(html: str) -> list:
    """
    Filas de la tabla del correo "Ticket Único" (INCIDENTE | FECHA | CLIENTE | CREDITO (HIJO)):
    [{'ticket': 'I-362277', 'fecha': '13/10/2026', 'cliente': '4826460', 'credito': '82508'}].
    Solo lee tablas sin tablas anidadas (las de diseño del correo envuelven a la de datos).
    """
    if not html:
        return []
    try:
        soup = make_soup(html)
    except Exception:
        return []
    filas = []
    for tbl in soup.find_all('table'):
        if tbl.find('table'):
            continue
        cols = None
        for tr in tbl.find_all('tr'):
            raw = [c.get_text(" ", strip=True) for c in tr.find_all(['td', 'th'])]
            if cols is None:
                cols = _columnas_concentrado([normalize_text(c) for c in raw])
                continue

            def val(campo):
                i = cols.get(campo)
                return raw[i] if i is not None and i < len(raw) else ''

            cli = first_digits(strip_letter_prefixes(val('cliente')))
            # La celda puede traer varios créditos del mismo cliente ("35237 y 35238")
            creditos = [n.lstrip('0') for n in re.findall(r'(?<!\d)\d{3,7}(?!\d)', strip_letter_prefixes(val('credito')))]
            if not cli and not creditos:
                continue
            for cred in creditos or ['']:
                filas.append({'ticket': ticket_key(val('ticket')) if re.search(r'\d', val('ticket')) else '',
                              'fecha': val('fecha'), 'cliente': cli, 'credito': cred})
    return filas


def _columnas_concentrado(encabezados: list) -> dict:
    """
    {'ticket', 'cliente', 'credito', 'fecha': índice} del renglón de encabezados, o None si no lo es. Con varias
    columnas de crédito se toma la del crédito renovado (hijo / a cancelar), no la del origen; con varias de cliente,
    la del número (no la del nombre).
    """
    idx, puntos = {}, {}

    def poner(campo, i, valor):
        if valor > puntos.get(campo, -1):
            idx[campo], puntos[campo] = i, valor

    for i, cell in enumerate(encabezados):
        if re.search(r'incidente|ticket|folio', cell):
            poner('ticket', i, 1)
        elif 'cliente' in cell or 'trabajador' in cell or cell in ('cuenta', 'no. cuenta'):
            poner('cliente', i, 0 if 'nombre' in cell else 2 if re.search(r'\bno\b|num|cuenta|#', cell) else 1)
        elif 'credito' in cell or 'renovacion' in cell or 'pagare' in cell:
            if re.search(r'padre|origen|anterior|original|antes', cell):
                poner('credito', i, 0)
            else:
                poner('credito', i, 2 if re.search(r'hijo|renov|cancel|nuevo|pagare', cell) else 1)
        elif 'fecha' in cell:
            poner('fecha', i, 1)
    return idx if {'cliente', 'credito'} <= idx.keys() else None


def extract_concentrado_intro(plain_body: str) -> str:
    """
    Texto del concentrado antes de la tabla ("Hola, Se envía “35” solicitudes…  FAVOR DE BLOQUEAR…"),
    con el número entre comillas cambiado por {total}, para reutilizarlo en el correo del día. '' si no aplica.
    """
    texto = str(plain_body or '').replace('\r\n', '\n')
    m = re.search(r'^\s*(INCIDENTE|TICKET|FOLIO)\b', texto, re.I | re.M) or re.search(r'\bI-\d{3,10}\b', texto)
    if not m:
        return ''
    intro = re.sub(r'\n{3,}', '\n\n', texto[:m.start()]).strip()
    if not intro or len(intro) > 1500 or 'solicitud' not in normalize_text(intro):
        return ''
    return re.sub(r'([“"«])\s*\d+\s*([”"»])', r'\g<1>{total}\g<2>', intro, count=1)


def _concentrado_filas_del_ticket(conc: dict, ticket: str) -> list:
    return [f for f in conc.get('filas') or [] if f.get('ticket') and first_digits(f['ticket']) == ticket]


def _row_match_concentrado(row: dict, concentrados: list):
    """
    Busca la renovación en los concentrados. Regla anti-omisión:
    - Si la fila tiene Cliente y Crédito, solo cuenta como enviada si AMBOS aparecen en el concentrado
      (que esté el ticket no basta: un ticket con 3 renovaciones pudo enviarse con solo 2).
      Si el concentrado trae su tabla (INCIDENTE | FECHA | CLIENTE | CREDITO), se compara renglón por renglón;
      el mismo ticket con el mismo crédito pero OTRO cliente NO cuenta: se envió con datos equivocados.
    - Si a la fila le falta el cliente o el crédito, basta con que aparezca el ticket.
    Retorna (tipo_match, concentrado) con tipo_match en {'PAR', 'TICKET', None}.
    """
    ticket_hit = None
    for conc in concentrados:
        tipo = _match_en_concentrado(row, conc)
        if tipo == 'PAR':
            return 'PAR', conc
        if tipo == 'TICKET' and ticket_hit is None:
            ticket_hit = conc
    if ticket_hit is not None:
        return 'TICKET', ticket_hit
    return None, None


def _junto_al_ticket(texto: str, ticket: str, numeros: list, renglones_tabla=()) -> bool:
    """
    ¿Vienen todos los `numeros` justo después del folio (antes del siguiente folio, a menos de 150 caracteres)?
    Se salta la aparición que es el renglón de la tabla (`renglones_tabla`: [(cliente, crédito)] de ese folio).
    """
    if not texto or not ticket or not numeros:
        return False
    for m in re.finditer(rf'(?<!\d)0*{re.escape(ticket)}(?!\d)', texto):
        tramo = re.split(r'\b[IR]\s*-\s*\d{3,}', texto[m.end():m.end() + 150], maxsplit=1)[0]
        orden = [n.lstrip('0') for n in re.findall(r'\d{3,10}', strip_letter_prefixes(tramo))]
        if any(c in orden[:4] and cr in orden[:4] for c, cr in renglones_tabla):
            continue
        if all(n in orden for n in numeros):
            return True
    return False


def _match_en_concentrado(row: dict, conc: dict):
    """'PAR' si la renovación (Cliente+Crédito) viene en este concentrado, 'TICKET' si solo viene su ticket, None si no."""
    cli = '' if str(row.get('Numero_Cliente', '')).startswith('SIN') else first_digits(row.get('Numero_Cliente'))
    cred = '' if str(row.get('Credito_Renovado_Hijo', '')).startswith('SIN') else first_digits(row.get('Credito_Renovado_Hijo'))
    ticket = first_digits(row.get('Numero_Ticket'))
    filas = conc.get('filas') or []
    for f in filas:
        if cli and cred and f['cliente'] == cli and f['credito'] == cred:
            return 'PAR'
        if not cli and ticket and cred and f['credito'] == cred and first_digits(f['ticket']) == ticket:
            return 'PAR'
        if cli and not cred and ticket and f['cliente'] == cli and first_digits(f['ticket']) == ticket:
            return 'PAR'  # Sin crédito: basta su ticket con su mismo cliente (no cualquier renglón del ticket)
    # El renglón viene escrito fuera de la tabla (otra lista, otro formato de quien lo mandó): «I-362605 6572804 35238»
    del_ticket = [(f['cliente'], f['credito']) for f in filas if f.get('ticket') and first_digits(f['ticket']) == ticket]
    if ticket and cred and _junto_al_ticket(conc.get('texto'), ticket, [n for n in (cli, cred) if n], del_ticket):
        return 'PAR'
    # Sin tabla legible, o el ticket no viene en ella: cuenta si el Cliente y el Crédito están en el correo
    en_la_tabla = ticket and any(first_digits(f['ticket']) == ticket for f in filas if f.get('ticket'))
    if not en_la_tabla and cli and cred and cli in conc['numeros'] and cred in conc['numeros']:
        return 'PAR'
    if ticket and ticket in conc['tickets']:
        return 'TICKET'
    return None


def _num(valor) -> str:
    """Dígitos de un cliente/crédito, '' si es 'SIN CLIENTE' / 'SIN CRÉDITO'."""
    return '' if str(valor or '').startswith('SIN') else first_digits(valor)


def error_en_ticket_unico(row: dict, filas_del_ticket: list, incidentes_de_cliente: dict) -> str:
    """
    El ticket único trae el incidente de la renovación pero con otros datos: explica qué está mal
    ("salió con Cliente 4530500 en lugar de 5864479 (ese cliente es del incidente I-361002)"). '' si no aplica.
    """
    cli, cred = _num(row.get('Numero_Cliente')), _num(row.get('Credito_Renovado_Hijo'))
    propio = row.get('Numero_Ticket')
    partes = []
    for f in filas_del_ticket:
        if cred and f['credito'] == cred and cli and f['cliente'] != cli:
            otro = sorted(t for t in incidentes_de_cliente.get(f['cliente'], ()) if t != propio)
            partes.append(f"salió con Cliente {f['cliente']} en lugar de {cli}"
                          + (f" (ese cliente es del incidente {', '.join(otro)})" if otro else ""))
        elif cli and f['cliente'] == cli and cred and f['credito'] != cred:
            partes.append(f"salió con Crédito {f['credito']} en lugar de {cred}")
    return "; ".join(partes)


def _fecha_concentrado(c: dict):
    return c.get('fecha_efectiva') or c.get('fecha_asunto') or (c['recibido'].date() if c.get('recibido') else None)


def _buscar_envio(row: dict, concentrados: list, llegada: datetime):
    """
    Busca en TODOS los concentrados (ordenados por fecha) el ticket único donde se envió la renovación.
    Retorna (concentrado_envio, enviado_antes_de_llegar, concentrado_con_solo_el_ticket).
    Se prefiere el primer envío posterior a la llegada; si solo aparece en uno anterior, es una
    renovación que ya se había mandado en otro ticket (duplicado).
    """
    incompleta = bool(row.get('Datos_Incompletos'))
    envios, ticket_hit = [], None
    for c in concentrados:
        tipo = _match_en_concentrado(row, c)
        # Con tabla legible, una renovación incompleta necesita su renglón (PAR); sin tabla basta el ticket
        if tipo == 'PAR' or (tipo == 'TICKET' and incompleta and not c.get('filas')):
            envios.append(c)
        elif tipo == 'TICKET' and ticket_hit is None:
            ticket_hit = c
    if not envios:
        return None, False, ticket_hit
    posteriores = [
        c for c in envios
        if (c.get('recibido') or datetime.combine(_fecha_concentrado(c) or llegada.date(), time(23, 59))) >= llegada
    ]
    return (posteriores[0], False, ticket_hit) if posteriores else (envios[-1], True, ticket_hit)


def concentrado_semanal(concentrados: list, rows: list, fecha_inicio: date, fecha_fin: date) -> dict:
    """
    Junta en una sola lista todo lo que salió en los tickets únicos desde `fecha_inicio` (incluye los posteriores al
    periodo: ahí llegan las que se mandaron tarde). Cada renglón dice si se encontró el correo del incidente en Outlook
    y si el mismo Cliente+Crédito salió en más de un ticket único. Ignora respuestas y reenvíos (RE:/FW:).
    Regresa {'filas': [...], 'sin_tabla': [asuntos de concentrados sin tabla legible]}.
    """
    pares, por_ticket, incidentes_de_cliente = {}, {}, {}
    for r in rows:
        cli, cred = _num(r.get('Numero_Cliente')), _num(r.get('Credito_Renovado_Hijo'))
        if cli and cred:
            pares.setdefault((cli, cred), [])
            if r['Numero_Ticket'] not in pares[(cli, cred)]:
                pares[(cli, cred)].append(r['Numero_Ticket'])
        if cli:
            incidentes_de_cliente.setdefault(cli, set()).add(r['Numero_Ticket'])
        por_ticket.setdefault(first_digits(r.get('Numero_Ticket')), []).append(r)

    filas, sin_tabla, vistos = [], [], {}
    for c in sorted(concentrados or [], key=lambda x: x.get('recibido') or datetime.min):
        dia = _fecha_concentrado(c)
        if not dia or dia < fecha_inicio or es_respuesta(c.get('asunto')):
            continue
        if not c.get('filas'):
            sin_tabla.append(c.get('asunto', ''))
            continue
        for f in c['filas']:
            solicitantes, error = pares.get((f['cliente'], f['credito']), []), ''
            num_ticket = first_digits(f['ticket'])
            propio = [t for t in solicitantes if first_digits(t) == num_ticket]
            incidente = (propio or solicitantes or [''])[0]
            if solicitantes and num_ticket and not propio:
                error = f"este Cliente+Crédito es del incidente {', '.join(solicitantes)}, no del {f['ticket']}"
            elif not incidente and num_ticket in por_ticket:
                incidente = por_ticket[num_ticket][0]['Numero_Ticket']
                for r in por_ticket[num_ticket]:
                    error = error_en_ticket_unico(r, [f], incidentes_de_cliente)
                    if error:
                        break
                error = error or f"el incidente {incidente} no tiene Cliente {f['cliente']} Crédito {f['credito']}"
            clave = (f['cliente'], f['credito'])
            vistos[clave] = vistos.get(clave, 0) + 1
            # Día en que se pidió (para la base): el de llegada del incidente; si no se encontró, puede ser del día
            # hábil anterior (llegó después del envío del ticket único de ese día)
            llegadas = [r.get('Fecha_Llegada_DT') or r.get('Fecha_DT') for r in por_ticket.get(first_digits(incidente), [])]
            llegada = next((x.date() if isinstance(x, datetime) else x for x in llegadas if x), None)
            filas.append({
                'llegada': llegada, 'desde': previous_business_day(dia),
                'fecha': dia, 'recibido': c.get('recibido'), 'asunto': c.get('asunto', ''),
                'ticket': f['ticket'], 'cliente': f['cliente'], 'credito': f['credito'],
                'incidente': incidente, 'error': error, 'en_periodo': dia <= fecha_fin,
                'duplicada': [t for t in solicitantes if t != incidente],
            })
    for f in filas:
        f['veces'] = vistos[(f['cliente'], f['credito'])]
    return {'filas': filas, 'sin_tabla': sin_tabla}


def resumen_por_dia(rows: list, concentrados: list, fecha_inicio: date, fecha_fin: date) -> list:
    """
    Por cada día de lunes a sábado del periodo (el sábado, por la guardia): ticket(s) único(s) que salieron y cómo
    quedó lo que debía ir en ellos.
    """
    dias, d = [], fecha_inicio
    while d <= fecha_fin:
        if d.weekday() < 6 or d in _HABILES_FORZADOS:
            envios = [c for c in concentrados if _fecha_concentrado(c) == d and not es_respuesta(c.get('asunto'))]
            suyas = [r for r in rows if r.get('Dia_Lote') == d]
            dias.append({
                'dia': d,
                'habil': es_habil(d),
                'envios': [(c['recibido'], c.get('remitente') or c.get('carpeta') or '', len(c.get('filas') or []))
                           for c in envios if c.get('recibido')],
                'debian': len(suyas),
                'a_tiempo': sum(1 for r in suyas if r['Lote_Estado'] in (LOTE_ENVIADA, LOTE_EN_CONCENTRADO)
                                and not r.get('Dias_Retraso') and not r.get('Posible_Duplicado')),
                'tarde': sum(1 for r in suyas if r.get('Dias_Retraso')),
                'no_enviadas': sum(1 for r in suyas if r.get('No_Enviada')),
                'errores': sum(1 for r in suyas if r.get('Error_Envio')),
                'sin_comprobar': sum(1 for r in suyas if r['Lote_Estado'] == LOTE_SIN_COMPROBAR),
                'pendientes': sum(1 for r in suyas if r['Lote_Estado'] in (LOTE_HOY, LOTE_REZAGADA) and not r.get('No_Enviada')),
            })
        d += timedelta(days=1)
    return dias


def clasificar_lotes(
    rows: list,
    fecha_inicio: date,
    fecha_fin: date,
    hora_corte: time,
    concentrados: list = None,
    ultimo_lote: dict = None,
    ahora: datetime = None,
    revision: bool = None,
    dia_lote: date = None,
    dias_de_corte: dict = None
) -> dict:
    """
    Concilia cada renovación contra TODOS los correos "Ticket Único" encontrados y le asigna
    'Lote_Estado', 'Lote_Etiqueta' (texto completo), 'Lote_Resumen' (texto corto para la tabla),
    'Dia_Lote' (día del ticket único en que debía ir), 'Enviado_En' / 'Enviado_Fecha' / 'Dias_Retraso'
    y 'No_Enviada'. Regresa un resumen para la interfaz.

    - Corte de cada día: la hora en que salió su ticket único (`cortes_reales`). Hoy, mientras no sale, todo entra
      hoy. Solo los días pasados sin ticket único encontrado usan `hora_corte` (hora habitual de respaldo).
    - Día del lote: el mismo día si llegó a más tardar al corte; si no, el día hábil siguiente. Si llegó en los
      MARGEN_ENVIO previos al envío y no viene, se toma como que llegó mientras se armaba (va al día siguiente).
    - Lote actual: el día hábil de `ahora` (por omisión, fecha_fin). Lo de días anteriores se revisa:
        * Viene en algún ticket único -> enviada (a tiempo, tarde o ya enviada antes en otro ticket).
        * Su día ya pasó, se encontró el ticket único de ese día y no viene -> NO se envió:
          si el incidente sigue abierto entra al lote de hoy (rezagada); si ya se cerró se marca aparte.
        * Llegó ayer después del envío (o en fin de semana) -> rezagada que va hoy.
        * Sin ticket único del día anterior: marca local "último lote copiado/exportado" o, sin marca, la hora habitual.
        * Días más viejos sin ticket único encontrado -> sin comprobar.
    Con un periodo de varios días (fecha_inicio < fecha_fin) funciona como revisión de envíos; `revision` lo fuerza.
    `dia_lote`: arma el lote de un día pasado tal como debía salir ese día (se compara solo contra los ticket únicos
    hasta ese día; si se mandó después, se anota).
    `dias_de_corte` ({día: texto}, generaciones de emisión y cierres de mes): ese día el corte es el último ticket
    único que salió; y hoy, si es día de corte, lo que llega después del ticket único se manda en otro ese mismo día.
    """
    dias_de_corte = dias_de_corte or {}
    concentrados = sorted(concentrados or [], key=lambda c: c.get('recibido') or datetime.min)
    normalizar_fechas_concentrados(concentrados)
    # Si salió un ticket único un día marcado como inhábil, ese día sí se trabajó
    _HABILES_FORZADOS.clear()
    _HABILES_FORZADOS.update(_fecha_concentrado(c) for c in concentrados
                             if c.get('recibido') and not es_respuesta(c.get('asunto'))
                             and c['recibido'].date() == _fecha_concentrado(c))
    pasado = dia_lote is not None
    if pasado:
        ahora = datetime.combine(dia_lote, time(23, 59, 59))
    ahora = ahora or datetime.combine(fecha_fin, time(0, 0))
    lote_actual = ahora.date() if es_habil(ahora.date()) else next_business_day(ahora.date())
    inicio_actual = datetime.combine(lote_actual, time(0, 0))
    hoy_real = lote_actual == datetime.now().date()
    corte_fin = datetime.combine(fecha_fin, hora_corte)
    dia_anterior = previous_business_day(lote_actual)
    cortes = cortes_reales(concentrados, dias_de_corte)
    corte_hoy = cortes.get(lote_actual)
    # Hoy es día de emisión o de cierre: lo que llegue después del ticket único va en otro ticket único hoy mismo
    corte_de_hoy = dias_de_corte.get(lote_actual) if not pasado else None
    envios_hoy = sorted(c['recibido'] for c in concentrados if c.get('recibido') and not es_respuesta(c.get('asunto'))
                        and _fecha_concentrado(c) == lote_actual and c['recibido'].date() == lote_actual)
    corte_anterior = cortes.get(dia_anterior) or datetime.combine(dia_anterior, hora_corte)
    hc = hora_corte.strftime('%H:%M')
    if revision is None:
        revision = fecha_inicio < fecha_fin or fecha_fin < lote_actual
    de_hoy = f"del {lote_actual:%d/%m}" if pasado else "de hoy"
    va_hoy = f"va el {lote_actual:%d/%m}" if pasado else "va hoy"
    # El lote de un día pasado se compara solo contra lo enviado hasta ese día
    posteriores = [c for c in concentrados if pasado and (_fecha_concentrado(c) or date.max) > lote_actual]
    para_buscar = [c for c in concentrados if c not in posteriores]

    def corte_del(d: date):
        if corte_de_hoy and d == lote_actual:
            return None
        if d in cortes:
            return cortes[d]
        return None if d >= lote_actual else datetime.combine(d, hora_corte)

    conc_previos = [c for c in concentrados if _fecha_concentrado(c) and dia_anterior <= _fecha_concentrado(c) < lote_actual]

    marca = None
    if ultimo_lote and ultimo_lote.get('hasta'):
        try:
            marca = datetime.fromisoformat(ultimo_lote['hasta'])
        except (TypeError, ValueError):
            marca = None
        if marca and not (datetime.combine(dia_anterior, time(0, 0)) <= marca < inicio_actual):
            marca = None

    if conc_previos:
        metodo = 'concentrado'
    elif marca:
        metodo = 'marca'
    else:
        metodo = 'horario'

    def hay_ticket_unico_del(dia: date) -> bool:
        fin = next_business_day(dia)
        return any(_fecha_concentrado(c) and dia <= _fecha_concentrado(c) < fin for c in concentrados)

    def asunto(c):
        return f"«{c['asunto'][:60]}»"

    def poner(r, estado, etiqueta, resumen):
        r['Lote_Estado'], r['Lote_Etiqueta'], r['Lote_Resumen'] = estado, etiqueta, resumen

    incidentes_de_cliente, incidentes_del_par = {}, {}
    for r in rows:
        if _num(r.get('Numero_Cliente')):
            incidentes_de_cliente.setdefault(_num(r.get('Numero_Cliente')), set()).add(r.get('Numero_Ticket'))
            incidentes_del_par.setdefault((_num(r.get('Numero_Cliente')), _num(r.get('Credito_Renovado_Hijo'))),
                                          set()).add(first_digits(r.get('Numero_Ticket')))

    ultima_considerada = None
    for r in rows:
        llegada = r.get('Fecha_Llegada_DT') or r.get('Fecha_DT')
        r['Fecha_Llegada_DT'] = llegada
        dia = dia_de_lote(llegada, corte_del(llegada.date()))
        r['Enviado_En'], r['Enviado_Fecha'], r['Dias_Retraso'], r['No_Enviada'] = '', None, 0, False
        r['Falta_Hoy'] = False
        r['Error_Envio'] = ''
        r['Posible_Duplicado'] = False
        r['Enviado_Despues'] = None

        conc, antes, ticket_hit = _buscar_envio(r, para_buscar, llegada)
        junto_al_envio = None
        if (not conc and dia == llegada.date() and dia in cortes and llegada > cortes[dia] - MARGEN_ENVIO
                and not (corte_de_hoy and dia == lote_actual)):
            # Llegó mientras se armaba el ticket único: no es un olvido, le toca el siguiente
            junto_al_envio = cortes[dia]
            dia = next_business_day(dia)
        r['Dia_Lote'] = dia
        r['En_Periodo'] = fecha_inicio <= dia <= fecha_fin
        if conc:
            f_env = _fecha_concentrado(conc)
            r['Enviado_En'], r['Enviado_Fecha'] = conc['asunto'], f_env
            # Cliente+Crédito correctos pero con el número de otro incidente
            propio = first_digits(r.get('Numero_Ticket'))
            par = (_num(r.get('Numero_Cliente')), _num(r.get('Credito_Renovado_Hijo')))
            # Otro folio que pide lo mismo (solicitud duplicada) no es un error del ticket único
            ajenos = sorted({f['ticket'] for f in conc.get('filas') or []
                             if (f['cliente'], f['credito']) == par and f.get('ticket')
                             and first_digits(f['ticket']) not in incidentes_del_par.get(par, {propio})})
            if ajenos and not any(first_digits(f['ticket']) == propio for f in conc.get('filas') or []
                                  if f['credito'] == _num(r.get('Credito_Renovado_Hijo'))):
                r['Error_Envio'] = f"salió con el incidente {', '.join(ajenos)} en lugar de {r.get('Numero_Ticket')}"
            if antes:
                # Ya había salido antes de que llegara: solicitud repetida, no entra al lote ni se da por validada
                r['Posible_Duplicado'] = True
                poner(r, LOTE_ENVIADA,
                      f"📨 Ya se había enviado en {asunto(conc)} del {f_env:%d/%m}, antes de que llegara este ticket "
                      f"(posible solicitud duplicada)",
                      f"Ya enviada antes ({f_env:%d/%m})  ·  posible duplicado")
                continue
            if llegada >= inicio_actual:
                poner(r, LOTE_EN_CONCENTRADO, f"📨 Ya en {asunto(conc)}", f"Ya en el ticket único {de_hoy}")
                continue
            if antes:
                poner(r, LOTE_ENVIADA,
                      f"📨 Ya se había enviado en {asunto(conc)} del {f_env:%d/%m} (antes de que llegara este ticket)",
                      f"Ya enviada antes ({f_env:%d/%m})")
            elif f_env > dia:
                r['Dias_Retraso'] = (f_env - dia).days
                poner(r, LOTE_ENVIADA,
                      f"📨 Enviada tarde en {asunto(conc)} del {f_env:%d/%m}; debía ir en el del {dia:%d/%m} "
                      f"({r['Dias_Retraso']} día(s) después)",
                      f"Enviada tarde {f_env:%d/%m}  ·  debía ir {dia:%d/%m}")
            else:
                poner(r, LOTE_ENVIADA, f"📨 Ya enviada en {asunto(conc)} del {f_env:%d/%m}", f"Enviada {f_env:%d/%m}")
            if ultima_considerada is None or llegada > ultima_considerada['Fecha_Llegada_DT']:
                ultima_considerada = r
            continue

        # No viene en ningún ticket único
        motivo_ticket = ''
        if ticket_hit:
            en_conc = _concentrado_filas_del_ticket(ticket_hit, first_digits(r.get('Numero_Ticket')))
            r['Error_Envio'] = error_en_ticket_unico(r, en_conc, incidentes_de_cliente)
            if r['Error_Envio']:
                motivo_ticket = (f"DATOS EQUIVOCADOS en el ticket único del {_fecha_concentrado(ticket_hit):%d/%m}: "
                                 f"{r['Error_Envio']}")
            else:
                motivo_ticket = "el ticket está en el concentrado pero NO este Cliente+Crédito"
                if en_conc:
                    motivo_ticket += "; ahí viene " + ", ".join(f"Cliente {f['cliente']} Crédito {f['credito']}" for f in en_conc)
        equivocada = "datos equivocados  ·  " if r['Error_Envio'] else ""

        if dia > lote_actual:
            if junto_al_envio:
                poner(r, LOTE_SIGUIENTE,
                      f"➡️ Llegó a las {llegada:%H:%M}, mientras se enviaba el ticket único {de_hoy} ({junto_al_envio:%H:%M}); "
                      f"va al lote del {dia:%d/%m}",
                      f"Llegó junto al envío  ·  va el {dia:%d/%m}")
            elif corte_hoy and llegada.date() == lote_actual:
                poner(r, LOTE_SIGUIENTE,
                      f"➡️ Llegó después del ticket único {de_hoy} ({corte_hoy:%H:%M}); va al lote del {dia:%d/%m}",
                      f"Después del envío  ·  va el {dia:%d/%m}")
            else:
                poner(r, LOTE_SIGUIENTE, f"➡️ Va al lote del {dia:%d/%m}", f"Va el {dia:%d/%m}")
        elif llegada >= inicio_actual:
            texto = "Lote de hoy" if not pasado and (hoy_real or lote_actual == fecha_fin) else f"Lote del {lote_actual:%d/%m}"
            if corte_de_hoy and corte_hoy and llegada > corte_hoy - MARGEN_ENVIO:
                # Día de emisión o de cierre: lo que llegó después del ticket único sale hoy en otro
                poner(r, LOTE_HOY,
                      f"📅 {texto}: llegó después del ticket único de las {corte_hoy:%H:%M} y hoy es "
                      f"{corte_de_hoy}: mándala HOY en otro ticket único (mañana ya no se puede cancelar)",
                      f"{texto}  ·  otro ticket único hoy (corte)")
            elif corte_hoy:
                # El ticket único de hoy ya salió y no la trae (llegó bastante antes del envío)
                r['Falta_Hoy'] = True
                detalle = motivo_ticket or f"llegó a las {llegada:%H:%M}"
                poner(r, LOTE_HOY,
                      f"⚠️ {texto}: falta en el ticket único que salió a las {corte_hoy:%H:%M} "
                      f"({detalle}; agrégala o mándala en el siguiente)",
                      f"{texto}  ·  {equivocada or 'falta '}en el de las {corte_hoy:%H:%M}")
            else:
                poner(r, LOTE_HOY, f"🟢 {texto}", texto)
        elif dia == lote_actual:
            # Llegó ayer después del corte o en fin de semana: le toca el lote de hoy
            if metodo == 'marca' and llegada <= marca:
                poner(r, LOTE_ENVIADA, f"📨 Enviada (antes de la marca {marca:%d/%m %H:%M})", "Enviada (marca local)")
                if ultima_considerada is None or llegada > ultima_considerada['Fecha_Llegada_DT']:
                    ultima_considerada = r
                continue
            if motivo_ticket:
                motivo, corto = motivo_ticket, ("datos equivocados en el ticket único" if r['Error_Envio']
                                                else "otros datos en el concentrado")
            elif junto_al_envio:
                motivo, corto = f"llegó mientras se enviaba el ticket único ({junto_al_envio:%H:%M})", "llegó junto al envío"
            elif not es_habil(llegada.date()):
                motivo, corto = ("llegó en fin de semana" if llegada.weekday() >= 5 else "llegó en día inhábil"), "día inhábil"
            elif llegada.date() in cortes:
                motivo, corto = f"llegó después del envío del ticket único ({cortes[llegada.date()]:%H:%M})", "después del envío"
            elif metodo == 'concentrado':
                motivo, corto = "no está en el concentrado", "no se envió"
            elif metodo == 'marca':
                motivo, corto = "después de la última marca", "después de la marca"
            else:
                motivo, corto = f"llegó después de las {hc}, hora habitual", "después de la hora habitual"
            poner(r, LOTE_REZAGADA, f"⏰ Rezagada {llegada:%d/%m %H:%M} ({motivo})",
                  f"Rezagada {llegada:%d/%m %H:%M}  ·  {corto}")
        elif hay_ticket_unico_del(dia) or ticket_hit:
            # Su ticket único ya salió y no la trae: NO se envió
            r['No_Enviada'] = True
            motivo = motivo_ticket or f"no viene en el ticket único del {dia:%d/%m}"
            cierre = r.get('Cierre_iTop') or {}
            if r.get('Cerrado_iTop'):
                f_c = cierre.get('fecha_dt')
                cuando = f" el {f_c:%d/%m}" if f_c else ""
                poner(r, LOTE_CERRADA_SIN_ENVIO,
                      f"❌ NO SE ENVIÓ: debía ir en el ticket único del {dia:%d/%m} ({motivo}) y el incidente ya se cerró{cuando}",
                      f"NO ENVIADA ({dia:%d/%m})  ·  {equivocada}cerrado{cuando}")
            else:
                poner(r, LOTE_REZAGADA,
                      f"❌ NO SE ENVIÓ: debía ir en el ticket único del {dia:%d/%m} ({motivo}); se suma al lote {de_hoy}",
                      f"NO ENVIADA ({dia:%d/%m})  ·  {equivocada}{va_hoy}")
        elif dia == dia_anterior and metodo == 'marca':
            if llegada <= marca:
                poner(r, LOTE_ENVIADA, f"📨 Enviada (antes de la marca {marca:%d/%m %H:%M})", "Enviada (marca local)")
                if ultima_considerada is None or llegada > ultima_considerada['Fecha_Llegada_DT']:
                    ultima_considerada = r
            else:
                poner(r, LOTE_REZAGADA, f"⏰ Rezagada {llegada:%d/%m %H:%M} (después de la última marca)",
                      f"Rezagada {llegada:%d/%m %H:%M}  ·  después de la marca")
        elif dia == dia_anterior:
            poner(r, LOTE_SIN_COMPROBAR,
                  f"❔ Sin comprobar: no se encontró el ticket único del {dia:%d/%m} (llegó antes de las {hc}); "
                  f"no hay prueba de que saliera. Revísalo o usa «Restaurar» para sumarla al lote",
                  f"Sin comprobar ({dia:%d/%m})  ·  sin ticket único")
        else:
            poner(r, LOTE_SIN_COMPROBAR,
                  f"❔ Sin comprobar: no se encontró el ticket único del {dia:%d/%m} y no viene en ningún otro",
                  f"Sin comprobar ({dia:%d/%m})")

    # El ticket único del lote ya salió: las rezagadas que no trae también le faltan
    if corte_hoy and not revision and not corte_de_hoy:
        for r in rows:
            if r['Lote_Estado'] == LOTE_REZAGADA and not r['Falta_Hoy']:
                r['Falta_Hoy'] = True
                r['Lote_Etiqueta'] += f" — tampoco viene en el ticket único {de_hoy} ({corte_hoy:%H:%M})"
                r['Lote_Resumen'] += f"  ·  falta en el de las {corte_hoy:%H:%M}"

    # Lote de un día pasado: lo que no salió ese día pero sí después se anota (no deja de ser un faltante de ese día)
    for r in rows:
        if posteriores and r['Lote_Estado'] not in (LOTE_ENVIADA, LOTE_EN_CONCENTRADO, LOTE_SIGUIENTE):
            post, _, _ = _buscar_envio(r, posteriores, r['Fecha_Llegada_DT'])
            if post:
                r['Enviado_Despues'] = _fecha_concentrado(post)
                r['Lote_Etiqueta'] += f" — se mandó después, en el ticket único del {r['Enviado_Despues']:%d/%m}"
                r['Lote_Resumen'] += f"  ·  salió el {r['Enviado_Despues']:%d/%m}"

    conc_ref = max(conc_previos, key=lambda c: c.get('recibido') or datetime.min) if conc_previos else None
    # En una revisión solo cuenta lo que debía ir en algún ticket único del periodo (no lo de la semana anterior)
    base = [r for r in rows if r['En_Periodo']] if revision else rows
    conteos = {}
    for r in base:
        conteos[r['Lote_Estado']] = conteos.get(r['Lote_Estado'], 0) + 1
    return {
        'metodo': metodo,
        'dia_anterior': dia_anterior,
        'corte_anterior': corte_anterior,
        'corte_fin': corte_fin,
        'hora_corte': hora_corte,
        'cortes': cortes,
        'corte_hoy': corte_hoy,
        'envios_hoy': envios_hoy,
        'corte_de_hoy': corte_de_hoy,
        'dias_de_corte': dias_de_corte,
        'faltan_hoy': sum(1 for r in rows if r.get('Falta_Hoy')),
        'fecha_inicio': fecha_inicio,
        'fecha_fin': fecha_fin,
        'lote_actual': lote_actual,
        'pasado': pasado,
        'revision': revision,
        'semana': concentrado_semanal(concentrados, rows, fecha_inicio, fecha_fin),
        'en_periodo': sum(1 for r in rows if r['En_Periodo']),
        'enviadas': sum(1 for r in base if r['Lote_Estado'] in (LOTE_ENVIADA, LOTE_EN_CONCENTRADO)),
        'concentrados': concentrados,
        'concentrado_previo': conc_ref,
        'marca': marca,
        'ultima_considerada': (
            {
                'ticket': ultima_considerada['Numero_Ticket'],
                'fecha': ultima_considerada['Fecha_Llegada_DT'],
                'cliente': ultima_considerada['Numero_Cliente'],
                'credito': ultima_considerada['Credito_Renovado_Hijo'],
            } if ultima_considerada else None
        ),
        'conteos': conteos,
        'tarde': sum(1 for r in base if r.get('Dias_Retraso')),
        'no_enviadas': sum(1 for r in base if r.get('No_Enviada')),
        'errores_envio': sum(1 for r in base if r.get('Error_Envio')),
        'duplicadas': sum(1 for r in base if r.get('Posible_Duplicado')),
        'por_dia': resumen_por_dia(base, concentrados, fecha_inicio, min(fecha_fin, lote_actual)) if revision else [],
        'avisos_fecha': [f"«{c['asunto']}»: {c['aviso_fecha']}" for c in concentrados if c.get('aviso_fecha')],
        'cerradas_sin_envio': sum(1 for r in base if r['Lote_Estado'] == LOTE_CERRADA_SIN_ENVIO),
        'sin_comprobar': sum(1 for r in base if r['Lote_Estado'] == LOTE_SIN_COMPROBAR),
    }


# ======================= CIERRE EN iTop ("Aviso de solución del Ticket") =======================

_CAMPOS_CIERRE = (
    ('estado', r'^estado$'),
    ('codigo', r'^codigo de solucion$'),
    ('motivo', r'^(solucion|objetivo de solucion|comentario de solucion|respuesta|motivo( de (cierre|solucion))?)$'),
    ('cerro', r'resuelto por|cerrado por|atendido por|^agente|tecnico|analista|especialista|^asignado a|responsable'),
    ('fecha', r'^fecha.*(solucion|resolucion|cierre)'),
    ('solicitante', r'^(nombre|solicitante)$'),
    ('categoria', r'^categoria$'),
)


def is_aviso_solucion(subject: str) -> str:
    """Folio si el asunto es un aviso de solución/cierre de iTop ('I-362386 Aviso de solución del Ticket'); '' si no."""
    s = normalize_text(subject)
    if not re.search(r'aviso de solucion|solucion del (ticket|reporte|incidente)|(ticket|reporte|incidente) \S+ (fue )?(resuelto|cerrado)', s):
        return ''
    m = re.search(r'\b([ir])\s*-\s*(\d{3,10})\b', s)
    return ticket_key(f"{m.group(1)}-{m.group(2)}") if m else ''


def parse_aviso_solucion(html_body: str, plain_body: str = '') -> dict:
    """
    Campos del correo de cierre de iTop (tabla Etiqueta | Valor): estado, motivo (Solución / Objetivo de solución),
    código de solución, quién lo cerró (si el correo lo trae), fecha, solicitante y categoría.
    """
    pares = []
    if html_body:
        try:
            for tr in make_soup(html_body).find_all('tr'):
                tds = tr.find_all(['td', 'th'], recursive=False) or tr.find_all(['td', 'th'])
                if len(tds) >= 2:
                    pares.append((tds[0].get_text(" ", strip=True), tds[1].get_text(" ", strip=True)))
        except Exception:
            pass
    if not pares and plain_body:
        lineas = [ln.strip() for ln in str(plain_body).replace('\r', '').split('\n')]
        for i, ln in enumerate(lineas):
            m = re.match(r'([^:\t]{3,40}?)\s*[:\t]\s*(.*)$', ln)
            if not m:
                continue
            valor = m.group(2).strip() or next((x for x in lineas[i + 1:i + 3] if x), '')
            pares.append((m.group(1), valor))

    campos = {}
    for etiqueta, valor in pares:
        clave = normalize_text(etiqueta).strip(' :#')
        valor = re.sub(r'\s+', ' ', valor).strip()
        if not valor or len(clave) > 40:
            continue
        for campo, patron in _CAMPOS_CIERRE:
            if campo not in campos and re.search(patron, clave):
                campos[campo] = valor
                break
    return campos


def resumen_cierre(cierre: dict, corto: bool = False) -> str:
    """
    'Resuelto 22/10 17:48 · 7353 Sammy Avila · Por instrucciones superiores…' para reportes.
    corto=True: 'Resuelto 22/10 17:48 · 7353 Sammy' (sin motivo, solo gafete y primer nombre) para la tabla.
    """
    if not cierre:
        return ''
    partes = [cierre.get('estado') or 'Cerrado']
    if cierre.get('fecha_dt'):
        partes[0] += f" {cierre['fecha_dt'].strftime('%d/%m %H:%M')}"
    if cierre.get('cerro'):
        cerro = cierre['cerro']
        if corto:
            palabras = cerro.split()
            n = 2 if palabras and palabras[0].isdigit() else 1
            cerro = " ".join(palabras[:n])
        partes.append(cerro)
    if cierre.get('motivo') and not corto:
        partes.append(cierre['motivo'])
    return "  ·  ".join(partes)
