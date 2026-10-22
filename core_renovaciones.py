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
from datetime import datetime
from bs4 import BeautifulSoup


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
                soup = BeautifulSoup(html_body, 'lxml')
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
            r'(cancel\w*\s+(de\s+|la\s+|el\s+|las\s+|los\s+)?(credito\s+renovado|renovaci|ren\b)|reversa\s+(de\s+|a\s+la\s+)?renovaci|numero\s+de\s+credito\s*[\(/]\s*renovaci|no\.?\s*de\s+renovaci|numero\s+de\s+renovaci)',
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
            r'\b(renovaci\w*|renovad\w*|renovo|renovar|ren)\b|credito\s+padre|credito\s+hijo|credito\s+original|credito\s+anterior|antes\s+de\s+renovar|operacion\s*\(\s*origen\s*\)',
            combined
        ))
        has_cancel_word = bool(re.search(
            r'\b(cancel\w*|reversa\w*|sin\s+efecto|baja)\b',
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
        for tbl in inner_tables:
            if len(tbl) < 2:
                continue

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
            after_hdr = re.split(r'Renovado\s+Sucursal\s+Fecha', plain_desc, flags=re.I)[-1]
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

        clean_desc = cls._clean_user_text_for_extraction(descripcion)
        full_context = f"{resumen or ''}\n{clean_desc or ''}"

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
            r'(?:n[úu]mero|num\.?|no\.?|#)\s*(?:de\s+)?(?:cliente|trabajador(?:a)?|fonacot|credere|cuenta)\s*[:\.\-#]*\s*(\d{5,10})\b',
            r'\b(?:cliente|trabajador(?:a)?|credere)\s*(?:no\.?|num\.?|#|:|\.)?\s*(?:[A-ZÁÉÍÓÚÑ]{2,20}\s+){0,4}[#:\.]?\s*(\d{5,10})\b',
            r'\bT\s*[\-\.]\s*(\d{5,10})\b',
            r'/\s*T\s+(\d{5,10})\b',
            r'[\+\-]\s*(\d{6,10})\s*[\+\-]\s*\d{4,7}\b',
            r'\b(?:cancelaci[óo]n\s+(?:de\s+)?renovaci(?:[óo]n|ones)|cancelaci[óo]n\s+(?:de\s+)?cr[ée]dito)\s*[\+\-:]?\s*(\d{6,10})\b',
            r'\bid\s+(\d{6,8})\s+cr[ée]dito\s+\d{4,7}\b'
        ]
        for pat in cliente_patterns:
            for m in re.finditer(pat, full_context, re.I):
                c_num = m.group(1).strip()
                # Excluir NSS (11 dígitos) y números de solicitud (empiezan con 507 o 508 de 9 dígitos)
                if len(c_num) == 9 and c_num.startswith(('507', '508')):
                    continue
                if cls._is_valid_number_token(c_num, 5, 10) and c_num not in clientes_found and c_num not in creditos_padre_found:
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
            r'\b(?:cr[ée]ditos?|renovaci(?:[óo]n|ones)|ren)\s*(?:(?:n[úu]mero|num\.?|no\.?|n°)\s*)?[:#\.]?\s*([C\-\d]+(?:\s*(?:,|y|Y|\-|\+)\s*[C\-\d]+)+)',
            r'\b(?:cr[ée]ditos?|renovaci(?:[óo]n|ones)|ren)\s*(?:(?:n[úu]mero|num\.?|no\.?|n°)\s*)?[:#\.]?\s*[C\-]?\s*(\d{3,7})\b',
            r'\bse\s+renov[óo]\s+al\s+(\d{4,7})\b'
        ]

        for pat in multi_credit_label_patterns:
            for m in re.finditer(pat, full_context, re.I):
                chunk = m.group(1)
                # Cortar el chunk si empieza otra etiqueta ("Motivo", "Gafete", "Cliente", "Trabajador", "Fecha", "Sucursal")
                chunk = re.split(
                    r'\b(motivo|gafete|cliente|trabajador|fecha|sucursal|estado|solicitud|nombre|antes|original|anterior)\b',
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

        # Patrón en Asunto/Resumen con signos '+' o '-': "+ 7239770+50074", "-8220974-60117", "+ 2431191 + 20757 + 20758"
        for m_plus in re.finditer(
            r'[\+\-]\s*(\d{5,10})\s*[\+\-]\s*(\d{3,7})\b(?:\s*(?:[\+\-,]|y)\s*(\d{3,7})\b)?',
            full_context,
            re.I
        ):
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
                n for n in re.findall(r'\b\d{4,10}\b', resumen or '')
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
            if len(creditos_renovados_found) == len(clientes_found):
                for idx, cli_str in enumerate(clientes_found):
                    pad_str = creditos_padre_found[idx] if idx < len(creditos_padre_found) else ""
                    disaggregated.append({
                        'cliente': cli_str,
                        'credito_renovado': creditos_renovados_found[idx],
                        'credito_padre': pad_str,
                        'sucursal': '',
                        'fecha_op': '',
                        'origen_extraccion': 'Múltiples Clientes en Ticket'
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
                        'origen_extraccion': 'Múltiples Clientes y Créditos'
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
                            'discrepancia_cliente': True
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
                        'origen_extraccion': 'Múltiples Clientes en Ticket'
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
                'Row_ID': f"{ticket_num}_{idx}_{cli}_{cred_ren}",
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
