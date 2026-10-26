#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Motor MAPI de Microsoft Outlook para el Validador y Concentrador Diario de Renovaciones FONACOT.
Incluye:
- Escaneo ultrarrápido multi-buzón (Stores / PST / Subcarpetas / Elementos eliminados)
- Regla de Nancy (2609) y detección de reasignaciones entre compañeros
- Modo Individual por Agente (Gafete o Nombre) y Modo Concentrador General de Cierre Diario
- Desglose automático 1 Ticket -> N Renovaciones (Cliente + Crédito Padre + Crédito Hijo/Renovado)
- Detección de renovaciones repetidas en distintos tickets (mismo Cliente + mismo Crédito)
- Exportación ejecutiva a Excel con hipervínculos directos a iTop
"""

import os
import re
import json
import shutil
import tempfile
from datetime import datetime, timedelta
from tkinter import filedialog
import win32com.client
import pythoncom
import pywintypes
import win32timezone  # Indispensable para fechas COM en ejecutable PyInstaller
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from core_renovaciones import (
    get_app_dir,
    normalize_text,
    make_soup,
    build_itop_url,
    parse_hora_corte,
    parse_fecha_asunto,
    previous_business_day,
    extract_concentrado_tokens,
    extract_concentrado_rows,
    extract_concentrado_intro,
    clasificar_lotes,
    ticket_key,
    is_aviso_solucion,
    parse_aviso_solucion,
    completar_con_imagenes,
    HORA_CORTE_DEFAULT,
    ASUNTO_CONCENTRADO_DEFAULT,
    RenovacionTicketParser
)

from ocr_imagenes import leer_correo, OCRNoDisponible

# Días hábiles hacia atrás que revisa el lote del día (lo no enviado sigue apareciendo una semana)
DIAS_ATRAS_LOTE = 5


class FonacotRenovacionesEngine:
    """Motor principal para validación y concentrado diario de cancelaciones de renovación."""

    SKIP_SYSTEM_FOLDERS = {
        'calendario', 'calendar', 'contactos', 'contacts', 'tareas', 'tasks',
        'notas', 'notes', 'diario', 'journal', 'archivos de datos', 'archivos',
        'problemas de sincronización', 'sync issues', 'errores locales',
        'local failures', 'conflictos', 'conflicts', 'errores del servidor',
        'server failures', 'rss feeds', 'fuentes rss', 'canales rss', 'suscripciones de rss',
        'suministro de noticias', 'working set', 'eventcheckpoints',
        'correo no deseado', 'junk email', 'junk', 'spam',
        'historial de conversaciones', 'conversation history', 'chat de equipo',
        'raíz de yammer', 'raiz de yammer', 'yammer root', 'inbound', 'outbound', 'feeds',
        'configuración de pasos rápidos', 'quick step settings',
        'conversation action settings', 'configuración de acciones de conversación',
        'externalcontacts', 'files', 'personmetadata',
        'social activity notifications', 'carpetas de búsqueda', 'search folders',
        'bandeja de salida', 'outbox', 'borradores', 'drafts', 'borrador'
    }

    def __init__(self):
        self.app_dir = get_app_dir()
        self.settings_file = os.path.join(self.app_dir, 'renovaciones_settings.json')
        self.settings = self.load_settings()

        self.outlook = None
        self.namespace = None
        self.current_user_email = "Usuario Outlook"
        self.available_folders_map = {}

        self.renovaciones_rows = []
        self.tickets_summary = []
        self.lote_info = None
        self.ocr_info = {'completadas': 0, 'sin_leer': [], 'error': ''}

    def load_settings(self) -> dict:
        default_settings = {
            'last_gafetes': '',
            'modo_concentrador_general': True,
            'incluir_subcarpetas': True,
            'incluir_eliminados': True,
            'ignorar_reasignados_a_otro': False,
            'ocultar_creditos_duplicados': False,
            'appearance_mode': 'Light'
        }
        if os.path.exists(self.settings_file):
            try:
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    saved = json.load(f)
                    default_settings.update(saved)
            except Exception:
                pass
        return default_settings

    def save_settings(self):
        try:
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(self.settings, f, indent=2, ensure_ascii=False)
        except Exception:
            pass

    def connect_outlook(self) -> bool:
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass

        try:
            try:
                self.outlook = win32com.client.GetActiveObject("Outlook.Application")
            except Exception:
                self.outlook = win32com.client.Dispatch("Outlook.Application")

            self.namespace = self.outlook.GetNamespace("MAPI")
            self.namespace.Logon("", "", False, False)

            detected_email = None
            try:
                accounts = self.namespace.Accounts
                for i in range(1, accounts.Count + 1):
                    acc = accounts.Item(i)
                    smtp = getattr(acc, 'SmtpAddress', '') or getattr(acc, 'DisplayName', '')
                    if 'fonacot' in smtp.lower():
                        detected_email = smtp
                        break
                if not detected_email and accounts.Count > 0:
                    detected_email = getattr(accounts.Item(1), 'SmtpAddress', '') or getattr(accounts.Item(1), 'DisplayName', '')
            except Exception:
                pass

            if not detected_email:
                try:
                    if self.namespace.Folders.Count > 0:
                        detected_email = str(self.namespace.Folders.Item(1).Name)
                except Exception:
                    pass

            if detected_email:
                self.current_user_email = detected_email
            return True
        except Exception:
            return False

    def _get_all_root_folders(self) -> list:
        if not self.namespace:
            self.connect_outlook()

        roots = []
        seen_ids = set()

        try:
            top_folders = self.namespace.Folders
            for i in range(1, top_folders.Count + 1):
                try:
                    rf = top_folders.Item(i)
                    name = str(getattr(rf, 'Name', f'Buzón {i}'))
                    if name.lower() in {'carpetas públicas', 'public folders'}:
                        continue
                    eid = str(getattr(rf, 'EntryID', name))
                    if eid not in seen_ids:
                        seen_ids.add(eid)
                        roots.append((name, rf))
                except Exception:
                    continue
        except Exception:
            pass

        try:
            stores = self.namespace.Stores
            for i in range(1, stores.Count + 1):
                try:
                    st = stores.Item(i)
                    rf = st.GetRootFolder()
                    name = str(getattr(st, 'DisplayName', '') or getattr(rf, 'Name', f'Almacén {i}'))
                    if name.lower() in {'carpetas públicas', 'public folders'}:
                        continue
                    eid = str(getattr(rf, 'EntryID', name))
                    if eid not in seen_ids:
                        seen_ids.add(eid)
                        roots.append((name, rf))
                except Exception:
                    continue
        except Exception:
            pass

        return roots

    def _collect_mail_folders_recursive(self, folder, parent_label="", max_depth=6, current_depth=0) -> list:
        results = []
        if not folder or current_depth > max_depth:
            return results

        try:
            folder_name = str(getattr(folder, 'Name', 'Carpeta')).strip()
            name_low = folder_name.lower()

            if name_low in self.SKIP_SYSTEM_FOLDERS:
                return results

            if current_depth > 0:
                try:
                    item_type = getattr(folder, 'DefaultItemType', 0)
                    if item_type != 0:
                        return results
                except Exception:
                    pass

            if current_depth == 0:
                full_label = parent_label if parent_label else folder_name
            else:
                full_label = f"{parent_label}/{folder_name}" if parent_label else folder_name
                results.append((full_label, folder))

            subfolders = getattr(folder, 'Folders', None)
            if subfolders and subfolders.Count > 0:
                for i in range(1, subfolders.Count + 1):
                    try:
                        sub = subfolders.Item(i)
                        results.extend(
                            self._collect_mail_folders_recursive(
                                sub,
                                parent_label=full_label,
                                max_depth=max_depth,
                                current_depth=current_depth + 1
                            )
                        )
                    except Exception:
                        continue
        except Exception:
            pass

        return results

    def list_all_user_mail_folders(self) -> list:
        if not self.connect_outlook():
            return []

        self.available_folders_map = {}
        options = [
            "🌐 Todas (Todos los buzones + Subcarpetas + Eliminados)",
            "📥 Bandeja de entrada (+ Todas sus subcarpetas)",
            "🗑️ Solo Elementos eliminados (+ Subcarpetas)"
        ]

        roots = self._get_all_root_folders()
        multi_store = len(roots) > 1

        for store_name, root_folder in roots:
            prefix = f"[{store_name}]" if multi_store else store_name
            all_in_store = self._collect_mail_folders_recursive(root_folder, parent_label=prefix, max_depth=6, current_depth=0)

            for full_label, f_obj in all_in_store:
                low_lbl = full_label.lower()
                if '/elementos enviados' in low_lbl or '/sent items' in low_lbl:
                    continue

                if 'eliminados' in low_lbl or 'deleted' in low_lbl:
                    display_key = f"🗑️ {full_label}"
                else:
                    display_key = f"📂 {full_label}"

                if display_key not in self.available_folders_map:
                    self.available_folders_map[display_key] = f_obj
                    options.append(display_key)

        return options

    def _resolve_folders_to_scan(self, selection: str, incluir_subcarpetas: bool, incluir_eliminados: bool) -> list:
        self.list_all_user_mail_folders()

        folders_to_scan = []
        seen_entry_ids = set()

        def add_folder(lbl, f_obj):
            try:
                eid = str(getattr(f_obj, 'EntryID', lbl))
            except Exception:
                eid = lbl
            if eid not in seen_entry_ids:
                seen_entry_ids.add(eid)
                folders_to_scan.append((lbl, f_obj))

        if selection.startswith("🌐"):
            for key, f_obj in self.available_folders_map.items():
                clean_lbl = re.sub(r'^[📂📁🗑️]\s*', '', key)
                low_lbl = clean_lbl.lower()
                is_deleted = ('eliminados' in low_lbl or 'deleted' in low_lbl)
                if is_deleted and not incluir_eliminados:
                    continue
                if not incluir_subcarpetas:
                    parts = clean_lbl.split('/')
                    if len(parts) > 2:
                        continue
                add_folder(clean_lbl, f_obj)

        elif selection.startswith("📥"):
            for key, f_obj in self.available_folders_map.items():
                clean_lbl = re.sub(r'^[📂📁🗑️]\s*', '', key)
                low_lbl = clean_lbl.lower()
                is_inbox_tree = ('bandeja de entrada' in low_lbl or 'inbox' in low_lbl)
                is_deleted = ('eliminados' in low_lbl or 'deleted' in low_lbl)

                if is_inbox_tree:
                    if not incluir_subcarpetas and not (low_lbl.endswith('bandeja de entrada') or low_lbl.endswith('inbox')):
                        continue
                    add_folder(clean_lbl, f_obj)
                elif is_deleted and incluir_eliminados:
                    add_folder(clean_lbl, f_obj)

        elif selection.startswith("🗑️ Solo"):
            for key, f_obj in self.available_folders_map.items():
                clean_lbl = re.sub(r'^[📂📁🗑️]\s*', '', key)
                low_lbl = clean_lbl.lower()
                if 'eliminados' in low_lbl or 'deleted' in low_lbl:
                    add_folder(clean_lbl, f_obj)

        else:
            target_f = self.available_folders_map.get(selection)
            if target_f:
                clean_name = re.sub(r'^[📂📁🗑️]\s*', '', selection)
                add_folder(clean_name, target_f)
                if incluir_subcarpetas:
                    for sub_lbl, sub_obj in self._collect_mail_folders_recursive(target_f, parent_label=clean_name, max_depth=4, current_depth=0):
                        add_folder(sub_lbl, sub_obj)

        return folders_to_scan

    @staticmethod
    def _to_naive_local_dt(outlook_dt) -> datetime:
        if outlook_dt is None:
            return None
        try:
            return datetime(
                outlook_dt.year,
                outlook_dt.month,
                outlook_dt.day,
                outlook_dt.hour,
                outlook_dt.minute,
                outlook_dt.second
            )
        except Exception:
            return None

    @staticmethod
    def _matches_agent_filter(gafete: str, agente: str, gafetes_filter: list) -> bool:
        if not gafetes_filter:
            return True
        g_clean = (gafete or "").strip().lower()
        ag_norm = normalize_text(agente or "")
        for token in gafetes_filter:
            t_clean = (token or "").strip().lower()
            if not t_clean:
                continue
            if t_clean == g_clean:
                return True
            t_norm = normalize_text(t_clean)
            if t_norm and (t_norm == g_clean or t_norm in ag_norm):
                return True
        return False

    def search_and_validate_renewals(
        self,
        fecha_inicio: datetime,
        fecha_fin: datetime,
        gafetes_filter: list = None,
        modo_concentrador_general: bool = False,
        carpeta_seleccionada: str = "🌐 Todas (Todos los buzones + Subcarpetas + Eliminados)",
        incluir_subcarpetas: bool = True,
        incluir_eliminados: bool = True,
        ignorar_reasignados_a_otro: bool = False,
        progress_callback=None,
        leer_imagenes: bool = True
    ) -> tuple:
        """
        Escanea Outlook, identifica todos los tickets de Cancelación/Reversa de Renovación
        y los desglosa en filas individuales (1 fila por cada renovación/crédito a cancelar).

        - Si `modo_concentrador_general` es True (o `gafetes_filter` está vacío):
          Consolida cada ticket una sola vez asignado a su agente responsable final
          (evitando duplicar el correo inicial de Nancy 2609 con la asignación al compañero).
        - Si se filtra por un gafete/nombre específico (ej. '2609', 'nancy', '7353', 'sammy'):
          Aplica la regla de Nancy (2609) y las alertas de reasignación entre compañeros.
        """
        if not self.connect_outlook():
            raise ConnectionError("No se pudo establecer conexión con Microsoft Outlook.")

        start_dt = datetime(fecha_inicio.year, fecha_inicio.month, fecha_inicio.day, 0, 0, 0)
        end_dt = datetime(fecha_fin.year, fecha_fin.month, fecha_fin.day, 23, 59, 59)
        extended_end_dt = end_dt + timedelta(hours=12)
        self.ocr_info = {'completadas': 0, 'sin_leer': [], 'error': ''}
        self._correos_por_ticket = {}  # todos los correos de cada incidente (las imágenes pueden venir en otro)

        effective_general_mode = modo_concentrador_general or not gafetes_filter

        folders_to_scan = self._resolve_folders_to_scan(
            selection=carpeta_seleccionada,
            incluir_subcarpetas=incluir_subcarpetas,
            incluir_eliminados=incluir_eliminados
        )

        # (ticket_num, gafete) -> candidate_meta
        raw_candidates_map = {}
        all_gafetes_by_ticket = {}  # ticket_num -> set(gafetes)
        assignments_history = {}    # ticket_num -> list of (rec_time, gafete, agente)
        avisos_solucion = {}        # ticket_key -> (rec_time, msg) del aviso de cierre más reciente

        total_folders = max(len(folders_to_scan), 1)
        scanned_folder_eids = set()

        for f_idx, (folder_label, folder_obj) in enumerate(folders_to_scan, 1):
            try:
                scanned_folder_eids.add(str(getattr(folder_obj, 'EntryID', folder_label)))
                items = folder_obj.Items
                total_in_folder = items.Count
                if total_in_folder == 0:
                    continue

                items.Sort("[ReceivedTime]", True)
            except Exception:
                continue

            scanned = 0
            consecutive_older = 0
            consecutive_errors = 0
            max_older_tolerance = 65

            if progress_callback:
                progress_callback(
                    f_idx,
                    total_folders + 1,
                    f"📂 Escaneando [{f_idx}/{total_folders}]: {folder_label} ({total_in_folder:,} correos)..."
                )

            for msg in items:
                scanned += 1
                if progress_callback and scanned % 40 == 0:
                    progress_callback(
                        f_idx,
                        total_folders + 1,
                        f"📂 [{folder_label}] Analizando correo {scanned:,} de {total_in_folder:,} (Tickets candidatos: {len(raw_candidates_map)})..."
                    )

                try:
                    rec_time = self._to_naive_local_dt(getattr(msg, 'ReceivedTime', None) or getattr(msg, 'CreationTime', None))
                    if not rec_time:
                        consecutive_errors += 1
                        if consecutive_errors > 40:
                            break
                        continue

                    consecutive_errors = 0

                    if rec_time < start_dt:
                        consecutive_older += 1
                        if consecutive_older > max_older_tolerance:
                            break
                        continue
                    else:
                        consecutive_older = 0

                    if getattr(msg, 'Class', 0) != 43:
                        continue

                    subject = str(getattr(msg, 'Subject', '') or '').strip()
                    if not subject:
                        continue

                    # Avisos de cierre de iTop: se toman aunque lleguen después del periodo (el ticket pudo cerrarse hoy)
                    t_cierre = is_aviso_solucion(subject)
                    if t_cierre:
                        if t_cierre not in avisos_solucion or rec_time > avisos_solucion[t_cierre][0]:
                            avisos_solucion[t_cierre] = (rec_time, msg)
                        continue

                    if rec_time > extended_end_dt:
                        continue

                    subj_low = subject.lower()
                    if 'i-' not in subj_low and 'ticket' not in subj_low:
                        continue
                    m_ticket = re.search(r'\b[IR]-\d{5,7}\b', subject, re.I)
                    if m_ticket:
                        correos = self._correos_por_ticket.setdefault(m_ticket.group(0).upper(), [])
                        if len(correos) < 8:
                            correos.append(msg)

                    is_assign, gafete, agente, ticket_num = RenovacionTicketParser.is_official_assignment_email(subject, "", "")
                    if not is_assign and ('asignado al agente' in subj_low or 'se te asign' in subj_low):
                        plain_fallback = str(getattr(msg, 'Body', '') or '')
                        is_assign, gafete, agente, ticket_num = RenovacionTicketParser.is_official_assignment_email(subject, plain_fallback, "")

                    if is_assign and ticket_num:
                        all_gafetes_by_ticket.setdefault(ticket_num, set()).add(gafete)
                        assignments_history.setdefault(ticket_num, []).append((rec_time, gafete, agente))

                        if rec_time <= end_dt:
                            if effective_general_mode or self._matches_agent_filter(gafete, agente, gafetes_filter):
                                map_key = (ticket_num, gafete)
                                if map_key not in raw_candidates_map:
                                    raw_candidates_map[map_key] = {
                                        '_msg': msg,
                                        'Asunto': subject,
                                        'Fecha_DT': rec_time,
                                        'Numero_Ticket': ticket_num,
                                        'Gafete': gafete,
                                        'Nombre_Asignado': agente,
                                        'Carpeta_Origen': folder_label
                                    }
                except Exception:
                    consecutive_errors += 1
                    if consecutive_errors > 40:
                        break
                    continue

        # Construir línea de tiempo limpia por cada folio
        clean_timeline_by_ticket = {}
        for t_num, raw_hist in assignments_history.items():
            sorted_hist = sorted(raw_hist, key=lambda x: x[0])
            compact_hist = []
            for dt_item, g_item, ag_item in sorted_hist:
                if not compact_hist or compact_hist[-1][1] != g_item:
                    compact_hist.append((dt_item, g_item, ag_item))
            clean_timeline_by_ticket[t_num] = compact_hist

        # Seleccionar qué candidatos sobreviven según el modo (Concentrador General vs Filtro por Agente)
        surviving_candidates = []

        if effective_general_mode:
            # En modo Concentrador General, queremos 1 solo registro por Numero_Ticket asignado al responsable final
            tickets_grouped = {}
            for cand in raw_candidates_map.values():
                t_num = cand['Numero_Ticket']
                tickets_grouped.setdefault(t_num, []).append(cand)

            for t_num, cands in tickets_grouped.items():
                timeline = clean_timeline_by_ticket.get(t_num, [])
                hist_str = " ➔ ".join(f"{dt.strftime('%d/%m %H:%M')} [{g} {ag}]" for dt, g, ag in timeline)
                non_nancy = [c for c in cands if c['Gafete'] != '2609']

                if non_nancy:
                    # Ordenar por fecha de asignación y tomar el último compañero asignado
                    non_nancy.sort(key=lambda x: x['Fecha_DT'])
                    chosen = non_nancy[-1]
                    distinct_non_nancy = list(dict.fromkeys(c['Gafete'] for c in non_nancy))
                    if len(distinct_non_nancy) > 1:
                        chosen['Requiere_Validacion_Agente'] = True
                        chosen['Alerta_Asignacion'] = f"⚠️ Reasignado entre compañeros ➔ Quedó con {chosen['Gafete']} {chosen['Nombre_Asignado']}"
                    else:
                        chosen['Requiere_Validacion_Agente'] = False
                        chosen['Alerta_Asignacion'] = f"✅ Asignado a {chosen['Gafete']} {chosen['Nombre_Asignado']}"
                else:
                    # Solo lo tiene Nancy (2609) -> Ella lo atendió directamente
                    chosen = cands[0]
                    chosen['Requiere_Validacion_Agente'] = False
                    chosen['Alerta_Asignacion'] = "✅ Atendido directamente por 2609 Nancy"

                chosen['Historial_Asignaciones'] = hist_str
                surviving_candidates.append(chosen)
        else:
            # Modo Individual por Agente (Gafete o Nombre)
            for cand in raw_candidates_map.values():
                t_num = cand['Numero_Ticket']
                my_gafete = cand['Gafete']
                timeline = clean_timeline_by_ticket.get(t_num, [])
                hist_str = " ➔ ".join(f"{dt.strftime('%d/%m %H:%M')} [{g} {ag}]" for dt, g, ag in timeline)
                cand['Historial_Asignaciones'] = hist_str

                # Regla Nancy (2609)
                if my_gafete == '2609':
                    other_gafetes = all_gafetes_by_ticket.get(t_num, set()) - {'2609'}
                    if other_gafetes:
                        continue
                    cand['Requiere_Validacion_Agente'] = False
                    cand['Reasignado_A_Otro'] = False
                    cand['Alerta_Asignacion'] = "✅ Atendido por 2609 (Sin reasignar)"
                    surviving_candidates.append(cand)
                    continue

                non_nancy = [(dt, g, ag) for dt, g, ag in timeline if g != '2609']
                distinct_non_nancy_gafetes = list(dict.fromkeys(g for _, g, _ in non_nancy))

                if len(distinct_non_nancy_gafetes) > 1 and non_nancy:
                    last_dt, last_g, last_ag = non_nancy[-1]
                    if last_g != my_gafete:
                        if ignorar_reasignados_a_otro:
                            continue
                        cand['Requiere_Validacion_Agente'] = True
                        cand['Reasignado_A_Otro'] = True
                        cand['Alerta_Asignacion'] = f"⚠️ REASIGNADO A: {last_g} {last_ag} ({last_dt.strftime('%d/%m %H:%M')}) — Validar si es tuyo"
                    else:
                        prev_agents = [f"{g} {ag}" for _, g, ag in non_nancy[:-1] if g != my_gafete]
                        prev_txt = ", ".join(dict.fromkeys(prev_agents))
                        cand['Requiere_Validacion_Agente'] = True
                        cand['Reasignado_A_Otro'] = False
                        cand['Alerta_Asignacion'] = f"⚠️ REASIGNADO DE: {prev_txt} ➔ Quedó contigo ({my_gafete})"
                else:
                    cand['Requiere_Validacion_Agente'] = False
                    cand['Reasignado_A_Otro'] = False
                    cand['Alerta_Asignacion'] = "✅ Asignación directa"

                surviving_candidates.append(cand)

        surviving_candidates.sort(key=lambda x: x['Fecha_DT'], reverse=True)

        # Ahora descargamos el HTMLBody/Body de los tickets candidatos y filtramos/desglosamos SOLO los de Renovación
        all_renewal_rows = []
        tickets_summary_list = []
        total_cands = len(surviving_candidates)

        for idx_c, cand in enumerate(surviving_candidates, 1):
            if progress_callback and (idx_c == 1 or idx_c % 5 == 0 or idx_c == total_cands):
                progress_callback(
                    total_folders,
                    total_folders + 1,
                    f"🧬 Analizando y desglosando renovaciones ({idx_c}/{total_cands} — {cand['Numero_Ticket']})..."
                )

            msg_obj = cand.get('_msg')
            plain_body = ""
            html_body = ""
            if msg_obj is not None:
                try:
                    plain_body = str(getattr(msg_obj, 'Body', '') or '')
                except Exception:
                    pass
                try:
                    html_body = str(getattr(msg_obj, 'HTMLBody', '') or '')
                except Exception:
                    pass

            disagg_rows = RenovacionTicketParser.parse_and_disaggregate_email(
                subject=cand['Asunto'],
                html_body=html_body,
                plain_body=plain_body,
                received_dt=cand['Fecha_DT'],
                ticket_num=cand['Numero_Ticket'],
                gafete=cand['Gafete'],
                nombre_asignado=cand['Nombre_Asignado'],
                carpeta_origen=cand['Carpeta_Origen']
            )

            if not disagg_rows:
                continue

            # Incidentes que mandan una captura de pantalla en vez de escribir Cliente y Crédito
            if leer_imagenes and msg_obj is not None and any(r['Datos_Incompletos'] for r in disagg_rows):
                if progress_callback:
                    progress_callback(total_folders, total_folders + 1,
                                      f"🖼️ Leyendo imágenes del incidente {cand['Numero_Ticket']}...")
                disagg_rows = self._completar_con_imagenes(msg_obj, disagg_rows, cand['Numero_Ticket'])

            # Hora en que el ticket llegó por primera vez (asignación inicial), no la de la última reasignación:
            # es la que decide si entró antes o después del corte.
            timeline = clean_timeline_by_ticket.get(cand['Numero_Ticket'], [])
            llegada_dt = timeline[0][0] if timeline else cand['Fecha_DT']

            for r in disagg_rows:
                r['Fecha_Llegada_DT'] = llegada_dt
                r['Requiere_Validacion_Agente'] = cand.get('Requiere_Validacion_Agente', False)
                r['Reasignado_A_Otro'] = cand.get('Reasignado_A_Otro', False)
                r['Alerta_Asignacion'] = cand.get('Alerta_Asignacion', '✅ Asignación directa')
                r['Historial_Asignaciones'] = cand.get('Historial_Asignaciones', '')
                all_renewal_rows.append(r)

            clientes_en_t = list(dict.fromkeys(r['Numero_Cliente'] for r in disagg_rows))
            creditos_en_t = list(dict.fromkeys(r['Credito_Renovado_Hijo'] for r in disagg_rows))
            padres_en_t = list(dict.fromkeys(r['Credito_Origen_Padre'] for r in disagg_rows if r['Credito_Origen_Padre'] != '-'))

            tickets_summary_list.append({
                'Fecha_DT': cand['Fecha_DT'],
                'Fecha': cand['Fecha_DT'].strftime('%d/%m/%Y %H:%M'),
                'Numero_Ticket': cand['Numero_Ticket'],
                'Gafete': cand['Gafete'],
                'Nombre_Asignado': cand['Nombre_Asignado'],
                'Cantidad_Renovaciones': len(disagg_rows),
                'Cantidad_Clientes': len(clientes_en_t),
                'Clientes_Lista': ", ".join(clientes_en_t),
                'Creditos_Renovados_Lista': ", ".join(creditos_en_t),
                'Creditos_Padre_Lista': ", ".join(padres_en_t) if padres_en_t else "-",
                'Subtipo_Renovacion': disagg_rows[0]['Subtipo_Renovacion'],
                'Resumen_iTop': disagg_rows[0]['Resumen_iTop'],
                'Alerta_Asignacion': cand.get('Alerta_Asignacion', ''),
                'Es_Multiple': len(disagg_rows) > 1,
                'Datos_Incompletos': any(r['Datos_Incompletos'] for r in disagg_rows)
            })

        # Detectar si un mismo (Numero_Cliente, Credito_Renovado_Hijo) fue reportado en más de 1 ticket distinto
        credit_key_tickets = {}
        for r in all_renewal_rows:
            c_cli = r['Numero_Cliente']
            c_ren = r['Credito_Renovado_Hijo']
            if c_cli != "SIN CLIENTE" and c_ren != "SIN CRÉDITO":
                credit_key_tickets.setdefault((c_cli, c_ren), []).append(r['Numero_Ticket'])

        for r in all_renewal_rows:
            c_cli = r['Numero_Cliente']
            c_ren = r['Credito_Renovado_Hijo']
            t_list = list(dict.fromkeys(credit_key_tickets.get((c_cli, c_ren), [])))
            if len(t_list) > 1:
                other_ts = [t for t in t_list if t != r['Numero_Ticket']]
                r['Ticket_Duplicado'] = True
                r['Otros_Tickets_Mismo_Credito'] = ", ".join(other_ts)
                r['Alerta_Renovacion'] = f"{r['Alerta_Renovacion']} | 🔄 Repetido en {', '.join(other_ts)}"
            else:
                r['Ticket_Duplicado'] = False
                r['Otros_Tickets_Mismo_Credito'] = ""

        self._aplicar_cierres_itop(all_renewal_rows, avisos_solucion, clean_timeline_by_ticket)

        self.renovaciones_rows = all_renewal_rows
        self.tickets_summary = tickets_summary_list
        return all_renewal_rows, tickets_summary_list

    def _completar_con_imagenes(self, msg, rows: list, ticket: str) -> list:
        """
        Lee con OCR las imágenes del correo y completa el dato que falta (ver core.completar_con_imagenes). Si en
        ese correo no hay imágenes, prueba con los demás correos del mismo incidente (la asignación original, la
        reasignación, la nota de iTop…). Si no se pudo, deja en 'Imagen_Info' por qué, para verlo en pantalla.
        """
        if self.ocr_info['error']:
            for r in rows:
                if r['Datos_Incompletos']:
                    r['Imagen_Info'] = f"No se pudieron leer las imágenes en este equipo: {self.ocr_info['error']}"
            return rows  # el OCR no está disponible en este equipo: no se reintenta en cada correo
        correos = [msg] + [m for m in getattr(self, '_correos_por_ticket', {}).get(ticket.upper(), []) if m is not msg]
        lecturas, n_imagenes, notas, error = [], 0, [], ''
        for correo in correos:
            try:
                leidas, n = leer_correo(correo, notas)
            except OCRNoDisponible as e:
                self.ocr_info['error'] = str(e)
                error = str(e)
                break
            except Exception as e:
                error = f"{type(e).__name__}: {e}"
                continue
            n_imagenes += n
            if leidas:
                lecturas = leidas
                break
        if lecturas:
            rows = completar_con_imagenes(rows, lecturas)
            self.ocr_info['completadas'] += sum(1 for r in rows if r.get('Dato_Imagen'))
        if any(r['Datos_Incompletos'] for r in rows):
            if n_imagenes:
                self.ocr_info['sin_leer'].append(ticket)
                info = (f"Se leyeron {n_imagenes} imagen(es) del incidente y no se encontró el No. de cliente ni el "
                        f"No. de crédito (revisa la imagen en el correo o en iTop)")
            elif error:
                info = f"No se pudieron leer las imágenes: {error}"
            else:
                info = (f"No se encontró ninguna imagen en {len(correos)} correo(s) del incidente"
                        + (f" ({'; '.join(dict.fromkeys(notas))})" if notas else "")
                        + ": si la captura solo está en iTop, revísala ahí")
            for r in rows:
                if r['Datos_Incompletos']:
                    r['Imagen_Info'] = info
        return rows

    @staticmethod
    def _aplicar_cierres_itop(rows: list, avisos_solucion: dict, timeline_by_ticket: dict):
        """
        Marca las renovaciones cuyo ticket ya tiene "Aviso de solución" de iTop: estado, motivo y quién lo cerró.
        Si el correo no dice quién, se toma el agente que tenía asignado el ticket al momento del cierre.
        """
        timelines = {ticket_key(t): tl for t, tl in timeline_by_ticket.items()}
        cierres = {}
        for r in rows:
            r['Cerrado_iTop'] = False
            r['Cierre_iTop'] = None
            t = ticket_key(r['Numero_Ticket'])
            if t not in avisos_solucion:
                continue
            if t not in cierres:
                rec_time, msg = avisos_solucion[t]
                try:
                    html = str(getattr(msg, 'HTMLBody', '') or '')
                except Exception:
                    html = ''
                try:
                    plain = str(getattr(msg, 'Body', '') or '')
                except Exception:
                    plain = ''
                cierre = parse_aviso_solucion(html, plain)
                cierre['fecha_dt'] = rec_time
                cierre['cerro_fuente'] = 'correo' if cierre.get('cerro') else ''
                if not cierre.get('cerro'):
                    previas = [(dt, g, ag) for dt, g, ag in timelines.get(t, []) if dt <= rec_time]
                    if previas:
                        _dt, g, ag = previas[-1]
                        cierre['cerro'] = f"{g} {ag}".strip()
                        cierre['cerro_fuente'] = 'asignacion'
                cierres[t] = cierre
            r['Cerrado_iTop'] = True
            r['Cierre_iTop'] = cierres[t]

    def export_to_excel(
        self,
        renovaciones_rows: list,
        tickets_summary: list = None,
        filename: str = None,
        parent_window=None
    ) -> str:
        """
        Exporta el reporte ejecutivo de Renovaciones a Excel (.xlsx) con 3 hojas:
        1. Desglose_Renovaciones_Cancelar (1 fila por cada crédito renovado a cancelar + link iTop)
        2. Concentrado_Limpio_CREDERE (Sin duplicados de Cliente+Crédito, listo para envío de fin de día)
        3. Resumen_Por_Ticket_y_Agente (Comparativa Tickets recibidos vs Total Renovaciones reales)
        """
        if not renovaciones_rows:
            return ""

        if not filename:
            fecha_archivo = datetime.now().strftime('%d-%m-%Y_%H-%M-%S')
            dialog_kwargs = {
                'defaultextension': ".xlsx",
                'filetypes': [("Archivos de Excel", "*.xlsx"), ("Todos los archivos", "*.*")],
                'title': "Guardar Reporte Validador y Concentrador de Renovaciones FONACOT",
                'initialfile': f"Concentrado_Renovaciones_FONACOT_{fecha_archivo}.xlsx"
            }
            if parent_window is not None:
                dialog_kwargs['parent'] = parent_window
            filename = filedialog.asksaveasfilename(**dialog_kwargs)

        if not filename:
            return ""

        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            # HOJA 1: Desglose completo de renovaciones
            df_desglose = pd.DataFrame(renovaciones_rows)
            df_desglose['Estado_Check'] = df_desglose['Validado_Check'].apply(lambda v: "✅ VALIDADO" if v else "⏳ Pendiente")
            df_desglose['Contador_Ticket'] = df_desglose.apply(
                lambda r: f"{r['Indice_En_Ticket']} de {r['Total_En_Ticket']}", axis=1
            )
            cols_h1 = [
                'Estado_Check', 'Fecha', 'Numero_Ticket', 'Contador_Ticket', 'Alerta_Renovacion',
                'Numero_Cliente', 'Credito_Renovado_Hijo', 'Credito_Origen_Padre',
                'Gafete', 'Nombre_Asignado', 'Alerta_Asignacion',
                'Subtipo_Renovacion', 'Motivo', 'Sucursal', 'Fecha_Operacion',
                'Resumen_iTop', 'Carpeta_Origen'
            ]
            for c in cols_h1:
                if c not in df_desglose.columns:
                    df_desglose[c] = ""
            df_h1 = df_desglose[cols_h1].copy()
            df_h1.columns = [
                '¿Validado?', 'Fecha Recepción', 'Ticket iTop', 'Renovación # en Ticket', 'Alerta de Desglose',
                'No. Cliente (Cuenta)', 'Crédito Renovado / Hijo (Cancelar)', 'Crédito Origen / Padre',
                'Gafete', 'Agente Asignado', 'Validación Asignación',
                'Tipo de Solicitud', 'Motivo Cancelación', 'Sucursal / Localidad', 'Fecha Operación',
                'Resumen del Ticket', 'Carpeta Outlook'
            ]
            df_h1.to_excel(writer, sheet_name='Desglose_Renovaciones', index=False)

            # HOJA 2: Concentrado Limpio Único por (No. Cliente + Crédito Renovado) para envío de cierre del día
            unique_credits_rows = []
            seen_credit_keys = set()
            for r in renovaciones_rows:
                c_cli = r.get('Numero_Cliente', '')
                c_ren = r.get('Credito_Renovado_Hijo', '')
                dedup_key = (c_cli, c_ren) if (c_cli != 'SIN CLIENTE' and c_ren != 'SIN CRÉDITO') else r.get('Row_ID')
                if dedup_key not in seen_credit_keys:
                    seen_credit_keys.add(dedup_key)
                    unique_credits_rows.append({
                        'No. Cliente (Cuenta)': c_cli,
                        'Crédito Origen (Padre)': r.get('Credito_Origen_Padre', '-'),
                        'Crédito Renovado (Hijo a Cancelar)': c_ren,
                        'Sucursal / Localidad': r.get('Sucursal', ''),
                        'Fecha Operación / Correo': r.get('Fecha_Operacion') or r.get('Fecha', '')[:10],
                        'Ticket iTop': r.get('Numero_Ticket', ''),
                        'Gafete Responsable': r.get('Gafete', ''),
                        'Agente Responsable': r.get('Nombre_Asignado', ''),
                        'Motivo': r.get('Motivo', ''),
                        'Alerta': r.get('Alerta_Renovacion', '')
                    })
            if unique_credits_rows:
                df_h2 = pd.DataFrame(unique_credits_rows)
                df_h2.to_excel(writer, sheet_name='Concentrado_Cierre_Dia', index=False)

            # HOJA 3: Resumen por Agente y por Ticket
            if tickets_summary:
                df_h3 = pd.DataFrame(tickets_summary)
                cols_h3 = [
                    'Fecha', 'Numero_Ticket', 'Gafete', 'Nombre_Asignado',
                    'Cantidad_Renovaciones', 'Cantidad_Clientes',
                    'Clientes_Lista', 'Creditos_Renovados_Lista', 'Creditos_Padre_Lista',
                    'Subtipo_Renovacion', 'Resumen_iTop', 'Alerta_Asignacion'
                ]
                for c in cols_h3:
                    if c not in df_h3.columns:
                        df_h3[c] = ""
                df_h3 = df_h3[cols_h3].copy()
                df_h3.columns = [
                    'Fecha Asignación', 'Ticket iTop', 'Gafete', 'Agente Asignado',
                    'Total Renovaciones en Ticket', 'Total Clientes en Ticket',
                    'Números de Cliente', 'Créditos Renovados (Hijos)', 'Créditos Origen (Padres)',
                    'Tipo', 'Resumen del Ticket', 'Estado de Asignación'
                ]
                df_h3.to_excel(writer, sheet_name='Resumen_Por_Ticket', index=False)

            # Resumen Estadístico por Compañero
            agent_stats = {}
            for r in renovaciones_rows:
                ag_key = f"{r.get('Gafete', '')} - {r.get('Nombre_Asignado', '')}"
                st = agent_stats.setdefault(ag_key, {
                    'Agente': ag_key,
                    'Tickets de Renovación': set(),
                    'Total Créditos Renovados a Cancelar': 0,
                    'Tickets con Múltiples Renovaciones': set(),
                    'Registros con Datos Incompletos': 0
                })
                st['Tickets de Renovación'].add(r['Numero_Ticket'])
                st['Total Créditos Renovados a Cancelar'] += 1
                if r.get('Es_Multiple'):
                    st['Tickets con Múltiples Renovaciones'].add(r['Numero_Ticket'])
                if r.get('Datos_Incompletos'):
                    st['Registros con Datos Incompletos'] += 1

            if agent_stats:
                stats_rows = []
                for st in agent_stats.values():
                    stats_rows.append({
                        'Gafete y Agente': st['Agente'],
                        'Tickets de Renovación Asignados': len(st['Tickets de Renovación']),
                        'Total Créditos Renovados a Cancelar': st['Total Créditos Renovados a Cancelar'],
                        'Diferencia (Renovaciones Extra Ocultas)': st['Total Créditos Renovados a Cancelar'] - len(st['Tickets de Renovación']),
                        'Tickets con 2 o más Renovaciones': len(st['Tickets con Múltiples Renovaciones']),
                        'Alertas por Datos Faltantes': st['Registros con Datos Incompletos']
                    })
                df_h4 = pd.DataFrame(stats_rows)
                df_h4.to_excel(writer, sheet_name='Totales_Por_Agente', index=False)

            # Estilos profesionales e hipervínculos a iTop
            header_font = Font(name="Segoe UI", bold=True, color="FFFFFF", size=11)
            header_fill = PatternFill(start_color="1A365D", end_color="1A365D", fill_type="solid")
            header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
            thin_border = Border(
                left=Side(style='thin', color='CBD5E0'),
                right=Side(style='thin', color='CBD5E0'),
                top=Side(style='thin', color='CBD5E0'),
                bottom=Side(style='thin', color='CBD5E0')
            )
            font_link = Font(name="Segoe UI", size=10, color="0000FF", underline="single")

            for sheet_name, worksheet in writer.sheets.items():
                worksheet.row_dimensions[1].height = 28
                for col_num in range(1, worksheet.max_column + 1):
                    cell = worksheet.cell(row=1, column=col_num)
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = header_align

                ticket_col_idx = None
                alerta_col_idx = None
                for c_idx in range(1, worksheet.max_column + 1):
                    hdr_val = str(worksheet.cell(row=1, column=c_idx).value or '').strip()
                    if hdr_val == 'Ticket iTop':
                        ticket_col_idx = c_idx
                    elif hdr_val in ('Alerta de Desglose', 'Alerta', 'Total Renovaciones en Ticket'):
                        alerta_col_idx = c_idx

                for row_num in range(2, worksheet.max_row + 1):
                    worksheet.row_dimensions[row_num].height = 22
                    even_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid") if row_num % 2 == 0 else None

                    row_alert_txt = str(worksheet.cell(row=row_num, column=alerta_col_idx).value or '') if alerta_col_idx else ''
                    row_special_fill = None
                    if 'FALTA DATO' in row_alert_txt or 'SIN CRÉDITO' in row_alert_txt:
                        row_special_fill = PatternFill(start_color="FED7D7", end_color="FED7D7", fill_type="solid")
                    elif 'MÚLTIPLE' in row_alert_txt or (alerta_col_idx and row_alert_txt.isdigit() and int(row_alert_txt) > 1):
                        row_special_fill = PatternFill(start_color="FEEBC8", end_color="FEEBC8", fill_type="solid")

                    for col_num in range(1, worksheet.max_column + 1):
                        c = worksheet.cell(row=row_num, column=col_num)
                        c.font = Font(name="Segoe UI", size=10)
                        c.border = thin_border
                        if row_special_fill:
                            c.fill = row_special_fill
                        elif even_fill:
                            c.fill = even_fill

                    if ticket_col_idx:
                        t_cell = worksheet.cell(row=row_num, column=ticket_col_idx)
                        t_val = str(t_cell.value or '').strip()
                        if t_val.startswith('I-'):
                            t_cell.hyperlink = build_itop_url(t_val, self.settings.get('itop_url_template', ''))
                            t_cell.font = font_link

                for col_idx in range(1, worksheet.max_column + 1):
                    col_letter = get_column_letter(col_idx)
                    max_len = 0
                    for row_idx in range(1, worksheet.max_row + 1):
                        val = str(worksheet.cell(row=row_idx, column=col_idx).value or "")
                        if len(val) > max_len:
                            max_len = len(val)
                    worksheet.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 46)

                if worksheet.max_row >= 1 and worksheet.max_column >= 1:
                    last_col_letter = get_column_letter(worksheet.max_column)
                    worksheet.auto_filter.ref = f"A1:{last_col_letter}{worksheet.max_row}"
                worksheet.freeze_panes = "A2"

        return filename

    def get_accounts(self) -> list:
        """Retorna la lista de carpetas/opciones de buzón disponibles en Outlook."""
        return self.list_all_user_mail_folders()

    @staticmethod
    def create_draft_email(subject: str, html_fragment: str, to: str = "", cc: str = "") -> None:
        """
        Abre en Outlook un correo NUEVO (borrador, no se envía) con la tabla del lote al inicio del cuerpo,
        conservando la firma predeterminada del usuario. Se crea una conexión COM propia porque se llama
        desde el hilo de la interfaz y la del motor pertenece al hilo de extracción.
        """
        from lote_utils import insert_into_html_body
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.Subject = subject
        if to:
            mail.To = to
        if cc:
            mail.CC = cc
        mail.Display()  # Al mostrarlo, Outlook agrega la firma al HTMLBody
        mail.HTMLBody = insert_into_html_body(str(mail.HTMLBody or ''), html_fragment)

    @staticmethod
    def _attachments_text(msg) -> str:
        return FonacotRenovacionesEngine._attachments_content(msg)[0]

    @staticmethod
    def _attachments_content(msg) -> tuple:
        """
        (texto, html) de los adjuntos Excel/CSV del concentrado (a veces la relación viene en un .xlsx).
        El html trae cada hoja como tabla para leerla renglón por renglón igual que la tabla del correo.
        """
        chunks, tablas = [], []
        try:
            attachments = msg.Attachments
            count = attachments.Count
        except Exception:
            return "", ""
        if not count:
            return "", ""
        tmp_dir = tempfile.mkdtemp(prefix="fonacot_conc_")
        try:
            for i in range(1, count + 1):
                try:
                    att = attachments.Item(i)
                    name = str(getattr(att, 'FileName', '') or '')
                    ext = os.path.splitext(name)[1].lower()
                    if ext not in ('.xlsx', '.xlsm', '.xls', '.csv', '.txt'):
                        continue
                    path = os.path.join(tmp_dir, f"{i}{ext}")
                    att.SaveAsFile(path)
                    if ext in ('.csv', '.txt'):
                        with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                            chunks.append(f.read())
                        try:
                            df = pd.read_csv(path, header=None, dtype=str, sep=None, engine='python', encoding_errors='ignore')
                            tablas.append(df.fillna('').to_html(header=False, index=False))
                        except Exception:
                            pass
                    else:
                        sheets = pd.read_excel(path, sheet_name=None, header=None, dtype=str)
                        for df in sheets.values():
                            chunks.append(df.fillna('').to_csv(sep=' ', index=False, header=False))
                            tablas.append(df.fillna('').to_html(header=False, index=False))
                except Exception:
                    continue
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        return "\n".join(chunks), "\n".join(tablas)

    def find_concentrado_emails(
        self,
        start_dt: datetime,
        end_dt: datetime,
        asunto_clave: str = ASUNTO_CONCENTRADO_DEFAULT,
        progress_callback=None
    ) -> list:
        """
        Busca en TODOS los buzones (incluye Elementos enviados) los correos cuyo asunto contiene `asunto_clave`
        ("ticket unico", sin importar acentos/mayúsculas) recibidos/enviados entre start_dt y end_dt.
        Retorna: [{'asunto', 'fecha_asunto', 'recibido', 'remitente', 'carpeta', 'tickets', 'numeros'}]
        """
        if not self.connect_outlook():
            return []
        clave = normalize_text(asunto_clave or ASUNTO_CONCENTRADO_DEFAULT)
        if not clave:
            return []

        folders = []
        seen = set()
        for store_name, root in self._get_all_root_folders():
            for lbl, f_obj in self._collect_mail_folders_recursive(root, parent_label=store_name, max_depth=6, current_depth=0):
                try:
                    eid = str(getattr(f_obj, 'EntryID', lbl))
                except Exception:
                    eid = lbl
                if eid not in seen:
                    seen.add(eid)
                    folders.append((lbl, f_obj))

        found = {}
        for f_idx, (lbl, f_obj) in enumerate(folders, 1):
            if progress_callback and f_idx % 5 == 1:
                progress_callback(f"📨 Buscando concentrado '{asunto_clave}' en {lbl}...")
            try:
                items = f_obj.Items
                if items.Count == 0:
                    continue
                items.Sort("[ReceivedTime]", True)
            except Exception:
                continue

            consecutive_older = 0
            for msg in items:
                try:
                    rec = self._to_naive_local_dt(
                        getattr(msg, 'ReceivedTime', None) or getattr(msg, 'SentOn', None) or getattr(msg, 'CreationTime', None)
                    )
                    if not rec:
                        continue
                    if rec < start_dt:
                        consecutive_older += 1
                        if consecutive_older > 30:
                            break
                        continue
                    consecutive_older = 0
                    if rec > end_dt or getattr(msg, 'Class', 0) != 43:
                        continue
                    subject = str(getattr(msg, 'Subject', '') or '').strip()
                    if clave not in normalize_text(subject):
                        continue

                    # El mismo correo puede estar en Enviados y en la Bandeja de varios compañeros
                    # (con uno o dos minutos de diferencia): mismo asunto a menos de 5 min = mismo correo
                    asunto_norm = normalize_text(subject)
                    if any(k[0] == asunto_norm and abs((k[1] - rec).total_seconds()) <= 300 for k in found):
                        continue
                    dedup_key = (asunto_norm, rec)

                    body = str(getattr(msg, 'Body', '') or '')
                    try:
                        html = str(getattr(msg, 'HTMLBody', '') or '')
                        html_txt = make_soup(html).get_text(" ", strip=True) if html else ""
                    except Exception:
                        html_txt = ""
                    adj_texto, adj_html = self._attachments_content(msg)
                    full_text = f"{body}\n{html_txt}\n{adj_texto}"
                    tickets, numeros = extract_concentrado_tokens(full_text)
                    filas, vistas = [], set()
                    for f in extract_concentrado_rows(html) + extract_concentrado_rows(adj_html):
                        # Otro nombre: `clave` es la palabra del asunto y se sigue usando con los demás correos
                        llave_fila = (f['ticket'], f['cliente'], f['credito'])
                        if llave_fila not in vistas:  # la misma relación en el cuerpo y en el adjunto
                            vistas.add(llave_fila)
                            filas.append(f)
                    tickets.update(f['ticket'].split('-', 1)[-1] for f in filas if f['ticket'])

                    found[dedup_key] = {
                        'asunto': subject,
                        'fecha_asunto': parse_fecha_asunto(subject, ref_year=rec.year),
                        'recibido': rec,
                        'remitente': str(getattr(msg, 'SenderName', '') or ''),
                        'carpeta': lbl,
                        'tickets': tickets,
                        'numeros': numeros,
                        'filas': filas,
                        'texto': full_text,  # para los renglones escritos fuera de la tabla
                        'intro': extract_concentrado_intro(body),
                    }
                except Exception:
                    continue

        return sorted(found.values(), key=lambda c: c['recibido'])

    def extract_renewals(
        self,
        account_name: str,
        start_dt: datetime,
        end_dt: datetime,
        mode: str = "CONCENTRADO",
        agent_filter: str = "",
        progress_callback=None,
        aplicar_corte: bool = True,
        hora_corte: str = HORA_CORTE_DEFAULT,
        asunto_concentrado: str = ASUNTO_CONCENTRADO_DEFAULT,
        ultimo_lote: dict = None,
        revision: bool = None,
        dias_atras: int = DIAS_ATRAS_LOTE,
        leer_imagenes: bool = True
    ) -> list:
        """
        Envoltorio compatible con la GUI principal.
        `mode`: "CONCENTRADO" (todos los compañeros consolidados) o "AGENTE" (por Gafete o Nombre).
        Con `aplicar_corte`, escanea desde el día hábil anterior, busca el correo concentrado ("Ticket Único")
        y clasifica cada renovación en lote de hoy / rezagada / ya enviada / siguiente lote (ver clasificar_lotes).
        El resumen queda en `self.lote_info`.
        """
        modo_general = (mode == "CONCENTRADO")
        gafetes_list = []
        if not modo_general and agent_filter:
            gafetes_list = [tok.strip() for tok in re.split(r'[,;]+', agent_filter) if tok.strip()]

        carpeta_sel = account_name if account_name else "🌐 Todas (Todos los buzones + Subcarpetas + Eliminados)"

        def _wrapped_progress(curr, total, msg):
            if progress_callback:
                frac = (curr / total) if total else 0.5
                progress_callback(msg, frac)

        scan_start = start_dt
        if aplicar_corte:
            # Lote del día: se revisan varios días hábiles atrás para que lo que no se envió no se pierda después de
            # un día. Revisión: basta el día anterior al periodo (por lo que llegó tras el corte)
            dia = start_dt.date()
            for _ in range(1 if revision else max(1, int(dias_atras))):
                dia = previous_business_day(dia)
            scan_start = datetime.combine(dia, datetime.min.time())

        self.lote_info = None
        rows, _summary = self.search_and_validate_renewals(
            fecha_inicio=scan_start,
            fecha_fin=end_dt,
            gafetes_filter=gafetes_list,
            modo_concentrador_general=modo_general,
            carpeta_seleccionada=carpeta_sel,
            incluir_subcarpetas=True,
            incluir_eliminados=True,
            ignorar_reasignados_a_otro=False,
            progress_callback=_wrapped_progress,
            leer_imagenes=leer_imagenes
        )

        if aplicar_corte:
            if progress_callback:
                progress_callback(f"📨 Buscando correos '{asunto_concentrado}' desde {scan_start.strftime('%d/%m/%Y')}...", 0.95)
            # El concentrado de hoy puede enviarse ya entrada la noche: se busca hasta el fin del día "hasta" + 12 h.
            # Al revisar días pasados se busca hasta hoy: lo que se mandó días después cuenta como enviado (tarde).
            fin_busqueda = max(datetime.combine(end_dt.date(), datetime.max.time()) + timedelta(hours=12), datetime.now())
            concentrados = self.find_concentrado_emails(
                scan_start,
                fin_busqueda,
                asunto_concentrado,
                progress_callback=(lambda m: progress_callback(m, 0.97)) if progress_callback else None
            )
            # Lote del día de una fecha pasada: se arma como debía salir ese día
            dia_lote = end_dt.date() if revision is False and end_dt.date() < datetime.now().date() else None
            self.lote_info = clasificar_lotes(
                rows,
                fecha_inicio=start_dt.date(),
                fecha_fin=end_dt.date(),
                hora_corte=parse_hora_corte(hora_corte),
                concentrados=concentrados,
                ultimo_lote=ultimo_lote,
                ahora=datetime.now(),
                revision=revision,
                dia_lote=dia_lote
            )
        return rows


# Alias de compatibilidad
OutlookRenovacionesEngine = FonacotRenovacionesEngine


