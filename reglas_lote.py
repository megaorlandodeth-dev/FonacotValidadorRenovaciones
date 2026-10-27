#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Qué se manda en el ticket único y cómo se explica cada renovación (sin interfaz: se puede probar solo).

Reglas del lote (después de clasificar_lotes, la base de datos y el calendario de emisión):
- Ya pasó la generación de la emisión (o ya está en CR_CEDULAS_DETALLE): ya se emitió, no se puede cancelar. Sale del
  lote y queda solo como informativa.
- Pasó el cierre de mes: sigue en el lote, pero solo se cancela con autorización.
- El incidente ya se resolvió en iTop: no se manda, salvo que se haya cerrado por error (se cerró diciendo que es de
  otro día, la base dice que es del mismo día, no se mandó y la renovación no está bloqueada: TZ_LOCK = 0).
- La base dice «No procede, la renovación es de otro día»: no se manda.

Desglose del ticket: un veredicto (Lista, Revisar, Corregir, No se manda, Ya enviada…), qué hacer, y las
comprobaciones con la evidencia de cada una (correo, base de datos, calendario, ticket único e iTop).
"""

import re
import unicodedata

from core_renovaciones import (
    LOTE_HOY, LOTE_REZAGADA, LOTE_ENVIADA, LOTE_EN_CONCENTRADO, LOTE_SIGUIENTE, LOTE_CERRADA_SIN_ENVIO,
    LOTE_SIN_COMPROBAR,
)
from sql_renovaciones import (
    SQL_OK, SQL_MESA_CREDITO, SQL_SE_ENVIA, SQL_SUGERENCIA, SQL_SIN_DATOS, SQL_ES_PADRE, SQL_PADRE_CORREGIDO, SQL_NO_PROCEDE, SQL_PROBLEMAS,
    SQL_NO_EXISTE, SQL_SIN_RENOVACIONES, SQL_OTRA_FECHA,
    LEYENDA_NO_PROCEDE,
)
from lote_utils import pl

# Estados de lote que no van en el ticket único de hoy (se ven en su filtro; el validador puede «Restaurar»)
LOTE_EXCLUIDOS = (LOTE_ENVIADA, LOTE_SIGUIENTE, LOTE_CERRADA_SIN_ENVIO, LOTE_SIN_COMPROBAR)
LOTE_POR_ENVIAR = (LOTE_HOY, LOTE_REZAGADA)

INFO_EMISION, INFO_CERRADA, INFO_NO_PROCEDE = "emision", "cerrada", "no_procede"

# Veredictos
LISTA, REVISAR, CORREGIR, AUTORIZACION = "lista", "revisar", "corregir", "autorizacion"
NO_SE_MANDA, ENVIADA, SIGUIENTE, DESCARTADA, SIN_COMPROBAR = "no_se_manda", "enviada", "siguiente", "descartada", "sin_comprobar"

# Comprobaciones
OK, AVISO, MAL, INFO = "ok", "aviso", "mal", "info"
ICONOS = {OK: "✔", AVISO: "⚠", MAL: "✖", INFO: "•"}

# Motivos de cierre que dicen que la renovación es de otro día (o que solo se cancela el mismo día)
_DICE_OTRO_DIA = re.compile(r"otro dia|dia (anterior|distinto|diferente)|mismo dia|no procede|fuera de (tiempo|horario)")


def _sin_acentos(texto: str) -> str:
    t = unicodedata.normalize("NFKD", str(texto or "").lower())
    return "".join(c for c in t if not unicodedata.combining(c))


def _limpio(texto) -> str:
    """Sin emojis ni espacios dobles (los textos del motor traen íconos)."""
    t = re.sub("[\U0001F000-\U0001FAFF\u2300-\u23FF\u2600-\u27BF\uFE0F]", "", str(texto or ""))
    t = re.sub(r"\s*\((\d{1,2}:\d{2})\)\)", r" de las \1)", t)
    return re.sub(r"\s{2,}", " ", t).strip(" ·-")


def _f(d, formato="%d/%m/%Y") -> str:
    return d.strftime(formato) if d else ""


def _su_ticket(r: dict) -> str:
    return f"el ticket único del {_f(r['Dia_Lote'])}" if r.get("Dia_Lote") else "el ticket único de su día"


def _cierre_txt(r: dict) -> str:
    c = r.get("Cierre_iTop") or {}
    partes = [f"se resolvió el {_f(c.get('fecha_dt'), '%d/%m %H:%M')}" if c.get("fecha_dt") else "se resolvió"]
    if c.get("cerro"):
        partes.append(f"lo cerró {c['cerro']}")
    texto = " (".join(partes[:1] + [", ".join(partes[1:])]) + ")" if len(partes) > 1 else partes[0]
    if c.get("motivo"):
        texto += f": «{c['motivo']}»"
    return texto


def _cierre_corto(r: dict) -> str:
    """«26/10 13:27, María» (el motivo completo está en «Cierre en iTop»)."""
    c = r.get("Cierre_iTop") or {}
    return ", ".join(t for t in (_f(c.get("fecha_dt"), "%d/%m %H:%M"), c.get("cerro") or "") if t) or "sin fecha"


def _emitida(r: dict) -> bool:
    return r.get("Bloqueo_Tipo") == "emision" or bool(r.get("SQL_Cedula"))


def cerrado_por_error(r: dict) -> tuple:
    """
    (texto, seguro) si el incidente se cerró sin mandar la renovación diciendo que es de otro día y la base dice lo
    contrario. seguro=False cuando no se pudo leer TZ_LOCK (posible error: revísalo). ('', False) si no aplica o si ya
    se emitió (entonces el cierre no le costó nada a nadie: ya no se podía cancelar).
    """
    if not r.get("Cerrado_iTop") or r.get("Lote_Estado") not in LOTE_POR_ENVIAR + (LOTE_CERRADA_SIN_ENVIO,):
        return "", False
    if _emitida(r):
        return "", False
    cierre = r.get("Cierre_iTop") or {}
    motivo = _sin_acentos(f"{cierre.get('motivo', '')} {cierre.get('codigo', '')}")
    if not _DICE_OTRO_DIA.search(motivo) or r.get("SQL_Estado") not in SQL_SE_ENVIA:
        return "", False
    lock = r.get("SQL_Tz_Lock")
    if lock not in (None, 0):
        return "", False  # La renovación ya está bloqueada: el cierre estuvo bien
    alta = r.get("SQL_Fecha_Alta")
    base = (f"la base dice que la renovación {r.get('Credito_Renovado_Hijo')} se dio de alta el {_f(alta)}"
            + (" (el mismo día que llegó el incidente)" if r.get("SQL_Estado") == SQL_OK
               else " y la manda la Mesa de Crédito: sí va" if r.get("SQL_Estado") == SQL_MESA_CREDITO
               else " y la generó la APP con plazo menor: sí va"))
    if lock == 0:
        return (f"CERRADO POR ERROR: el incidente {_cierre_txt(r)}, pero {base}, no se mandó en el ticket único y no "
                f"está bloqueada (TZ_LOCK = 0). Va en el lote."), True
    return (f"POSIBLE CIERRE POR ERROR: el incidente {_cierre_txt(r)}, pero {base} y no se mandó en el ticket único. "
            f"No se pudo leer TZ_LOCK para confirmar si ya está bloqueada: revísalo."), False


def aplicar_reglas(rows: list) -> dict:
    """
    Decide qué ya no se manda, en este orden (lo primero que aplica gana; lo demás ya no cuenta):
    1. Ya se emitió (calendario contra el día en que saldría o CR_CEDULAS_DETALLE): no se puede cancelar. Aunque el
       incidente se haya cerrado por error o no haya salido en su día, ya no hay nada que hacer.
    2. El incidente se resolvió en iTop: no se manda, salvo que se haya cerrado por error ('Cerrado_Por_Error': va hoy)
       o que lo parezca y falte TZ_LOCK ('Posible_Cierre_Error': se revisa antes, no va sola).
    3. La base dice «No procede»: no se manda.
    4. Pasó el cierre de mes ('Requiere_Autorizacion'): sigue en el lote, solo con autorización.
    """
    conteo = {INFO_EMISION: 0, INFO_CERRADA: 0, INFO_NO_PROCEDE: 0, "cerrado_por_error": 0, "posible_cierre_error": 0,
              "autorizacion": 0}
    for r in rows:
        r["Informativa"], r["Informativa_Tipo"] = "", ""
        r["Cerrado_Por_Error"], r["Posible_Cierre_Error"], r["Requiere_Autorizacion"] = "", "", ""
        r["No_Enviada_Emitida"] = False
        estado = r.get("Lote_Estado")
        if estado not in LOTE_POR_ENVIAR + (LOTE_CERRADA_SIN_ENVIO,):
            continue
        if _emitida(r):
            r["Informativa_Tipo"] = INFO_EMISION
            r["Informativa"] = (r.get("Bloqueo_Emision") if r.get("Bloqueo_Tipo") == "emision"
                                else "YA NO SE PUEDE CANCELAR: la renovación ya está en la cédula (CR_CEDULAS_DETALLE): "
                                     "ya se emitió")
            if r.get("No_Enviada") or r.get("Cerrada_Sin_Envio"):
                # No salió en su día y ya se emitió: se perdió la cancelación (ya no se suma a ningún lote)
                r["No_Enviada"], r["Cerrada_Sin_Envio"], r["No_Enviada_Emitida"] = False, False, True
                if r.get("Dia_Lote"):
                    r["Lote_Resumen"] = f"No salió el {_f(r['Dia_Lote'], '%d/%m')}  ·  ya se emitió"
            conteo[INFO_EMISION] += 1
            continue
        texto, seguro = cerrado_por_error(r)
        if seguro:
            r["Cerrado_Por_Error"] = texto
            conteo["cerrado_por_error"] += 1
        elif texto:
            r["Posible_Cierre_Error"] = texto
            conteo["posible_cierre_error"] += 1
        if estado == LOTE_CERRADA_SIN_ENVIO:
            if r["Cerrado_Por_Error"]:
                r["No_Enviada"], r["Cerrada_Sin_Envio"] = True, False
                r["Lote_Resumen"] = f"NO SALIÓ ({_f(r.get('Dia_Lote'), '%d/%m')})  ·  cerrado por error  ·  va hoy"
            elif r["Posible_Cierre_Error"]:
                r["No_Enviada"], r["Cerrada_Sin_Envio"] = False, False
                r["Lote_Resumen"] = "¿Cerrado por error?  ·  falta TZ_LOCK"
                continue
            else:
                # No salió en su día pero el incidente ya se resolvió (y no por error): no es un faltante, no se manda
                r["No_Enviada"], r["Cerrada_Sin_Envio"] = False, True
                r["Informativa_Tipo"] = INFO_CERRADA
                r["Informativa"] = f"No salió en {_su_ticket(r)}, pero el incidente ya {_cierre_txt(r)}: no se manda"
                r["Lote_Etiqueta"] = f"🔒 {r['Informativa']}"
                r["Lote_Resumen"] = "Resuelto en iTop  ·  no se manda"
                conteo[INFO_CERRADA] += 1
                continue
        if r.get("Cerrado_iTop") and not r["Cerrado_Por_Error"]:
            if not r["Posible_Cierre_Error"]:
                r["Informativa_Tipo"] = INFO_CERRADA
                r["Informativa"] = f"El incidente ya {_cierre_txt(r)}: no se manda en el ticket único"
        elif r.get("SQL_Estado") == SQL_NO_PROCEDE:
            r["Informativa_Tipo"] = INFO_NO_PROCEDE
            r["Informativa"] = f"{LEYENDA_NO_PROCEDE}: no se manda en el ticket único"
        elif r.get("Bloqueo_Tipo") == "cierre":
            r["Requiere_Autorizacion"] = r.get("Bloqueo_Emision") or "Pasó el cierre de mes: solo con autorización"
            conteo["autorizacion"] += 1
        if r["Informativa_Tipo"]:
            conteo[r["Informativa_Tipo"]] += 1
            if r.get("No_Enviada"):
                # No salió en su día, pero ya no se manda: no es un faltante ni dice «NO SALIÓ» en la tabla
                r["No_Enviada"] = False
                r["Lote_Resumen"] = (f"No salió el {_f(r.get('Dia_Lote'), '%d/%m')}  ·  "
                                     + ("resuelto en iTop" if r["Informativa_Tipo"] == INFO_CERRADA else "no procede"))
    conteo["informativas"] = conteo[INFO_EMISION] + conteo[INFO_CERRADA] + conteo[INFO_NO_PROCEDE]
    return conteo


def recontar_envios(info: dict, rows: list):
    """
    Después de aplicar_reglas: lo que no salió en su día pero ya se resolvió en iTop o ya se emitió deja de contar como
    «NO enviada» (en el resumen, en los avisos y en la pestaña «Por día»). Solo cuenta lo que todavía se puede mandar.
    """
    if not info:
        return
    base = [r for r in rows if r.get("En_Periodo", True)] if info.get("revision") else rows
    info["no_enviadas"] = sum(1 for r in base if r.get("No_Enviada"))
    info["cerradas_sin_envio"] = sum(1 for r in base if r.get("Cerrada_Sin_Envio"))
    info["emitidas_sin_envio"] = sum(1 for r in base if r.get("No_Enviada_Emitida"))
    for d in info.get("por_dia") or []:
        suyas = [r for r in base if r.get("Dia_Lote") == d["dia"]]
        d["no_enviadas"] = sum(1 for r in suyas if r.get("No_Enviada"))
        d["cerradas"] = sum(1 for r in suyas if r.get("Cerrada_Sin_Envio"))


def va_en_lote(r: dict) -> bool:
    """¿Va en el ticket único de hoy? (sin contar lo que el usuario descartó o metió a mano)."""
    if r.get("Informativa") or r.get("Posible_Cierre_Error"):
        return False
    if r.get("Cerrado_Por_Error"):
        return True
    return r.get("Lote_Estado") not in LOTE_EXCLUIDOS


def es_del_dia(r: dict, desde) -> bool:
    """¿Le toca el ticket único de `desde` o uno posterior? (sin `desde`, o si no tiene día: sí)."""
    return desde is None or not r.get("Dia_Lote") or r["Dia_Lote"] >= desde


def grupos_del_dia(filas: list, descartadas=(), incluidas=(), solo_periodo: bool = False, desde=None) -> dict:
    """
    Los grupos que usan a la vez los avisos y los indicadores (así no se contradicen):
      'cerradas_error': se cerraron por error y sí van hoy;
      'posibles_cierre': ¿se cerraron por error? (falta TZ_LOCK), fuera del lote;
      'no_se_mandan': informativas fuera del lote (emitidas, resueltas en iTop, no proceden), y
      'por_tipo': {Informativa_Tipo: cuántas} de esas.
    No cuenta lo descartado. `incluidas`: lo que se metió a mano (va aunque las reglas digan que no).
    `solo_periodo`: en la revisión semanal, solo lo del periodo.
    `desde`: en el lote del día, «no se mandan» solo cuenta las de ese ticket único o de después (las de días
    anteriores se ven aparte, en «No se mandan (todas)», de todos los días).
    """
    descartadas, incluidas = set(descartadas), set(incluidas)
    vivas = [r for r in filas if r.get("Row_ID") not in descartadas
             and (not solo_periodo or r.get("En_Periodo", True))]

    def fuera(r):
        return not va_en_lote(r) and r.get("Row_ID") not in incluidas

    no_se_mandan = [r for r in vivas if r.get("Informativa") and fuera(r) and es_del_dia(r, desde)]
    por_tipo = {}
    for r in no_se_mandan:
        por_tipo[r.get("Informativa_Tipo")] = por_tipo.get(r.get("Informativa_Tipo"), 0) + 1
    return {
        "cerradas_error": [r for r in vivas if r.get("Cerrado_Por_Error") and not fuera(r)],
        "posibles_cierre": [r for r in vivas if r.get("Posible_Cierre_Error") and fuera(r)],
        "no_se_mandan": no_se_mandan,
        "por_tipo": por_tipo,
    }


def texto_no_se_mandan(por_tipo: dict) -> str:
    """«1 ya se emitió · 2 resueltas en iTop» ('' si no hay)."""
    textos = ((INFO_EMISION, "# ya se [emitió|emitieron]"), (INFO_CERRADA, "# [resuelta|resueltas] en iTop"),
              (INFO_NO_PROCEDE, "# no [procede|proceden]"))
    return " · ".join(pl(por_tipo[t], txt) for t, txt in textos if por_tipo.get(t))


# ======================= Desglose del ticket =======================

def _check_correo(r: dict) -> tuple:
    cli, cred = r.get("Numero_Cliente", ""), r.get("Credito_Renovado_Hijo", "")
    sin_cli, sin_cred = str(cli).upper().startswith("SIN"), str(cred).upper().startswith("SIN")
    if sin_cli or sin_cred:
        falta = "el número de cliente y el crédito" if sin_cli and sin_cred else ("el número de cliente" if sin_cli else "el crédito")
        return MAL, f"Correo: falta {falta}. Búscalo en el correo original o en iTop."
    if r.get("Credito_Padre_Enviado"):
        return AVISO, (f"Correo: trae el crédito padre {r['Credito_Padre_Enviado']}; en el lote va su renovación {cred}. "
                       f"Valida que sea la correcta.")
    if r.get("Dato_Imagen"):
        return AVISO, f"Correo: leído de una captura de pantalla ({_limpio(r['Dato_Imagen'])}). Compáralo con la imagen."
    if "DISCREPANCIA" in str(r.get("Alerta_Renovacion", "")):
        return AVISO, f"Correo: trae dos números de cliente para este crédito ({cli}). Confirma cuál es."
    if r.get("Clientes_Sin_Leer"):
        nums = r["Clientes_Sin_Leer"]
        uno = len(nums) == 1
        return AVISO, (f"Correo: también trae {'el número' if uno else 'los números'} {', '.join(nums[:4])}"
                       f"{'…' if len(nums) > 4 else ''}, que no se {'leyó' if uno else 'leyeron'} como cliente: "
                       f"revisa que no falte una renovación.")
    parte = f" (renovación {r.get('Indice_En_Ticket')} de {r.get('Total_En_Ticket')})" if r.get("Es_Multiple") else ""
    return OK, f"Correo: cliente {cli} y crédito {cred}{parte}."


def _checks_base(r: dict, sql_activo: bool) -> list:
    estado, detalle = r.get("SQL_Estado"), _limpio(r.get("SQL_Detalle")).rstrip(".")
    checks = []
    if not estado:
        checks.append((INFO, "Base de datos: no se validó" + (
            "." if sql_activo else " (actívala en Configuración → Validaciones automáticas para confirmarla).")))
    elif estado in SQL_SE_ENVIA:
        checks.append((OK, f"Base de datos: {detalle}."))
    elif estado in (SQL_SUGERENCIA, SQL_PADRE_CORREGIDO):
        checks.append((AVISO, f"Base de datos: {detalle}."))
    elif estado == SQL_ES_PADRE:
        checks.append((MAL, f"Base de datos: {detalle}."))
    elif estado in SQL_PROBLEMAS:
        checks.append((MAL, f"Base de datos: {detalle}."))
    elif estado != SQL_SIN_DATOS:
        checks.append((INFO, f"Base de datos: {detalle}."))
    for texto in r.get("SQL_Clientes_Sin_Renov") or []:
        checks.append((AVISO, f"Base de datos: {texto}."))
    return checks


def _check_calendario(r: dict) -> tuple:
    corto = _limpio(r.get("Bloqueo_Corto") or r.get("Bloqueo_Emision"))
    if r.get("Bloqueo_Tipo") == "emision":
        return MAL, f"Calendario: ya se emitió: {corto}. No se puede cancelar."
    if r.get("SQL_Cedula"):
        return MAL, "Calendario: ya está en la cédula (CR_CEDULAS_DETALLE): ya se emitió. No se puede cancelar."
    if r.get("Sin_Calendario"):
        return AVISO, (f"Calendario: no tiene la fecha de la emisión de {r['Sin_Calendario']}: no se sabe si ya se "
                       f"emitió. Revisa el calendario de FONACOT.")
    if r.get("Bloqueo_Tipo") == "cierre":
        return AVISO, f"Calendario: pasó el cierre de mes: {corto}. Solo con autorización."
    alta = r.get("SQL_Fecha_Alta") or r.get("Fecha_Alta_Correo")
    desde = f"el alta ({_f(alta, '%d/%m')})" if alta else "que llegó"
    return OK, f"Calendario: desde {desde} no ha pasado ninguna emisión ni cierre de mes."


def _check_envio(r: dict) -> tuple:
    estado, etiqueta = r.get("Lote_Estado"), _limpio(r.get("Lote_Etiqueta"))
    dia = _f(r.get("Dia_Lote"), "%d/%m")
    if r.get("Error_Envio"):
        return MAL, f"Ticket único: salió con datos equivocados ({r['Error_Envio']})."
    if r.get("No_Enviada_Emitida"):
        return INFO, f"Ticket único: no salió el {dia}, que era su día, y ya se emitió."
    if estado == LOTE_ENVIADA:
        tarde = f", {r['Dias_Retraso']} {'día' if r['Dias_Retraso'] == 1 else 'días'} tarde" if r.get("Dias_Retraso") else ""
        if r.get("Posible_Duplicado"):
            return AVISO, (f"Ticket único: ya había salido el {_f(r.get('Enviado_Fecha'), '%d/%m')}, antes de que "
                           f"llegara este incidente (¿solicitud duplicada?).")
        return INFO, f"Ticket único: salió el {_f(r.get('Enviado_Fecha'), '%d/%m')}{tarde}."
    if estado == LOTE_EN_CONCENTRADO:
        return OK, "Ticket único: ya va en el de hoy."
    if r.get("No_Enviada") or r.get("Cerrado_Por_Error"):
        return AVISO, f"Ticket único: no salió el {dia}, que era su día: va hoy."
    if estado == LOTE_HOY:
        if r.get("Falta_Hoy"):
            return AVISO, f"Ticket único: {etiqueta}."
        return OK, "Ticket único: va en el de hoy" + (" (en otro ticket único: hoy es día de corte)." if "otro ticket" in etiqueta else ".")
    if estado in (LOTE_REZAGADA, LOTE_SIGUIENTE, LOTE_SIN_COMPROBAR):
        return (INFO if estado == LOTE_SIGUIENTE else AVISO), f"Ticket único: {etiqueta}."
    if estado == LOTE_CERRADA_SIN_ENVIO:
        return INFO, f"Ticket único: no salió el {dia}, que era su día; el incidente ya se resolvió."
    return INFO, f"Ticket único: {etiqueta or 'sin datos del envío'}."


def _checks_itop(r: dict) -> list:
    checks = []
    if r.get("Cerrado_Por_Error"):
        checks.append((AVISO, f"iTop: se cerró por error ({_cierre_corto(r)}): dice que es de otro día, pero la base "
                              f"dice que no y no está bloqueada (TZ_LOCK = 0)."))
    elif r.get("Posible_Cierre_Error"):
        checks.append((AVISO, f"iTop: POSIBLE CIERRE POR ERROR ({_cierre_corto(r)}): dice que es de otro día y la base "
                              f"dice que no; falta TZ_LOCK para confirmarlo."))
    elif r.get("Cerrado_iTop"):
        checks.append((INFO, f"iTop: resuelto ({_cierre_corto(r)})."))
    else:
        checks.append((OK, "iTop: el incidente sigue abierto."))
    if r.get("Requiere_Validacion_Agente") or r.get("Reasignado_A_Otro"):
        checks.append((AVISO, f"Asignación: {_limpio(r.get('Alerta_Asignacion'))}."))
    if r.get("Ticket_Duplicado"):
        checks.append((AVISO, "Duplicado: el mismo crédito viene en otro incidente."))
    return checks


def comprobaciones(r: dict, sql_activo: bool = True) -> list:
    """[(estado, texto)]: correo, base de datos, calendario, ticket único e iTop, con la evidencia de cada una."""
    return ([_check_correo(r)] + _checks_base(r, sql_activo) + [_check_calendario(r), _check_envio(r)]
            + _checks_itop(r))


def _fuente(texto: str) -> str:
    return texto.split(":", 1)[0]


def _priorizar(checks: list, decide: str, va: bool) -> tuple:
    """
    Primero la comprobación que decide el veredicto; luego lo que todavía hay que atender (✖ y ⚠); al final lo demás.
    Si no va en el ticket único, lo que no decide ya no importa: queda como dato (•), sin ✖ ni ⚠ que lo contradigan.
    Regresa (comprobaciones ordenadas, cuántas son las principales).
    """
    principal = [c for c in checks if decide and _fuente(c[1]) == decide][:1]
    resto = [c for c in checks if c not in principal]
    if not va:
        resto = [(INFO if e in (MAL, AVISO) else e, t) for e, t in resto]
    atender = sorted((c for c in resto if c[0] in (MAL, AVISO)), key=lambda c: c[0] != MAL)
    otras = [c for c in resto if c[0] not in (MAL, AVISO)]
    return principal + atender + otras, len(principal) + len(atender)


def _accion_corregir(r: dict, checks: list) -> str:
    if r.get("SQL_Estado") == SQL_ES_PADRE:
        return (f"Cambia el crédito {r.get('Credito_Renovado_Hijo')} por la renovación "
                f"{r.get('SQL_Renovacion_Hija') or '(ver base de datos)'} antes de mandarla.")
    if r.get("Datos_Incompletos"):
        return "Completa el dato que falta (correo original o iTop) antes de mandarla."
    if r.get("Error_Envio"):
        return "Salió con datos equivocados: avisa a quien mandó el ticket único para corregirlo."
    if r.get("SQL_Estado") in SQL_PROBLEMAS:
        return "La base no la confirma: verifica cliente y crédito con la sucursal antes de mandarla."
    malas = [t for e, t in checks if e == MAL]
    return f"Revisa: {malas[0]}" if malas else "Revísala antes de mandarla."


def veredicto(r: dict, en_lote: bool = None, validada: bool = False, descartada: bool = False,
              sql_activo: bool = True) -> dict:
    """
    {'nivel', 'corto' (tabla), 'titulo' (banner), 'accion' (qué hacer), 'checks', 'principales'}.
    Una sola decisión, por prioridad: descartada > salió con datos equivocados > ya se emitió / resuelta / no procede >
    ¿cerrada por error? (falta TZ_LOCK) > ya enviada / siguiente / sin comprobar > corregir > autorización (se suma a
    lo que siga) > no salió en su día (va hoy) > revisar > lista. Las comprobaciones siguen esa decisión: la que decide va
    primero y, si no va en el ticket único, las demás no la contradicen. 'principales' = cuántas se muestran de entrada.
    `en_lote`: si va en el ticket único de hoy (por omisión, va_en_lote).
    """
    checks = comprobaciones(r, sql_activo)
    en_lote = va_en_lote(r) if en_lote is None else en_lote
    estado = r.get("Lote_Estado")

    def v(nivel, corto, titulo, accion, decide=None, va=True):
        ordenadas, n = _priorizar(checks, decide, va)
        return dict(nivel=nivel, corto=corto, titulo=titulo, accion=accion, checks=ordenadas, principales=n)

    if descartada:
        aun_asi = {INFO_EMISION: "ya se emitió", INFO_CERRADA: "el incidente ya se resolvió",
                   INFO_NO_PROCEDE: "no procede"}.get(r.get("Informativa_Tipo"))
        return v(DESCARTADA, "Descartada", "Descartada",
                 f"La descartaste. De todos modos no se manda: {aun_asi}." if aun_asi
                 else "La descartaste: no va en el lote. «Restaurar» la regresa.", va=False)
    if r.get("Error_Envio") and estado in (LOTE_ENVIADA, LOTE_EN_CONCENTRADO):
        return v(CORREGIR, "Corregir", "Salió con datos equivocados",
                 f"Avisa a quien mandó el ticket único para corregirlo: {r['Error_Envio']}.", "Ticket único")
    if r.get("Informativa") and not en_lote:
        tipo = r.get("Informativa_Tipo")
        titulo, accion, decide = {
            INFO_EMISION: ("No se manda: ya se emitió",
                           "Nada que hacer: ya no se puede cancelar. Si el incidente sigue abierto, respóndelo.",
                           "Calendario"),
            INFO_CERRADA: ("No se manda: resuelta en iTop", "Nada que hacer.", "iTop"),
            INFO_NO_PROCEDE: ("No se manda: no procede",
                              f"Si el incidente sigue abierto, respóndelo con: «{LEYENDA_NO_PROCEDE}».",
                              "Base de datos"),
        }.get(tipo, ("No se manda", "Nada que hacer.", None))
        return v(NO_SE_MANDA, "No se manda", titulo, accion, decide, va=False)
    if r.get("Posible_Cierre_Error") and not en_lote:
        return v(REVISAR, "¿Cerrada por error?", "¿Se cerró por error?",
                 "No se pudo leer TZ_LOCK: confirma en la base si la renovación está bloqueada. Si no lo está "
                 "(TZ_LOCK = 0), «Restaurar» la manda hoy.", "iTop", va=False)
    if not en_lote:
        if estado == LOTE_ENVIADA:
            if r.get("Posible_Duplicado"):
                return v(REVISAR, "¿Duplicada?", "¿Solicitud duplicada?",
                         f"Ya había salido el {_f(r.get('Enviado_Fecha'))}, antes de que llegara este incidente: "
                         f"confirma con la sucursal si es otra solicitud.", "Ticket único", va=False)
            return v(ENVIADA, "Enviada", "Ya enviada",
                     f"Nada que hacer: salió en el ticket único del {_f(r.get('Enviado_Fecha'))}.", "Ticket único", va=False)
        if estado == LOTE_SIGUIENTE:
            return v(SIGUIENTE, "Siguiente", "Va en el siguiente ticket único",
                     f"No va hoy: va en el ticket único del {_f(r.get('Dia_Lote'))}.", "Ticket único", va=False)
        if estado == LOTE_CERRADA_SIN_ENVIO:
            return v(NO_SE_MANDA, "No se manda", "No salió y ya se cerró",
                     "No salió en su día y el incidente ya se cerró: no entra al lote («Restaurar» la incluye).",
                     "iTop", va=False)
        return v(SIN_COMPROBAR, "Sin comprobar", "Sin comprobar",
                 "No se encontró el ticket único de su día: revisa si salió («Restaurar» la suma al lote).",
                 "Ticket único", va=False)
    if estado == LOTE_EN_CONCENTRADO:
        return v(ENVIADA, "Enviada", "Ya va en el ticket único de hoy",
                 "Nada que hacer: ya salió en el ticket único de hoy.", "Ticket único")

    # Va en el ticket único de hoy
    autorizacion = "Solo con autorización (pasó el cierre de mes). " if r.get("Requiere_Autorizacion") else ""
    malas = [t for e, t in checks if e == MAL]
    if malas:
        return v(CORREGIR, "Corregir", "Corrige antes de enviar", autorizacion + _accion_corregir(r, checks),
                 _fuente(malas[0]))
    atrasada = bool(r.get("No_Enviada") or r.get("Cerrado_Por_Error"))
    va_hoy = "No salió en su día: mándala hoy" + (
        f" y avisa a {(r.get('Cierre_iTop') or {}).get('cerro') or 'quien lo cerró'} que se cerró por error"
        if r.get("Cerrado_Por_Error") else "") + "."
    if autorizacion:
        return v(AUTORIZACION, "Autorización", "Necesita autorización",
                 "Pasó el cierre de mes: mándala solo si ya tienes la autorización." + (f" {va_hoy}" if atrasada else ""),
                 "Calendario")
    # Lo de «no salió en su día» o «se cerró por error» no es un dato que corregir: va hoy (al validarla, queda lista)
    otras = [t for e, t in checks if e == AVISO and not (atrasada and _fuente(t) in ("Ticket único", "iTop"))]
    if atrasada and not otras:
        if validada:
            return v(LISTA, "✓ Validada", "Lista · no salió en su día", "Va hoy (no salió en su día). Ya la validaste.",
                     "Ticket único")
        return v(REVISAR, "Atrasada · va hoy", "No salió en su día: va hoy", va_hoy, "Ticket único")
    if otras or atrasada:
        aviso = otras[0].split(": ", 1)[-1] if otras else ""
        return v(REVISAR, "Revisar", "Revisa antes de enviar",
                 (f"{va_hoy} Revisa también: {aviso}" if atrasada else f"Revisa: {aviso}"),
                 _fuente(otras[0]) if otras else "Ticket único")
    return v(LISTA, "✓ Validada" if validada else "Lista", "Lista para enviar",
             ("Va en el ticket único de hoy. Ya la validaste." if validada
              else "Va en el ticket único de hoy: si todo cuadra, márcala como validada (Espacio)."))


# ======================= Comparación en tres columnas =======================

def _falta(valor) -> bool:
    return not valor or str(valor).strip() in ("-", "") or str(valor).upper().startswith("SIN")


def _v(valor) -> str:
    return "falta" if _falta(valor) else str(valor)


def comparacion(r: dict, en_lote: bool = None) -> list:
    """
    Lo que dice cada fuente, lado a lado: [{'dato', 'incidente', 'base', 'ticket', 'estado'}] para el cliente, el
    crédito y la fecha. 'estado' (OK/AVISO/MAL/INFO) dice si las tres cuadran, si se corrigió algo o si no cuadran.
    """
    en_lote = va_en_lote(r) if en_lote is None else en_lote
    estado = r.get("SQL_Estado")
    cli, cred = r.get("Numero_Cliente"), r.get("Credito_Renovado_Hijo")
    cli_inc = r.get("Cliente_Enviado", cli) if "Cliente_Enviado" in r else cli
    cred_inc = r.get("Credito_Padre_Enviado") or cred
    alta, llegada = r.get("SQL_Fecha_Alta"), r.get("Fecha_Llegada_DT") or r.get("Fecha_DT")
    confirma = estado in SQL_SE_ENVIA + (SQL_PADRE_CORREGIDO,)
    sin_base = not estado or estado == SQL_SIN_DATOS

    # Columna de la base de datos
    if sin_base:
        b_cli = b_cred = b_fecha = "sin validar"
    elif confirma:
        b_cli, b_cred, b_fecha = _v(cli), _v(cred), _f(alta) or "—"
        if r.get("Credito_Padre_Enviado"):
            b_cred = f"{cred} (renovación del padre {r['Credito_Padre_Enviado']})"
    elif estado == SQL_ES_PADRE:
        b_cli, b_cred, b_fecha = _v(cli), f"{cred} es el padre de {r.get('SQL_Renovacion_Hija') or '?'}", "—"
    elif estado == SQL_NO_EXISTE:
        b_cli, b_cred, b_fecha = _v(cli), "no existe para ese cliente", "—"
    elif estado == SQL_SIN_RENOVACIONES:
        b_cli, b_cred, b_fecha = "sin renovaciones", "—", "—"
    elif estado in (SQL_OTRA_FECHA, SQL_NO_PROCEDE):
        b_cli, b_cred, b_fecha = _v(cli), _v(cred), f"{_f(alta) or '?'} (otro día)"
    else:  # sugerencia u otro: el detalle lo explica
        b_cli, b_cred, b_fecha = _v(cli), "ver detalle", _f(alta) or "—"

    # Columna del ticket único
    if en_lote:
        t_cli, t_cred = _v(cli), _v(cred)
        atrasada = r.get("No_Enviada") or r.get("Cerrado_Por_Error")
        t_fecha = (f"hoy (le tocaba el {_f(r.get('Dia_Lote'))})" if atrasada and r.get("Dia_Lote")
                   else f"hoy ({_f(r.get('Dia_Lote'))})" if r.get("Dia_Lote") else "hoy")
    elif r.get("Lote_Estado") in (LOTE_ENVIADA, LOTE_EN_CONCENTRADO):
        t_cli, t_cred, t_fecha = _v(cli), _v(cred), f"salió el {_f(r.get('Enviado_Fecha')) or 'día de hoy'}"
    else:
        t_cli = t_cred = t_fecha = "no se manda"

    def est(inc, ticket, base_ok):
        if _falta(inc) and _falta(ticket):
            return MAL
        if sin_base:
            return INFO if not _falta(ticket) else MAL
        if not base_ok:
            return MAL
        return OK if str(inc) == str(ticket) or not en_lote else AVISO

    llegada = llegada.date() if hasattr(llegada, "date") else llegada
    return [
        dict(dato="Cliente", incidente=_v(cli_inc), base=b_cli, ticket=t_cli,
             estado=est(cli_inc, cli, estado not in (SQL_SIN_RENOVACIONES,))),
        dict(dato="Crédito", incidente=_v(cred_inc), base=b_cred, ticket=t_cred,
             estado=est(cred_inc, cred, confirma or estado in (SQL_OTRA_FECHA, SQL_NO_PROCEDE))),
        dict(dato="Fecha", incidente=f"llegó {_f(llegada)}" if llegada else "—", base=b_fecha, ticket=t_fecha,
             estado=(INFO if sin_base else OK if confirma else MAL if estado in (SQL_OTRA_FECHA, SQL_NO_PROCEDE)
                     else INFO)),
    ]
