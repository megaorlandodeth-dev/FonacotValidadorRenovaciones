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
    SQL_OK, SQL_APP_PLAZO, SQL_SUGERENCIA, SQL_SIN_DATOS, SQL_ES_PADRE, SQL_PADRE_CORREGIDO, SQL_NO_PROCEDE, SQL_PROBLEMAS,
    SQL_NO_EXISTE, SQL_SIN_RENOVACIONES, SQL_OTRA_FECHA,
    LEYENDA_NO_PROCEDE,
)

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
    t = re.sub(r"[\U0001F000-\U0001FAFF☀-➿️]", "", str(texto or ""))
    return re.sub(r"\s{2,}", " ", t).strip(" ·-")


def _f(d, formato="%d/%m/%Y") -> str:
    return d.strftime(formato) if d else ""


def _cierre_txt(r: dict) -> str:
    c = r.get("Cierre_iTop") or {}
    partes = [f"se resolvió el {_f(c.get('fecha_dt'), '%d/%m %H:%M')}" if c.get("fecha_dt") else "se resolvió"]
    if c.get("cerro"):
        partes.append(f"lo cerró {c['cerro']}")
    texto = " (".join(partes[:1] + [", ".join(partes[1:])]) + ")" if len(partes) > 1 else partes[0]
    if c.get("motivo"):
        texto += f": «{c['motivo']}»"
    return texto


def cerrado_por_error(r: dict) -> tuple:
    """
    (texto, seguro) si el incidente se cerró sin mandar la renovación diciendo que es de otro día y la base dice lo
    contrario. seguro=False cuando no se pudo leer TZ_LOCK (posible error: revísalo). ('', False) si no aplica.
    """
    if not r.get("Cerrado_iTop") or r.get("Lote_Estado") not in LOTE_POR_ENVIAR + (LOTE_CERRADA_SIN_ENVIO,):
        return "", False
    cierre = r.get("Cierre_iTop") or {}
    motivo = _sin_acentos(f"{cierre.get('motivo', '')} {cierre.get('codigo', '')}")
    if not _DICE_OTRO_DIA.search(motivo) or r.get("SQL_Estado") not in (SQL_OK, SQL_APP_PLAZO):
        return "", False
    lock = r.get("SQL_Tz_Lock")
    if lock not in (None, 0):
        return "", False  # La renovación ya está bloqueada: el cierre estuvo bien
    alta = r.get("SQL_Fecha_Alta")
    base = (f"la base dice que la renovación {r.get('Credito_Renovado_Hijo')} se dio de alta el {_f(alta)}"
            + (" (el mismo día que llegó el incidente)" if r.get("SQL_Estado") == SQL_OK
               else " y la generó la APP con plazo menor: sí va"))
    if lock == 0:
        return (f"CERRADO POR ERROR: el incidente {_cierre_txt(r)}, pero {base}, no se mandó en el ticket único y no "
                f"está bloqueada (TZ_LOCK = 0). Va en el lote."), True
    return (f"POSIBLE CIERRE POR ERROR: el incidente {_cierre_txt(r)}, pero {base} y no se mandó en el ticket único. "
            f"No se pudo leer TZ_LOCK para confirmar si ya está bloqueada: revísalo."), False


def aplicar_reglas(rows: list) -> dict:
    """
    Pone en cada renovación:
    - 'Informativa' (texto) e 'Informativa_Tipo' si ya no se manda (emisión, resuelta en iTop o no procede);
    - 'Cerrado_Por_Error' (texto) si el incidente se cerró por error (entonces sí va en el lote);
    - 'Posible_Cierre_Error' si parece cerrado por error pero falta TZ_LOCK para confirmarlo;
    - 'Requiere_Autorizacion' si pasó el cierre de mes (sigue en el lote).
    """
    conteo = {INFO_EMISION: 0, INFO_CERRADA: 0, INFO_NO_PROCEDE: 0, "cerrado_por_error": 0, "autorizacion": 0}
    for r in rows:
        r["Informativa"], r["Informativa_Tipo"] = "", ""
        r["Cerrado_Por_Error"], r["Posible_Cierre_Error"], r["Requiere_Autorizacion"] = "", "", ""
        texto, seguro = cerrado_por_error(r)
        if seguro:
            r["Cerrado_Por_Error"] = texto
            conteo["cerrado_por_error"] += 1
        elif texto:
            r["Posible_Cierre_Error"] = texto
        if r.get("Lote_Estado") not in LOTE_POR_ENVIAR and not r["Cerrado_Por_Error"]:
            continue
        if r.get("Bloqueo_Tipo") == "emision" or r.get("SQL_Cedula"):
            r["Informativa_Tipo"] = INFO_EMISION
            r["Informativa"] = (r.get("Bloqueo_Emision") if r.get("Bloqueo_Tipo") == "emision"
                                else "YA NO SE PUEDE CANCELAR: la renovación ya está en la cédula (CR_CEDULAS_DETALLE): "
                                     "ya se emitió")
        elif r.get("Cerrado_iTop") and not r["Cerrado_Por_Error"]:
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
    conteo["informativas"] = conteo[INFO_EMISION] + conteo[INFO_CERRADA] + conteo[INFO_NO_PROCEDE]
    return conteo


def va_en_lote(r: dict) -> bool:
    """¿Va en el ticket único de hoy? (sin contar lo que el usuario descartó o metió a mano)."""
    if r.get("Informativa"):
        return False
    if r.get("Cerrado_Por_Error"):
        return True
    return r.get("Lote_Estado") not in LOTE_EXCLUIDOS


# ======================= Desglose del ticket =======================

def _check_correo(r: dict) -> tuple:
    cli, cred = r.get("Numero_Cliente", ""), r.get("Credito_Renovado_Hijo", "")
    sin_cli, sin_cred = str(cli).upper().startswith("SIN"), str(cred).upper().startswith("SIN")
    if sin_cli or sin_cred:
        falta = "el número de cliente y el crédito" if sin_cli and sin_cred else ("el número de cliente" if sin_cli else "el crédito")
        return MAL, f"Correo: falta {falta}. Búscalo en el correo original (abajo) o en iTop."
    if r.get("Credito_Padre_Enviado"):
        return AVISO, (f"Correo: trae el crédito padre {r['Credito_Padre_Enviado']}; en el lote va su renovación {cred} "
                       f"(cliente {cli}). Valida que sea la correcta.")
    if r.get("Dato_Imagen"):
        return AVISO, f"Correo: leído de una captura de pantalla ({_limpio(r['Dato_Imagen'])}). Compáralo con la imagen."
    if "DISCREPANCIA" in str(r.get("Alerta_Renovacion", "")):
        return AVISO, f"Correo: trae dos números de cliente para este crédito ({cli}). Confirma cuál es."
    parte = (f" Es la renovación {r.get('Indice_En_Ticket')} de {r.get('Total_En_Ticket')} de este incidente."
             if r.get("Es_Multiple") else "")
    return OK, f"Correo: cliente {cli} y crédito {cred} vienen escritos en el incidente.{parte}"


def _checks_base(r: dict, sql_activo: bool) -> list:
    estado, detalle = r.get("SQL_Estado"), _limpio(r.get("SQL_Detalle")).rstrip(".")
    checks = []
    if not estado:
        checks.append((INFO, "Base de datos: no se validó" + (
            "." if sql_activo else " (actívala en Configuración → Validaciones automáticas para confirmarla).")))
    elif estado in (SQL_OK, SQL_APP_PLAZO):
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
    if r.get("Bloqueo_Tipo") == "emision":
        return MAL, f"Calendario: {_limpio(r['Bloqueo_Emision'])}."
    if r.get("SQL_Cedula"):
        return MAL, "Calendario: la renovación ya está en la cédula (CR_CEDULAS_DETALLE): ya se emitió."
    if r.get("Bloqueo_Tipo") == "cierre":
        return AVISO, f"Calendario: {_limpio(r['Bloqueo_Emision'])}"
    desde = f"el alta ({_f(r['SQL_Fecha_Alta'])})" if r.get("SQL_Fecha_Alta") else "que llegó el incidente"
    return OK, f"Calendario: desde {desde} no ha pasado ninguna emisión ni cierre de mes: se puede cancelar."


def _check_envio(r: dict) -> tuple:
    estado, etiqueta = r.get("Lote_Estado"), _limpio(r.get("Lote_Etiqueta"))
    if r.get("Error_Envio"):
        return MAL, f"Ticket único: salió con datos equivocados ({r['Error_Envio']}). Hay que corregirlo."
    if estado == LOTE_ENVIADA:
        tarde = f", {r['Dias_Retraso']} día(s) tarde" if r.get("Dias_Retraso") else ""
        if r.get("Posible_Duplicado"):
            return AVISO, f"Ticket único: ya había salido el {_f(r.get('Enviado_Fecha'))}, antes de que llegara este incidente (posible solicitud duplicada)."
        return INFO, f"Ticket único: ya salió en el del {_f(r.get('Enviado_Fecha'))}{tarde}."
    if estado == LOTE_EN_CONCENTRADO:
        return OK, "Ticket único: ya va en el de hoy."
    if estado == LOTE_HOY:
        if r.get("Falta_Hoy"):
            return AVISO, f"Ticket único: {etiqueta}."
        return OK, "Ticket único: va en el de hoy" + (" (en otro ticket único: hoy es día de corte)." if "otro ticket" in etiqueta else ".")
    if estado == LOTE_REZAGADA:
        if r.get("No_Enviada"):
            return MAL, f"Ticket único: no salió en el del {_f(r.get('Dia_Lote'))}, que era su día. Va hoy."
        return AVISO, f"Ticket único: {etiqueta}."
    if estado == LOTE_SIGUIENTE:
        return INFO, f"Ticket único: {etiqueta}."
    if estado == LOTE_CERRADA_SIN_ENVIO:
        return MAL, f"Ticket único: {etiqueta}."
    if estado == LOTE_SIN_COMPROBAR:
        return AVISO, f"Ticket único: {etiqueta}."
    return INFO, f"Ticket único: {etiqueta or 'sin datos del envío'}."


def _checks_itop(r: dict) -> list:
    checks = []
    if r.get("Cerrado_Por_Error"):
        checks.append((MAL, f"iTop: {r['Cerrado_Por_Error']}"))
    elif r.get("Posible_Cierre_Error"):
        checks.append((AVISO, f"iTop: {r['Posible_Cierre_Error']}"))
    elif r.get("Cerrado_iTop"):
        checks.append((INFO, f"iTop: el incidente {_cierre_txt(r)}."))
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


def _accion_corregir(r: dict, checks: list) -> str:
    if r.get("SQL_Estado") == SQL_ES_PADRE:
        return (f"Cambia el crédito {r.get('Credito_Renovado_Hijo')} por la renovación "
                f"{r.get('SQL_Renovacion_Hija') or '(ver base de datos)'} antes de mandarla.")
    if r.get("Datos_Incompletos"):
        return "Completa el dato que falta (correo original o iTop) antes de mandarla."
    if r.get("Error_Envio"):
        return "Salió con datos equivocados: avisa a quien mandó el ticket único para corregirlo."
    if r.get("Cerrado_Por_Error"):
        return "Mándala hoy aunque el incidente esté cerrado (se cerró por error) y avisa a quien lo cerró."
    if r.get("No_Enviada"):
        return "No salió en su día: mándala hoy."
    if r.get("SQL_Estado") in SQL_PROBLEMAS:
        return "La base no la confirma: verifica cliente y crédito con la sucursal antes de mandarla."
    malas = [t for e, t in checks if e == MAL]
    return f"Revisa: {malas[0]}" if malas else "Revísala antes de mandarla."


def veredicto(r: dict, en_lote: bool = None, validada: bool = False, descartada: bool = False,
              sql_activo: bool = True) -> dict:
    """
    {'nivel', 'corto' (para la tabla), 'titulo' (banner), 'accion' (qué hacer), 'checks'}.
    `en_lote`: si va en el ticket único de hoy (por omisión, va_en_lote).
    """
    checks = comprobaciones(r, sql_activo)
    en_lote = va_en_lote(r) if en_lote is None else en_lote
    estado = r.get("Lote_Estado")
    if descartada:
        return dict(nivel=DESCARTADA, corto="Descartada", titulo="Descartada",
                    accion="La descartaste: no va en el lote. «Restaurar» la regresa.", checks=checks)
    if r.get("Informativa") and not en_lote:
        tipo = r.get("Informativa_Tipo")
        titulo = {INFO_EMISION: "Ya se emitió · no se manda", INFO_CERRADA: "Incidente resuelto · no se manda",
                  INFO_NO_PROCEDE: "No procede · no se manda"}.get(tipo, "No se manda")
        accion = {INFO_EMISION: "No la mandes: ya pasó la emisión y no se puede cancelar. Solo es informativa.",
                  INFO_CERRADA: "No la mandes: el incidente ya se resolvió en iTop.",
                  INFO_NO_PROCEDE: f"No la mandes. Si el incidente sigue abierto, respóndelo con: «{LEYENDA_NO_PROCEDE}»."
                  }.get(tipo, "No la mandes.")
        return dict(nivel=NO_SE_MANDA, corto="No se manda", titulo=titulo, accion=accion, checks=checks)
    if not en_lote:
        if estado == LOTE_ENVIADA:
            return dict(nivel=ENVIADA, corto="Enviada", titulo="Ya enviada",
                        accion=f"Nada que hacer: salió en el ticket único del {_f(r.get('Enviado_Fecha'))}.", checks=checks)
        if estado == LOTE_SIGUIENTE:
            return dict(nivel=SIGUIENTE, corto="Siguiente", titulo="Va en el siguiente ticket único",
                        accion=f"No va hoy: va en el ticket único del {_f(r.get('Dia_Lote'))}.", checks=checks)
        if estado == LOTE_CERRADA_SIN_ENVIO:
            return dict(nivel=NO_SE_MANDA, corto="No se manda", titulo="No se envió y ya se cerró",
                        accion="No salió en su día y el incidente ya se cerró: no entra al lote («Restaurar» la incluye).",
                        checks=checks)
        return dict(nivel=SIN_COMPROBAR, corto="Sin comprobar", titulo="Sin comprobar",
                    accion="No se encontró el ticket único de su día: revisa si salió («Restaurar» la suma al lote).",
                    checks=checks)
    if estado == LOTE_EN_CONCENTRADO:
        return dict(nivel=ENVIADA, corto="Enviada", titulo="Ya va en el ticket único de hoy",
                    accion="Nada que hacer: ya salió en el ticket único de hoy.", checks=checks)
    peor = MAL if any(e == MAL for e, _ in checks) else (AVISO if any(e == AVISO for e, _ in checks) else OK)
    if peor == MAL:
        return dict(nivel=CORREGIR, corto="Corregir", titulo="Corrige antes de enviar",
                    accion=_accion_corregir(r, checks), checks=checks)
    if r.get("Requiere_Autorizacion"):
        return dict(nivel=AUTORIZACION, corto="Autorización", titulo="Necesita autorización",
                    accion="Pasó el cierre de mes: mándala solo si ya tienes la autorización.", checks=checks)
    if peor == AVISO:
        avisos = [t.split(": ", 1)[-1] for e, t in checks if e == AVISO]
        return dict(nivel=REVISAR, corto="Revisar", titulo="Revisa antes de enviar",
                    accion=f"Revisa: {avisos[0]}" if avisos else "Revísala antes de mandarla.", checks=checks)
    return dict(nivel=LISTA, corto="✓ Validada" if validada else "Lista", titulo="Lista para enviar",
                accion=("Va en el ticket único de hoy. Ya la validaste." if validada
                        else "Va en el ticket único de hoy: si todo cuadra, márcala como validada (Espacio)."),
                checks=checks)


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
    confirma = estado in (SQL_OK, SQL_APP_PLAZO, SQL_PADRE_CORREGIDO)
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
        t_fecha = f"hoy ({_f(r.get('Dia_Lote'))})" if r.get("Dia_Lote") else "hoy"
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
