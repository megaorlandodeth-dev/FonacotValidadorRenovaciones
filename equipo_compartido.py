#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Trabajo en equipo por una carpeta compartida (sin servidor ni base de datos nueva):
- Cada analista publica en la carpeta un archivo marcas_<gafete>.json con las renovaciones que validó y las dudas
  que le pasa al validador, más su avance del día.
- El validador lee todos los archivos y ve en su lote quién marcó qué y a qué hora.

Cada quien escribe solo su archivo (se escribe a un temporal y luego se reemplaza, para que nadie lea un archivo a
medias), así que no hay choques entre equipos. Funciona con una carpeta de red (\\\\servidor\\compartida), OneDrive o
SharePoint sincronizado.
"""

import json
import os
import re
import socket
from datetime import date, datetime

VERSION = 1
PREFIJO = "marcas_"
VALIDADA, DUDA = "validada", "duda"
# Un analista sin publicar en este tiempo se muestra como «sin señal»
MINUTOS_SIN_SENAL = 30


def _limpio(gafete: str) -> str:
    return re.sub(r"[^0-9A-Za-z_-]", "", str(gafete or "")) or "sin_gafete"


def archivo_de(carpeta: str, gafete: str) -> str:
    return os.path.join(carpeta, f"{PREFIJO}{_limpio(gafete)}.json")


def carpeta_valida(carpeta: str) -> str:
    """'' si se puede escribir en la carpeta; si no, el motivo."""
    if not carpeta:
        return "No hay carpeta compartida configurada."
    if not os.path.isdir(carpeta):
        return f"No se encuentra la carpeta {carpeta}."
    prueba = os.path.join(carpeta, f".prueba_{_limpio(socket.gethostname())}.tmp")
    try:
        with open(prueba, "w", encoding="utf-8") as f:
            f.write("ok")
        os.remove(prueba)
    except OSError as e:
        return f"No se puede escribir en la carpeta: {e.strerror or e}."
    return ""


def marca(row: dict, estado: str, hora: datetime = None, nota: str = "", llaves=()) -> dict:
    """Lo que se publica de una renovación (con datos para reconocerla aunque la llave no coincida)."""
    return {
        "estado": estado,
        "hora": (hora or datetime.now()).isoformat(timespec="seconds"),
        "nota": nota or "",
        "ticket": row.get("Numero_Ticket", ""),
        "cliente": str(row.get("Numero_Cliente", "")),
        "credito": str(row.get("Credito_Renovado_Hijo", "")),
        "padre_enviado": str(row.get("Credito_Padre_Enviado", "") or ""),
        "llaves": list(llaves),
    }


def publicar(carpeta: str, gafete: str, nombre: str, dia: date, marcas: dict, avance: dict = None) -> str:
    """
    Escribe marcas_<gafete>.json (completo, reemplazando el anterior). `marcas`: {llave: marca(...)}.
    `avance`: {'total': n, 'listas': n, 'revisar': n, 'validadas': n}. Regresa '' si se publicó o el error.
    """
    datos = {
        "version": VERSION, "gafete": str(gafete), "nombre": nombre or "", "equipo": socket.gethostname(),
        "actualizado": datetime.now().isoformat(timespec="seconds"), "dia": dia.isoformat() if dia else "",
        "avance": avance or {}, "marcas": marcas,
    }
    destino = archivo_de(carpeta, gafete)
    temporal = f"{destino}.{_limpio(socket.gethostname())}.tmp"
    try:
        with open(temporal, "w", encoding="utf-8") as f:
            json.dump(datos, f, ensure_ascii=False, indent=1)
        os.replace(temporal, destino)
    except OSError as e:
        try:
            os.remove(temporal)
        except OSError:
            pass
        return f"No se pudo publicar en la carpeta compartida: {e.strerror or e}."
    return ""


def leer_equipo(carpeta: str, dia: date, ahora: datetime = None) -> dict:
    """
    Lee los archivos de todos los analistas para el día `dia`:
    {'analistas': [{gafete, nombre, equipo, actualizado (datetime), sin_senal, avance, marcas: n}],
     'marcas': {llave: marca + gafete/nombre}, 'errores': [texto]}.
    Se ignoran los archivos de otro día; si dos analistas marcan la misma renovación gana la marca más reciente.
    """
    ahora = ahora or datetime.now()
    res = {"analistas": [], "marcas": {}, "errores": []}
    try:
        nombres = sorted(n for n in os.listdir(carpeta) if n.startswith(PREFIJO) and n.endswith(".json"))
    except OSError as e:
        res["errores"].append(f"No se puede leer la carpeta compartida: {e.strerror or e}.")
        return res
    for n in nombres:
        try:
            with open(os.path.join(carpeta, n), "r", encoding="utf-8") as f:
                datos = json.load(f)
            if datos.get("dia") != (dia.isoformat() if dia else ""):
                continue
            actualizado = datetime.fromisoformat(datos.get("actualizado"))
        except (OSError, ValueError, TypeError, AttributeError) as e:
            res["errores"].append(f"{n}: no se pudo leer ({e.__class__.__name__}).")
            continue
        gafete, nombre = str(datos.get("gafete", "")), datos.get("nombre", "")
        marcas = datos.get("marcas") or {}
        res["analistas"].append({
            "gafete": gafete, "nombre": nombre, "equipo": datos.get("equipo", ""), "actualizado": actualizado,
            "sin_senal": (ahora - actualizado).total_seconds() > MINUTOS_SIN_SENAL * 60,
            "avance": datos.get("avance") or {}, "marcas": len(marcas),
        })
        for llave, m in marcas.items():
            if not isinstance(m, dict):
                continue
            m = dict(m, gafete=gafete, nombre=nombre)
            previa = res["marcas"].get(llave)
            if previa is None or str(m.get("hora", "")) > str(previa.get("hora", "")):
                res["marcas"][llave] = m
    return res


def marca_de_fila(llaves, marcas: dict):
    """La marca del equipo para una renovación, buscando por cualquiera de sus llaves (None si nadie la marcó)."""
    for llave in llaves:
        if llave in marcas:
            return marcas[llave]
    for m in marcas.values():
        if set(m.get("llaves") or ()) & set(llaves):
            return m
    return None
