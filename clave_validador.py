#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Contraseña del perfil Validador.

No se guarda la contraseña: solo su huella (PBKDF2-SHA256 con sal) en validador_clave.json, junto a la app. Ese archivo
se reparte con la app para que en todos los equipos se pida la misma contraseña. Quien la sabe puede pedir que se
recuerde en su equipo: entonces va al Administrador de credenciales de Windows (de su usuario de Windows).
"""

import hashlib
import hmac
import json
import os
import secrets

from core_renovaciones import get_app_dir, guardar_json

# Junto al .exe (o al script), no dentro de la carpeta interna del ejecutable
ARCHIVO = os.path.join(get_app_dir(), "validador_clave.json")
# Si falta el archivo, solo se puede crear la contraseña abriendo la app con esta opción (quien la instala)
OPCION_CREAR = "--crear-clave-validador"
ITERACIONES = 200_000
MINIMO = 6
_CREDENCIAL = "FonacotValidadorRenovaciones/perfil_validador"


def _huella(clave: str, sal: bytes, iteraciones: int) -> str:
    return hashlib.pbkdf2_hmac("sha256", clave.encode("utf-8"), sal, iteraciones).hex()


def _leer(archivo: str = None) -> dict:
    try:
        with open(archivo or ARCHIVO, "r", encoding="utf-8") as f:
            datos = json.load(f)
        return datos if datos.get("huella") and datos.get("sal") else {}
    except (OSError, ValueError, AttributeError):
        return {}


def existe(archivo: str = None) -> bool:
    return bool(_leer(archivo))


def crear(clave: str, archivo: str = None) -> str:
    """Guarda la huella de la contraseña. Regresa '' o el motivo por el que no se pudo."""
    if len(clave or "") < MINIMO:
        return f"La contraseña debe tener al menos {MINIMO} caracteres."
    sal = secrets.token_bytes(16)
    datos = {"version": 1, "algoritmo": "pbkdf2_sha256", "iteraciones": ITERACIONES, "sal": sal.hex(),
             "huella": _huella(clave, sal, ITERACIONES)}
    try:
        guardar_json(archivo or ARCHIVO, datos, indent=1)
    except OSError as e:
        return f"No se pudo guardar la contraseña: {e.strerror or e}."
    return ""


def verificar(clave: str, archivo: str = None) -> bool:
    datos = _leer(archivo)
    if not datos or not clave:
        return False
    try:
        calculada = _huella(clave, bytes.fromhex(datos["sal"]), int(datos.get("iteraciones") or ITERACIONES))
    except (ValueError, TypeError):
        return False
    return hmac.compare_digest(calculada, datos["huella"])


# ---- Recordar en este equipo (Administrador de credenciales de Windows) ----

def recordada() -> str:
    try:
        import win32cred
        cred = win32cred.CredRead(_CREDENCIAL, win32cred.CRED_TYPE_GENERIC)
        blob = cred.get("CredentialBlob") or b""
        return blob.decode("utf-16-le") if isinstance(blob, bytes) else str(blob)
    except Exception:
        return ""


def recordar(clave: str) -> bool:
    try:
        import win32cred
        win32cred.CredWrite({
            "Type": win32cred.CRED_TYPE_GENERIC, "TargetName": _CREDENCIAL, "UserName": "validador",
            "CredentialBlob": clave, "Persist": win32cred.CRED_PERSIST_LOCAL_MACHINE,
            "Comment": "Validador de Renovaciones FONACOT (perfil validador)"}, 0)
        return True
    except Exception:
        return False


def olvidar():
    try:
        import win32cred
        win32cred.CredDelete(_CREDENCIAL, win32cred.CRED_TYPE_GENERIC, 0)
    except Exception:
        pass
