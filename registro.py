#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Registro de errores en validador.log, junto a la app (hasta 1 MB y 3 respaldos). Si algo falla en la extracción, en
las reglas o en un botón, aquí queda el detalle para revisarlo; la app sigue trabajando.
"""

import logging
import os
from logging.handlers import RotatingFileHandler

ARCHIVO = "validador.log"
_log = None


def ruta() -> str:
    from core_renovaciones import get_app_dir
    return os.path.join(get_app_dir(), ARCHIVO)


def log() -> logging.Logger:
    """El registro de la app (se crea la primera vez). Si no se puede escribir el archivo, no registra nada."""
    global _log
    if _log is None:
        _log = logging.getLogger("validador")
        _log.setLevel(logging.INFO)
        _log.propagate = False
        try:
            manejador = RotatingFileHandler(ruta(), maxBytes=1_000_000, backupCount=3, encoding="utf-8")
            manejador.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s", "%Y-%m-%d %H:%M:%S"))
            _log.addHandler(manejador)
        except OSError:
            _log.addHandler(logging.NullHandler())
    return _log


def error(contexto: str):
    """Registra la excepción que se está atendiendo (llamar dentro de un except)."""
    log().exception(contexto)
