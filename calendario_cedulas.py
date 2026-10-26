#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Calendario de Emisión de Cédulas de FONACOT: una renovación ya no se puede cancelar si después de que se dio de alta
pasó la generación de una emisión o un cierre de mes (el último día de cada mes).

Las fechas de «Generación de Emisión» vienen de https://www.fonacot.gob.mx/empresa/Paginas/calendariocedulas.aspx.
Las de 2026 van incluidas; una vez por semana se intenta leer la página para tomar las nuevas (si no hay internet,
se usan las que ya se tienen). Se guardan en calendario_cedulas.json, junto a la configuración.
"""

import calendar
import json
import os
import re
import unicodedata
from datetime import date, datetime, timedelta

URL_CALENDARIO = "https://www.fonacot.gob.mx/empresa/Paginas/calendariocedulas.aspx"
DIAS_ENTRE_ACTUALIZACIONES = 7

_MESES = ("enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre",
          "noviembre", "diciembre")
_MESES_CORTOS = {m[:3]: i for i, m in enumerate(_MESES, 1)}

# (año, mes de la emisión): día en que se generó. Ejercicio 2026 (última actualización de la página: 03/06/2026)
EMISIONES_INCLUIDAS = {
    (2026, 1): date(2025, 12, 14), (2026, 2): date(2026, 1, 25), (2026, 3): date(2026, 2, 24),
    (2026, 4): date(2026, 3, 24), (2026, 5): date(2026, 4, 23), (2026, 6): date(2026, 5, 24),
    (2026, 7): date(2026, 6, 24), (2026, 8): date(2026, 7, 23), (2026, 9): date(2026, 8, 24),
    (2026, 10): date(2026, 9, 23), (2026, 11): date(2026, 10, 25), (2026, 12): date(2026, 11, 24),
    (2027, 1): date(2026, 12, 13),
}

_emisiones = dict(EMISIONES_INCLUIDAS)
_cargado = False


def _archivo() -> str:
    from core_renovaciones import get_app_dir
    return os.path.join(get_app_dir(), "calendario_cedulas.json")


def _sin_acentos(texto: str) -> str:
    texto = unicodedata.normalize("NFKD", str(texto or "").replace("​", ""))
    return "".join(c for c in texto if not unicodedata.combining(c)).lower()


def leer_pagina(html: str) -> dict:
    """{(año, mes): generación} de la tabla de la página («NOVIEMBRE '26 | dom, 25-oct '26 | …»)."""
    texto = re.sub(r"<[^>]+>", " ", html or "")
    texto = re.sub(r"&nbsp;|&#160;", " ", texto)
    texto = re.sub(r"&#39;|&#x27;|&rsquo;|’|‘", "'", texto)
    texto = re.sub(r"\s+", " ", _sin_acentos(texto))
    emisiones = {}
    patron = (r"\b(" + "|".join(_MESES) + r")\s*'\s*(\d{2})\s+[a-z]{2,4},?\s*(\d{1,2})\s*-\s*([a-z]{3})\s*'\s*(\d{2})")
    for m in re.finditer(patron, texto):
        mes, anio = _MESES.index(m.group(1)) + 1, 2000 + int(m.group(2))
        if m.group(4) not in _MESES_CORTOS:
            continue
        try:
            emisiones[(anio, mes)] = date(2000 + int(m.group(5)), _MESES_CORTOS[m.group(4)], int(m.group(3)))
        except ValueError:
            continue
    return emisiones


def cargar():
    """Junta las fechas incluidas con las que se bajaron antes."""
    global _cargado
    _cargado = True
    try:
        with open(_archivo(), encoding="utf-8") as f:
            datos = json.load(f)
        for llave, valor in (datos.get("emisiones") or {}).items():
            anio, mes = (int(x) for x in llave.split("-"))
            _emisiones[(anio, mes)] = date.fromisoformat(valor)
    except Exception:
        pass


def actualizar_calendario(forzar: bool = False, timeout: int = 8) -> bool:
    """Lee la página de FONACOT (una vez por semana). True si se tomaron fechas nuevas. Nunca lanza errores."""
    if not _cargado:
        cargar()
    try:
        if not forzar and os.path.exists(_archivo()):
            with open(_archivo(), encoding="utf-8") as f:
                ultima = datetime.fromisoformat(json.load(f).get("actualizado", "2000-01-01T00:00:00"))
            if datetime.now() - ultima < timedelta(days=DIAS_ENTRE_ACTUALIZACIONES):
                return False
    except Exception:
        pass
    try:
        import urllib.request
        pedido = urllib.request.Request(URL_CALENDARIO, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(pedido, timeout=timeout) as resp:
            crudo = resp.read()
            juego = resp.headers.get_content_charset() or "utf-8"
        try:
            html = crudo.decode(juego)
        except (UnicodeDecodeError, LookupError):
            html = crudo.decode("cp1252", errors="replace")
        nuevas = leer_pagina(html)
    except Exception:
        nuevas = {}
    _emisiones.update(nuevas)
    try:
        with open(_archivo(), "w", encoding="utf-8") as f:
            json.dump({"actualizado": datetime.now().isoformat(timespec="seconds"), "fuente": URL_CALENDARIO,
                       "emisiones": {f"{a}-{m:02d}": g.isoformat() for (a, m), g in sorted(_emisiones.items())}},
                      f, indent=2, ensure_ascii=False)
    except Exception:
        pass
    return bool(nuevas)


def emisiones() -> dict:
    if not _cargado:
        cargar()
    return dict(_emisiones)


def ultimo_dia_del_mes(d: date) -> date:
    return date(d.year, d.month, calendar.monthrange(d.year, d.month)[1])


def corte_que_bloquea(alta: date, llegada: date):
    """
    El primer corte que pasó entre el alta de la renovación y la llegada del incidente: (fecha, texto), o None si se
    puede cancelar. Lo que cuenta es cuándo se pidió: si el incidente llegó a tiempo (el día de la renovación, aunque
    después del envío del ticket único) se manda al día siguiente aunque en medio haya una emisión o un cierre de mes.
    - Generación de una emisión: bloquea lo dado de alta ese día o antes y pedido después.
    - Cierre de mes (último día del mes): bloquea lo del mes pedido ya en el mes siguiente.
    """
    if not alta or not llegada or alta >= llegada:
        return None
    cortes = []
    for (anio, mes), gen in emisiones().items():
        if alta <= gen < llegada:
            cortes.append((gen, f"la generación de la emisión de {_MESES[mes - 1]} {anio} ({gen:%d/%m/%Y})"))
    cierre = ultimo_dia_del_mes(alta)
    if cierre < llegada:
        cortes.append((cierre, f"el cierre de mes ({cierre:%d/%m/%Y})"))
    return min(cortes) if cortes else None


def marcar_bloqueos(rows: list) -> dict:
    """
    Pone 'Bloqueo_Emision' (texto) en las renovaciones que ya no se pueden cancelar: dadas de alta ('SQL_Fecha_Alta',
    de la base) antes de una emisión o de un cierre de mes que pasó antes de que llegara el incidente. Sin la fecha de
    alta no se marca (se necesita la validación en la base).
    """
    n = 0
    for r in rows:
        r['Bloqueo_Emision'] = ''
        alta = r.get('SQL_Fecha_Alta')
        llegada = r.get('Fecha_Llegada_DT') or r.get('Fecha_DT')
        llegada = llegada.date() if isinstance(llegada, datetime) else llegada
        corte = corte_que_bloquea(alta, llegada)
        if corte:
            r['Bloqueo_Emision'] = (f"YA NO SE PUEDE CANCELAR: se dio de alta el {alta:%d/%m/%Y} y el incidente llegó "
                                    f"el {llegada:%d/%m/%Y}, después de {corte[1]}")
            n += 1
    return {'bloqueadas': n}
