#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Validación de las renovaciones contra la base de datos (Oracle, tabla TOPAZ.CR_RENOVACIONES).

Solo consulta (SELECT) tres columnas: FECHA_ALTA, CUENTA (No. de cliente) y NO_RENOVACION (crédito renovado):
- ¿el cliente tiene una renovación con ese número y se dio de alta el día del incidente?
- si al incidente le falta el crédito: la renovación que el cliente tiene ese día;
- si le falta el cliente: de qué cliente es esa renovación.

La contraseña no se guarda en la configuración: se pide en pantalla y, si el usuario quiere, se guarda en el
Administrador de credenciales de Windows.
"""

import re
from datetime import date, datetime, timedelta

TABLA_DEFAULT = "TOPAZ.CR_RENOVACIONES"
DIAS_ANTES = 7          # la renovación pudo darse de alta unos días antes de que llegara el incidente
LOTE_CONSULTA = 500     # valores por consulta (Oracle admite hasta 1000 en un IN)
TIEMPO_MAX_MS = 60000

SQL_OK = "OK"
SQL_OTRA_FECHA = "OTRA_FECHA"
SQL_NO_EXISTE = "NO_EXISTE"
SQL_SIN_RENOVACIONES = "SIN_RENOVACIONES"
SQL_SUGERENCIA = "SUGERENCIA"
SQL_SIN_DATOS = "SIN_DATOS"
# Estados que indican que el dato del incidente no cuadra con la base
SQL_PROBLEMAS = (SQL_OTRA_FECHA, SQL_NO_EXISTE, SQL_SIN_RENOVACIONES)

_CREDENCIAL = "FonacotValidadorRenovaciones/SQL"
_cliente_oracle_iniciado = False


class ErrorSQL(Exception):
    """No se pudo conectar o consultar: la extracción sigue y solo se avisa."""


def _num(valor) -> str:
    """Dígitos sin ceros a la izquierda ('' si no hay número): '0010430' y 10430 son el mismo."""
    digitos = re.sub(r'\D', '', str(valor if valor is not None else ''))
    return str(int(digitos)) if digitos else ''


def _fecha(valor):
    if isinstance(valor, datetime):
        return valor.date()
    return valor if isinstance(valor, date) else None


# ======================= Conexión =======================

def tabla_valida(tabla: str) -> str:
    """El nombre de la tabla va dentro del SQL (no admite parámetro): solo ESQUEMA.TABLA con letras, números y _."""
    tabla = (tabla or TABLA_DEFAULT).strip()
    if not re.fullmatch(r'[A-Za-z][A-Za-z0-9_$#]*(\.[A-Za-z][A-Za-z0-9_$#]*)?', tabla):
        raise ErrorSQL(f"Nombre de tabla no válido: «{tabla}» (usa ESQUEMA.TABLA)")
    return tabla


def armar_dsn(servidor: str, oracledb=None) -> str:
    """
    Acepta lo que muestra SQL Developer: «host:puerto/servicio», «host:puerto:SID», un alias de tnsnames
    o el descriptor completo «(DESCRIPTION=…)».
    """
    s = (servidor or "").strip()
    m = re.fullmatch(r'([^:/\s()]+):(\d+):([A-Za-z0-9_$#.]+)', s)
    if m and oracledb is not None:
        return oracledb.makedsn(m.group(1), int(m.group(2)), sid=m.group(3))
    return s


def _mensaje(e: Exception) -> str:
    """Texto del error en una línea, sin el identificador interno de la conexión (la causa viene en la 2a línea)."""
    lineas = [re.sub(r'\s*\(CONNECTION_ID=[^)]*\)', '', ln).strip() for ln in str(e).strip().splitlines()]
    return "  ·  ".join(ln for ln in lineas if ln)[:300] or type(e).__name__


def conectar(servidor: str, usuario: str, password: str, cliente_oracle: str = ""):
    """
    Conexión de solo consulta. Usa el modo «thin» de python-oracledb (no requiere instalar Oracle Client); si la
    base exige el modo clásico, se puede indicar la carpeta del Oracle Instant Client en `cliente_oracle`.
    """
    global _cliente_oracle_iniciado
    try:
        import oracledb
    except ImportError:
        raise ErrorSQL("Falta el paquete «oracledb». Instálalo con:  python -m pip install oracledb")
    if not (servidor and usuario and password):
        raise ErrorSQL("Faltan servidor, usuario o contraseña de la base de datos")
    if cliente_oracle and not _cliente_oracle_iniciado:
        try:
            oracledb.init_oracle_client(lib_dir=cliente_oracle)
            _cliente_oracle_iniciado = True
        except Exception as e:
            raise ErrorSQL(f"No se pudo usar el Oracle Client de «{cliente_oracle}»: {e}")
    try:
        conn = oracledb.connect(user=usuario, password=password, dsn=armar_dsn(servidor, oracledb),
                                tcp_connect_timeout=10)
    except Exception as e:
        texto = _mensaje(e)
        if "DPY-3015" in texto or "DPY-3010" in texto:
            texto += ("  →  Esta base pide el modo clásico de Oracle: indica en la configuración "
                      "(«sql_cliente_oracle») la carpeta del Oracle Instant Client.")
        raise ErrorSQL(texto)
    try:
        conn.call_timeout = TIEMPO_MAX_MS
    except Exception:
        pass
    return conn


def probar_conexion(servidor: str, usuario: str, password: str, tabla: str = TABLA_DEFAULT, cliente_oracle: str = "") -> str:
    """Conecta y lee un renglón de la tabla. Devuelve el texto para mostrar; lanza ErrorSQL si algo falla."""
    tabla = tabla_valida(tabla)
    conn = conectar(servidor, usuario, password, cliente_oracle)
    try:
        cur = conn.cursor()
        try:
            cur.execute(f"SELECT FECHA_ALTA, CUENTA, NO_RENOVACION FROM {tabla} WHERE ROWNUM = 1")
            cur.fetchall()
        except Exception as e:
            raise ErrorSQL(f"Conectó, pero no se pudo leer {tabla}: {_mensaje(e)}")
        finally:
            cur.close()
    finally:
        conn.close()
    return f"Conexión correcta: se puede leer {tabla}"


def _lotes(valores: list):
    for i in range(0, len(valores), LOTE_CONSULTA):
        yield valores[i:i + LOTE_CONSULTA]


def consultar_por_clientes(conn, tabla: str, clientes) -> list:
    """[(fecha_alta, cuenta, no_renovacion)] de todas las renovaciones de esos clientes."""
    tabla = tabla_valida(tabla)
    registros = []
    cur = conn.cursor()
    try:
        for grupo in _lotes(sorted(set(c for c in clientes if c))):
            binds = {f"c{i}": v for i, v in enumerate(grupo)}
            cur.execute(f"SELECT FECHA_ALTA, CUENTA, NO_RENOVACION FROM {tabla} "
                        f"WHERE CUENTA IN ({', '.join(':' + k for k in binds)})", binds)
            registros.extend(cur.fetchall())
    finally:
        cur.close()
    return [(_fecha(f), _num(c), _num(n)) for f, c, n in registros]


def consultar_por_creditos(conn, tabla: str, creditos, desde: date, hasta: date) -> list:
    """[(fecha_alta, cuenta, no_renovacion)] de esas renovaciones dadas de alta entre `desde` y `hasta`."""
    tabla = tabla_valida(tabla)
    registros = []
    cur = conn.cursor()
    try:
        for grupo in _lotes(sorted(set(c for c in creditos if c))):
            binds = {f"r{i}": v for i, v in enumerate(grupo)}
            binds.update(desde=datetime.combine(desde, datetime.min.time()),
                         hasta=datetime.combine(hasta + timedelta(days=1), datetime.min.time()))
            marcas = ', '.join(':' + k for k in binds if k.startswith('r'))
            cur.execute(f"SELECT FECHA_ALTA, CUENTA, NO_RENOVACION FROM {tabla} "
                        f"WHERE NO_RENOVACION IN ({marcas}) AND FECHA_ALTA >= :desde AND FECHA_ALTA < :hasta", binds)
            registros.extend(cur.fetchall())
    finally:
        cur.close()
    return [(_fecha(f), _num(c), _num(n)) for f, c, n in registros]


# ======================= Validación (sin base de datos: se puede probar sola) =======================

def _dato(valor) -> str:
    v = str(valor or '')
    return '' if v.upper().startswith('SIN ') else _num(v)


def _referencia(r: dict):
    llegada = r.get('Fecha_Llegada_DT') or r.get('Fecha_DT')
    return llegada.date() if isinstance(llegada, datetime) else (llegada or date.today())


def validar_par(cliente: str, credito: str, ref: date, por_cliente: dict, por_credito: dict) -> tuple:
    """
    (estado, detalle) de un Cliente + Crédito contra lo que trajo la base. `ref`: día en que debía hacerse la
    renovación (llegada del incidente o día del ticket único). `por_cliente`: {cuenta: [(fecha, no_renov)]},
    `por_credito`: {no_renov: [(fecha, cuenta)]}.
    """
    desde, hasta = ref - timedelta(days=DIAS_ANTES), ref + timedelta(days=1)

    def lista(regs):
        return ", ".join(f"{n} ({f:%d/%m/%Y})" for f, n in sorted(regs, key=lambda x: x[0] or date.min, reverse=True)[:4])

    def de_otro_cliente(cred):
        otros = sorted({c for f, c in por_credito.get(cred, []) if f and desde <= f <= hasta and c != cliente})
        return f" La renovación {cred} de esos días es del cliente {', '.join(otros)}." if otros else ""

    renov = por_cliente.get(cliente, []) if cliente else []
    cercanas = [(f, n) for f, n in renov if f and desde <= f <= hasta]
    if cliente and credito:
        iguales = [f for f, n in renov if n == credito]
        if iguales:
            alta = max(f for f in iguales if f) if any(iguales) else None
            if alta and desde <= alta <= hasta:
                cuando = "el mismo día" if alta == ref else f"el {alta:%d/%m/%Y}"
                return SQL_OK, f"Renovación {credito} del cliente {cliente} dada de alta {cuando}"
            otra = f" Ese día tiene: {lista(cercanas)}." if cercanas else ""
            return SQL_OTRA_FECHA, (f"La renovación {credito} del cliente {cliente} es del "
                                    f"{alta:%d/%m/%Y}, no de cuando llegó el incidente ({ref:%d/%m/%Y}).{otra}"
                                    if alta else f"La renovación {credito} no tiene fecha de alta")
        if renov:
            sugerencia = f" ¿Será {', '.join(n for _, n in cercanas)}?" if cercanas else ""
            return SQL_NO_EXISTE, (f"El cliente {cliente} no tiene la renovación {credito}. Sus renovaciones: "
                                   f"{lista(renov)}.{sugerencia}{de_otro_cliente(credito)}")
        return SQL_SIN_RENOVACIONES, f"El cliente {cliente} no tiene renovaciones.{de_otro_cliente(credito)}"
    if cliente:
        if cercanas:
            return SQL_SUGERENCIA, f"Falta el crédito: el cliente tiene la renovación {lista(cercanas)}"
        if renov:
            return SQL_NO_EXISTE, f"El cliente {cliente} no tiene renovaciones de esos días (la última: {lista(renov[:1])})"
        return SQL_SIN_RENOVACIONES, f"El cliente {cliente} no tiene renovaciones"
    if credito:
        duenos = sorted({(f, c) for f, c in por_credito.get(credito, []) if f and desde <= f <= hasta})
        if duenos:
            return SQL_SUGERENCIA, "Falta el cliente: la renovación {} es del cliente {}".format(
                credito, ", ".join(f"{c} (alta {f:%d/%m/%Y})" for f, c in duenos))
        return SQL_NO_EXISTE, f"No hay una renovación {credito} dada de alta cerca del {ref:%d/%m/%Y}"
    return SQL_SIN_DATOS, ""


def validar_contra_base(conn, tabla: str, rows: list, filas_enviadas: list = ()) -> dict:
    """
    Consulta la base y marca cada renovación con 'SQL_Estado' y 'SQL_Detalle'. Los renglones de los tickets
    únicos (`filas_enviadas`, con 'cliente', 'credito', 'fecha') se marcan con 'sql_estado' y 'sql_detalle'.
    Devuelve el resumen {'ok': True, 'conteos': {estado: n}, 'problemas': n}.
    """
    pares = [(_dato(r.get('Numero_Cliente')), _dato(r.get('Credito_Renovado_Hijo')), _referencia(r)) for r in rows]
    pares_env = [(_dato(f.get('cliente')), _dato(f.get('credito')), f.get('fecha') or date.today()) for f in filas_enviadas]
    todos = pares + pares_env

    por_cliente = {}
    for f, c, n in consultar_por_clientes(conn, tabla, [c for c, _, _ in todos]):
        por_cliente.setdefault(c, []).append((f, n))
    # Por crédito: los que no traen cliente y los que el cliente no tiene (¿de quién es esa renovación?)
    buscar = [(cr, ref) for c, cr, ref in todos
              if cr and (not c or cr not in {n for _, n in por_cliente.get(c, [])})]
    por_credito = {}
    if buscar:
        refs = [ref for _, ref in buscar]
        for f, c, n in consultar_por_creditos(conn, tabla, [cr for cr, _ in buscar],
                                              min(refs) - timedelta(days=DIAS_ANTES), max(refs) + timedelta(days=1)):
            por_credito.setdefault(n, []).append((f, c))

    conteos = {}
    for r, (c, cr, ref) in zip(rows, pares):
        r['SQL_Estado'], r['SQL_Detalle'] = validar_par(c, cr, ref, por_cliente, por_credito)
        conteos[r['SQL_Estado']] = conteos.get(r['SQL_Estado'], 0) + 1
    for f, (c, cr, ref) in zip(filas_enviadas, pares_env):
        f['sql_estado'], f['sql_detalle'] = validar_par(c, cr, ref, por_cliente, por_credito)
    return {'ok': True, 'conteos': conteos, 'problemas': sum(conteos.get(e, 0) for e in SQL_PROBLEMAS),
            'enviadas_con_problema': sum(1 for f in filas_enviadas if f.get('sql_estado') in SQL_PROBLEMAS)}


def texto_corto(estado: str) -> str:
    return {SQL_OK: "✔ Existe", SQL_OTRA_FECHA: "⚠ Otra fecha", SQL_NO_EXISTE: "✖ No existe",
            SQL_SIN_RENOVACIONES: "✖ Sin renovaciones", SQL_SUGERENCIA: "💡 Sugerencia", SQL_SIN_DATOS: "—"}.get(estado, "")


# ======================= Contraseña en el Administrador de credenciales de Windows =======================

def _llave(usuario: str, servidor: str) -> str:
    return f"{_CREDENCIAL}/{usuario}@{servidor}"


def leer_password(usuario: str, servidor: str) -> str:
    try:
        import win32cred
        cred = win32cred.CredRead(_llave(usuario, servidor), win32cred.CRED_TYPE_GENERIC)
        blob = cred.get('CredentialBlob') or b''
        return blob.decode('utf-16-le') if isinstance(blob, bytes) else str(blob)
    except Exception:
        return ""


def guardar_password(usuario: str, servidor: str, password: str) -> bool:
    try:
        import win32cred
        win32cred.CredWrite({
            'Type': win32cred.CRED_TYPE_GENERIC, 'TargetName': _llave(usuario, servidor), 'UserName': usuario,
            'CredentialBlob': password, 'Persist': win32cred.CRED_PERSIST_LOCAL_MACHINE,
            'Comment': "Validador de Renovaciones FONACOT (consulta a CR_RENOVACIONES)"}, 0)
        return True
    except Exception:
        return False


def borrar_password(usuario: str, servidor: str):
    try:
        import win32cred
        win32cred.CredDelete(_llave(usuario, servidor), win32cred.CRED_TYPE_GENERIC, 0)
    except Exception:
        pass
