#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Validación de las renovaciones contra la base de datos (Oracle, tabla TOPAZ.CR_RENOVACIONES).

Solo consulta (SELECT). Primera consulta: FECHA_ALTA, CUENTA (No. de cliente) y NO_RENOVACION (crédito renovado):
- ¿el cliente tiene una renovación con ese número y se dio de alta el día del incidente?
- si al incidente le falta el crédito: la renovación que el cliente tiene ese día;
- si le falta el cliente: de qué cliente es esa renovación.
Segunda consulta (SALDOS, CR_DISPOSICIONES, CR_CEDULAS_DETALLE), para las de días anteriores: solo se envían si las
generó la APP (sucursal 104) con un plazo menor; si no, «No procede, la renovación es de otro día». Las que manda la
Mesa de Crédito (renovaciones de otros días) sí se envían: solo las detiene la emisión o el cierre de mes (calendario).
También avisa si la renovación ya emitió (está en CR_CEDULAS_DETALLE).
Si el crédito que manda el incidente no es una renovación del cliente, se busca como crédito padre (OPERACION): si lo
es, se avisa cuál es su renovación (la que hay que cancelar).

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
SQL_APP_PLAZO = "APP_PLAZO"      # de otro día, pero la generó la APP con plazo menor: sí se envía
SQL_NO_PROCEDE = "NO_PROCEDE"    # de otro día: «No procede, la renovación es de otro día»
SQL_MESA_CREDITO = "MESA_CREDITO"  # de otro día, pero la manda la Mesa de Crédito: sí se envía
SQL_ES_PADRE = "ES_PADRE"        # mandaron el crédito padre (origen) en vez de la renovación
SQL_PADRE_CORREGIDO = "PADRE_CORREGIDO"  # traía el padre: se puso su renovación (hijo) como crédito a cancelar; validar
# Estados que indican que el dato del incidente no cuadra con la base
SQL_PROBLEMAS = (SQL_OTRA_FECHA, SQL_NO_EXISTE, SQL_SIN_RENOVACIONES, SQL_NO_PROCEDE, SQL_ES_PADRE)
# La base confirma la renovación y sí se envía
SQL_SE_ENVIA = (SQL_OK, SQL_APP_PLAZO, SQL_MESA_CREDITO)
LEYENDA_NO_PROCEDE = "No procede, la renovación es de otro día"
SUCURSAL_APP = 104

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


def servidor_de(host: str, puerto="", servicio: str = "") -> str:
    """Los tres campos de SQL Developer (Nombre del host, Puerto, Nombre del servicio) -> «host:puerto/servicio»."""
    host, servicio = (host or "").strip(), (servicio or "").strip()
    puerto = str(puerto or "").strip()
    if re.search(r'[:/\s,;]', host):
        # Pegaron «host:puerto/servicio» (o «1521/fonatop») en el campo del host
        h, p, s = separar_servidor(host)
        host, puerto, servicio = h, puerto or p, servicio or s
    puerto = puerto or "1521"
    if not host:
        return ""
    return f"{host}:{puerto}/{servicio}" if servicio else f"{host}:{puerto}"


def separar_servidor(servidor: str) -> tuple:
    """(host, puerto, servicio) de «host:puerto/servicio» (o con espacios, comas o diagonales). Al revés de servidor_de."""
    s = (servidor or "").strip()
    m = re.fullmatch(r'([^:/\s(),;]+)(?:[\s:,;/]+(\d{2,5}))?(?:[\s:,;/]+([A-Za-z0-9_$#.\-]+))?', s)
    if not m:
        return s, "", ""
    host, puerto, servicio = m.group(1), m.group(2) or "", m.group(3) or ""
    if host.isdigit():
        # «1521/fonatop» (sin host): el número es el puerto, no el servidor
        host, puerto = "", puerto or host
    return host, puerto, servicio


def host_invalido(servidor: str) -> str:
    """El problema si al servidor le falta el host (p. ej. sólo «1521/fonatop»); '' si se ve bien."""
    s = (servidor or "").strip()
    if not s or s.startswith("("):
        return ""
    if not separar_servidor(s)[0]:
        puerto = re.split(r'[:/\s,;]', s)[0]
        return ("Falta el Nombre del host (el de SQL Developer, p. ej. scanfonacot.fonacot.gob.mx): "
                f"«{puerto}» es el puerto, no el servidor.")
    return ""


def armar_dsn(servidor: str, oracledb=None) -> str:
    """
    Acepta «host:puerto/servicio» (lo de SQL Developer), «host:puerto:SID», el host, puerto y servicio separados por
    espacios o comas, un alias de tnsnames o el descriptor completo «(DESCRIPTION=…)».
    """
    s = (servidor or "").strip()
    if s.startswith("("):
        return s
    m = re.fullmatch(r'([^:/\s()]+):(\d+):([A-Za-z0-9_$#.\-]+)', s)
    if m and oracledb is not None:
        return oracledb.makedsn(m.group(1), int(m.group(2)), sid=m.group(3))
    host, puerto, servicio = separar_servidor(s)
    if host == s and re.search(r'[:/\s,;]', s):
        return s  # no se entendió: se le pasa a Oracle tal cual
    # «scanfonacot.fonacot.gob.mx 1521 fonatop»: sin esto, Oracle lo busca como alias en tnsnames.ora (DPY-4000)
    if puerto or "." in host:
        return servidor_de(host, puerto, servicio)
    return s


def _buscar_cliente_oracle() -> str:
    """Carpeta de un Oracle Client instalado (la de SQL Developer o la del sistema), para el modo clásico. '' si no hay."""
    import glob
    import os
    candidatos = []
    for var in ("ORACLE_HOME",):
        if os.environ.get(var):
            candidatos += [os.path.join(os.environ[var], "bin"), os.environ[var]]
    candidatos += [p for p in os.environ.get("PATH", "").split(os.pathsep) if p]
    for patron in (r"C:\app\*\product\*\*\bin", r"C:\app\client\*\product\*\*\bin", r"C:\oracle\*\bin",
                   r"C:\oracle\product\*\*\bin", r"C:\instantclient*", r"C:\oracle\instantclient*"):
        candidatos += sorted(glob.glob(patron), reverse=True)
    for carpeta in candidatos:
        if os.path.isfile(os.path.join(carpeta, "oci.dll")):
            return carpeta
    return ""


# oracledb 4 ya no acepta Oracle Client anteriores a 19 (el de SQL Developer de FONACOT es 12.2): la 3.4 sí
ORACLEDB_COMPATIBLE = "oracledb>=3.4,<4"


def _explicar_cliente_viejo(texto: str, carpeta: str) -> str:
    """DPI-1050 («Oracle Client library is at version 12.2 but version 19.1 or higher is needed»): qué hacer."""
    if "DPI-1050" not in texto:
        return texto
    try:
        import oracledb
        version = oracledb.__version__
    except Exception:
        version = "?"
    return (f"El Oracle Client de este equipo ({carpeta or 'el instalado'}) es anterior a 19 y la versión {version} "
            f"de oracledb ya no lo acepta.  →  Cierra la app y ábrela con Ejecutar_Validador_Renovaciones.bat (instala "
            f"la versión compatible) o ejecuta:  python -m pip install \"{ORACLEDB_COMPATIBLE}\"  ·  Otra opción: "
            f"instalar Oracle Instant Client 19 o más reciente e indicar su carpeta en «sql_cliente_oracle».")


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
    problema = host_invalido(servidor)
    if problema:
        raise ErrorSQL(problema)
    if cliente_oracle and not _cliente_oracle_iniciado:
        try:
            oracledb.init_oracle_client(lib_dir=cliente_oracle)
            _cliente_oracle_iniciado = True
        except Exception as e:
            raise ErrorSQL(_explicar_cliente_viejo(f"No se pudo usar el Oracle Client de «{cliente_oracle}»: {e}",
                                                   cliente_oracle))
    dsn = armar_dsn(servidor, oracledb)
    try:
        conn = oracledb.connect(user=usuario, password=password, dsn=dsn, tcp_connect_timeout=10)
    except Exception as e:
        texto = _mensaje(e)
        clasico = "DPY-3015" in texto or "DPY-3010" in texto
        encontrado = _buscar_cliente_oracle() if clasico and not _cliente_oracle_iniciado else ""
        if encontrado:
            # La base pide el modo clásico: se usa el Oracle Client que ya está instalado (el de SQL Developer)
            try:
                oracledb.init_oracle_client(lib_dir=encontrado)
                _cliente_oracle_iniciado = True
                conn = oracledb.connect(user=usuario, password=password, dsn=dsn)
            except Exception as e2:
                raise ErrorSQL(_explicar_cliente_viejo(f"{_mensaje(e2)}  (con el Oracle Client de {encontrado})",
                                                       encontrado))
        else:
            if clasico:
                texto += ("  →  Esta base pide el modo clásico de Oracle y no se encontró un Oracle Client: indica "
                          "su carpeta en la configuración («sql_cliente_oracle»).")
            elif "DPY-4000" in texto or "DPY-4026" in texto or "DPY-4027" in texto:
                texto += "  →  Revisa Host, Puerto y Servicio (como en SQL Developer: Básico, Nombre del servicio)."
            elif "getaddrinfo" in texto or "11001" in texto or "DPY-6005" in texto:
                texto += (f"  →  No se encontró o no responde el servidor «{separar_servidor(servidor)[0]}»: revisa el "
                          "Nombre del host y que estés en la red de FONACOT (VPN).")
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


def consultar_por_padres(conn, tabla: str, creditos) -> list:
    """[(fecha_alta, cuenta, no_renovacion, operacion)] de las renovaciones cuyo crédito padre (OPERACION) es alguno de esos."""
    tabla = tabla_valida(tabla)
    registros = []
    cur = conn.cursor()
    try:
        for grupo in _lotes(sorted(set(c for c in creditos if c))):
            binds = {f"p{i}": v for i, v in enumerate(grupo)}
            cur.execute(f"SELECT FECHA_ALTA, CUENTA, NO_RENOVACION, OPERACION FROM {tabla} "
                        f"WHERE OPERACION IN ({', '.join(':' + k for k in binds)})", binds)
            registros.extend(cur.fetchall())
    finally:
        cur.close()
    return [(_fecha(f), _num(c), _num(n), _num(o)) for f, c, n, o in registros]


COLUMNAS_TZ_LOCK = ("TZ_LOCK", "TZLOCK")


def consultar_tz_lock(conn, tabla: str, clientes) -> dict:
    """
    {(cuenta, no_renovacion): TZ_LOCK} de las renovaciones de esos clientes. TZ_LOCK distinto de 0 = la renovación ya
    está bloqueada en la base. Se prueba TZ_LOCK y TZLOCK (si la columna no existe se usa el siguiente nombre).
    """
    tabla = tabla_valida(tabla)
    clientes = sorted(set(c for c in clientes if c))
    ultimo = None
    for columna in COLUMNAS_TZ_LOCK:
        cur = conn.cursor()
        try:
            res = {}
            for grupo in _lotes(clientes):
                binds = {f"c{i}": v for i, v in enumerate(grupo)}
                cur.execute(f"SELECT CUENTA, NO_RENOVACION, {columna} FROM {tabla} "
                            f"WHERE CUENTA IN ({', '.join(':' + k for k in binds)})", binds)
                for cuenta, renov, lock in cur.fetchall():
                    res[(_num(cuenta), _num(renov))] = _entero(lock)
            return res
        except Exception as e:
            ultimo = e
            if "00904" not in str(e):  # ORA-00904: la columna no existe con ese nombre
                raise
        finally:
            cur.close()
    raise ultimo


def _esquema(tabla: str) -> str:
    """'TOPAZ.CR_RENOVACIONES' -> 'TOPAZ.' (las demás tablas están en el mismo esquema)."""
    return tabla.split(".")[0] + "." if "." in tabla else ""


def consultar_detalle(conn, tabla: str, clientes) -> dict:
    """
    La segunda consulta (para las renovaciones de días anteriores), la misma lógica que la de SQL Developer:
    {(cuenta, no_renovacion): {'sucursal', 'plazo_renovacion' (SALDOS.PLAZO_ORIGINAL de la renovación),
    'plazo_disposicion' (CR_DISPOSICIONES.PLAZOMESES), 'cedula' (está en CR_CEDULAS_DETALLE)}}.
    """
    tabla = tabla_valida(tabla)
    e = _esquema(tabla)
    detalle = {}
    cur = conn.cursor()
    try:
        for grupo in _lotes(sorted(set(c for c in clientes if c))):
            binds = {f"c{i}": v for i, v in enumerate(grupo)}
            cur.execute(
                f"SELECT c.CUENTA, c.NO_RENOVACION, MAX(c.SUCURSAL_EJERCE), "
                f"(SELECT MAX(s.PLAZO_ORIGINAL) FROM {e}SALDOS s WHERE s.CUENTA = c.CUENTA AND s.OPERACION = c.NO_RENOVACION), "
                f"MAX(b.PLAZOMESES), MAX(CASE WHEN d.OPERACION IS NOT NULL THEN 1 ELSE 0 END) "
                f"FROM {e}SALDOS a "
                f"INNER JOIN {e}CR_DISPOSICIONES b ON a.CUENTA = b.IDTRABAJADOR AND b.NROPRESTAMO = a.OPERACION "
                f"INNER JOIN {tabla} c ON a.CUENTA = c.CUENTA AND c.OPERACION = a.OPERACION "
                f"LEFT JOIN {e}CR_CEDULAS_DETALLE d ON c.CUENTA = d.CUENTA AND c.NO_RENOVACION = d.OPERACION "
                f"WHERE a.PRODUCTO > '99' AND c.CUENTA IN ({', '.join(':' + k for k in binds)}) "
                f"GROUP BY c.CUENTA, c.NO_RENOVACION", binds)
            for cuenta, renov, sucursal, plazo_r, plazo_d, cedula in cur.fetchall():
                detalle[(_num(cuenta), _num(renov))] = {
                    'sucursal': _num(sucursal), 'plazo_renovacion': _entero(plazo_r),
                    'plazo_disposicion': _entero(plazo_d), 'cedula': bool(cedula)}
    finally:
        cur.close()
    return detalle


def _entero(valor):
    try:
        return int(valor) if valor is not None else None
    except (TypeError, ValueError):
        return None


def decidir_otro_dia(credito: str, alta: date, ref: date, det: dict) -> tuple:
    """
    Renovación dada de alta antes del día del incidente: se envía solo si la generó la APP (sucursal 104) con un plazo
    menor al de la disposición; si no, «No procede, la renovación es de otro día». (estado, detalle)
    """
    cuando = f"se dio de alta el {alta:%d/%m/%Y} y el incidente llegó el {ref:%d/%m/%Y}"
    if not det:
        return SQL_OTRA_FECHA, (f"La renovación {credito} {cuando}. No se encontró en SALDOS / CR_DISPOSICIONES "
                                f"para saber si la generó la APP: revísala a mano.")
    app = det['sucursal'] == str(SUCURSAL_APP)
    pr, pd = det['plazo_renovacion'], det['plazo_disposicion']
    plazos = f"plazo de la renovación {pr if pr is not None else '?'} vs. disposición {pd if pd is not None else '?'} meses"
    if app and pr is not None and pd is not None and pr < pd:
        return SQL_APP_PLAZO, (f"De otro día pero se envía: la renovación {credito} {cuando}, la generó la APP "
                               f"(sucursal {SUCURSAL_APP}) con plazo menor ({plazos})")
    motivo = f"la generó la APP pero el plazo no es menor ({plazos})" if app else f"no es de la APP (sucursal {det['sucursal'] or '?'})"
    return SQL_NO_PROCEDE, f"{LEYENDA_NO_PROCEDE}: {credito} {cuando}; {motivo}"


# ======================= Validación (sin base de datos: se puede probar sola) =======================

def _dato(valor) -> str:
    """El número del campo (el primero: «1184282 (⚠️ En texto también: 3386153)» -> '1184282'); '' si es «SIN …»."""
    v = str(valor or '')
    m = re.search(r'\d+', v)
    return '' if v.upper().startswith('SIN ') or not m else _num(m.group(0))


def clientes_sin_renovaciones(r: dict, por_cliente: dict) -> list:
    """
    Incidente con varios números de cliente: [(cliente, texto)] de los que no cuentan con renovaciones (o no de esos
    días). El cliente de la propia fila no se repite (ya lo dice su estado).
    """
    ref = _referencia(r)
    desde, hasta = ref - timedelta(days=DIAS_ANTES), ref + timedelta(days=1)
    propio = _dato(r.get('Numero_Cliente'))
    res = []
    for cli in dict.fromkeys(_dato(c) for c in r.get('Clientes_Ticket') or []):
        if not cli or cli == propio:
            continue
        renov = por_cliente.get(cli, [])
        if not renov:
            res.append((cli, f"El cliente {cli} (también viene en el incidente) no cuenta con renovaciones"))
        elif not any(f and desde <= f <= hasta for f, _ in renov):
            ultima = max(renov, key=lambda x: x[0] or date.min)
            res.append((cli, f"El cliente {cli} (también viene en el incidente) no cuenta con renovaciones de esos días"
                             + (f" (la última: {ultima[1]} del {ultima[0]:%d/%m/%Y})" if ultima[0] else "")))
    return res


def _referencia(r: dict):
    llegada = r.get('Fecha_Llegada_DT') or r.get('Fecha_DT')
    return llegada.date() if isinstance(llegada, datetime) else (llegada or date.today())


def _alta(cliente: str, credito: str, por_cliente: dict):
    """Fecha de alta de esa renovación del cliente (la más reciente si hay varias). None si no está."""
    fechas = [f for f, n in por_cliente.get(cliente, []) if n == credito and f] if cliente and credito else []
    return max(fechas) if fechas else None


def hijas_del_padre(cliente: str, credito: str, por_padre: dict) -> list:
    """[(fecha, cuenta, no_renovacion)] de las renovaciones de ese crédito padre (del cliente, si se conoce), la más reciente primero."""
    hijas = [(f, c, n) for f, c, n in (por_padre or {}).get(credito, []) if not cliente or c == cliente]
    return sorted(hijas, key=lambda x: x[0] or date.min, reverse=True)


def _texto_padre(credito: str, hijas: list) -> str:
    f, c, n = hijas[0]
    otras = f" (también: {', '.join(h[2] for h in hijas[1:4])})" if len(hijas) > 1 else ""
    alta = f", alta {f:%d/%m/%Y}" if f else ""
    return (f"OJO: {credito} es el crédito padre (origen), no la renovación. La renovación que hay que cancelar es la "
            f"{n} del cliente {c}{alta}{otras}")


def validar_par(cliente: str, credito: str, ref: date, por_cliente: dict, por_credito: dict,
                por_padre: dict = None) -> tuple:
    """
    (estado, detalle) de un Cliente + Crédito contra lo que trajo la base. `ref`: día en que debía hacerse la
    renovación (llegada del incidente o día del ticket único). `por_cliente`: {cuenta: [(fecha, no_renov)]},
    `por_credito`: {no_renov: [(fecha, cuenta)]}, `por_padre`: {operacion: [(fecha, cuenta, no_renov)]}.
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
            alta = _alta(cliente, credito, por_cliente)
            if alta == ref:
                return SQL_OK, f"Renovación {credito} del cliente {cliente} dada de alta el mismo día ({alta:%d/%m/%Y})"
            # De otro día: validar_contra_base hace la segunda consulta (¿la generó la APP con plazo menor?)
            otra = f" Ese día tiene: {lista([(f, n) for f, n in cercanas if f == ref])}." if any(
                f == ref and n != credito for f, n in cercanas) else ""
            return SQL_OTRA_FECHA, (f"La renovación {credito} del cliente {cliente} es del "
                                    f"{alta:%d/%m/%Y}, no de cuando llegó el incidente ({ref:%d/%m/%Y}).{otra}"
                                    if alta else f"La renovación {credito} no tiene fecha de alta")
        hijas = hijas_del_padre(cliente, credito, por_padre)
        if hijas:
            return SQL_ES_PADRE, _texto_padre(credito, hijas)
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
        hijas = hijas_del_padre("", credito, por_padre)
        if hijas:
            return SQL_ES_PADRE, _texto_padre(credito, hijas)
        return SQL_NO_EXISTE, f"No hay una renovación {credito} dada de alta cerca del {ref:%d/%m/%Y}"
    return SQL_SIN_DATOS, ""


def opciones_correccion(cliente: str, credito: str, ref: date, por_cliente: dict, por_credito: dict) -> list:
    """
    Renovaciones de la base que podrían ser la buena, para «Corregir dato»: [(cliente, renovación, fecha de alta)].
    Las del cliente de esos días (las del mismo día primero) y, si falta el cliente, de quién es ese crédito.
    """
    desde, hasta = ref - timedelta(days=DIAS_ANTES), ref + timedelta(days=1)
    res = []
    if cliente:
        cercanas = [(f, n) for f, n in por_cliente.get(cliente, []) if f and desde <= f <= hasta and n != credito]
        for f, n in sorted(cercanas, key=lambda x: (x[0] != ref, -x[0].toordinal())):
            res.append((cliente, n, f))
    elif credito:
        res += [(c, credito, f) for f, c in sorted(por_credito.get(credito, []), key=lambda x: x[0] or date.min,
                                                     reverse=True) if f and desde <= f <= hasta]
    return list(dict.fromkeys(res))[:4]


def corregir_padre(r: dict, hijas: list) -> str:
    """
    El incidente trae el crédito padre: el crédito a cancelar pasa a ser su renovación (el hijo), para que el ticket único
    no salga con el padre (lo rechazarían). Si faltaba el cliente, se pone el de la renovación. Queda la leyenda de
    validar. Regresa el nuevo estado (SQL_PADRE_CORREGIDO).
    """
    fecha, cuenta, hija = hijas[0]
    fecha = _fecha(fecha)
    padre = _dato(r.get('Credito_Renovado_Hijo'))
    r['Credito_Padre_Enviado'] = padre
    r['Credito_Renovado_Hijo'] = hija
    if str(r.get('Credito_Origen_Padre') or '-').strip() in ('', '-'):
        r['Credito_Origen_Padre'] = padre
    puso_cliente = not _dato(r.get('Numero_Cliente'))
    if puso_cliente:
        r['Cliente_Enviado'] = r.get('Numero_Cliente') or ''
        r['Numero_Cliente'] = cuenta
    if not str(r.get('Numero_Cliente', '')).upper().startswith('SIN') and not str(hija).upper().startswith('SIN'):
        r['Datos_Incompletos'] = False
    otras = f" Ese padre también tiene: {', '.join(h[2] for h in hijas[1:4])}; confirma cuál es." if len(hijas) > 1 else ""
    r['SQL_Fecha_Alta'] = fecha or r.get('SQL_Fecha_Alta')
    r['SQL_Estado'] = SQL_PADRE_CORREGIDO
    r['SQL_Detalle'] = (f"VALIDAR: el incidente trae el crédito padre {padre}; se puso su renovación {hija} como crédito a "
                        f"cancelar (cliente {cuenta}" + (f", alta {fecha:%d/%m/%Y}" if fecha else "") + ")"
                        + (" y el número de cliente, que faltaba" if puso_cliente else "") + f".{otras}")
    return SQL_PADRE_CORREGIDO


def validar_contra_base(conn, tabla: str, rows: list, filas_enviadas: list = ()) -> dict:
    """
    Consulta la base y marca cada renovación con 'SQL_Estado' y 'SQL_Detalle'. Los renglones de los tickets
    únicos (`filas_enviadas`, con 'cliente', 'credito', 'fecha') se marcan con 'sql_estado' y 'sql_detalle'.
    Devuelve el resumen {'ok': True, 'conteos': {estado: n}, 'problemas': n}.
    """
    pares = [(_dato(r.get('Numero_Cliente')), _dato(r.get('Credito_Renovado_Hijo')), _referencia(r)) for r in rows]
    pares_env = [(_dato(f.get('cliente')), _dato(f.get('credito')), f.get('llegada') or f.get('fecha') or date.today())
                 for f in filas_enviadas]
    todos = pares + pares_env

    por_cliente = {}
    # Incidentes con varios números de cliente: también se consultan los demás (¿alguno no tiene renovaciones?)
    extra = [_dato(c) for r in rows for c in r.get('Clientes_Ticket') or []]
    for f, c, n in consultar_por_clientes(conn, tabla, [c for c, _, _ in todos] + extra):
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

    # El crédito que el cliente no tiene como renovación, ¿es un crédito padre? (mandaron el origen en vez de la renovación)
    por_padre, error_padre = {}, ''
    padres = sorted({cr for c, cr, _ in todos if cr and cr not in {n for _, n in por_cliente.get(c, [])}})
    if padres:
        try:
            for f, c, n, o in consultar_por_padres(conn, tabla, padres):
                por_padre.setdefault(o, []).append((f, c, n))
        except Exception as e:
            error_padre = _mensaje(e)

    # Renglón del ticket único sin su incidente: lo del día hábil anterior que llegó después del envío también es «del día»
    for i, f in enumerate(filas_enviadas, len(pares)):
        c, cr, ref = todos[i]
        alta = _alta(c, cr, por_cliente)
        if not f.get('llegada') and f.get('desde') and alta and f['desde'] <= alta <= ref:
            todos[i] = (c, cr, alta)
    resultados = [validar_par(c, cr, ref, por_cliente, por_credito, por_padre) for c, cr, ref in todos]

    # Segunda consulta: las de días anteriores (¿APP con plazo menor?) y si ya emitió (CR_CEDULAS_DETALLE)
    con_alta = [(c, cr) for c, cr, _ in todos if _alta(c, cr, por_cliente)]
    detalle, error_detalle = {}, ''
    if con_alta:
        try:
            detalle = consultar_detalle(conn, tabla, [c for c, _ in con_alta])
        except Exception as e:
            error_detalle = _mensaje(e)
    for i, ((c, cr, ref), (estado, texto)) in enumerate(zip(todos, resultados)):
        alta = _alta(c, cr, por_cliente)
        if estado == SQL_OTRA_FECHA and alta and alta < ref:
            if i < len(rows) and rows[i].get('Mesa_Credito'):
                estado, texto = SQL_MESA_CREDITO, (
                    f"De otro día pero se envía: la renovación {cr} se dio de alta el {alta:%d/%m/%Y} y el incidente "
                    f"llegó el {ref:%d/%m/%Y}; la manda la Mesa de Crédito (sus renovaciones de otros días sí se "
                    f"cancelan si no les tocó la emisión ni el cierre de mes)")
            elif error_detalle:
                texto += f" No se pudo hacer la segunda consulta (APP / plazo): {error_detalle}"
            else:
                pista = re.search(r' Ese día tiene: .*$', texto)  # «¿no será la de ese día?»
                estado, texto = decidir_otro_dia(cr, alta, ref, detalle.get((c, cr)))
                if estado != SQL_APP_PLAZO and pista:
                    texto += "." + pista.group(0).rstrip(".") + "."
        resultados[i] = (estado, texto, alta, (detalle.get((c, cr)) or {}).get('cedula', False))

    # TZ_LOCK de las renovaciones que existen (distinto de 0 = ya bloqueada): sirve para saber si un incidente cerrado
    # sin enviar se cerró por error
    tz_lock, error_tz_lock = {}, ''
    if con_alta:
        try:
            tz_lock = consultar_tz_lock(conn, tabla, [c for c, _ in con_alta])
        except Exception as e:
            error_tz_lock = _mensaje(e)

    conteos = {}
    for r, (c, cr, ref), (estado, texto, alta, cedula) in zip(rows, pares, resultados):
        r['SQL_Estado'], r['SQL_Detalle'], r['SQL_Fecha_Alta'] = estado, texto, alta
        r['SQL_Opciones'] = [] if estado in SQL_SE_ENVIA else opciones_correccion(c, cr, ref, por_cliente, por_credito)
        r['SQL_Tz_Lock'] = tz_lock.get((c, cr)) if alta else None
        r['SQL_Clientes_Sin_Renov'] = [t for _, t in clientes_sin_renovaciones(r, por_cliente)]
        r['SQL_Cedula'] = "Esta renovación ha emitido, favor de validar (está en CR_CEDULAS_DETALLE)" if cedula else ""
        hijas = hijas_del_padre(_dato(r.get('Numero_Cliente')), _dato(r.get('Credito_Renovado_Hijo')), por_padre) \
            if estado == SQL_ES_PADRE else []
        r['SQL_Renovacion_Hija'] = hijas[0][2] if hijas else ''
        if hijas:
            estado = corregir_padre(r, hijas)
        conteos[estado] = conteos.get(estado, 0) + 1
    for f, (estado, texto, alta, cedula) in zip(filas_enviadas, resultados[len(rows):]):
        f['sql_estado'], f['sql_detalle'] = estado, texto
    return {'ok': True, 'conteos': conteos, 'problemas': sum(conteos.get(e, 0) for e in SQL_PROBLEMAS),
            'enviadas_con_problema': sum(1 for f in filas_enviadas if f.get('sql_estado') in SQL_PROBLEMAS),
            'error_detalle': error_detalle, 'error_padre': error_padre, 'error_tz_lock': error_tz_lock,
            'clientes_sin_renov': sum(1 for r in rows if r.get('SQL_Clientes_Sin_Renov')),
            'padres': conteos.get(SQL_ES_PADRE, 0) + conteos.get(SQL_PADRE_CORREGIDO, 0),
            'cedulas': sum(1 for *_, ced in resultados[:len(rows)] if ced)}


def texto_corto(estado: str) -> str:
    return {SQL_OK: "✔ Existe", SQL_OTRA_FECHA: "⚠ Otra fecha", SQL_NO_EXISTE: "✖ No existe",
            SQL_SIN_RENOVACIONES: "✖ Sin renovaciones", SQL_SUGERENCIA: "💡 Sugerencia", SQL_SIN_DATOS: "—",
            SQL_APP_PLAZO: "✔ APP, plazo menor", SQL_MESA_CREDITO: "✔ Mesa de Crédito",
            SQL_NO_PROCEDE: "✖ No procede (otro día)",
            SQL_ES_PADRE: "⚠ Es el padre", SQL_PADRE_CORREGIDO: "⚠ Era el padre"}.get(estado, "")


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
