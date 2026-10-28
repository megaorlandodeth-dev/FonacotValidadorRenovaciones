#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
v3.3: etiquetas del crédito con una aclaración entre paréntesis («NUMERO DE CREDITO (REN) : 57815», I-362949).
"""

import os
import sys
import unittest
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core_renovaciones import RenovacionTicketParser as Parser  # noqa: E402


def leer(texto):
    rows = Parser.parse_and_disaggregate_email("CANCELAR RENOVACION", "", texto, datetime(2026, 10, 28, 10, 3),
                                               "I-362949", "2609", "N")
    return [(r['Numero_Cliente'], r['Credito_Renovado_Hijo'], r['Datos_Incompletos']) for r in rows]


class TestEtiquetaConParentesis(unittest.TestCase):
    def test_numero_de_credito_ren(self):
        texto = ("Buen día,\nSolicito su apoyo para cancelar la siguiente renovación:\nCLIENTE: 7438176\n"
                 "NOMBRE: JONATHAN ORLANDO GOMEZ NUÑEZ\nNUMERO DE CREDITO (REN) : 57815\nIMSS: 43089066336\n"
                 "MOTIVO: EL SISTEMA LIMITA DESCUENTO, NO PERMITE TRAMITAR NUEVO CREDITO\nGRACIAS, SALUDOS.")
        self.assertEqual(leer(texto), [("7438176", "57815", False)])

    def test_otras_aclaraciones(self):
        for etiqueta in ("NUMERO DE CREDITO (RENOV.):", "No. de credito (hijo)", "CREDITO (REN):",
                         "RENOVACION (CREDITO NUEVO):", "CREDITO (RENOVACION)"):
            with self.subTest(etiqueta=etiqueta):
                self.assertEqual(leer(f"CLIENTE: 7438176\n{etiqueta} 57815\ncancelar renovacion"),
                                 [("7438176", "57815", False)])

    def test_varios_creditos(self):
        self.assertEqual(leer("CLIENTE: 7438176\nCREDITOS (REN): 57815 Y 57816\ncancelar renovacion"),
                         [("7438176", "57815", False), ("7438176", "57816", False)])

    def test_numero_entre_parentesis_no_se_brinca(self):
        # El paréntesis solo se salta si trae letras: «crédito (57815)» sigue leyéndose
        self.assertEqual(leer("CLIENTE: 7438176\nCREDITO (57815)\ncancelar renovacion")[0][1], "57815")


class TestConexionSinCryptography(unittest.TestCase):
    """3.3.1: el .exe de la 3.3 no traía cryptography (DPY-3016): se usa el Oracle Client instalado."""

    def test_dpy3016_usa_el_oracle_client(self):
        import types
        from unittest import mock
        import sql_renovaciones as sql
        llamadas = []

        def connect(**kw):
            llamadas.append("connect")
            if len(llamadas) == 1:
                raise Exception("DPY-3016: python-oracledb thin mode cannot be used because the cryptography "
                                "package cannot be imported")
            return types.SimpleNamespace()

        falso = types.SimpleNamespace(connect=connect, makedsn=lambda h, p, service_name: f"{h}:{p}/{service_name}",
                                      init_oracle_client=lambda lib_dir: llamadas.append(f"cliente {lib_dir}"),
                                      __version__="3.4.2")
        with mock.patch.dict(sys.modules, {"oracledb": falso}), \
                mock.patch.object(sql, "_clientes_oracle", return_value=[r"C:\oracle\bin"]), \
                mock.patch.object(sql, "_cliente_oracle_iniciado", False):
            sql.conectar("scanfonacot.fonacot.gob.mx:1521/fonatop", "usuario", "clave")
        self.assertEqual(llamadas, ["connect", r"cliente C:\oracle\bin", "connect"])

    def test_autoprueba_detecta_que_falta(self):
        import types
        from unittest import mock
        import autoprueba

        def oracledb_que_responde(error):
            def connect(**kw):
                raise Exception(error)
            return types.SimpleNamespace(connect=connect, __version__="3.4.2")

        sin_cifrado = oracledb_que_responde("DPY-3016: python-oracledb thin mode cannot be used because the "
                                            "cryptography package cannot be imported")
        completo = oracledb_que_responde("DPY-6005: cannot connect to database. [WinError 10061] refused")
        with mock.patch.dict(sys.modules, {"oracledb": sin_cifrado}):
            nivel, detalle = autoprueba._oracle_thin()
        self.assertEqual(nivel, autoprueba.FALLA)
        self.assertIn("DPY-3016", detalle)
        with mock.patch.dict(sys.modules, {"oracledb": completo}):
            self.assertEqual(autoprueba._oracle_thin()[0], autoprueba.OK)


def _oracledb_falso(llamadas, error_thin, clientes_malos=()):
    """oracledb de mentira: el modo thin falla con error_thin; init_oracle_client falla en las carpetas malas."""
    import types

    def connect(**kw):
        llamadas.append("connect")
        if len(llamadas) == 1:
            raise Exception(error_thin)
        return types.SimpleNamespace()

    def init_oracle_client(lib_dir):
        llamadas.append(f"cliente {lib_dir}")
        if lib_dir in clientes_malos:
            raise Exception("DPI-1047: Cannot locate a 64-bit Oracle Client library: is not a valid Win32 application")

    return types.SimpleNamespace(connect=connect, init_oracle_client=init_oracle_client, __version__="3.4.2",
                                 makedsn=lambda h, p, service_name: f"{h}:{p}/{service_name}")


def _oci_dll(carpeta, maquina=0x8664):
    """Un oci.dll de mentira con la cabecera PE de 64 bits (0x8664) o de 32 (0x14C)."""
    import struct
    os.makedirs(carpeta, exist_ok=True)
    datos = bytearray(512)
    datos[:2] = b"MZ"
    struct.pack_into("<I", datos, 0x3C, 0x80)
    datos[0x80:0x84] = b"PE\0\0"
    struct.pack_into("<H", datos, 0x84, maquina)
    with open(os.path.join(carpeta, "oci.dll"), "wb") as f:
        f.write(bytes(datos))


class TestModoClasico(unittest.TestCase):
    """3.3.2: la base de FONACOT pide el modo clásico (DPY-3015) y en el equipo no se encontraba el Oracle Client."""

    def conectar(self, error_thin, clientes, clientes_malos=()):
        from unittest import mock
        import sql_renovaciones as sql
        llamadas = []
        with mock.patch.dict(sys.modules, {"oracledb": _oracledb_falso(llamadas, error_thin, clientes_malos)}), \
                mock.patch.object(sql, "_clientes_oracle", return_value=list(clientes)), \
                mock.patch.object(sql, "_cliente_oracle_iniciado", False):
            try:
                sql.conectar("scanfonacot.fonacot.gob.mx:1521/fonatop", "FNT_USUARIO", "clave")
                error = ""
            except sql.ErrorSQL as e:
                error = str(e)
        return llamadas, error

    def test_verificador_10g_usa_el_cliente(self):
        llamadas, error = self.conectar("DPY-3015: password verifier type 0x939 is not supported by python-oracledb "
                                        "in thin mode", [r"C:\app\oracle_cliente"])
        self.assertEqual((llamadas, error), (["connect", r"cliente C:\app\oracle_cliente", "connect"], ""))

    def test_si_un_cliente_no_sirve_prueba_el_siguiente(self):
        llamadas, error = self.conectar("DPY-3010: connections to this database server version are not supported",
                                        [r"C:\cliente32", r"D:\app\client\maria\product\12.2.0\client_1\bin"],
                                        clientes_malos=(r"C:\cliente32",))
        self.assertEqual(error, "")
        self.assertEqual(llamadas[-2:], [r"cliente D:\app\client\maria\product\12.2.0\client_1\bin", "connect"])

    def test_cifrado_de_red_y_conexion_cerrada_tambien(self):
        for error_thin in ("DPY-3001: Native Network Encryption and Data Integrity is only supported in thick mode",
                           "DPY-4011: the database or network closed the connection"):
            with self.subTest(error_thin=error_thin[:8]):
                llamadas, error = self.conectar(error_thin, [r"C:\oc"])
                self.assertEqual((llamadas, error), (["connect", r"cliente C:\oc", "connect"], ""))

    def test_sin_cliente_lo_explica(self):
        llamadas, error = self.conectar("DPY-3015: password verifier type 0x939 is not supported", [])
        self.assertIn("modo clásico", error)
        self.assertIn("Oracle Instant Client 19", error)

    def test_contrasena_incorrecta_no_cambia_de_modo(self):
        llamadas, error = self.conectar("ORA-01017: invalid username/password; logon denied", [r"C:\oc"])
        self.assertEqual(llamadas, ["connect"])
        self.assertIn("Usuario o contraseña incorrectos", error)

    def test_servidor_que_no_responde(self):
        llamadas, error = self.conectar("DPY-6005: cannot connect to database. timed out", [r"C:\oc"])
        self.assertEqual(llamadas, ["connect"])
        self.assertIn("scanfonacot.fonacot.gob.mx", error)
        self.assertIn("VPN", error)


class TestDireccionModoClasico(unittest.TestCase):
    """En modo clásico va la dirección completa: el sqlnet.ora de un Oracle Client instalado puede no admitir la corta."""

    def test_dsn_clasico(self):
        import sql_renovaciones as sql
        self.assertEqual(sql.dsn_clasico("scanfonacot.fonacot.gob.mx:1521/fonatop"),
                         "(DESCRIPTION=(TRANSPORT_CONNECT_TIMEOUT=10)(ADDRESS=(PROTOCOL=TCP)"
                         "(HOST=scanfonacot.fonacot.gob.mx)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=fonatop)))")
        for igual in ("(DESCRIPTION=(ADDRESS=(HOST=x)(PORT=1)))", "FONATOP", "host:1521"):
            self.assertEqual(sql.dsn_clasico(igual), igual)

    def test_conectar_usa_la_direccion_completa_en_modo_clasico(self):
        import types
        from unittest import mock
        import sql_renovaciones as sql
        dsns = []

        def connect(**kw):
            dsns.append(kw["dsn"])
            if len(dsns) == 1:
                raise Exception("DPY-3015: password verifier type 0x939 is not supported")
            return types.SimpleNamespace()

        falso = types.SimpleNamespace(connect=connect, init_oracle_client=lambda lib_dir: None, __version__="3.4.2",
                                      makedsn=lambda h, p, service_name: f"{h}:{p}/{service_name}")
        with mock.patch.dict(sys.modules, {"oracledb": falso}), \
                mock.patch.object(sql, "_clientes_oracle", return_value=[r"C:\oc"]), \
                mock.patch.object(sql, "_cliente_oracle_iniciado", False):
            sql.conectar("scanfonacot.fonacot.gob.mx:1521/fonatop", "FNT_USUARIO", "clave")
            sql.conectar("scanfonacot.fonacot.gob.mx:1521/fonatop", "FNT_USUARIO", "clave")
        self.assertEqual(dsns[0], "scanfonacot.fonacot.gob.mx:1521/fonatop")
        self.assertTrue(all(d.startswith("(DESCRIPTION=") for d in dsns[1:]), dsns)
        self.assertEqual(len(dsns), 3)


class TestBuscarClienteOracle(unittest.TestCase):
    """El Oracle Client de SQL Developer puede estar en D: (perfiles en D:\\Users) y solo sirve el de 64 bits."""

    def buscar(self, raiz):
        from unittest import mock
        import core_renovaciones
        import sql_renovaciones as sql
        app = os.path.join(raiz, "ValidadorRenovaciones")
        os.makedirs(app, exist_ok=True)
        with mock.patch.object(sql, "_unidades_fijas", return_value=[raiz]), \
                mock.patch.object(sql, "_homes_del_registro", return_value=[]), \
                mock.patch.object(core_renovaciones, "get_app_dir", return_value=app), \
                mock.patch.dict(os.environ, {"PATH": "", "ORACLE_HOME": ""}):
            return [os.path.relpath(c, raiz) for c in sql._clientes_oracle()]

    def test_cliente_en_otro_disco(self):
        import tempfile
        with tempfile.TemporaryDirectory() as raiz:
            _oci_dll(os.path.join(raiz, "app", "client", "maria.chavez", "product", "12.2.0", "client_1", "bin"))
            _oci_dll(os.path.join(raiz, "instantclient_12_2"), maquina=0x14C)  # de 32 bits: no sirve
            self.assertEqual(self.buscar(raiz),
                             [os.path.join("app", "client", "maria.chavez", "product", "12.2.0", "client_1", "bin")])

    def test_primero_el_que_trae_la_app(self):
        import tempfile
        with tempfile.TemporaryDirectory() as raiz:
            _oci_dll(os.path.join(raiz, "app", "client", "x", "product", "12.2.0", "client_1", "bin"))
            _oci_dll(os.path.join(raiz, "ValidadorRenovaciones", "oracle_cliente", "instantclient_19_32"))
            self.assertEqual(self.buscar(raiz)[0],
                             os.path.join("ValidadorRenovaciones", "oracle_cliente", "instantclient_19_32"))


class TestAutopruebaServidor(unittest.TestCase):
    def test_puerto_abierto_y_cerrado(self):
        import socket
        import autoprueba
        servidor = socket.socket()
        servidor.bind(("127.0.0.1", 0))
        servidor.listen(1)
        puerto = servidor.getsockname()[1]
        try:
            nivel, detalle = autoprueba._servidor({"sql_host": "127.0.0.1", "sql_puerto": str(puerto),
                                                   "sql_servicio": "fonatop"})
            self.assertEqual(nivel, autoprueba.OK, detalle)
        finally:
            servidor.close()
        self.assertEqual(autoprueba._servidor({"sql_host": "127.0.0.1", "sql_puerto": str(puerto),
                                               "sql_servicio": "fonatop"})[0], autoprueba.AVISO)
        self.assertEqual(autoprueba._servidor({})[0], autoprueba.AVISO)


if __name__ == "__main__":
    unittest.main()
