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
                mock.patch.object(sql, "_buscar_cliente_oracle", return_value=r"C:\oracle\bin"), \
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


if __name__ == "__main__":
    unittest.main()
