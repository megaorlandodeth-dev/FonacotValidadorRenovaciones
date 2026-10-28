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


if __name__ == "__main__":
    unittest.main()
