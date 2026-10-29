#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
v3.3.5: «cancelar el crédito 67200 y regresar al 67774» (I-363065) es reversar la renovación aunque no diga
«renovación»: se cancela 67200 y 67774 es el crédito origen. Un «cancelar el crédito» sin nada de renovación sigue fuera.
"""

import os
import sys
import unittest
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core_renovaciones import RenovacionTicketParser as Parser  # noqa: E402


def leer(desc, asunto="cancelación"):
    rows = Parser.parse_and_disaggregate_email(asunto, "", desc, datetime(2026, 10, 28, 15, 0), "I-363065", "7261",
                                               "Moises")
    return [(r['Numero_Cliente'], r['Credito_Renovado_Hijo'], r['Credito_Origen_Padre']) for r in rows]


class TestRegresarAlCreditoAnterior(unittest.TestCase):
    def test_i363065(self):
        self.assertEqual(leer("Buena tarde me podrán apoyar en cancelar el crédito 67200 y regresar al 67774 del cliente "
                              "6626931 ya que por tope de buro no le arroja capacidad"),
                         [('6626931', '67200', '67774')])

    def test_otras_formas(self):
        for desc in ("Favor de cancelar el crédito 67200 del cliente 6626931 y reactivar el crédito 67774",
                     "Cliente 6626931: cancelar el crédito 67200 y volver al crédito 67774",
                     "Cancelar crédito 67200 y regresar a su crédito anterior 67774, cliente 6626931"):
            self.assertEqual(leer(desc), [('6626931', '67200', '67774')], desc)

    def test_tope_de_buro_es_motivo_de_capacidad(self):
        self.assertEqual(leer("Cancelar la renovación 67200 del cliente 6626931, tope de buro"),
                         [('6626931', '67200', '-')])

    def test_cancelar_credito_sin_renovacion_sigue_fuera(self):
        for desc in ("Cancelar el crédito 67200 cliente 6626931 y regresar el dinero al trabajador",
                     "Cancelación del crédito 67200 cliente 6626931 por error de captura"):
            self.assertEqual(leer(desc), [], desc)


if __name__ == '__main__':
    unittest.main()
