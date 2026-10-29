#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
v3.3.7: «crédito renovado 7933850 cliente GUERRERO… CREDITO REN 69988» (I-363102): el número va antes de la palabra
«cliente»; salían dos filas SIN CLIENTE (7933850 y 69988). Solo cambia cuando no hay cliente y hay varios créditos.
"""

import os
import sys
import unittest
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core_renovaciones import RenovacionTicketParser as Parser  # noqa: E402


def leer(desc, asunto="Cancelación de renovación"):
    rows = Parser.parse_and_disaggregate_email(asunto, "", desc, datetime(2026, 10, 29, 9, 0), "I-363102", "7261",
                                               "Moises")
    return [(r['Numero_Cliente'], r['Credito_Renovado_Hijo'], r['Credito_Origen_Padre']) for r in rows]


class TestClienteAntesDeLaPalabra(unittest.TestCase):
    def test_i363102(self):
        self.assertEqual(leer("Buenos días favor cancelar crédito renovado\n7933850 cliente\nGUERRERO LOPEZ JAIME ANDREY\n"
                              "CREDITO REN 69988\nQUEDO A SUS ORDENES"),
                         [('7933850', '69988', '-')])

    def test_los_formatos_de_siempre_no_cambian(self):
        self.assertEqual(leer("Cancelar renovación 69988 cliente 7933850"), [('7933850', '69988', '-')])
        self.assertEqual(leer("Favor de cancelar la renovacion 69988 del cliente 7933850"), [('7933850', '69988', '-')])
        self.assertEqual(leer("cancelar renovaciones 69988 y 70001 cliente 7933850"),
                         [('7933850', '69988', '-'), ('7933850', '70001', '-')])

    def test_un_credito_de_5_digitos_antes_de_cliente_sigue_siendo_credito(self):
        self.assertEqual(leer("Cancelar renovacion 69988 y 70001 del cliente GUERRERO"),
                         [('SIN CLIENTE', '69988', '-'), ('SIN CLIENTE', '70001', '-')])


if __name__ == '__main__':
    unittest.main()
