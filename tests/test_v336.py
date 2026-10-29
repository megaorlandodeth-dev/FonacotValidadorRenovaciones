#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
v3.3.6: en el ticket único (copiar filas, copiar lote y el correo) la FECHA es la de alta de la renovación (FECHA_ALTA
de la primera consulta a la base), no la de llegada del incidente ni la de hoy.
"""

import os
import sys
import unittest
from datetime import date, datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from lote_utils import build_ticket_unico_html, build_ticket_unico_tsv  # noqa: E402


def fila(**extra):
    r = {'Numero_Ticket': 'I-363065', 'Fecha_Llegada_DT': datetime(2026, 10, 28, 15, 0),
         'Fecha_DT': datetime(2026, 10, 28, 15, 0), 'Fecha': '28/10/2026 15:00', 'Numero_Cliente': '6626931',
         'Credito_Renovado_Hijo': '67200'}
    r.update(extra)
    return r


class TestFechaDelTicketUnico(unittest.TestCase):
    def test_va_la_fecha_de_alta_de_la_base(self):
        rows = [fila(SQL_Fecha_Alta=date(2026, 10, 23))]
        self.assertEqual(build_ticket_unico_tsv(rows).splitlines()[1], "I-363065\t23/10/2026\t6626931\t67200")
        html = build_ticket_unico_html(rows, intro="Hola {total}")
        self.assertIn(">23/10/2026<", html)
        self.assertNotIn("28/10/2026", html)

    def test_sin_base_la_del_correo_y_si_no_la_de_llegada(self):
        self.assertIn("\t22/10/2026\t", build_ticket_unico_tsv([fila(Fecha_Alta_Correo=date(2026, 10, 22))]))
        self.assertIn("\t28/10/2026\t", build_ticket_unico_tsv([fila()]))

    def test_la_base_manda_sobre_el_correo(self):
        r = fila(SQL_Fecha_Alta=date(2026, 10, 23), Fecha_Alta_Correo=date(2026, 10, 22))
        self.assertIn("\t23/10/2026\t", build_ticket_unico_tsv([r]))


if __name__ == '__main__':
    unittest.main()
