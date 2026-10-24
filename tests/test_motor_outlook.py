#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Búsqueda de los correos «ticket único» en un buzón simulado (sin Outlook)."""

import os
import sys
import unittest
from datetime import datetime, date

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from outlook_renovaciones_engine import FonacotRenovacionesEngine  # noqa: E402


def tabla(*filas):
    renglones = "".join(f"<tr><td>{t}</td><td>{f}</td><td>{c}</td><td>{r}</td></tr>" for t, f, c, r in filas)
    return ("<html><body><table border='1'><tr><td>INCIDENTE</td><td>FECHA</td><td>CLIENTE</td>"
            f"<td>CREDITO (HIJO)</td></tr>{renglones}</table></body></html>")


class Adjuntos:
    Count = 0

    def __iter__(self):
        return iter(())


class Correo:
    Class = 43

    def __init__(self, asunto, recibido, html="", remitente="Javier Servin"):
        self.Subject = asunto
        self.ReceivedTime = recibido
        self.HTMLBody = html
        self.Body = ""
        self.SenderName = remitente
        self.Attachments = Adjuntos()


class Elementos(list):
    @property
    def Count(self):
        return len(self)

    def Sort(self, _campo, descendente=False):
        super().sort(key=lambda m: m.ReceivedTime, reverse=descendente)


class Carpeta:
    def __init__(self, correos):
        self.Items = Elementos(correos)
        self.EntryID = str(id(self))


def motor_con(correos):
    motor = FonacotRenovacionesEngine.__new__(FonacotRenovacionesEngine)
    carpeta = Carpeta(correos)
    motor.connect_outlook = lambda: True
    motor._get_all_root_folders = lambda: [("Buzón", carpeta)]
    motor._collect_mail_folders_recursive = lambda root, **k: [("Buzón/Bandeja", root)]
    return motor


class TestBuscarTicketsUnicos(unittest.TestCase):
    def test_encuentra_todos_los_de_la_semana(self):
        # Regresión v2.4: después del primer ticket único con tabla se dejaban de reconocer los demás
        correos = [
            Correo(f"Cancelación de renovaciones – ticket único {d:02d}-09-2026", datetime(2026, 9, d, 17, 45),
                   tabla((f"I-36000{d}", f"{d}/09/2026", "4363419", f"1762{d}")))
            for d in (28, 29, 30)
        ] + [
            Correo("Cancelación de renovaciones – ticket único 22-10-2026", datetime(2026, 10, 22, 17, 59),
                   tabla(("I-362409", "22/10/2026", "7891353", "4127"))),
            Correo("Junta del viernes", datetime(2026, 10, 1, 9, 0)),
        ]
        concs = motor_con(correos).find_concentrado_emails(datetime(2026, 9, 28), datetime(2026, 10, 23, 23, 59))
        self.assertEqual([c['recibido'].date() for c in concs],
                         [date(2026, 9, 28), date(2026, 9, 29), date(2026, 9, 30), date(2026, 10, 22)])
        self.assertEqual(concs[0]['filas'][0]['cliente'], "4363419")
        self.assertEqual(concs[0]['filas'][0]['credito'], "176228")

    def test_misma_copia_en_dos_carpetas_cuenta_una_vez(self):
        html = tabla(("I-362409", "22/10/2026", "7891353", "4127"))
        correos = [Correo("Cancelación de renovaciones – ticket único 22-10-2026", datetime(2026, 10, 22, 17, 59), html),
                   Correo("Cancelación de renovaciones – ticket único 22-10-2026", datetime(2026, 10, 22, 18, 0), html)]
        concs = motor_con(correos).find_concentrado_emails(datetime(2026, 10, 22), datetime(2026, 10, 23))
        self.assertEqual(len(concs), 1)


if __name__ == '__main__':
    unittest.main()
