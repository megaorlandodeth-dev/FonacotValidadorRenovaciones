#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Correo con una captura de pantalla adjunta → OCR de Windows → renovación completa.
Se salta si el equipo no tiene Pillow o el OCR de Windows (no es parte de la lógica, es del entorno).
"""

import os
import shutil
import sys
import tempfile
import unittest
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from outlook_renovaciones_engine import FonacotRenovacionesEngine  # noqa: E402

try:
    from PIL import Image, ImageDraw, ImageFont
except ImportError:  # pragma: no cover
    Image = None


def captura(ruta):
    """Pantalla parecida a «Consulta de Crédito»: etiquetas blancas sobre gris y valores en cajas blancas."""
    im = Image.new("RGB", (900, 260), (236, 236, 236))
    d = ImageDraw.Draw(im)
    f = ImageFont.truetype(r"C:\Windows\Fonts\arial.ttf", 15)
    fb = ImageFont.truetype(r"C:\Windows\Fonts\arialbd.ttf", 15)
    d.text((10, 10), "250013, Consulta de Crédito, Transacción 3 - Consulta Saldos", font=f, fill="black")
    for y, etiqueta, valor in ((50, "No. Cliente", "4363419"), (100, "No. Credito", "17626"), (140, "Producto", "301")):
        d.rectangle((10, y, 160, y + 24), fill=(110, 110, 110))
        d.text((18, y + 4), etiqueta, font=fb, fill="white")
        d.rectangle((170, y, 310, y + 24), fill="white", outline=(150, 150, 150))
        d.text((300 - d.textlength(valor, font=f), y + 4), valor, font=f, fill="black")
    d.text((330, 54), "GALVAN REYNA IVAN ALEJANDRO", font=f, fill="black")
    im.save(ruta)


class Adjunto:
    def __init__(self, ruta):
        self.ruta, self.FileName, self.Size = ruta, "image001.png", os.path.getsize(ruta)

    def SaveAsFile(self, destino):
        shutil.copy(self.ruta, destino)


class Adjuntos:
    def __init__(self, lista):
        self.lista, self.Count = lista, len(lista)

    def Item(self, i):
        return self.lista[i - 1]


class Correo:
    def __init__(self, adjuntos):
        self.Attachments, self.HTMLBody = Adjuntos(adjuntos), "<html><body><img src='cid:image001.png'></body></html>"


@unittest.skipUnless(sys.platform == "win32" and Image is not None and os.path.exists(r"C:\Windows\Fonts\arial.ttf"),
                     "requiere Windows con Pillow")
class TestCapturaDeVerdad(unittest.TestCase):
    def test_incidente_que_solo_trae_la_imagen(self):
        tmp = tempfile.mkdtemp()
        try:
            ruta = os.path.join(tmp, "consulta.png")
            captura(ruta)
            motor = FonacotRenovacionesEngine.__new__(FonacotRenovacionesEngine)
            motor.ocr_info = {'completadas': 0, 'sin_leer': [], 'error': ''}
            fila = {'Row_ID': 'I-362557_x_1', 'Numero_Ticket': 'I-362557', 'Numero_Cliente': "SIN CLIENTE",
                    'Credito_Renovado_Hijo': "SIN CRÉDITO", 'Datos_Incompletos': True, 'Fecha_DT': datetime(2026, 10, 23, 16, 1),
                    'Indice_En_Ticket': 1, 'Total_En_Ticket': 1, 'Es_Multiple': False,
                    'Alerta_Renovacion': "🚨 FALTA DATO", 'Nivel_Alerta': "incompleto"}
            filas = motor._completar_con_imagenes(Correo([Adjunto(ruta)]), [fila], 'I-362557')
            if motor.ocr_info['error']:
                self.skipTest(f"OCR de Windows no disponible: {motor.ocr_info['error']}")
            self.assertEqual((filas[0]['Numero_Cliente'], filas[0]['Credito_Renovado_Hijo']), ("4363419", "17626"))
            self.assertFalse(filas[0]['Datos_Incompletos'])
            self.assertEqual(motor.ocr_info['completadas'], 1)
        finally:
            shutil.rmtree(tmp, ignore_errors=True)


if __name__ == '__main__':
    unittest.main()
