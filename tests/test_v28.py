#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
v2.8: crédito antes del cliente con la etiqueta partida en dos renglones, «clienta», revisión de lunes a sábado,
recordatorio del calendario de emisión, Oracle Client 12.2 y de dónde salen las imágenes de un incidente.
"""

import os
import sys
import tempfile
import unittest
from datetime import date
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import ocr_imagenes  # noqa: E402
from calendario_cedulas import cortes_alrededor  # noqa: E402
from core_renovaciones import RenovacionTicketParser as P, resumen_por_dia  # noqa: E402
from outlook_renovaciones_engine import FonacotRenovacionesEngine  # noqa: E402
from sql_renovaciones import _explicar_cliente_viejo  # noqa: E402

CREDITO_ANTES = (
    "Buenas tardes.\nAtentamente solicito de su apoyo para cancelar la renovación del crédito\nnumero: 51937,\n"
    "otorgado a la trabajadora\n.1993891. Edith Sanchez Ramirez.\nLo anterior debido a que el\n"
    "CREDITO SE REALIZO DESDE LA APLICACIÓN Y LA TRABAJADORA TRAMITO A 6 MESES"
)


def pares(texto):
    return [(r['cliente'], r['credito_renovado'])
            for r in P.extract_disaggregated_renewals('CANCELACION DE RENOVACION', texto, [])]


class TestCreditoAntesDelCliente(unittest.TestCase):
    def test_etiqueta_en_un_renglon_y_numero_en_el_otro(self):
        self.assertEqual(pares(CREDITO_ANTES), [("1993891", "51937")])

    def test_clienta(self):
        self.assertEqual(pares("Se solicita cancelar el crédito 51937 de la clienta 1993891"), [("1993891", "51937")])

    def test_sin_etiqueta_numero_el_renglon_de_abajo_no_es_credito(self):
        texto = "SEGÚN NUESTRO SISTEMA DE CREDITO\nPEREZ DIAZ CARLOS ERNESTO\n3326312\nC/33363"
        self.assertEqual(pares(texto), [("3326312", "33363")])


class TestSemanaConSabado(unittest.TestCase):
    def test_el_sabado_sale_en_por_dia(self):
        dias = resumen_por_dia([], [], date(2026, 10, 19), date(2026, 10, 24))
        self.assertEqual([d['dia'].weekday() for d in dias], [0, 1, 2, 3, 4, 5])


class TestRecordatorioEmision(unittest.TestCase):
    def test_proximos_cortes(self):
        ultimo, proximos = cortes_alrededor(date(2026, 10, 26))
        self.assertEqual(ultimo[0], date(2026, 10, 25))
        self.assertEqual([f for f, _ in proximos], [date(2026, 10, 31), date(2026, 11, 24)])
        self.assertIn("cierre de mes", proximos[0][1])

    def test_el_dia_del_corte_todavia_cuenta(self):
        _, proximos = cortes_alrededor(date(2026, 10, 31))
        self.assertEqual(proximos[0][0], date(2026, 10, 31))


class TestOracleClienteViejo(unittest.TestCase):
    def test_explica_que_hacer(self):
        texto = _explicar_cliente_viejo("DPI-1050: Oracle Client library is at version 12.2.0 but version 19.1.0 "
                                        "or higher is needed", r"C:\app\client\x\product\12.2.0\client_1\bin")
        self.assertIn('oracledb>=3.4,<4', texto)
        self.assertIn("Ejecutar_Validador_Renovaciones.bat", texto)

    def test_otros_errores_no_cambian(self):
        self.assertEqual(_explicar_cliente_viejo("ORA-01017: invalid username/password", ""),
                         "ORA-01017: invalid username/password")


class Adjuntos:
    def __init__(self, lista):
        self.lista, self.Count = lista, len(lista)

    def Item(self, i):
        return self.lista[i - 1]


class Correo:
    def __init__(self, html="", adjuntos=(), session=None):
        self.HTMLBody, self.Attachments, self.Session = html, Adjuntos(list(adjuntos)), session


class CorreoAdjunto:
    """Correo de la sucursal reenviado como adjunto (.msg)."""
    Type, FileName, Size = 5, "Solicitud.msg", 50000

    def SaveAsFile(self, destino):
        with open(destino, "wb") as f:
            f.write(b"msg")


class AdjuntoImagen:
    Type, FileName = 1, "captura.png"

    def __init__(self):
        self.Size = 20000

    def SaveAsFile(self, destino):
        with open(destino, "wb") as f:
            f.write(b"\x89PNG" + b"0" * 20000)


class Respuesta:
    def __init__(self, tipo, datos):
        self.headers, self._datos = mock.Mock(), datos
        self.headers.get_content_type.return_value = tipo

    def read(self, _n=None):
        return self._datos

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False


class TestImagenesDelIncidente(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()

    def test_imagen_dentro_de_un_correo_reenviado(self):
        interno = Correo(adjuntos=[AdjuntoImagen()])
        interno.Close = mock.Mock()
        sesion = mock.Mock()
        sesion.OpenSharedItem.return_value = interno
        rutas = ocr_imagenes.guardar_imagenes_del_correo(Correo(adjuntos=[CorreoAdjunto()], session=sesion), self.tmp)
        self.assertEqual(len(rutas), 1)
        self.assertTrue(rutas[0].endswith("captura.png"))
        interno.Close.assert_called_once_with(1)

    def test_imagen_ligada_a_itop(self):
        html = ('<img src="https://itop.fonacot.gob.mx/pages/ajax.document.php?operation=download_inlineimage'
                '&amp;id=5&amp;s=abc"><img src="https://tracker.example.com/pixel.png">')
        with mock.patch("urllib.request.urlopen", return_value=Respuesta("image/png", b"0" * 20000)) as abrir:
            rutas = ocr_imagenes.guardar_imagenes_del_correo(Correo(html), self.tmp)
        self.assertEqual(len(rutas), 1)
        self.assertEqual(abrir.call_count, 1)  # nunca se baja nada de servidores de fuera
        self.assertIn("&id=5&s=abc", abrir.call_args[0][0].full_url)

    def test_itop_pide_sesion(self):
        notas = []
        html = '<img src="https://itop.fonacot.gob.mx/pages/ajax.document.php?id=5">'
        with mock.patch("urllib.request.urlopen", return_value=Respuesta("text/html", b"<html>login</html>")):
            self.assertEqual(ocr_imagenes.guardar_imagenes_del_correo(Correo(html), self.tmp, notas), [])
        self.assertIn("iniciar sesión", notas[0])

    def _fila(self):
        return {'Datos_Incompletos': True, 'Numero_Cliente': 'SIN CLIENTE', 'Credito_Renovado_Hijo': 'SIN CRÉDITO'}

    def test_busca_en_los_demas_correos_del_incidente(self):
        motor = FonacotRenovacionesEngine.__new__(FonacotRenovacionesEngine)
        motor.ocr_info = {'completadas': 0, 'sin_leer': [], 'error': ''}
        sin_imagen, con_imagen = Correo(), Correo()
        motor._correos_por_ticket = {'I-362700': [sin_imagen, con_imagen]}
        vistos = []

        def leer(correo, notas=None, ref=None):
            vistos.append(correo)
            return ([{'cliente': '4937968', 'credito': '85277', 'imagen': 'x.png', 'texto': '', 'dudosa': False}], 1) \
                if correo is con_imagen else ([], 0)
        with mock.patch("outlook_renovaciones_engine.leer_correo", side_effect=leer), \
                mock.patch("outlook_renovaciones_engine.completar_con_imagenes",
                           side_effect=lambda rows, lect: [dict(r, Datos_Incompletos=False, Dato_Imagen='ok') for r in rows]):
            filas = motor._completar_con_imagenes(sin_imagen, [self._fila()], 'I-362700')
        self.assertEqual(vistos, [sin_imagen, con_imagen])
        self.assertFalse(filas[0]['Datos_Incompletos'])

    def test_dice_por_que_no_se_leyo(self):
        motor = FonacotRenovacionesEngine.__new__(FonacotRenovacionesEngine)
        motor.ocr_info = {'completadas': 0, 'sin_leer': [], 'error': ''}
        motor._correos_por_ticket = {}
        with mock.patch("outlook_renovaciones_engine.leer_correo", return_value=([], 0)):
            filas = motor._completar_con_imagenes(Correo(), [self._fila()], 'I-362701')
        self.assertIn("No se encontró ninguna imagen", filas[0]['Imagen_Info'])


if __name__ == "__main__":
    unittest.main()
