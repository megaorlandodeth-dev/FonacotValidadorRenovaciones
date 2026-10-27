#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
v3.2: lectura de correos que antes se perdían o salían mal (auditoría, puntos 22 a 26): tablas en texto, listas con
viñetas, «e», «/» o en columna, «Crédito a cancelar», el crédito padre, números de la plantilla, folios y firmas,
números con separadores y la ruta HTML (la que corre con Outlook).
"""

import os
import sys
import unittest
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core_renovaciones import RenovacionTicketParser as Parser  # noqa: E402

A, B = ('4919283', '71470'), ('5864479', '56572')


def par(*creds, cli='4919283'):
    return [(cli, c) for c in creds]


def filas(texto, resumen="CANCELACION DE RENOVACION", html=""):
    rows = Parser.parse_and_disaggregate_email(resumen, html, texto, datetime(2026, 10, 27, 9), "I-1", "2609", "N")
    return [(r['Numero_Cliente'], r['Credito_Renovado_Hijo']) for r in rows]


def html(descripcion):
    return ("<html><body><table><tr><td>Incidente #</td><td>I-1</td></tr>"
            "<tr><td>Resumen</td><td>CANCELACION DE RENOVACION</td></tr>"
            f"<tr><td>Descripción</td><td>{descripcion}</td></tr></table></body></html>")


TABLA = ("<table><tr><td>Cliente</td><td>Crédito</td></tr><tr><td>4919283</td><td>71470</td></tr>"
         "<tr><td>5864479</td><td>56572</td></tr></table>")


class TestTextoPlano(unittest.TestCase):
    CASOS = {
        # Tablas en texto: todas las filas
        'CLIENTE CREDITO\n4919283 71470\n5864479 56572': [A, B],
        'Cliente\tCrédito\n4919283\t71470\n5864479\t56572': [A, B],
        'Cliente\nCrédito\n4919283\n71470\n5864479\n56572': [A, B],
        'CLIENTE CREDITO\n4919283 71470\n5864479 56572\n5864479 56573': [A, B, ('5864479', '56573')],
        # Listas
        'Cliente: 4919283\nCréditos:\n- 71470\n- 71471': par('71470', '71471'),
        'Cliente 4919283 cancelar renovaciones\n- 71470\n- 71471': par('71470', '71471'),
        'Cliente 4919283 cancelar créditos 71470 e 81471': par('71470', '81471'),
        'Cliente 4919283 créditos 71470/71471': par('71470', '71471'),
        'Cliente 4919283 cancelar las renovaciones:\n71470\n71471\n71472': par('71470', '71471', '71472'),
        'Cliente 4919283\nCréditos:\n• 71470\n• 71471': par('71470', '71471'),
        'Cliente 4919283\nRenovaciones a cancelar:\n1. 71470\n2. 71471': par('71470', '71471'),
        'Cliente 4919283 créditos 45375,45377': par('45375', '45377'),
        'Cliente 4919283 crédito 71470 del 27/10/2026': [A],
        # El crédito escrito antes de su cliente
        'Crédito 71470 del cliente 4919283\nCrédito 56572 del cliente 5864479\nCrédito 56573 del cliente 5864479':
            [A, B, ('5864479', '56573')],
        # «a cancelar»
        'Cliente 4919283\nCrédito a cancelar: 71470': [A],
        'Cliente: 4919283 Renovación a cancelar 71470': [A],
        'Cliente 4919283 crédito renovado a cancelar: 71470': [A],
        # El crédito padre no se manda
        'Cliente 4919283 el credito 65000 se renovó al 71470 cancelar': [A],
        'Cliente 4919283 el crédito 65000 se renovó con el crédito 71470': [A],
        'Cliente 4919283 cancelar renovación 65000 -> 71470': [A],
        'Cliente 4919283 cancelar la renovación 71470 del crédito 65000': [A],
        'Cliente 4919283 -> 71470': [A],
        # Números que antes se ignoraban por coincidir con la plantilla (si vienen junto a «crédito», sí son créditos)
        'Cliente 4919283 crédito 2026': par('2026'),
        'Cliente 4919283 crédito 1003': par('1003'),
        'Cliente 4919283 créditos 452 y 681': par('452', '681'),
        'Cliente 4919283 crédito 85000': par('85000'),
        # Números que no son ni cliente ni crédito
        'Cliente 4919283 crédito 71470\nSaludos\nNancy López\nAnalista de Crédito\n2609': [A],
        'Cliente 4919283 crédito 71470 porque bajó su capacidad de crédito\n15000': [A],
        'Cliente 4919283 crédito 71470 folio de solicitud de credito:\n123456': [A],
        'Cliente 4919283 crédito 71470\nDepto. de créditos\n5512-3456': [A],
        'Cliente 4919283 crédito 71470 relacionado con R-123456': [A],
        'Cliente 4919283 crédito 71470 tr. 202610': [A],
        'Cliente 4919283 crédito 71470 cta 12345678': [A],
        'Cliente 4919283 crédito 71470\nAtte.\nMaria Fernanda Ruiz 123456': [A],
        'Cliente 4919283 crédito 71470\nCentro de trabajo Grupo Bimbo 1234567': [A],
        # Números con separadores
        'Cliente: 4 919 283 crédito 71470': [A],
        'Cliente: 4919-283 crédito 71470': [A],
        'Cliente: 4,919,283 crédito 71470': [A],
        'Cliente: 4919283 crédito 71.470': [A],
    }

    def test_casos(self):
        for texto, esperado in self.CASOS.items():
            with self.subTest(texto=texto):
                self.assertEqual(filas(texto), esperado)

    def test_gafete_en_el_asunto_no_es_credito(self):
        self.assertEqual(filas("", "CANCELACION DE RENOVACION GAF 2609 4919283 71470"), [A])
        self.assertEqual(filas("Cancelar renovación", "CANCELACION DE RENOVACION GAF 2609 4919283 71470"), [A])


class TestHTML(unittest.TestCase):
    def test_tabla_dentro_de_pre(self):
        # lxml saca la tabla del <pre>: se lee toda la celda de la Descripción
        self.assertEqual(filas("", html=html('<pre>cancelar renovación' + TABLA + '</pre>')), [A, B])

    def test_tabla_dentro_de_div(self):
        self.assertEqual(filas("", html=html('<div class="HTML"><p>cancelar renovación</p>' + TABLA + '</div>')), [A, B])

    def test_etiquetas_en_linea_no_parten_el_renglon(self):
        for envoltura in ('<pre>{}</pre>', '<div><p>{}</p></div>'):
            desc = envoltura.format('cancelar la renovación del cliente <span>4919283</span>, crédito '
                                    '<span>71470</span> y <span>71471</span>')
            self.assertEqual(filas("", html=html(desc)), par('71470', '71471'))
        self.assertEqual(filas("", html=html('<div><p>Cancelar renovación cliente <b>4919283</b> crédito <b>71470</b>'
                                             '</p></div>')), [A])

    def test_respaldo_en_el_texto_plano(self):
        self.assertEqual(filas('Cancelar renovación Cliente 4919283 crédito 71470',
                               html=html('<div>Cancelar renovación</div>')), [A])

    def test_firma_en_html(self):
        self.assertEqual(filas("", html=html('<pre>Cliente 4919283 crédito 71470\nAnalista de <span>Crédito</span>'
                                             '<br><span>2609</span></pre>')), [A])

    def test_texto_html(self):
        from core_renovaciones import make_soup
        soup = make_soup('<div><p>uno <b>dos</b>\n tres</p><p>cuatro<br>cinco</p><table><tr><td>a</td><td>b</td></tr>'
                         '</table></div>')
        self.assertEqual(Parser._texto_html(soup.div), "uno dos tres\ncuatro\ncinco\na\tb")


if __name__ == "__main__":
    unittest.main()
