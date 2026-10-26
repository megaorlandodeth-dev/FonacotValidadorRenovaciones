#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
v2.7: correos sin etiquetas (un trabajador por bloque), espacios duros y Ñ compuesta, números que se confundían
entre cliente y crédito, host de la base guardado con el puerto y reintento del calendario de cédulas.
"""

import os
import sys
import tempfile
import unittest
from datetime import datetime
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import calendario_cedulas as cal  # noqa: E402
from core_renovaciones import RenovacionTicketParser as P, expand_abbreviated_labels  # noqa: E402
from sql_renovaciones import armar_dsn, host_invalido, separar_servidor, servidor_de  # noqa: E402

DOS_TRABAJADORES = (
    "FAVOR DE APOYAR CON LA CANCELACION DE LAS SIGUIENTES RENOVACIONES, YA QUE SU CAPACIDAD DE CREDITO ESTA "
    "TOPADA SEGÚN NUESTRO SISTEMA DE CREDITO\n"
    "PEREZ DIAZ CARLOS ERNESTO\n3326312\nC/33363\n"
    "JUAREZ RUIZ JOSE ANGEL\n1280433\n33143"
)

NOMBRE_Y_CREDITOS_ABAJO = (
    "Buen día,\nPara que el trabajador no se vea afectado en el crédito.\n"
    "LUIS RENE MENDEZ ZUÑIGA 6650919\nCREDITO RENOVADO\n. 53290 Y 61780\n"
    "SUC. LAZARO CARDENAS\nANALISTA:    SUSAN MICHELLE LOPEZ MEJIA\nGAFETE:      7093"
)


def pares(texto, resumen="CANCELACION DE RENOVACION"):
    return [(r['cliente'], r['credito_renovado']) for r in P.extract_disaggregated_renewals(resumen, texto, [])]


class TestCorreosSinEtiquetas(unittest.TestCase):
    def test_un_trabajador_por_bloque(self):
        self.assertEqual(pares(DOS_TRABAJADORES), [("3326312", "33363"), ("1280433", "33143")])

    def test_con_saltos_de_windows_y_espacios_duros(self):
        texto = DOS_TRABAJADORES.replace("\n", "\r\n").replace(" ", "\xa0")
        self.assertEqual(pares(texto), [("3326312", "33363"), ("1280433", "33143")])

    def test_nombre_con_cliente_y_creditos_abajo(self):
        esperado = [("6650919", "53290"), ("6650919", "61780")]
        self.assertEqual(pares(NOMBRE_Y_CREDITOS_ABAJO), esperado)
        self.assertEqual(pares(NOMBRE_Y_CREDITOS_ABAJO.replace("ZUÑIGA 6650919", "ZUÑIGA\xa06650919")), esperado)
        self.assertEqual(pares(NOMBRE_Y_CREDITOS_ABAJO.replace("Ñ", "Ñ")), esperado)
        # En negritas: el número queda en el renglón de abajo
        self.assertEqual(pares(NOMBRE_Y_CREDITOS_ABAJO.replace("ZUÑIGA 6650919", "ZUÑIGA\n6650919")), esperado)

    def test_diagonal_despues_de_la_abreviatura(self):
        self.assertEqual(expand_abbreviated_labels("C/33363"), "Credito: 33363")

    def test_un_nombre_suelto_no_arma_bloques(self):
        texto = "Cliente: 5864479\nCredito: 56572 y 56574\nSaludos\nJUAN PEREZ LOPEZ"
        self.assertEqual(pares(texto), [("5864479", "56572"), ("5864479", "56574")])


class TestNumerosConfundidos(unittest.TestCase):
    def test_credito_no_es_cliente(self):
        self.assertEqual(pares("Cancelar renovación del trabajador con número de crédito 61780"), [("SIN CLIENTE", "61780")])
        self.assertEqual(pares("CANCELAR RENOVACION DEL TRABAJADOR CON NUMERO DE CREDITO 61780"),
                         [("SIN CLIENTE", "61780")])

    def test_varios_clientes_en_el_orden_del_correo(self):
        texto = "Cliente: 1111111 Credito: 22222\nCliente: 3333333 Credito renovado: 44444"
        self.assertEqual(pares(texto), [("1111111", "22222"), ("3333333", "44444")])

    def test_cliente_con_etiqueta_y_otro_junto_al_nombre(self):
        texto = "Cliente 3326312 credito 33363\nJUAREZ RUIZ JOSE ANGEL 1280433 credito 33143"
        self.assertEqual(pares(texto), [("3326312", "33363"), ("1280433", "33143")])

    def test_ticket_y_gafete_del_asunto(self):
        filas = P.parse_and_disaggregate_email("El Ticket I-362000 ha sido asignado agente: 7353", "",
                                               "Favor de cancelar la renovación", datetime(2026, 10, 26, 9),
                                               "I-362000", "7353", "Sammy")
        self.assertEqual([(r['Numero_Cliente'], r['Credito_Renovado_Hijo']) for r in filas],
                         [("SIN CLIENTE", "SIN CRÉDITO")])


class TestHostDeLaBase(unittest.TestCase):
    def test_puerto_guardado_como_host(self):
        self.assertEqual(separar_servidor("1521/fonatop"), ("", "1521", "fonatop"))
        self.assertIn("Nombre del host", host_invalido(servidor_de("1521", "1521", "fonatop") or "1521"))
        self.assertEqual(host_invalido("scanfonacot.fonacot.gob.mx:1521/fonatop"), "")

    def test_servidor_completo_pegado_en_el_host(self):
        self.assertEqual(servidor_de("scanfonacot.fonacot.gob.mx:1521/fonatop", "", ""),
                         "scanfonacot.fonacot.gob.mx:1521/fonatop")
        self.assertEqual(armar_dsn(servidor_de("host.fonacot.gob.mx", "1521", "fona-top")),
                         "host.fonacot.gob.mx:1521/fona-top")


class TestCalendario(unittest.TestCase):
    def test_sin_internet_se_reintenta(self):
        carpeta = tempfile.mkdtemp()
        archivo = os.path.join(carpeta, "calendario_cedulas.json")
        with mock.patch.object(cal, "_archivo", return_value=archivo), \
                mock.patch("urllib.request.urlopen", side_effect=OSError("sin red")):
            self.assertFalse(cal.actualizar_calendario(forzar=True))
        self.assertFalse(os.path.exists(archivo))  # no se anota como actualizado: se vuelve a intentar


if __name__ == "__main__":
    unittest.main()
