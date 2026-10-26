#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
v2.9: crédito padre en vez de la renovación (base de datos), días de corte con varios tickets únicos (cuenta el
último; lo que llega después ya no se cancela) y el lote copiado con el mismo formato del ticket único.
"""

import os
import sys
import unittest
from datetime import date, datetime, time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import calendario_cedulas as cal  # noqa: E402
import sql_renovaciones as sql  # noqa: E402
from core_renovaciones import clasificar_lotes, cortes_reales, LOTE_HOY, LOTE_ENVIADA  # noqa: E402
from lote_utils import build_ticket_unico_texto, build_ticket_unico_html  # noqa: E402
from test_sql_y_ocr import Conexion, fila  # noqa: E402

CIERRE = date(2026, 10, 30)  # se usa como día de corte en las pruebas (viernes)


def concentrado(fecha, recibido, filas):
    tabla = "".join(f"<tr><td>{t}</td><td>{c}</td><td>{cr}</td></tr>" for t, c, cr in filas)
    return {'asunto': f"Cancelación de renovaciones – ticket único {fecha:%d-%m-%Y}", 'fecha_asunto': fecha,
            'recibido': recibido, 'remitente': 'Javier', 'carpeta': 'Bandeja',
            'html': f"<table><tr><td>INCIDENTE</td><td>CLIENTE</td><td>CREDITO</td></tr>{tabla}</table>",
            'tickets': [t for t, _, _ in filas], 'numeros': [n for _, c, cr in filas for n in (c, cr)]}


def renovacion(ticket, cli, cred, llegada):
    r = fila(ticket, cli, cred, llegada)
    r.update(Cerrado_iTop=False, Cierre_iTop=None, Gafete="2609", Nombre_Asignado="Nancy")
    return r


class TestCreditoPadre(unittest.TestCase):
    def test_mandan_el_padre_con_su_cliente(self):
        r = fila("I-1", "1993891", "51937")
        resumen = sql.validar_contra_base(Conexion(), "TOPAZ.CR_RENOVACIONES", [r])
        self.assertEqual(r['SQL_Estado'], sql.SQL_ES_PADRE)
        self.assertEqual(r['SQL_Renovacion_Hija'], "52011")
        self.assertIn("51937 es el crédito padre", r['SQL_Detalle'])
        self.assertIn("la 52011 del cliente 1993891", r['SQL_Detalle'])
        self.assertIn(sql.SQL_ES_PADRE, sql.SQL_PROBLEMAS)
        self.assertEqual(resumen['padres'], 1)

    def test_padre_sin_cliente(self):
        r = fila("I-2", "SIN CLIENTE", "51937")
        sql.validar_contra_base(Conexion(), "TOPAZ.CR_RENOVACIONES", [r])
        self.assertEqual(r['SQL_Estado'], sql.SQL_ES_PADRE)
        self.assertEqual(r['SQL_Renovacion_Hija'], "52011")

    def test_consulta_con_variables_y_solo_si_hace_falta(self):
        conn = Conexion()
        sql.validar_contra_base(conn, "TOPAZ.CR_RENOVACIONES", [fila("I-3", "1993891", "51937")])
        texto, binds = next(c for c in conn.consultas if "WHERE OPERACION IN" in c[0])
        self.assertTrue(texto.startswith("SELECT FECHA_ALTA, CUENTA, NO_RENOVACION, OPERACION FROM TOPAZ.CR_RENOVACIONES"))
        self.assertNotIn("51937", texto)
        self.assertEqual(binds, {'p0': '51937'})
        # Si el crédito sí es una renovación del cliente no se busca como padre
        conn = Conexion()
        sql.validar_contra_base(conn, "TOPAZ.CR_RENOVACIONES", [fila("I-4", "1184282", "10430")])
        self.assertFalse(any("WHERE OPERACION IN" in c[0] for c in conn.consultas))


class TestDiaDeCorte(unittest.TestCase):
    def test_cuenta_el_ultimo_ticket_unico_del_dia_de_corte(self):
        concs = [concentrado(CIERRE, datetime(2026, 10, 30, 17, 50), []),
                 concentrado(CIERRE, datetime(2026, 10, 30, 19, 10), [])]
        self.assertEqual(cortes_reales(concs)[CIERRE], datetime(2026, 10, 30, 17, 50))
        self.assertEqual(cortes_reales(concs, {CIERRE: "el cierre"})[CIERRE], datetime(2026, 10, 30, 19, 10))

    def test_hoy_es_corte_lo_que_llega_despues_va_en_otro_hoy(self):
        r = renovacion("I-10", "1184282", "10430", datetime(2026, 10, 30, 18, 30))
        info = clasificar_lotes([r], CIERRE, CIERRE, time(17, 50),
                                [concentrado(CIERRE, datetime(2026, 10, 30, 17, 50), [])],
                                ahora=datetime(2026, 10, 30, 19, 0), revision=False, dias_de_corte={CIERRE: "el cierre de mes"})
        self.assertEqual(r['Lote_Estado'], LOTE_HOY)
        self.assertEqual(r['Dia_Lote'], CIERRE)
        self.assertFalse(r['Falta_Hoy'])
        self.assertIn("otro ticket único", r['Lote_Etiqueta'])
        self.assertEqual(info['corte_de_hoy'], "el cierre de mes")
        self.assertEqual(info['envios_hoy'], [datetime(2026, 10, 30, 17, 50)])

    def test_dia_normal_lo_que_llega_despues_va_al_siguiente(self):
        r = renovacion("I-11", "1184282", "10430", datetime(2026, 10, 29, 18, 30))
        clasificar_lotes([r], date(2026, 10, 29), date(2026, 10, 29), time(17, 50),
                         [concentrado(date(2026, 10, 29), datetime(2026, 10, 29, 17, 50), [])],
                         ahora=datetime(2026, 10, 29, 19, 0), revision=False, dias_de_corte={CIERRE: "el cierre"})
        self.assertNotEqual(r['Lote_Estado'], LOTE_HOY)
        self.assertEqual(r['Dia_Lote'], CIERRE)

    def test_dia_de_corte_pasado_lo_del_segundo_ticket_unico_va_a_tiempo(self):
        r = renovacion("I-12", "1184282", "10430", datetime(2026, 10, 30, 18, 30))
        concs = [concentrado(CIERRE, datetime(2026, 10, 30, 17, 50), [("I-9", "1000000", "10000")]),
                 concentrado(CIERRE, datetime(2026, 10, 30, 19, 10), [("I-12", "1184282", "10430")])]
        clasificar_lotes([r], date(2026, 10, 26), date(2026, 11, 2), time(17, 50), concs,
                         ahora=datetime(2026, 11, 2, 10, 0), dias_de_corte={CIERRE: "el cierre"})
        self.assertEqual(r['Lote_Estado'], LOTE_ENVIADA)
        self.assertEqual(r['Dias_Retraso'], 0)
        self.assertEqual(r['Dia_Lote'], CIERRE)


class TestYaSeEmitio(unittest.TestCase):
    def fila_lote(self, llegada, **extra):
        r = {'SQL_Fecha_Alta': date(2026, 10, 30), 'Fecha_Llegada_DT': llegada}
        r.update(extra)
        return r

    def test_salio_en_un_ticket_unico_del_dia_del_cierre(self):
        r = self.fila_lote(datetime(2026, 10, 30, 18, 30), Enviado_Fecha=date(2026, 10, 30), Dia_Lote=date(2026, 10, 30))
        cal.marcar_bloqueos([r], date(2026, 11, 2))
        self.assertEqual(r['Bloqueo_Emision'], '')

    def test_llego_despues_del_ultimo_ticket_unico_del_cierre(self):
        r = self.fila_lote(datetime(2026, 10, 30, 19, 30), Dia_Lote=date(2026, 11, 2), Lote_Estado="REZAGADA")
        cal.marcar_bloqueos([r], date(2026, 11, 2))
        self.assertIn("sale en el ticket único del 02/11/2026, después del cierre de mes (31/10/2026)",
                      r['Bloqueo_Emision'])
        self.assertIn("ya se emitió", r['Bloqueo_Emision'])

    def test_va_en_el_lote_de_hoy_que_es_el_dia_del_cierre(self):
        r = self.fila_lote(datetime(2026, 10, 30, 19, 30), Dia_Lote=date(2026, 10, 30), Lote_Estado="HOY")
        cal.marcar_bloqueos([r], date(2026, 10, 30))
        self.assertEqual(r['Bloqueo_Emision'], '')

    def test_cortes_entre(self):
        cortes = cal.cortes_entre(date(2026, 10, 20), date(2026, 11, 30))
        self.assertEqual(sorted(cortes), [date(2026, 10, 25), date(2026, 10, 31), date(2026, 11, 24), date(2026, 11, 30)])
        self.assertIn("emisión de noviembre", cortes[date(2026, 10, 25)])


class TestCopiarComoTicketUnico(unittest.TestCase):
    ROWS = [{'Numero_Ticket': 'I-1', 'Fecha_Llegada_DT': datetime(2026, 10, 26, 9, 0), 'Numero_Cliente': '1184282',
             'Credito_Renovado_Hijo': '10430'}]

    def test_texto_con_saludo_y_tabla(self):
        texto = build_ticket_unico_texto(self.ROWS, "Hola,\n\nSe envía “{total}” solicitudes")
        self.assertTrue(texto.startswith("Hola,\n\nSe envía “1” solicitudes\n\nINCIDENTE\tFECHA\tCLIENTE\tCREDITO (HIJO)"))
        self.assertTrue(texto.endswith("I-1\t26/10/2026\t1184282\t10430"))

    def test_html_igual_que_el_correo(self):
        html = build_ticket_unico_html(self.ROWS, intro="Hola,\n\nSe envía “{total}” solicitudes")
        self.assertIn("Se envía “1” solicitudes", html)
        self.assertIn("CREDITO (HIJO)", html)


if __name__ == "__main__":
    unittest.main()
