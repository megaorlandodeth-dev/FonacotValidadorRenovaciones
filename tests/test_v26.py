#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
v2.6: cliente escrito junto al nombre, ticket único con otro formato (el de la guardia del sábado), encabezados del
concentrado y calendario de emisión de cédulas.
"""

import os
import sys
import unittest
from datetime import date, datetime, time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))  # test_sql_y_ocr (base simulada)

import calendario_cedulas as cal  # noqa: E402
from core_renovaciones import (  # noqa: E402
    RenovacionTicketParser as P,
    clasificar_lotes,
    extract_concentrado_rows,
    extract_concentrado_tokens,
    parse_fecha_asunto,
    LOTE_ENVIADA,
    LOTE_HOY,
)

CORREO_SIN_ETIQUETA = """Buen día,
Por medio del presente, solicito de favor sea cancelada la renovación que se realizó al siguiente trabajador
Para que el trabajador no se vea afectado en el crédito.
LUIS RENE MENDEZ ZUÑIGA 6650919
CREDITO RENOVADO
. 53290 Y 61780
SUC. LAZARO CARDENAS
ANALISTA:   SUSAN MICHELLE LOPEZ MEJIA"""


def renovaciones(ticket, desc, llegada):
    rows = P.parse_and_disaggregate_email("CANCELACION DE RENOVACION", "", desc, llegada, ticket, "7799",
                                          "Citlali Rodriguez Rosas")
    for r in rows:
        r['Fecha_Llegada_DT'] = llegada
        r['Cerrado_iTop'], r['Cierre_iTop'] = False, None
    return rows


def concentrado(asunto, recibido, html, remitente="Citlali Rodriguez Rosas"):
    from core_renovaciones import make_soup
    texto = make_soup(html).get_text(" ", strip=True)
    tickets, numeros = extract_concentrado_tokens(texto)
    filas = extract_concentrado_rows(html)
    tickets.update(f['ticket'].split('-', 1)[-1] for f in filas if f['ticket'])
    return {'asunto': asunto, 'fecha_asunto': parse_fecha_asunto(asunto, recibido.year), 'recibido': recibido,
            'remitente': remitente, 'carpeta': 'Bandeja de entrada', 'tickets': tickets, 'numeros': numeros,
            'filas': filas, 'texto': texto}


class TestClienteJuntoAlNombre(unittest.TestCase):
    def test_correo_de_la_sucursal_lazaro_cardenas(self):
        rows = renovaciones("I-362639", CORREO_SIN_ETIQUETA, datetime(2026, 10, 26, 9, 0))
        self.assertEqual([(r['Numero_Cliente'], r['Credito_Renovado_Hijo']) for r in rows],
                         [("6650919", "53290"), ("6650919", "61780")])
        self.assertFalse(any(r['Datos_Incompletos'] for r in rows))

    def test_no_confunde_fechas_folios_ni_telefonos(self):
        texto = ("Fecha 26102026 Juan Perez\nFolio 7654321 Juan Perez\nTel 5512345678 Ana Lopez\n"
                 "CREDITO RENOVADO 1234567")
        self.assertEqual(P._clientes_junto_al_nombre(texto), [])
        self.assertEqual(P._clientes_junto_al_nombre("6650919 Luis Rene Mendez"), ["6650919"])


class TestTicketUnicoDeOtroCompanero(unittest.TestCase):
    """El sábado hace guardia Citlali y ella manda el ticket único; no siempre con la misma tabla."""
    AHORA = datetime(2026, 10, 26, 11, 47)

    def lote(self, html):
        rows = (renovaciones("I-362605", "Cliente 6572804 creditos 35237 y 35238", datetime(2026, 10, 24, 9, 30))
                + renovaciones("I-362595", "Cliente 607804 credito 24621", datetime(2026, 10, 24, 10, 0))
                + renovaciones("I-362485", "Cliente 7982848 credito 27179", datetime(2026, 10, 23, 18, 30))
                + renovaciones("I-362700", "Cliente 1184282 credito 10430", datetime(2026, 10, 26, 9, 0)))
        conc = concentrado("Cancelación de renovaciones – ticket único 24-10-2026", datetime(2026, 10, 24, 15, 0), html)
        info = clasificar_lotes(rows, date(2026, 10, 26), date(2026, 10, 26), time(17, 50), [conc],
                                ahora=self.AHORA, revision=False)
        return rows, info

    def test_tabla_y_lista_en_el_mismo_correo(self):
        # Un renglón en tabla y el resto escritos abajo: antes salían «NO SE ENVIÓ … NO este Cliente+Crédito»
        html = """<html><body><p>Hola, se envían “4” solicitudes</p>
        <table><tr><td>INCIDENTE</td><td>FECHA</td><td>CLIENTE</td><td>CREDITO</td></tr>
        <tr><td>I-362605</td><td>24/10/2026</td><td>6572804</td><td>35237</td></tr></table>
        <p>I-362605 6572804 35238</p><p>I-362595 607804 24621</p><p>I-362485 7982848 27179</p></body></html>"""
        rows, info = self.lote(html)
        estados = {(r['Numero_Ticket'], r['Credito_Renovado_Hijo']): r['Lote_Estado'] for r in rows}
        self.assertEqual(estados[("I-362605", "35237")], LOTE_ENVIADA)
        self.assertEqual(estados[("I-362605", "35238")], LOTE_ENVIADA)
        self.assertEqual(estados[("I-362595", "24621")], LOTE_ENVIADA)
        self.assertEqual(estados[("I-362485", "27179")], LOTE_ENVIADA)
        self.assertEqual(estados[("I-362700", "10430")], LOTE_HOY)
        self.assertEqual(info['no_enviadas'], 0)

    def test_varios_creditos_en_una_celda_y_columnas_de_origen(self):
        html = """<table><tr><th>No. Incidente</th><th>Nombre del cliente</th><th>No. Cliente</th>
        <th>Crédito origen</th><th>Crédito renovado</th></tr>
        <tr><td>I-362605</td><td>JUAN PEREZ</td><td>6572804</td><td>11111</td><td>35237 y 35238</td></tr>
        <tr><td>I-362595</td><td>ANA LOPEZ</td><td>0607804</td><td>22222</td><td>24621</td></tr>
        <tr><td>I-362485</td><td>LUIS DIAZ</td><td>7982848</td><td>33333</td><td>27179</td></tr></table>"""
        filas = extract_concentrado_rows(html)
        self.assertEqual([(f['cliente'], f['credito']) for f in filas],
                         [("6572804", "35237"), ("6572804", "35238"), ("607804", "24621"), ("7982848", "27179")])
        rows, info = self.lote(html)
        self.assertEqual(info['no_enviadas'], 0)

    def test_si_el_ticket_viene_en_la_tabla_con_otro_credito_sigue_siendo_error(self):
        html = """<table><tr><td>INCIDENTE</td><td>CLIENTE</td><td>CREDITO</td></tr>
        <tr><td>I-362595</td><td>607804</td><td>24620</td></tr></table><p>Saludos 24621</p>"""
        rows, info = self.lote(html)
        r = next(r for r in rows if r['Numero_Ticket'] == "I-362595")
        self.assertTrue(r['No_Enviada'])
        self.assertEqual(r['Error_Envio'], "salió con Crédito 24620 en lugar de 24621")


class TestCalendarioDeEmision(unittest.TestCase):
    def test_emision_de_noviembre(self):
        # La de noviembre se generó el domingo 25/10/2026: lo dado de alta antes y pedido después ya no se cancela
        self.assertEqual(cal.corte_que_bloquea(date(2026, 10, 23), date(2026, 10, 26))[0], date(2026, 10, 25))
        self.assertIsNone(cal.corte_que_bloquea(date(2026, 10, 26), date(2026, 10, 26)))

    def test_cierre_de_mes(self):
        corte = cal.corte_que_bloquea(date(2026, 10, 27), date(2026, 11, 2))
        self.assertEqual(corte[0], date(2026, 10, 31))
        self.assertIn("cierre de mes", corte[1])
        self.assertIsNone(cal.corte_que_bloquea(date(2026, 10, 27), date(2026, 10, 30)))

    def test_pedida_a_tiempo_se_manda_al_dia_siguiente_sin_leyenda(self):
        # Llegó el día de la renovación pero después del ticket único (17:56): sale al siguiente día hábil aunque en
        # medio pase el cierre de mes o la emisión
        rows = [{'SQL_Fecha_Alta': date(2026, 10, 30), 'Fecha_Llegada_DT': datetime(2026, 10, 30, 18, 2)},
                {'SQL_Fecha_Alta': date(2026, 10, 24), 'Fecha_Llegada_DT': datetime(2026, 10, 24, 15, 30)},
                {'SQL_Fecha_Alta': date(2026, 10, 23), 'Fecha_Llegada_DT': datetime(2026, 10, 26, 9, 0)},
                {'Fecha_Llegada_DT': datetime(2026, 10, 26, 9, 0)}]    # sin la fecha de la base: no se marca
        res = cal.marcar_bloqueos(rows)
        self.assertEqual([r['Bloqueo_Emision'] == '' for r in rows], [True, True, False, True])
        self.assertEqual(res['bloqueadas'], 1)
        self.assertIn("se dio de alta el 23/10/2026 y el incidente llegó el 26/10/2026", rows[2]['Bloqueo_Emision'])
        self.assertIn("emisión de noviembre 2026 (25/10/2026)", rows[2]['Bloqueo_Emision'])

    def test_rezagada_del_dia_anterior_completa(self):
        # Ticket único del viernes 30/10 a las 17:56; el incidente llega 17:58 y va en el lote del lunes 02/11
        import sql_renovaciones as sql
        from test_sql_y_ocr import Conexion, fila
        conc = concentrado("Ticket único 30-10-2026", datetime(2026, 10, 30, 17, 56),
                           "<table><tr><td>INCIDENTE</td><td>CLIENTE</td><td>CREDITO</td></tr>"
                           "<tr><td>I-1</td><td>1111111</td><td>11111</td></tr></table>")
        r = fila("I-2", "7982848", "27179", datetime(2026, 10, 30, 17, 58))
        r['Cerrado_iTop'], r['Cierre_iTop'] = False, None
        clasificar_lotes([r], date(2026, 11, 2), date(2026, 11, 2), time(17, 50), [conc],
                         ahora=datetime(2026, 11, 2, 10, 0), revision=False)
        self.assertEqual(r['Lote_Estado'], "REZAGADA")
        self.assertFalse(r['No_Enviada'])
        # La base: dada de alta el día que llegó -> sin «No procede» ni bloqueo por el cierre del 31/10
        import test_sql_y_ocr
        test_sql_y_ocr.BASE.append((datetime(2026, 10, 30), '7982848', 27179))
        try:
            sql.validar_contra_base(Conexion(), "TOPAZ.CR_RENOVACIONES", [r])
        finally:
            test_sql_y_ocr.BASE.pop()
        self.assertEqual(r['SQL_Estado'], sql.SQL_OK)
        cal.marcar_bloqueos([r])
        self.assertEqual(r['Bloqueo_Emision'], '')

    def test_renglon_del_ticket_unico_enviado_al_dia_siguiente(self):
        import sql_renovaciones as sql
        from test_sql_y_ocr import Conexion
        # 27179 se dio de alta el 22/10 y salió en el ticket único del 23/10: con o sin su incidente, es del día
        con_incidente = {'cliente': '7982848', 'credito': '27179', 'fecha': date(2026, 10, 23),
                         'llegada': date(2026, 10, 22), 'desde': date(2026, 10, 22)}
        sin_incidente = {'cliente': '7982848', 'credito': '27179', 'fecha': date(2026, 10, 23),
                         'llegada': None, 'desde': date(2026, 10, 22)}
        sql.validar_contra_base(Conexion(), "TOPAZ.CR_RENOVACIONES", [], [con_incidente, sin_incidente])
        self.assertEqual((con_incidente['sql_estado'], sin_incidente['sql_estado']), (sql.SQL_OK, sql.SQL_OK))

    def test_lee_la_tabla_de_la_pagina(self):
        html = """<table><tr><td>Emisión de:</td><td>Generación de Emisión</td><td>Publicación</td><td>F.L.P</td></tr>
        <tr><td>FEBRERO '27</td><td>dom, 24-ene &#39;27</td><td>lun, 25-ene '27</td><td>vie, 05-mar '27</td></tr>
        <tr><td>MARZO '27</td><td>mié, 24-feb ​'27</td><td>jue, 25-feb '27</td><td>jue, 08-abr '27</td></tr></table>"""
        self.assertEqual(cal.leer_pagina(html), {(2027, 2): date(2027, 1, 24), (2027, 3): date(2027, 2, 24)})

    def test_las_fechas_incluidas_son_las_de_la_pagina(self):
        self.assertEqual(cal.EMISIONES_INCLUIDAS[(2026, 11)], date(2026, 10, 25))
        self.assertEqual(len(cal.EMISIONES_INCLUIDAS), 13)


if __name__ == '__main__':
    unittest.main()
