# -*- coding: utf-8 -*-
"""
Pruebas del formato real del correo "Ticket Único" y de los avisos de cierre de iTop (no requieren Outlook).
"""
import os
import sys
import unittest
from datetime import datetime, date, time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core_renovaciones import (  # noqa: E402
    extract_concentrado_rows,
    extract_concentrado_intro,
    extract_concentrado_tokens,
    parse_fecha_asunto,
    clasificar_lotes,
    is_aviso_solucion,
    parse_aviso_solucion,
    resumen_cierre,
    ticket_key,
    LOTE_ENVIADA, LOTE_REZAGADA,
)
from lote_utils import build_ticket_unico_html, build_ticket_unico_tsv, lote_corto, INTRO_TICKET_UNICO_DEFAULT  # noqa: E402
from outlook_renovaciones_engine import FonacotRenovacionesEngine  # noqa: E402

# Reproduce el correo "Cancelación de renovaciones – ticket único 22-10-2026"
# (tabla de datos dentro de una tabla de diseño, como la arma Outlook)
CONCENTRADO_HTML = """
<html><body>
<p>Hola,</p>
<p>Se envía “35” solicitudes de cancelaciones para su atención, favor de asignar a incidentes.</p>
<table><tr><td>
  <table border="1">
    <tr><td><b>INCIDENTE</b></td><td><b>FECHA</b></td><td><b>CLIENTE</b></td><td><b>CREDITO (HIJO)</b></td></tr>
    <tr><td><a href="#">I-362277</a></td><td>13/10/2026</td><td>4826460</td><td>82508</td></tr>
    <tr><td><a href="#">I-362359</a></td><td>22/10/2026</td><td>167461</td><td>99316</td></tr>
    <tr><td><a href="#">I-362359</a></td><td>22/10/2026</td><td>167461</td><td>99315</td></tr>
    <tr><td><a href="#">I-362377</a></td><td>22/10/2026</td><td>6994333</td><td>681</td></tr>
    <tr><td><a href="#">I-362386</a></td><td>22/10/2026</td><td>3499810</td><td>89615</td></tr>
  </table>
</td></tr></table>
</body></html>
"""

CONCENTRADO_PLAIN = (
    "Hola,\r\n\r\nSe envía “35” solicitudes de cancelaciones para su atención, favor de asignar a incidentes.\r\n\r\n"
    "FAVOR DE BLOQUEAR LOS REGISTROS DE TODOS LOS CRÉDITOS A CANCELAR\r\n\r\n"
    "INCIDENTE\tFECHA\tCLIENTE\tCREDITO (HIJO)\r\nI-362277\t13/10/2026\t4826460\t82508\r\n"
)

AVISO_HTML = """
<html><body><p>Apreciable Erika Rocío Navarro Vigueras</p>
<table>
 <tr><td>Ticket #</td><td>I-362386</td></tr>
 <tr><td>Resumen:</td><td>CANCELACION RENOVACION</td></tr>
 <tr><td>Nombre:</td><td>Erika Rocío Navarro Vigueras</td></tr>
 <tr><td>Prioridad:</td><td>Alta</td></tr>
 <tr><td>Estado:</td><td>Resuelto</td></tr>
 <tr><td>Categoría:</td><td>Mesa de Control de Crédito</td></tr>
 <tr><td>Objetivo de solución:</td><td>Por instrucciones superiores, las cancelaciones de renovación solo podrán
   realizarse el mismo día de dicha renovación</td></tr>
</table></body></html>
"""


def fila(ticket, cli, cred, llegada, **extra):
    r = {'Row_ID': f"{ticket}_{cli}_{cred}", 'Numero_Ticket': ticket, 'Numero_Cliente': cli, 'Credito_Renovado_Hijo': cred,
         'Credito_Origen_Padre': '-', 'Datos_Incompletos': False, 'Fecha_Llegada_DT': llegada, 'Fecha_DT': llegada,
         'Gafete': '7353', 'Nombre_Asignado': 'SAMMY'}
    r.update(extra)
    return r


def concentrado():
    tickets, numeros = extract_concentrado_tokens(CONCENTRADO_HTML)
    return {'asunto': 'Cancelación de renovaciones – ticket único 22-10-2026',
            'fecha_asunto': parse_fecha_asunto('Cancelación de renovaciones – ticket único 22-10-2026', 2026),
            'recibido': datetime(2026, 10, 22, 18, 2), 'remitente': 'Javier Servin Jordan', 'carpeta': 'Bandeja',
            'tickets': tickets, 'numeros': numeros, 'filas': extract_concentrado_rows(CONCENTRADO_HTML)}


class TestConcentradoReal(unittest.TestCase):
    def test_fecha_del_asunto_con_guiones(self):
        self.assertEqual(parse_fecha_asunto('Cancelación de renovaciones – ticket único 22-10-2026'), date(2026, 10, 22))

    def test_lee_la_tabla_renglon_por_renglon(self):
        filas = extract_concentrado_rows(CONCENTRADO_HTML)
        self.assertEqual(len(filas), 5)
        self.assertEqual(filas[0], {'ticket': 'I-362277', 'fecha': '13/10/2026', 'cliente': '4826460', 'credito': '82508'})
        self.assertEqual(filas[3]['credito'], '681')  # créditos de 3 dígitos

    def test_par_exacto_y_credito_corto(self):
        rows = [fila("I-362377", "6994333", "681", datetime(2026, 10, 22, 11, 0)),
                fila("I-362359", "167461", "99315", datetime(2026, 10, 22, 12, 0)),
                fila("I-362359", "167461", "99317", datetime(2026, 10, 22, 12, 0))]
        clasificar_lotes(rows, date(2026, 10, 23), date(2026, 10, 23), time(17, 50), [concentrado()])
        self.assertEqual([r['Lote_Estado'] for r in rows], [LOTE_ENVIADA, LOTE_ENVIADA, LOTE_REZAGADA])

    def test_rezagada_dice_que_trae_el_concentrado(self):
        r = fila("I-362386", "3499810", "89651", datetime(2026, 10, 22, 14, 53))
        clasificar_lotes([r], date(2026, 10, 23), date(2026, 10, 23), time(17, 50), [concentrado()])
        self.assertEqual(r['Lote_Estado'], LOTE_REZAGADA)
        self.assertIn("salió con Crédito 89615 en lugar de 89651", r['Lote_Etiqueta'])

    def test_mismo_ticket_y_credito_con_otro_cliente_es_error(self):
        # Antes se toleraba como "dígito mal capturado"; ahora se marca: el ticket único salió con otro cliente
        r = fila("I-362277", "4826406", "82508", datetime(2026, 10, 22, 9, 0))
        clasificar_lotes([r], date(2026, 10, 23), date(2026, 10, 23), time(17, 50), [concentrado()])
        self.assertEqual(r['Lote_Estado'], LOTE_REZAGADA)
        self.assertEqual(r['Error_Envio'], "salió con Cliente 4826460 en lugar de 4826406")
        self.assertIn("datos equivocados", r['Lote_Resumen'])

    def test_intro_reutilizable(self):
        intro = extract_concentrado_intro(CONCENTRADO_PLAIN)
        self.assertTrue(intro.startswith("Hola,"))
        self.assertIn("“{total}” solicitudes", intro)
        self.assertIn("FAVOR DE BLOQUEAR", intro)
        self.assertNotIn("INCIDENTE", intro)


class TestCorreoTicketUnico(unittest.TestCase):
    def test_html_con_total_y_ligas(self):
        rows = [fila("I-362277", "4826460", "82508", datetime(2026, 10, 13, 9, 0)),
                fila("I-362285", "7354077", "89097", datetime(2026, 10, 21, 9, 0))]
        html = build_ticket_unico_html(rows, INTRO_TICKET_UNICO_DEFAULT, url_fn=lambda t: f"https://itop/{t}")
        self.assertIn("“2” solicitudes", html)
        for h in ("INCIDENTE", "FECHA", "CLIENTE", "CREDITO (HIJO)"):
            self.assertIn(h, html)
        self.assertIn('<a href="https://itop/I-362277"', html)
        self.assertIn("13/10/2026", html)
        self.assertEqual(extract_concentrado_rows(html)[1]['credito'], '89097')  # el correo de hoy se lee mañana

    def test_tsv(self):
        tsv = build_ticket_unico_tsv([fila("I-362277", "4826460", "82508", datetime(2026, 10, 13, 9, 0))]).splitlines()
        self.assertEqual(tsv, ["INCIDENTE\tFECHA\tCLIENTE\tCREDITO (HIJO)", "I-362277\t13/10/2026\t4826460\t82508"])


class TestCierreItop(unittest.TestCase):
    def test_asunto(self):
        self.assertEqual(is_aviso_solucion("I-362386 Aviso de solución del Ticket"), "I-362386")
        self.assertEqual(is_aviso_solucion("RE: I-0362386 Aviso de solucion del ticket"), "I-362386")
        self.assertEqual(is_aviso_solucion("El Ticket I-362386 se le ha asignado al agente: 7353 SAMMY"), "")
        self.assertEqual(ticket_key("i-0362386"), "I-362386")

    def test_campos(self):
        c = parse_aviso_solucion(AVISO_HTML)
        self.assertEqual(c['estado'], "Resuelto")
        self.assertTrue(c['motivo'].startswith("Por instrucciones superiores"))
        self.assertTrue(c['motivo'].endswith("mismo día de dicha renovación"))
        self.assertEqual(c['solicitante'], "Erika Rocío Navarro Vigueras")
        self.assertNotIn('cerro', c)  # este aviso no dice quién lo cerró

    def test_campos_texto_plano(self):
        c = parse_aviso_solucion("", "Estado:\tResuelto\nSolución:\nNo procede, la renovación es de otro día\nAgente: 7353 Sammy")
        self.assertEqual(c['estado'], "Resuelto")
        self.assertEqual(c['motivo'], "No procede, la renovación es de otro día")
        self.assertEqual(c['cerro'], "7353 Sammy")

    def test_aplica_cierre_y_quien_lo_tenia(self):
        class Msg:
            HTMLBody = AVISO_HTML
            Body = ""
        rows = [fila("I-362386", "3499810", "89651", datetime(2026, 10, 22, 14, 53)),
                fila("I-362400", "1111111", "22222", datetime(2026, 10, 22, 15, 0))]
        timeline = {"I-362386": [(datetime(2026, 10, 22, 14, 53), '2609', 'NANCY'),
                                 (datetime(2026, 10, 22, 15, 10), '7353', 'Sammy Orlando Avila Coria'),
                                 (datetime(2026, 10, 23, 12, 0), '7341', 'MARIA')]}
        avisos = {"I-362386": (datetime(2026, 10, 22, 17, 48), Msg())}
        FonacotRenovacionesEngine._aplicar_cierres_itop(rows, avisos, timeline)
        self.assertTrue(rows[0]['Cerrado_iTop'])
        self.assertEqual(rows[0]['Cierre_iTop']['cerro'], "7353 Sammy Orlando Avila Coria")
        self.assertEqual(rows[0]['Cierre_iTop']['cerro_fuente'], "asignacion")
        self.assertFalse(rows[1]['Cerrado_iTop'])
        texto = resumen_cierre(rows[0]['Cierre_iTop'])
        self.assertTrue(texto.startswith("Resuelto 22/10 17:48  ·  7353 Sammy"))
        self.assertEqual(resumen_cierre(rows[0]['Cierre_iTop'], corto=True), "Resuelto 22/10 17:48  ·  7353 Sammy")

    def test_textos_cortos_de_la_tabla(self):
        self.assertEqual(
            lote_corto("Rezagada 22/10 14:53 (el incidente viene en el ticket único pero NO este Cliente+Crédito; ahí viene Cliente 1 Crédito 2)"),
            "Rezagada 22/10 14:53  ·  otros datos en el ticket único")
        self.assertEqual(lote_corto("Rezagada 22/10 12:56 (no viene en el ticket único)"), "Rezagada 22/10 12:56  ·  no salió")
        self.assertEqual(lote_corto("Lote de hoy"), "Lote de hoy")


if __name__ == "__main__":
    unittest.main()
