# -*- coding: utf-8 -*-
"""
Pruebas del motor de parseo y de la lógica de corte/rezagadas (no requieren Outlook).
Ejecutar desde la carpeta del proyecto:  python -m unittest discover -s tests -v
"""
import os
import sys
import unittest
from datetime import datetime, date, time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core_renovaciones import (  # noqa: E402
    RenovacionTicketParser as P,
    expand_abbreviated_labels,
    build_itop_url,
    build_itop_search_url,
    parse_hora_corte,
    parse_fecha_asunto,
    previous_business_day,
    extract_concentrado_tokens,
    clasificar_lotes,
    LOTE_HOY, LOTE_REZAGADA, LOTE_ENVIADA, LOTE_EN_CONCENTRADO, LOTE_SIGUIENTE,
)


def pares(resumen, descripcion="", tablas=None):
    return [(x['cliente'], x['credito_renovado']) for x in P.extract_disaggregated_renewals(resumen, descripcion, tablas or [])]


class TestAbreviaturas(unittest.TestCase):
    def test_letra_pegada_cliente_y_creditos(self):
        self.assertEqual(
            pares("CANCELACION DE RENOVACION C5864479 C56572 Y 56574"),
            [('5864479', '56572'), ('5864479', '56574')]
        )

    def test_t_y_r_separados(self):
        self.assertEqual(
            pares("CANCELACION DE RENOVACION", "Favor de cancelar la renovacion T 4530500 R 91258"),
            [('4530500', '91258')]
        )

    def test_cte_cred_con_punto(self):
        self.assertEqual(
            pares("CANCELACION DE RENOVACION", "Cte. 3386153 Cred. 61116, 61117"),
            [('3386153', '61116'), ('3386153', '61117')]
        )

    def test_guion(self):
        self.assertEqual(pares("CANCELACION RENOVACION C-5864479 C-56572"), [('5864479', '56572')])

    def test_abreviado_se_clasifica_como_renovacion(self):
        self.assertTrue(P.classify_renewal_ticket("CANC REN CTE 6173151 CRED 57137", "", [])[0])
        self.assertTrue(P.classify_renewal_ticket("CANC. RENOV. C5864479 C56572", "", [])[0])

    def test_no_toca_palabras_completas_ni_folios(self):
        txt = "Ticket I-361515 credito 56572 relacionado con R-012345 CP 06760"
        self.assertEqual(expand_abbreviated_labels(txt), txt)

    def test_tabla_credere_con_prefijos(self):
        tbl = [["Cuenta", "Operación (Origen)", "Renovado", "Sucursal", "Fecha"],
               ["C5864479", "C34232", "C57137", "123", "22/10/2026"]]
        recs = P._extract_crede_table_records([tbl], "")
        self.assertEqual([(r['cliente'], r['credito_padre'], r['credito_renovado']) for r in recs],
                         [('5864479', '34232', '57137')])

    def test_tabla_sucursal_con_prefijos(self):
        tbl = [["CLIENTE", "CREDITO", "FECHA"], ["T-4530500", "R91258", "22/10/2026"]]
        recs = P._extract_crede_table_records([tbl], "")
        self.assertEqual([(r['cliente'], r['credito_renovado']) for r in recs], [('4530500', '91258')])


class TestCasosHistoricos(unittest.TestCase):
    """Casos documentados en PLAN_DE_TRABAJO_RENOVACIONES.md: no deben romperse."""

    def test_i359504_dos_creditos(self):
        self.assertEqual(pares("CANCELACION DE RENOVACION CLIENTE 5864479 CREDITO 56572 Y 56574"),
                         [('5864479', '56572'), ('5864479', '56574')])

    def test_i359527_anterior_nuevo(self):
        r = P.extract_disaggregated_renewals(
            "CANCELACION DE RENOVACION", "CLIENTE: 6173151 CREDITO ANTERIOR: 34232 CREDITO NUEVO: 57137", [])
        self.assertEqual((r[0]['cliente'], r[0]['credito_renovado'], r[0]['credito_padre']), ('6173151', '57137', '34232'))

    def test_i359736_ya_renovado(self):
        r = P.extract_disaggregated_renewals(
            "CANCELACION DE RENOVACION",
            "Cliente: 4919283 Ya renovado: 71470 (favor de cancelar este) Antes de renovar: 81106", [])
        self.assertEqual((r[0]['cliente'], r[0]['credito_renovado'], r[0]['credito_padre']), ('4919283', '71470', '81106'))

    def test_i362259_pagare(self):
        r = P.extract_disaggregated_renewals(
            "CANCELACION DE RENOVACION",
            "cancelar la renovación del crédito 9291 con no. de pagaré 91258, del cliente con no. Fonacot 4530500", [])
        self.assertEqual((r[0]['cliente'], r[0]['credito_renovado'], r[0]['credito_padre']), ('4530500', '91258', '9291'))

    def test_signos_mas(self):
        self.assertEqual(pares("CANCELACION DE RENOVACION + 7239770+50074"), [('7239770', '50074')])


class TestItop(unittest.TestCase):
    def test_incident(self):
        self.assertEqual(build_itop_url("I-361515"),
                         "https://mservicio.fonacot.gob.mx/pages/UI.php?operation=details&class=Incident&id=361515")

    def test_ceros_y_userrequest(self):
        self.assertIn("class=UserRequest&id=12345", build_itop_url("R-012345"))
        self.assertIn("class=Incident&id=12345", build_itop_url("I-012345"))

    def test_plantilla(self):
        self.assertEqual(build_itop_url("I-361515", "https://x/UI.php?c={clase}&id={id}&r={ref}"),
                         "https://x/UI.php?c=Incident&id=361515&r=I-361515")

    def test_busqueda(self):
        self.assertTrue(build_itop_search_url("I-361515").endswith("operation=full_text&text=I-361515"))


class TestUtilidadesCorte(unittest.TestCase):
    def test_hora_corte(self):
        self.assertEqual(parse_hora_corte("17:50"), time(17, 50))
        self.assertEqual(parse_hora_corte("5:50 pm"), time(17, 50))
        self.assertEqual(parse_hora_corte("1750"), time(17, 50))
        self.assertEqual(parse_hora_corte("basura"), time(17, 50))

    def test_fecha_asunto(self):
        self.assertEqual(parse_fecha_asunto("TICKET UNICO 22/10/2026"), date(2026, 10, 22))
        self.assertEqual(parse_fecha_asunto("Ticket Único 22-10-26"), date(2026, 10, 22))
        self.assertEqual(parse_fecha_asunto("RE: Ticket único 22 de octubre", 2026), date(2026, 10, 22))
        self.assertIsNone(parse_fecha_asunto("Ticket único"))

    def test_dia_habil_anterior(self):
        self.assertEqual(previous_business_day(date(2026, 10, 23)), date(2026, 10, 22))  # viernes -> jueves
        self.assertEqual(previous_business_day(date(2026, 10, 26)), date(2026, 10, 23))  # lunes -> viernes

    def test_tokens_concentrado(self):
        tickets, nums = extract_concentrado_tokens("I-361515\tC5864479\t56572\nI-0362000 4530500 91258")
        self.assertEqual(tickets, {'361515', '362000'})
        self.assertTrue({'5864479', '56572', '4530500', '91258'} <= nums)


def fila(ticket, cli, cred, llegada, incompleta=False):
    return {
        'Row_ID': f"{ticket}_{cli}_{cred}", 'Numero_Ticket': ticket, 'Numero_Cliente': cli,
        'Credito_Renovado_Hijo': cred, 'Fecha_DT': llegada, 'Fecha_Llegada_DT': llegada,
        'Datos_Incompletos': incompleta,
    }


def concentrado(asunto, recibido, texto):
    tickets, nums = extract_concentrado_tokens(texto)
    return {'asunto': asunto, 'fecha_asunto': parse_fecha_asunto(asunto, recibido.year), 'recibido': recibido,
            'remitente': 'Compañero', 'carpeta': 'Bandeja', 'tickets': tickets, 'numeros': nums}


class TestClasificarLotes(unittest.TestCase):
    HOY = date(2026, 10, 23)
    CORTE = time(17, 50)

    def test_con_concentrado(self):
        rows = [
            fila("I-361001", "5864479", "56572", datetime(2026, 10, 22, 10, 0)),   # en concentrado -> enviada
            fila("I-361001", "5864479", "56574", datetime(2026, 10, 22, 10, 0)),   # mismo ticket pero NO enviada
            fila("I-361002", "4530500", "91258", datetime(2026, 10, 22, 18, 5)),   # después del corte de ayer
            fila("I-361003", "3386153", "61116", datetime(2026, 10, 23, 9, 0)),    # hoy
            fila("I-361004", "6173151", "57137", datetime(2026, 10, 23, 18, 1)),   # hoy después del corte
        ]
        conc = concentrado("TICKET UNICO 22/10/2026", datetime(2026, 10, 22, 17, 52), "I-361001 5864479 56572")
        info = clasificar_lotes(rows, self.HOY, self.HOY, self.CORTE, [conc])
        self.assertEqual([r['Lote_Estado'] for r in rows],
                         [LOTE_ENVIADA, LOTE_REZAGADA, LOTE_REZAGADA, LOTE_HOY, LOTE_SIGUIENTE])
        self.assertIn("NO este Cliente+Crédito", rows[1]['Lote_Etiqueta'])
        self.assertEqual(info['metodo'], 'concentrado')
        self.assertEqual(info['ultima_considerada']['ticket'], "I-361001")

    def test_sin_concentrado_usa_horario(self):
        rows = [
            fila("I-361001", "5864479", "56572", datetime(2026, 10, 22, 17, 49)),
            fila("I-361002", "4530500", "91258", datetime(2026, 10, 22, 17, 51)),
        ]
        info = clasificar_lotes(rows, self.HOY, self.HOY, self.CORTE, [])
        self.assertEqual(info['metodo'], 'horario')
        self.assertEqual([r['Lote_Estado'] for r in rows], [LOTE_ENVIADA, LOTE_REZAGADA])

    def test_marca_local(self):
        rows = [
            fila("I-361001", "5864479", "56572", datetime(2026, 10, 22, 17, 55)),
            fila("I-361002", "4530500", "91258", datetime(2026, 10, 22, 18, 30)),
        ]
        info = clasificar_lotes(rows, self.HOY, self.HOY, self.CORTE, [], {'hasta': '2026-10-22T17:55'})
        self.assertEqual(info['metodo'], 'marca')
        self.assertEqual([r['Lote_Estado'] for r in rows], [LOTE_ENVIADA, LOTE_REZAGADA])

    def test_lunes_incluye_fin_de_semana(self):
        lunes = date(2026, 10, 26)
        rows = [
            fila("I-361001", "5864479", "56572", datetime(2026, 10, 23, 16, 0)),  # viernes, enviada
            fila("I-361002", "4530500", "91258", datetime(2026, 10, 24, 11, 0)),  # sábado
        ]
        conc = concentrado("Ticket único 23/10/2026", datetime(2026, 10, 23, 17, 50), "5864479 56572")
        clasificar_lotes(rows, lunes, lunes, self.CORTE, [conc])
        self.assertEqual([r['Lote_Estado'] for r in rows], [LOTE_ENVIADA, LOTE_REZAGADA])

    def test_hoy_ya_en_concentrado(self):
        rows = [fila("I-361003", "3386153", "61116", datetime(2026, 10, 23, 9, 0))]
        conc = concentrado("TICKET UNICO 23/10/2026", datetime(2026, 10, 23, 17, 50), "3386153 61116")
        clasificar_lotes(rows, self.HOY, self.HOY, self.CORTE, [conc])
        self.assertEqual(rows[0]['Lote_Estado'], LOTE_EN_CONCENTRADO)

    def test_incompleta_basta_ticket(self):
        rows = [fila("I-361009", "SIN CLIENTE", "SIN CRÉDITO", datetime(2026, 10, 22, 12, 0), incompleta=True)]
        conc = concentrado("TICKET UNICO 22/10/2026", datetime(2026, 10, 22, 17, 52), "I-361009")
        clasificar_lotes(rows, self.HOY, self.HOY, self.CORTE, [conc])
        self.assertEqual(rows[0]['Lote_Estado'], LOTE_ENVIADA)


if __name__ == "__main__":
    unittest.main()
