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
    first_digits,
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
    LOTE_CERRADA_SIN_ENVIO, LOTE_SIN_COMPROBAR,
    set_dias_inhabiles, parse_dias_inhabiles, feriados_oficiales, next_business_day, cortes_reales,
    normalizar_fechas_concentrados,
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

    def test_asterisco_i362479(self):
        desc = ("Buenas tardes,\nSolicito su apoyo para la cancelación de la renovación siguiente:\nSaludos.\n"
                "Cancelación de renovación\n1\nEnviar por correo electrónico al 111\n"
                "El asunto en el correo electrónico que contenga:\n“Cancelación de crédito 1184282*10430\n"
                "Redacción sugerida:\nSolicito su apoyo para la cancelación de la renovación siguiente")
        self.assertEqual(pares("Cancelación de renovación", desc), [('1184282', '10430')])
        self.assertEqual(pares("1184282 * 10430 * 10431"), [('1184282', '10430'), ('1184282', '10431')])


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
            fila("I-361004", "6173151", "57137", datetime(2026, 10, 23, 18, 1)),   # hoy 18:01, aún no sale el de hoy
        ]
        conc = concentrado("TICKET UNICO 22/10/2026", datetime(2026, 10, 22, 17, 52), "I-361001 5864479 56572")
        info = clasificar_lotes(rows, self.HOY, self.HOY, self.CORTE, [conc])
        self.assertEqual([r['Lote_Estado'] for r in rows],
                         [LOTE_ENVIADA, LOTE_REZAGADA, LOTE_REZAGADA, LOTE_HOY, LOTE_HOY])
        self.assertIn("NO este Cliente+Crédito", rows[1]['Lote_Etiqueta'])
        self.assertEqual(info['metodo'], 'concentrado')
        self.assertEqual(info['ultima_considerada']['ticket'], "I-361001")

    def test_sin_concentrado_no_se_da_por_enviada(self):
        # Sin ticket único del día anterior no hay prueba: queda «sin comprobar» (antes se daba por enviada)
        rows = [
            fila("I-361001", "5864479", "56572", datetime(2026, 10, 22, 17, 49)),
            fila("I-361002", "4530500", "91258", datetime(2026, 10, 22, 17, 51)),
        ]
        info = clasificar_lotes(rows, self.HOY, self.HOY, self.CORTE, [])
        self.assertEqual(info['metodo'], 'horario')
        self.assertEqual([r['Lote_Estado'] for r in rows], [LOTE_SIN_COMPROBAR, LOTE_REZAGADA])

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


class TestRevisionEnvios(unittest.TestCase):
    """Periodo de varios días: cada renovación contra todos los tickets únicos enviados después."""
    CORTE = time(17, 50)
    AHORA = datetime(2026, 10, 23, 10, 0)
    CONC_30 = staticmethod(lambda: concentrado(
        "Cancelación de renovaciones – ticket único 30-09-2026", datetime(2026, 9, 30, 17, 45), "I-359504 5864479 56572"))

    def revisar(self, rows, concs, inicio=date(2026, 9, 30), fin=date(2026, 10, 23)):
        return clasificar_lotes(rows, inicio, fin, self.CORTE, concs, ahora=self.AHORA)

    def test_i359504_no_enviada_y_cerrada(self):
        rows = [fila("I-359504", "5864479", "56572", datetime(2026, 9, 30, 10, 0)),
                fila("I-359504", "5864479", "56574", datetime(2026, 9, 30, 10, 0))]
        rows[1].update(Cerrado_iTop=True, Cierre_iTop={'estado': 'Resuelto', 'fecha_dt': datetime(2026, 10, 1, 11, 6)})
        info = self.revisar(rows, [self.CONC_30()])
        self.assertEqual(rows[0]['Lote_Estado'], LOTE_ENVIADA)
        self.assertEqual(rows[0]['Lote_Resumen'], "Enviada 30/09")
        self.assertEqual(rows[1]['Lote_Estado'], LOTE_CERRADA_SIN_ENVIO)
        self.assertTrue(rows[1]['No_Enviada'])
        self.assertIn("debía ir en el ticket único del 30/09", rows[1]['Lote_Etiqueta'])
        self.assertIn("se cerró el 01/10", rows[1]['Lote_Etiqueta'])
        self.assertEqual((info['no_enviadas'], info['cerradas_sin_envio'], info['revision']), (1, 1, True))

    def test_no_enviada_abierta_va_hoy(self):
        rows = [fila("I-359504", "5864479", "56574", datetime(2026, 9, 30, 10, 0))]
        self.revisar(rows, [self.CONC_30()])
        self.assertEqual(rows[0]['Lote_Estado'], LOTE_REZAGADA)
        self.assertTrue(rows[0]['No_Enviada'])
        self.assertEqual(rows[0]['Lote_Resumen'], "NO ENVIADA (30/09)  ·  va hoy")

    def test_lunes_enviada_el_martes_no_cuenta_como_faltante(self):
        rows = [fila("I-362001", "4530500", "91258", datetime(2026, 10, 19, 11, 0))]
        concs = [concentrado("Ticket único 19-10-2026", datetime(2026, 10, 19, 17, 50), "I-362000 1111111 22222"),
                 concentrado("Ticket único 20-10-2026", datetime(2026, 10, 20, 17, 50), "I-362001 4530500 91258")]
        info = self.revisar(rows, concs, date(2026, 10, 19), date(2026, 10, 19))
        self.assertEqual(rows[0]['Lote_Estado'], LOTE_ENVIADA)
        self.assertFalse(rows[0]['No_Enviada'])
        self.assertEqual(rows[0]['Dias_Retraso'], 1)
        self.assertIn("Enviada tarde", rows[0]['Lote_Etiqueta'])
        self.assertEqual(info['tarde'], 1)

    def test_lunes_sin_enviar_en_ningun_dia(self):
        rows = [fila("I-362001", "4530500", "91258", datetime(2026, 10, 19, 11, 0))]
        concs = [concentrado("Ticket único 19-10-2026", datetime(2026, 10, 19, 17, 50), "I-362000 1111111 22222"),
                 concentrado("Ticket único 21-10-2026", datetime(2026, 10, 21, 17, 50), "I-362100 3333333 44444")]
        self.revisar(rows, concs, date(2026, 10, 19), date(2026, 10, 19))
        self.assertTrue(rows[0]['No_Enviada'])
        self.assertEqual(rows[0]['Dia_Lote'], date(2026, 10, 19))

    def test_dia_sin_ticket_unico_queda_sin_comprobar(self):
        rows = [fila("I-361500", "4530500", "91258", datetime(2026, 10, 14, 11, 0))]
        concs = [concentrado("Ticket único 22-10-2026", datetime(2026, 10, 22, 17, 50), "I-362100 3333333 44444")]
        self.revisar(rows, concs)
        self.assertEqual(rows[0]['Lote_Estado'], LOTE_SIN_COMPROBAR)
        self.assertFalse(rows[0]['No_Enviada'])

    def test_ya_enviada_antes_en_otro_ticket(self):
        rows = [fila("I-362300", "4530500", "91258", datetime(2026, 10, 21, 9, 0))]
        concs = [concentrado("Ticket único 20-10-2026", datetime(2026, 10, 20, 17, 50), "I-362001 4530500 91258"),
                 concentrado("Ticket único 21-10-2026", datetime(2026, 10, 21, 17, 50), "I-362100 3333333 44444")]
        self.revisar(rows, concs)
        self.assertEqual(rows[0]['Lote_Estado'], LOTE_ENVIADA)
        self.assertIn("antes de que llegara", rows[0]['Lote_Etiqueta'])

    def test_hoy_sigue_siendo_lote_de_hoy_en_revision(self):
        rows = [fila("I-362500", "4530500", "91258", datetime(2026, 10, 23, 9, 0)),
                fila("I-362501", "1234567", "11111", datetime(2026, 10, 23, 18, 0))]
        self.revisar(rows, [self.CONC_30()])
        self.assertEqual([r['Lote_Estado'] for r in rows], [LOTE_HOY, LOTE_HOY])


class TestDatosEquivocados(unittest.TestCase):
    """El ticket único trae el incidente pero con el cliente (o crédito, o folio) de otro."""
    AHORA = datetime(2026, 10, 23, 10, 0)

    @staticmethod
    def conc_tabla(fecha, recibido, filas):
        return {'asunto': f"Ticket único {fecha:%d-%m-%Y}", 'fecha_asunto': fecha, 'recibido': recibido,
                'remitente': '', 'carpeta': 'Enviados', 'tickets': {first_digits(f[0]) for f in filas},
                'numeros': {n for f in filas for n in f[1:]},
                'filas': [{'ticket': f[0], 'fecha': '', 'cliente': f[1], 'credito': f[2]} for f in filas]}

    def test_incidente_con_cliente_de_otro_incidente(self):
        rows = [fila("I-361001", "5864479", "56572", datetime(2026, 10, 20, 10, 0)),
                fila("I-361002", "4530500", "91258", datetime(2026, 10, 20, 11, 0))]
        conc = self.conc_tabla(date(2026, 10, 20), datetime(2026, 10, 20, 17, 50),
                               [("I-361001", "4530500", "56572"), ("I-361002", "4530500", "91258")])
        info = clasificar_lotes(rows, date(2026, 10, 19), date(2026, 10, 23), time(17, 50), [conc], ahora=self.AHORA)
        self.assertTrue(rows[0]['No_Enviada'])
        self.assertEqual(rows[0]['Error_Envio'],
                         "salió con Cliente 4530500 en lugar de 5864479 (ese cliente es del incidente I-361002)")
        self.assertIn("DATOS EQUIVOCADOS", rows[0]['Lote_Etiqueta'])
        self.assertEqual(rows[1]['Lote_Estado'], LOTE_ENVIADA)
        self.assertEqual(info['errores_envio'], 1)
        renglon = info['semana']['filas'][0]
        self.assertEqual(renglon['incidente'], "I-361001")
        self.assertIn("salió con Cliente 4530500 en lugar de 5864479", renglon['error'])

    def test_cliente_y_credito_bien_pero_con_otro_folio(self):
        rows = [fila("I-361001", "5864479", "56572", datetime(2026, 10, 20, 10, 0))]
        conc = self.conc_tabla(date(2026, 10, 20), datetime(2026, 10, 20, 17, 50), [("I-361009", "5864479", "56572")])
        info = clasificar_lotes(rows, date(2026, 10, 19), date(2026, 10, 23), time(17, 50), [conc], ahora=self.AHORA)
        self.assertEqual(rows[0]['Lote_Estado'], LOTE_ENVIADA)
        self.assertEqual(rows[0]['Error_Envio'], "salió con el incidente I-361009 en lugar de I-361001")
        self.assertIn("es del incidente I-361001", info['semana']['filas'][0]['error'])


class TestRevisionSemanal(unittest.TestCase):
    def test_concentrado_y_solo_cuenta_la_semana(self):
        rows = [fila("I-360900", "1111111", "11111", datetime(2026, 10, 16, 10, 0)),   # viernes anterior: fuera
                fila("I-361001", "5864479", "56572", datetime(2026, 10, 19, 10, 0)),   # lunes, enviada martes
                fila("I-361002", "4530500", "91258", datetime(2026, 10, 20, 10, 0))]   # martes, nunca enviada
        concs = [TestDatosEquivocados.conc_tabla(date(2026, 10, 16), datetime(2026, 10, 16, 17, 0), [("I-360800", "2222222", "22222")]),
                 TestDatosEquivocados.conc_tabla(date(2026, 10, 19), datetime(2026, 10, 19, 17, 0), [("I-360999", "3333333", "33333")]),
                 TestDatosEquivocados.conc_tabla(date(2026, 10, 20), datetime(2026, 10, 20, 17, 0), [("I-361001", "5864479", "56572")])]
        info = clasificar_lotes(rows, date(2026, 10, 19), date(2026, 10, 23), time(17, 50), concs,
                                ahora=datetime(2026, 10, 23, 10, 0))
        self.assertFalse(rows[0]['En_Periodo'])
        self.assertEqual((info['en_periodo'], info['enviadas'], info['no_enviadas'], info['tarde']), (2, 1, 1, 1))
        semana = info['semana']['filas']
        self.assertEqual([(f['ticket'], f['incidente']) for f in semana], [("I-360999", ""), ("I-361001", "I-361001")])


class TestLoteDeDiaPasado(unittest.TestCase):
    def test_lote_del_martes_visto_el_viernes(self):
        martes = date(2026, 10, 20)
        rows = [fila("I-361001", "5864479", "56572", datetime(2026, 10, 19, 18, 30)),  # lunes tras el envío -> rezagada
                fila("I-361002", "4530500", "91258", datetime(2026, 10, 20, 10, 0)),   # en el ticket del martes
                fila("I-361003", "7239770", "50074", datetime(2026, 10, 20, 11, 0)),   # faltó; salió el jueves
                fila("I-361004", "6173151", "57137", datetime(2026, 10, 20, 19, 0))]   # tras el envío -> miércoles
        concs = [TestDatosEquivocados.conc_tabla(date(2026, 10, 19), datetime(2026, 10, 19, 17, 0), [("I-360000", "1000000", "10000")]),
                 TestDatosEquivocados.conc_tabla(martes, datetime(2026, 10, 20, 17, 30), [("I-361002", "4530500", "91258")]),
                 TestDatosEquivocados.conc_tabla(date(2026, 10, 22), datetime(2026, 10, 22, 17, 30),
                                                 [("I-361003", "7239770", "50074"), ("I-361001", "5864479", "56572")])]
        info = clasificar_lotes(rows, martes, martes, time(17, 50), concs, ahora=datetime(2026, 10, 23, 10, 0),
                                revision=False, dia_lote=martes)
        self.assertFalse(info['revision'])
        self.assertEqual([r['Lote_Estado'] for r in rows], [LOTE_REZAGADA, LOTE_EN_CONCENTRADO, LOTE_HOY, LOTE_SIGUIENTE])
        self.assertEqual(rows[1]['Lote_Resumen'], "Ya en el ticket único del 20/10")
        self.assertIn("Lote del 20/10", rows[2]['Lote_Resumen'])
        self.assertIn("salió el 22/10", rows[2]['Lote_Resumen'])
        self.assertTrue(rows[2]['Falta_Hoy'])
        self.assertEqual(rows[0]['Enviado_Despues'], date(2026, 10, 22))
        # La rezagada del lunes tampoco vino en el ticket único del martes: también le faltó a ese lote
        self.assertTrue(rows[0]['Falta_Hoy'])
        self.assertIn("falta en el de las 17:30", rows[0]['Lote_Resumen'])
        self.assertEqual(info['faltan_hoy'], 2)

    def test_revision_forzada_en_un_solo_dia(self):
        rows = [fila("I-361002", "4530500", "91258", datetime(2026, 10, 23, 10, 0))]
        info = clasificar_lotes(rows, date(2026, 10, 23), date(2026, 10, 23), time(17, 50), [],
                                ahora=datetime(2026, 10, 23, 12, 0), revision=True)
        self.assertTrue(info['revision'])


class TestAuditoria(unittest.TestCase):
    """Escenarios encontrados en la auditoría: festivos, fechas de asunto, sin crédito, duplicados."""
    conc = staticmethod(lambda *a, **k: TestDatosEquivocados.conc_tabla(*a, **k))

    def tearDown(self):
        set_dias_inhabiles(())

    def test_feriados_y_dias_capturados(self):
        self.assertIn(date(2026, 11, 16), feriados_oficiales(2026))  # tercer lunes de noviembre
        self.assertEqual(next_business_day(date(2026, 11, 13)), date(2026, 11, 17))
        set_dias_inhabiles(parse_dias_inhabiles("02/11/2026, 12-12-2026"))
        self.assertEqual(next_business_day(date(2026, 10, 30)), date(2026, 11, 3))

    def test_dia_inhabil_con_ticket_unico_si_se_trabajo(self):
        rows = [fila("I-1", "5864479", "56572", datetime(2026, 11, 16, 10, 0))]
        concs = [self.conc(date(2026, 11, 16), datetime(2026, 11, 16, 17, 0), [("I-1", "5864479", "56572")])]
        clasificar_lotes(rows, date(2026, 11, 16), date(2026, 11, 20), time(17, 50), concs,
                         ahora=datetime(2026, 11, 20, 10, 0))
        self.assertEqual(rows[0]['Dia_Lote'], date(2026, 11, 16))
        self.assertEqual(rows[0]['Lote_Resumen'], "Enviada 16/11")

    def test_asunto_con_fecha_de_ayer_por_error(self):
        concs = [self.conc(date(2026, 10, 21), datetime(2026, 10, 21, 17, 30), []),
                 self.conc(date(2026, 10, 21), datetime(2026, 10, 22, 17, 52), [])]  # el del 22 dice "21"
        normalizar_fechas_concentrados(concs)
        self.assertEqual(concs[1]['fecha_efectiva'], date(2026, 10, 22))
        self.assertEqual(cortes_reales(concs), {date(2026, 10, 21): datetime(2026, 10, 21, 17, 30),
                                                date(2026, 10, 22): datetime(2026, 10, 22, 17, 52)})

    def test_complemento_no_mueve_el_corte(self):
        concs = [self.conc(date(2026, 10, 21), datetime(2026, 10, 21, 17, 30), []),
                 self.conc(date(2026, 10, 21), datetime(2026, 10, 21, 18, 40), [])]
        self.assertEqual(cortes_reales(concs)[date(2026, 10, 21)], datetime(2026, 10, 21, 17, 30))

    def test_sin_credito_necesita_su_cliente(self):
        rows = [fila("I-5", "1111111", "SIN CRÉDITO", datetime(2026, 10, 20, 10, 0), incompleta=True),
                fila("I-5", "2222222", "33333", datetime(2026, 10, 20, 10, 0))]
        concs = [self.conc(date(2026, 10, 20), datetime(2026, 10, 20, 17, 30), [("I-5", "2222222", "33333")])]
        clasificar_lotes(rows, date(2026, 10, 20), date(2026, 10, 23), time(17, 50), concs, ahora=datetime(2026, 10, 23, 10, 0))
        self.assertTrue(rows[0]['No_Enviada'])
        self.assertEqual(rows[1]['Lote_Estado'], LOTE_ENVIADA)

    def test_solicitud_repetida_hoy_no_se_autovalida(self):
        rows = [fila("I-200", "5864479", "56572", datetime(2026, 10, 23, 9, 0))]
        concs = [self.conc(date(2026, 10, 20), datetime(2026, 10, 20, 17, 30), [("I-100", "5864479", "56572")])]
        clasificar_lotes(rows, date(2026, 10, 23), date(2026, 10, 23), time(17, 50), concs, ahora=datetime(2026, 10, 23, 10, 0))
        self.assertEqual(rows[0]['Lote_Estado'], LOTE_ENVIADA)
        self.assertTrue(rows[0]['Posible_Duplicado'])
        self.assertIn("posible duplicado", rows[0]['Lote_Resumen'])

    def test_dos_incidentes_piden_lo_mismo_no_es_dato_equivocado(self):
        rows = [fila("I-100", "5864479", "56572", datetime(2026, 10, 20, 9, 0)),
                fila("I-200", "5864479", "56572", datetime(2026, 10, 20, 11, 0))]
        concs = [self.conc(date(2026, 10, 20), datetime(2026, 10, 20, 17, 30),
                           [("I-100", "5864479", "56572"), ("I-200", "5864479", "56572")])]
        info = clasificar_lotes(rows, date(2026, 10, 19), date(2026, 10, 23), time(17, 50), concs,
                                ahora=datetime(2026, 10, 23, 10, 0))
        self.assertEqual([r['Error_Envio'] for r in rows], ['', ''])
        self.assertEqual([f['error'] for f in info['semana']['filas']], ['', ''])
        self.assertEqual(info['semana']['filas'][1]['incidente'], "I-200")

    def test_resumen_por_dia(self):
        rows = [fila("I-1", "5864479", "56572", datetime(2026, 10, 19, 10, 0)),
                fila("I-2", "4530500", "91258", datetime(2026, 10, 19, 11, 0))]
        concs = [self.conc(date(2026, 10, 19), datetime(2026, 10, 19, 17, 30), [("I-1", "5864479", "56572")])]
        info = clasificar_lotes(rows, date(2026, 10, 19), date(2026, 10, 23), time(17, 50), concs,
                                ahora=datetime(2026, 10, 23, 10, 0))
        lunes = info['por_dia'][0]
        self.assertEqual((lunes['dia'], lunes['debian'], lunes['a_tiempo'], lunes['no_enviadas']),
                         (date(2026, 10, 19), 2, 1, 1))
        self.assertEqual(len(info['por_dia']), 5)


class TestCorteReal(unittest.TestCase):
    """El corte de cada día es la hora en que salió su ticket único, no una hora fija."""
    CORTE = time(17, 50)
    AHORA = datetime(2026, 10, 23, 20, 0)

    def clasificar(self, rows, concs):
        return clasificar_lotes(rows, date(2026, 10, 23), date(2026, 10, 23), self.CORTE, concs, ahora=self.AHORA)

    def test_ticket_de_ayer_salio_tarde(self):
        # Ayer el ticket único salió a las 19:10: lo de las 18:30 sí debía ir (y va), lo de las 19:30 va hoy
        rows = [fila("I-361001", "5864479", "56572", datetime(2026, 10, 22, 18, 30)),
                fila("I-361002", "4530500", "91258", datetime(2026, 10, 22, 19, 30))]
        conc = concentrado("Ticket único 22-10-2026", datetime(2026, 10, 22, 19, 10), "I-361001 5864479 56572")
        info = self.clasificar(rows, [conc])
        self.assertEqual([r['Lote_Estado'] for r in rows], [LOTE_ENVIADA, LOTE_REZAGADA])
        self.assertIn("después del envío del ticket único (19:10)", rows[1]['Lote_Etiqueta'])
        self.assertEqual(info['corte_anterior'], datetime(2026, 10, 22, 19, 10))

    def test_ticket_de_ayer_salio_temprano(self):
        # Salió a las 16:00: lo de las 17:00 no es un olvido, va hoy (antes, con 17:50 fijo, se daba por enviada)
        rows = [fila("I-361002", "4530500", "91258", datetime(2026, 10, 22, 17, 0))]
        conc = concentrado("Ticket único 22-10-2026", datetime(2026, 10, 22, 16, 0), "I-361001 5864479 56572")
        self.clasificar(rows, [conc])
        self.assertEqual(rows[0]['Lote_Estado'], LOTE_REZAGADA)
        self.assertFalse(rows[0]['No_Enviada'])

    def test_llego_mientras_se_armaba(self):
        # 10 min antes del envío y no viene: llegó mientras se armaba, no se marca como NO ENVIADA
        rows = [fila("I-361002", "4530500", "91258", datetime(2026, 10, 22, 17, 40))]
        conc = concentrado("Ticket único 22-10-2026", datetime(2026, 10, 22, 17, 50), "I-361001 5864479 56572")
        self.clasificar(rows, [conc])
        self.assertEqual(rows[0]['Lote_Estado'], LOTE_REZAGADA)
        self.assertFalse(rows[0]['No_Enviada'])
        self.assertIn("junto al envío", rows[0]['Lote_Resumen'])

    def test_hoy_antes_y_despues_del_envio(self):
        rows = [fila("I-362500", "4530500", "91258", datetime(2026, 10, 23, 9, 0)),   # viene en el de hoy
                fila("I-362501", "1234567", "11111", datetime(2026, 10, 23, 10, 0)),  # no viene: falta
                fila("I-362502", "7654321", "22222", datetime(2026, 10, 23, 19, 0))]  # después del envío
        conc = concentrado("Ticket único 23-10-2026", datetime(2026, 10, 23, 18, 40), "I-362500 4530500 91258")
        info = self.clasificar(rows, [conc])
        self.assertEqual([r['Lote_Estado'] for r in rows], [LOTE_EN_CONCENTRADO, LOTE_HOY, LOTE_SIGUIENTE])
        self.assertTrue(rows[1]['Falta_Hoy'])
        self.assertIn("falta en el de las 18:40", rows[1]['Lote_Resumen'])
        self.assertEqual(rows[2]['Dia_Lote'], date(2026, 10, 26))  # viernes después del envío -> lunes
        self.assertEqual((info['corte_hoy'], info['faltan_hoy']), (datetime(2026, 10, 23, 18, 40), 1))

    def test_respuesta_no_mueve_el_corte(self):
        rows = [fila("I-361002", "4530500", "91258", datetime(2026, 10, 22, 19, 0))]
        concs = [concentrado("Ticket único 22-10-2026", datetime(2026, 10, 22, 17, 50), "I-361001 5864479 56572"),
                 concentrado("RE: Ticket único 22-10-2026", datetime(2026, 10, 23, 9, 0), "I-361001 5864479 56572")]
        self.clasificar(rows, concs)
        self.assertEqual(rows[0]['Lote_Estado'], LOTE_REZAGADA)
        self.assertFalse(rows[0]['No_Enviada'])


if __name__ == "__main__":
    unittest.main()
