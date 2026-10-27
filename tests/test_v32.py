#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
v3.2: lo que no salió en su día pero ya se resolvió en iTop no cuenta como «NO enviada» (salvo cierre por error)
y lo demás de la auditoría de la v3.2.
"""

import json
import os
import sys
import tempfile
import unittest
from datetime import date, datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core_renovaciones import RenovacionTicketParser as Parser  # noqa: E402
import reglas_lote as reglas  # noqa: E402
from test_v30 import cerrada, renovacion  # noqa: E402


class TestCerradasNoSonFaltantes(unittest.TestCase):
    def test_resuelta_en_itop_no_cuenta_como_no_enviada(self):
        r = cerrada("Ya no requiere", No_Enviada=True, Dia_Lote=date(2026, 10, 26))
        reglas.aplicar_reglas([r])
        self.assertFalse(r['No_Enviada'])
        self.assertTrue(r['Cerrada_Sin_Envio'])
        self.assertEqual(r['Informativa_Tipo'], reglas.INFO_CERRADA)
        self.assertIn("ticket único del 26/10/2026", r['Informativa'])
        self.assertEqual(reglas.veredicto(r)['nivel'], reglas.NO_SE_MANDA)
        nivel, texto = reglas._check_envio(r)
        self.assertEqual(nivel, reglas.INFO)
        self.assertIn("ya se resolvió", texto)

    def test_cerrada_por_error_si_es_no_enviada(self):
        # Dice que es de otro día, la base dice que no y no está bloqueada (TZ_LOCK = 0): sí va
        r = cerrada("Es de otro día", No_Enviada=True, Dia_Lote=date(2026, 10, 26))
        reglas.aplicar_reglas([r])
        self.assertTrue(r['Cerrado_Por_Error'])
        self.assertTrue(r['No_Enviada'])
        self.assertFalse(r['Cerrada_Sin_Envio'])
        self.assertEqual(reglas._check_envio(r)[0], reglas.AVISO)  # no es un error: va hoy

    def test_abierta_no_cambia(self):
        r = renovacion(No_Enviada=True)
        reglas.aplicar_reglas([r])
        self.assertTrue(r['No_Enviada'])
        self.assertFalse(r.get('Cerrada_Sin_Envio'))

    def test_recontar_envios(self):
        dia = date(2026, 10, 26)
        filas = [cerrada("Ya no requiere", No_Enviada=True, Dia_Lote=dia),
                 cerrada("Es de otro día", No_Enviada=True, Dia_Lote=dia),
                 renovacion(No_Enviada=True, Dia_Lote=date(2026, 10, 27)),
                 renovacion(No_Enviada=True, Dia_Lote=dia, En_Periodo=False)]
        reglas.aplicar_reglas(filas)
        info = {'revision': True, 'no_enviadas': 4, 'cerradas_sin_envio': 0,
                'por_dia': [{'dia': dia, 'no_enviadas': 3, 'cerradas': 0}]}
        reglas.recontar_envios(info, filas)
        self.assertEqual(info['no_enviadas'], 2)  # la cerrada por error y la abierta; la de fuera del periodo no
        self.assertEqual(info['cerradas_sin_envio'], 1)
        self.assertEqual(info['por_dia'][0], {'dia': dia, 'no_enviadas': 1, 'cerradas': 1})
        reglas.recontar_envios(None, filas)  # sin lote: no hace nada


def creditos(desc):
    rows = Parser.parse_and_disaggregate_email("CANCELACION DE RENOVACION", "", desc, datetime(2026, 10, 27, 9, 0),
                                               "I-1", "2609", "Nancy")
    return [(r['Numero_Cliente'], r['Credito_Renovado_Hijo']) for r in rows]


class TestCreditosEnElRenglonDeAbajo(unittest.TestCase):
    def test_lista_de_creditos_abajo_de_la_etiqueta(self):
        desc = ("Buen día , por medio del presente solicito cancelar las renovaciones de crédito\n"
                "45375, 45377,45380,\n"
                "a nombre de la trabajadora EDITH FABIOLA BARRON ORTEGA , numero de cliente 2996318.\nGracias")
        self.assertEqual(creditos(desc), [("2996318", "45375"), ("2996318", "45377"), ("2996318", "45380")])

    def test_un_credito_abajo(self):
        self.assertEqual(creditos("Favor de cancelar la renovación del crédito:\n71470\nCliente 4919283"),
                         [("4919283", "71470")])

    def test_cliente_abajo_no_es_credito(self):
        # El número de 7 dígitos del renglón de abajo es el cliente, no un crédito
        filas = creditos("No se ve en el sistema de crédito\n4919283\nfavor de cancelar la renovación 71470")
        self.assertEqual([cred for _cli, cred in filas], ["71470"])

    def test_firma_o_capacidad_abajo_no_son_creditos(self):
        # Solo es el último recurso: si ya hay crédito no se usa, y no se toma de renglones que son de otra cosa
        filas = creditos("Favor de cancelar la renovación 71470 del cliente 4919283\nAtentamente\n"
                         "Analista de Crédito\n2609\nSucursal Centro")
        self.assertEqual([cred for _cli, cred in filas], ["71470"])
        self.assertNotIn("15000", [cred for _cli, cred in creditos(
            "Cliente 4919283, favor de cancelar su renovación.\nSu capacidad de crédito\n15000\n")])


class TestUnaSolaDecision(unittest.TestCase):
    """v3.2: una sola decisión por prioridad; lo que no decide no la contradice."""

    def emitida_cerrada(self, **extra):
        return cerrada("Es de otro día", No_Enviada=True, Dia_Lote=date(2026, 10, 26), Bloqueo_Tipo='emision',
                       Bloqueo_Emision="YA NO SE PUEDE CANCELAR: ... ya se emitió",
                       Bloqueo_Corto="se dio de alta el 23/10 y la generación de la emisión de noviembre 2026 del 25/10 "
                                     "pasó antes del ticket único del 27/10", **extra)

    def test_ya_se_emitio_gana_al_cierre_por_error(self):
        r = self.emitida_cerrada()
        conteo = reglas.aplicar_reglas([r])
        self.assertEqual(r['Informativa_Tipo'], reglas.INFO_EMISION)
        self.assertEqual((r['Cerrado_Por_Error'], r['No_Enviada']), ('', False))
        self.assertTrue(r['No_Enviada_Emitida'])
        self.assertEqual(conteo['cerrado_por_error'], 0)
        self.assertIn("ya se emitió", r['Lote_Resumen'])
        ver = reglas.veredicto(r)
        self.assertEqual(ver['nivel'], reglas.NO_SE_MANDA)
        self.assertTrue(ver['checks'][0][1].startswith("Calendario: ya se emitió"))
        self.assertEqual(ver['principales'], 1)
        # Lo demás queda como dato: sin ✖ ni ⚠ que contradigan «no se manda»
        self.assertFalse([e for e, _t in ver['checks'][1:] if e in (reglas.MAL, reglas.AVISO)])
        self.assertNotIn("POR ERROR", " ".join(t for _e, t in ver['checks']))

    def test_descartada_y_ademas_emitida(self):
        r = self.emitida_cerrada()
        reglas.aplicar_reglas([r])
        ver = reglas.veredicto(r, descartada=True)
        self.assertEqual(ver['nivel'], reglas.DESCARTADA)
        self.assertIn("De todos modos no se manda: ya se emitió", ver['accion'])
        self.assertFalse([e for e, _t in ver['checks'] if e == reglas.AVISO])

    def test_salio_con_datos_equivocados_se_corrige(self):
        for estado in (reglas.LOTE_ENVIADA, reglas.LOTE_EN_CONCENTRADO):
            r = renovacion(Lote_Estado=estado, Enviado_Fecha=date(2026, 10, 26), Error_Envio="crédito 10431")
            reglas.aplicar_reglas([r])
            ver = reglas.veredicto(r)
            self.assertEqual(ver['nivel'], reglas.CORREGIR, estado)
            self.assertTrue(ver['checks'][0][1].startswith("Ticket único: salió con datos equivocados"))

    def test_posible_cierre_por_error_se_revisa(self):
        r = cerrada("Es de otro día", SQL_Tz_Lock=None, Dia_Lote=date(2026, 10, 26))
        conteo = reglas.aplicar_reglas([r])
        self.assertTrue(r['Posible_Cierre_Error'])
        self.assertEqual((r['Informativa'], r['No_Enviada'], r['Cerrada_Sin_Envio']), ('', False, False))
        self.assertEqual(conteo['posible_cierre_error'], 1)
        self.assertFalse(reglas.va_en_lote(r))
        ver = reglas.veredicto(r)
        self.assertEqual((ver['nivel'], ver['titulo']), (reglas.REVISAR, "¿Se cerró por error?"))
        self.assertIn("POSIBLE CIERRE POR ERROR", ver['checks'][0][1])

    def test_atrasada_con_autorizacion(self):
        r = renovacion(Lote_Estado=reglas.LOTE_REZAGADA, No_Enviada=True, Dia_Lote=date(2026, 10, 26),
                       Bloqueo_Tipo='cierre', Bloqueo_Emision="NECESITA AUTORIZACIÓN: ...", Bloqueo_Corto="x")
        reglas.aplicar_reglas([r])
        ver = reglas.veredicto(r)
        self.assertEqual(ver['nivel'], reglas.AUTORIZACION)
        self.assertIn("mándala hoy", ver['accion'])
        self.assertEqual(reglas._check_envio(r)[0], reglas.AVISO)

    def test_textos_cortos(self):
        r = cerrada("Ya no requiere " * 30, Dia_Lote=date(2026, 10, 26))
        reglas.aplicar_reglas([r])
        itop = [t for _e, t in reglas.comprobaciones(r) if t.startswith("iTop")][0]
        self.assertLess(len(itop), 80)  # el motivo completo está en «Cierre en iTop», no se repite
        self.assertEqual(reglas._limpio("⏰ salió (en el ticket único (17:56))"), "salió (en el ticket único de las 17:56)")

    def test_recontar_emitidas(self):
        filas = [self.emitida_cerrada(), renovacion(No_Enviada=True, Lote_Estado=reglas.LOTE_REZAGADA)]
        reglas.aplicar_reglas(filas)
        info = {'no_enviadas': 2, 'cerradas_sin_envio': 0, 'por_dia': []}
        reglas.recontar_envios(info, filas)
        self.assertEqual((info['no_enviadas'], info['emitidas_sin_envio']), (1, 1))

    def test_comparacion_de_la_atrasada(self):
        r = cerrada("Es de otro día", No_Enviada=True, Dia_Lote=date(2026, 10, 26))
        reglas.aplicar_reglas([r])
        self.assertIn("le tocaba el 26/10/2026", str(reglas.comparacion(r, en_lote=True)))


class TestCalendarioSinFecha(unittest.TestCase):
    """Si el calendario no tiene la siguiente emisión no se puede dar por buena: avisa."""

    def setUp(self):
        import calendario_cedulas as cal
        self.cal = cal
        self._original = cal.emisiones
        cal.emisiones = lambda: {(2026, 9): date(2026, 9, 18), (2026, 10): date(2026, 10, 20)}

    def tearDown(self):
        self.cal.emisiones = self._original

    def test_cubierto_por_el_calendario(self):
        self.assertEqual(self.cal.emision_sin_fecha(date(2026, 10, 10), date(2026, 10, 27)), "")
        self.assertEqual(self.cal.emision_sin_fecha(date(2026, 10, 10), date(2026, 11, 9)), "")

    def test_falta_la_siguiente_emision(self):
        self.assertEqual(self.cal.emision_sin_fecha(date(2026, 10, 10), date(2026, 11, 15)), "noviembre 2026")

    def test_sin_fechas_no_avisa(self):
        self.assertEqual(self.cal.emision_sin_fecha(None, date(2026, 11, 15)), "")
        self.assertEqual(self.cal.emision_sin_fecha(date(2026, 11, 20), date(2026, 11, 15)), "")

    def test_la_regla_avisa_sin_calendario(self):
        nivel, texto = reglas._check_calendario(renovacion(Sin_Calendario="noviembre 2026"))
        self.assertEqual(nivel, reglas.AVISO)
        self.assertIn("emisión de noviembre 2026", texto)
        self.assertIn("no se sabe si ya se emitió", texto)

    def test_ya_emitida_pesa_mas(self):
        r = renovacion(Sin_Calendario="noviembre 2026", Bloqueo_Tipo="emision", Bloqueo_Emision="emisión de octubre")
        self.assertEqual(reglas._check_calendario(r)[0], reglas.MAL)


class TestGuardadoSeguro(unittest.TestCase):
    def setUp(self):
        self.carpeta = tempfile.mkdtemp()
        self.ruta = os.path.join(self.carpeta, "datos.json")

    def test_guarda_y_no_deja_temporal(self):
        from core_renovaciones import guardar_json
        guardar_json(self.ruta, {"a": 1, "ñ": "sí"}, ensure_ascii=False)
        guardar_json(self.ruta, {"a": 2})
        with open(self.ruta, encoding="utf-8") as f:
            self.assertEqual(json.load(f), {"a": 2})
        self.assertEqual(os.listdir(self.carpeta), ["datos.json"])

    def test_si_falla_no_estropea_el_anterior(self):
        from core_renovaciones import guardar_json
        guardar_json(self.ruta, {"a": 1})
        with self.assertRaises(TypeError):
            guardar_json(self.ruta, {"a": object()})
        with open(self.ruta, encoding="utf-8") as f:
            self.assertEqual(json.load(f), {"a": 1})


class TestRegistro(unittest.TestCase):
    def test_registra_el_error_con_su_detalle(self):
        import logging
        import registro
        ruta = os.path.join(tempfile.mkdtemp(), registro.ARCHIVO)
        original, registro.ruta, registro._log = registro.ruta, (lambda: ruta), None
        logging.getLogger("validador").handlers.clear()
        try:
            try:
                int("x")
            except ValueError:
                registro.error("Leyendo el correo I-123")
        finally:
            for h in list(registro.log().handlers):
                h.close()
                registro.log().removeHandler(h)
            registro.ruta, registro._log = original, None
        with open(ruta, encoding="utf-8") as f:
            texto = f.read()
        self.assertIn("Leyendo el correo I-123", texto)
        self.assertIn("ValueError", texto)


class TestPlurales(unittest.TestCase):
    def test_concuerda_con_el_numero(self):
        from lote_utils import pl
        texto = "# [renovación ya se emitió|renovaciones ya se emitieron]"
        self.assertEqual(pl(1, texto), "1 renovación ya se emitió")
        self.assertEqual(pl(3, texto), "3 renovaciones ya se emitieron")
        self.assertEqual(pl(0, "# no [salió|salieron]"), "0 no salieron")
        self.assertEqual(pl(2, "Se [validará|validarán] # [fila|filas]."), "Se validarán 2 filas.")


class TestRedDeSeguridadClientes(unittest.TestCase):
    """Si el correo trae más números de 7 dígitos que clientes leídos, se avisa (quizá falta una renovación)."""

    def leer(self, texto):
        return Parser.parse_and_disaggregate_email("CANCELACION DE RENOVACION", "", texto, datetime(2026, 10, 27, 9),
                                                   "I-1", "2609", "N")

    def test_avisa_el_numero_que_no_se_leyo(self):
        rows = self.leer("Cliente 4919283 crédito 71470\nTambién el 6210399 por favor")
        self.assertEqual(rows[0]['Clientes_Sin_Leer'], ["6210399"])
        nivel, texto = reglas._check_correo(rows[0])
        self.assertEqual(nivel, reglas.AVISO)
        self.assertIn("el número 6210399, que no se leyó como cliente", texto)

    def test_no_avisa_telefonos_ni_centro_de_trabajo(self):
        rows = self.leer("Cliente 4919283 crédito 71470\nTel. 5551234\nCentro de trabajo Grupo Bimbo 1234567\n"
                         "Folio 7654321")
        self.assertEqual(rows[0]['Clientes_Sin_Leer'], [])
        self.assertEqual(reglas._check_correo(rows[0])[0], reglas.OK)

    def test_no_avisa_lo_ya_leido(self):
        rows = self.leer("Cliente 4919283 crédito 71470\nCliente 5864479 crédito 56572")
        self.assertEqual(len(rows), 2)
        self.assertTrue(all(r['Clientes_Sin_Leer'] == [] for r in rows))


class TestGruposDelDia(unittest.TestCase):
    """Avisos e indicadores salen de la misma función: no se contradicen."""

    def setUp(self):
        self.filas = [
            dict(Row_ID="a", Cerrado_Por_Error=True, Lote_Estado=reglas.LOTE_CERRADA_SIN_ENVIO),
            dict(Row_ID="b", Cerrado_Por_Error=True, Lote_Estado=reglas.LOTE_CERRADA_SIN_ENVIO),
            dict(Row_ID="c", Informativa=True, Informativa_Tipo=reglas.INFO_EMISION),
            dict(Row_ID="d", Informativa=True, Informativa_Tipo=reglas.INFO_CERRADA),
            dict(Row_ID="e", Informativa=True, Informativa_Tipo=reglas.INFO_CERRADA),
            dict(Row_ID="f", Informativa=True, Informativa_Tipo=reglas.INFO_CERRADA, En_Periodo=False),
            dict(Row_ID="g", Posible_Cierre_Error=True, Lote_Estado=reglas.LOTE_CERRADA_SIN_ENVIO),
        ]

    def ids(self, filas):
        return [r['Row_ID'] for r in filas]

    def test_grupos(self):
        g = reglas.grupos_del_dia(self.filas, descartadas={"b"}, incluidas={"d"})
        self.assertEqual(self.ids(g['cerradas_error']), ["a"])  # la descartada no cuenta
        self.assertEqual(self.ids(g['posibles_cierre']), ["g"])
        self.assertEqual(self.ids(g['no_se_mandan']), ["c", "e", "f"])  # la metida a mano sí va
        self.assertEqual(reglas.texto_no_se_mandan(g['por_tipo']), "1 ya se emitió · 2 resueltas en iTop")

    def test_revision_solo_lo_del_periodo(self):
        g = reglas.grupos_del_dia(self.filas, solo_periodo=True)
        self.assertNotIn("f", self.ids(g['no_se_mandan']))

    def test_sin_nada(self):
        self.assertEqual(reglas.texto_no_se_mandan({}), "")


class TestTextosDelRecuadro(unittest.TestCase):
    def test_emitida_dice_que_hacer(self):
        r = renovacion(Informativa=True, Informativa_Tipo=reglas.INFO_EMISION, Lote_Estado=reglas.LOTE_HOY)
        v = reglas.veredicto(r, en_lote=False)
        self.assertEqual(v['titulo'], "No se manda: ya se emitió")
        self.assertTrue(v['accion'].startswith("Nada que hacer"))

    def test_resuelta_en_itop(self):
        r = renovacion(Informativa=True, Informativa_Tipo=reglas.INFO_CERRADA)
        v = reglas.veredicto(r, en_lote=False)
        self.assertEqual((v['titulo'], v['accion']), ("No se manda: resuelta en iTop", "Nada que hacer."))


if __name__ == "__main__":
    unittest.main()
