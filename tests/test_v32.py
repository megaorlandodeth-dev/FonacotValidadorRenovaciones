#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
v3.2: lo que no salió en su día pero ya se resolvió en iTop no cuenta como «NO enviada» (salvo cierre por error)
y la contraseña del perfil validador (solo se guarda su huella).
"""

import os
import sys
import tempfile
import unittest
from datetime import date, datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import clave_validador  # noqa: E402
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


class TestClaveValidador(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.TemporaryDirectory()
        self.archivo = os.path.join(self.dir.name, "validador_clave.json")

    def tearDown(self):
        self.dir.cleanup()

    def test_crear_y_verificar(self):
        self.assertFalse(clave_validador.existe(self.archivo))
        self.assertEqual(clave_validador.crear("secreta1", self.archivo), "")
        self.assertTrue(clave_validador.existe(self.archivo))
        self.assertTrue(clave_validador.verificar("secreta1", self.archivo))
        self.assertFalse(clave_validador.verificar("secreta2", self.archivo))
        self.assertFalse(clave_validador.verificar("", self.archivo))
        with open(self.archivo, encoding="utf-8") as f:
            self.assertNotIn("secreta1", f.read())  # solo la huella

    def test_minimo_de_caracteres(self):
        self.assertTrue(clave_validador.crear("abc", self.archivo))
        self.assertFalse(clave_validador.existe(self.archivo))

    def test_sin_archivo_o_danado(self):
        self.assertFalse(clave_validador.verificar("secreta1", self.archivo))
        with open(self.archivo, "w", encoding="utf-8") as f:
            f.write("{no es json")
        self.assertFalse(clave_validador.existe(self.archivo))
        self.assertFalse(clave_validador.verificar("secreta1", self.archivo))


if __name__ == "__main__":
    unittest.main()
