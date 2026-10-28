#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
v3.0: lo que ya pasó la emisión es solo informativo (el cierre de mes pide autorización), lo resuelto en iTop no se
manda salvo que se haya cerrado por error (TZ_LOCK = 0), el veredicto del desglose, el crédito padre que se cambia por
su renovación y los incidentes con varios clientes donde alguno no tiene renovaciones.
"""

import os
import sys
import unittest
from datetime import date, datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import calendario_cedulas as cal  # noqa: E402
from core_renovaciones import RenovacionTicketParser as Parser, LOTE_ENVIADA  # noqa: E402
import reglas_lote as reglas  # noqa: E402
import sql_renovaciones as sql  # noqa: E402
from core_renovaciones import LOTE_HOY, LOTE_REZAGADA, LOTE_CERRADA_SIN_ENVIO  # noqa: E402
from test_sql_y_ocr import Conexion, Cursor, fila  # noqa: E402


def renovacion(**extra):
    """Renovación de hoy, con la base de datos que la confirma y el incidente abierto."""
    r = fila("I-1", "1184282", "10430", datetime(2026, 10, 23, 11, 0))
    r.update(Lote_Estado=LOTE_HOY, Lote_Etiqueta="Va en el ticket único de hoy", Cerrado_iTop=False, Cierre_iTop=None,
             SQL_Estado=sql.SQL_OK, SQL_Detalle="la renovación 10430 del cliente 1184282 se dio de alta el 23/10/2026",
             SQL_Fecha_Alta=date(2026, 10, 23), SQL_Tz_Lock=0, Bloqueo_Tipo='', Bloqueo_Emision='')
    r.update(extra)
    return r


def cerrada(motivo="Es de otro día", **extra):
    return renovacion(Lote_Estado=LOTE_CERRADA_SIN_ENVIO, Lote_Etiqueta="No se envió y ya se cerró", Cerrado_iTop=True,
                      Cierre_iTop={'motivo': motivo, 'cerro': 'Juan', 'fecha_dt': datetime(2026, 10, 23, 18, 0)},
                      **extra)


class TestEmisionYCierre(unittest.TestCase):
    def test_sin_base_usa_la_llegada_y_ya_se_emitio(self):
        # Llegó el viernes 23, la emisión de noviembre se generó el 25 y el lote es del lunes 26
        r = {'Fecha_Llegada_DT': datetime(2026, 10, 23, 19, 0), 'Dia_Lote': date(2026, 10, 26),
             'Lote_Estado': LOTE_REZAGADA}
        conteo = cal.marcar_bloqueos([r], date(2026, 10, 26))
        self.assertEqual(r['Bloqueo_Tipo'], 'emision')
        self.assertTrue(r['Bloqueo_Emision'].startswith("YA NO SE PUEDE CANCELAR: el incidente llegó el 23/10/2026"))
        self.assertIn("ya se emitió", r['Bloqueo_Emision'])
        self.assertEqual(conteo['emitidas'], 1)

    def test_emision_que_bloquea(self):
        gen = cal.emision_que_bloquea(date(2026, 10, 23), date(2026, 10, 26))
        self.assertEqual(gen[0], date(2026, 10, 25))
        self.assertIsNone(cal.emision_que_bloquea(date(2026, 10, 26), date(2026, 10, 27)))
        self.assertIsNone(cal.emision_que_bloquea(None, date(2026, 10, 27)))

    def test_emitida_es_informativa_y_no_va(self):
        r = renovacion(Lote_Estado=LOTE_REZAGADA, Bloqueo_Tipo='emision',
                       Bloqueo_Emision="YA NO SE PUEDE CANCELAR: ya se emitió")
        conteo = reglas.aplicar_reglas([r])
        self.assertEqual(r['Informativa_Tipo'], reglas.INFO_EMISION)
        self.assertFalse(reglas.va_en_lote(r))
        self.assertEqual(conteo['informativas'], 1)
        ver = reglas.veredicto(r)
        self.assertEqual(ver['nivel'], reglas.NO_SE_MANDA)
        self.assertIn("no se puede cancelar", ver['accion'])

    def test_en_la_cedula_tambien_es_informativa(self):
        r = renovacion(SQL_Cedula="Esta renovación ha emitido")
        reglas.aplicar_reglas([r])
        self.assertEqual(r['Informativa_Tipo'], reglas.INFO_EMISION)
        self.assertFalse(reglas.va_en_lote(r))

    def test_cierre_de_mes_pide_autorizacion_y_sigue_en_el_lote(self):
        r = renovacion(Bloqueo_Tipo='cierre', Bloqueo_Emision="NECESITA AUTORIZACIÓN: después del cierre de mes.")
        conteo = reglas.aplicar_reglas([r])
        self.assertTrue(r['Requiere_Autorizacion'])
        self.assertTrue(reglas.va_en_lote(r))
        self.assertEqual(conteo['autorizacion'], 1)
        self.assertEqual(reglas.veredicto(r)['nivel'], reglas.AUTORIZACION)

    def test_no_procede_no_se_manda(self):
        r = renovacion(SQL_Estado=sql.SQL_NO_PROCEDE, SQL_Detalle="renovación de hace meses")
        reglas.aplicar_reglas([r])
        self.assertEqual(r['Informativa_Tipo'], reglas.INFO_NO_PROCEDE)
        self.assertIn(sql.LEYENDA_NO_PROCEDE, reglas.veredicto(r)['accion'])


class TestCerradoPorError(unittest.TestCase):
    def test_dice_otro_dia_pero_no_esta_bloqueada(self):
        r = cerrada()
        conteo = reglas.aplicar_reglas([r])
        self.assertTrue(r['Cerrado_Por_Error'].startswith("CERRADO POR ERROR"))
        self.assertIn("TZ_LOCK = 0", r['Cerrado_Por_Error'])
        self.assertIn("«Es de otro día»", r['Cerrado_Por_Error'])
        self.assertEqual(r['Informativa'], '')
        self.assertTrue(reglas.va_en_lote(r))
        self.assertEqual(conteo['cerrado_por_error'], 1)
        ver = reglas.veredicto(r)
        # No hay dato que corregir: va hoy (atrasada) y se avisa a quien la cerró
        self.assertEqual(ver['nivel'], reglas.REVISAR)
        self.assertIn("mándala hoy", ver['accion'])
        self.assertIn("avisa a Juan que se cerró por error", ver['accion'])
        self.assertFalse([t for e, t in ver['checks'] if e == reglas.MAL])
        self.assertEqual(reglas.veredicto(r, validada=True)['nivel'], reglas.LISTA)

    def test_bloqueada_el_cierre_estuvo_bien(self):
        r = cerrada(SQL_Tz_Lock=7)
        reglas.aplicar_reglas([r])
        self.assertEqual((r['Cerrado_Por_Error'], r['Posible_Cierre_Error']), ('', ''))
        self.assertFalse(reglas.va_en_lote(r))
        self.assertEqual(reglas.veredicto(r)['nivel'], reglas.NO_SE_MANDA)

    def test_resuelta_hoy_es_informativa(self):
        # Sin TZ_LOCK (no se pudo leer): con TZ_LOCK = 0 sería «¿Resuelto sin cancelar?» (v3.3.4)
        r = renovacion(Cerrado_iTop=True, Cierre_iTop={'motivo': 'Se canceló en sucursal', 'cerro': 'Ana'},
                       SQL_Tz_Lock=None)
        reglas.aplicar_reglas([r])
        self.assertEqual(r['Informativa_Tipo'], reglas.INFO_CERRADA)
        self.assertFalse(reglas.va_en_lote(r))
        self.assertIn("lo cerró Ana", r['Informativa'])

    def test_sin_tz_lock_solo_avisa(self):
        r = cerrada(SQL_Tz_Lock=None)
        reglas.aplicar_reglas([r])
        self.assertEqual(r['Cerrado_Por_Error'], '')
        self.assertTrue(r['Posible_Cierre_Error'].startswith("POSIBLE CIERRE POR ERROR"))
        self.assertFalse(reglas.va_en_lote(r))
        self.assertTrue(any(e == reglas.AVISO and "POSIBLE" in t for e, t in reglas.comprobaciones(r)))

    def test_cerrada_por_otro_motivo_no_es_error(self):
        r = cerrada(motivo="Se canceló por sucursal")
        reglas.aplicar_reglas([r])
        self.assertEqual(r['Cerrado_Por_Error'], '')
        self.assertFalse(reglas.va_en_lote(r))

    def test_la_base_no_la_confirma_no_es_error(self):
        r = cerrada(SQL_Estado=sql.SQL_OTRA_FECHA)
        reglas.aplicar_reglas([r])
        self.assertEqual(r['Cerrado_Por_Error'], '')


class TestVeredicto(unittest.TestCase):
    def test_lista(self):
        r = renovacion()
        reglas.aplicar_reglas([r])
        ver = reglas.veredicto(r)
        self.assertEqual(ver['nivel'], reglas.LISTA)
        self.assertTrue(all(e in (reglas.OK, reglas.INFO) for e, _ in ver['checks']))
        fuentes = [t.split(":")[0] for _, t in ver['checks']]
        self.assertEqual(fuentes[:5], ["Correo", "Base de datos", "Calendario", "Ticket único", "iTop"])
        self.assertEqual(reglas.veredicto(r, validada=True)['corto'], "✓ Validada")

    def test_falta_dato_es_corregir(self):
        r = renovacion(Numero_Cliente="SIN CLIENTE", Datos_Incompletos=True)
        reglas.aplicar_reglas([r])
        ver = reglas.veredicto(r)
        self.assertEqual(ver['nivel'], reglas.CORREGIR)
        self.assertIn("Completa el dato", ver['accion'])

    def test_padre_corregido_pide_validar(self):
        r = renovacion(Credito_Renovado_Hijo="52011", Credito_Padre_Enviado="51937", SQL_Estado=sql.SQL_PADRE_CORREGIDO,
                       SQL_Detalle="VALIDAR: el incidente trae el crédito padre 51937")
        reglas.aplicar_reglas([r])
        ver = reglas.veredicto(r)
        self.assertEqual(ver['nivel'], reglas.REVISAR)
        self.assertIn("crédito padre 51937", ver['accion'])

    def test_descartada_y_enviada(self):
        r = renovacion()
        self.assertEqual(reglas.veredicto(r, descartada=True)['nivel'], reglas.DESCARTADA)
        r = renovacion(Lote_Estado=LOTE_ENVIADA, Enviado_Fecha=date(2026, 10, 22))
        reglas.aplicar_reglas([r])
        self.assertEqual(reglas.veredicto(r)['nivel'], reglas.ENVIADA)


class TestBaseDeDatos(unittest.TestCase):
    def test_tz_lock_de_la_renovacion(self):
        r = fila("I-1", "1184282", "10430")
        resumen = sql.validar_contra_base(Conexion(), "TOPAZ.CR_RENOVACIONES", [r])
        self.assertEqual(r['SQL_Tz_Lock'], 0)
        self.assertEqual(resumen['error_tz_lock'], '')

    def test_tz_lock_con_otro_nombre_de_columna(self):
        class SinGuion(Cursor):
            def execute(cur, texto, binds):
                if "TZ_LOCK" in texto:
                    raise Exception("ORA-00904: \"TZ_LOCK\": identificador no válido")
                if "TZLOCK" in texto:
                    texto = texto.replace("TZLOCK", "TZ_LOCK")
                super().execute(texto, binds)

        class Con(Conexion):
            def cursor(self):
                return SinGuion(self.consultas)

        self.assertEqual(sql.consultar_tz_lock(Con(), "TOPAZ.CR_RENOVACIONES", ["4363419"]), {('4363419', '17626'): 7})

    def test_cliente_del_incidente_sin_renovaciones(self):
        r = fila("I-5", "1184282", "10430")
        r['Clientes_Ticket'] = ["1184282", "9999999"]
        resumen = sql.validar_contra_base(Conexion(), "TOPAZ.CR_RENOVACIONES", [r])
        self.assertEqual(r['SQL_Clientes_Sin_Renov'],
                         ["El cliente 9999999 (también viene en el incidente) no cuenta con renovaciones"])
        self.assertEqual(resumen['clientes_sin_renov'], 1)
        self.assertTrue(any(e == reglas.AVISO and "9999999" in t for e, t in reglas.comprobaciones(r)))

    def test_cliente_con_renovaciones_viejas(self):
        r = fila("I-6", "1184282", "10430")
        r['Clientes_Ticket'] = ["1184282", "555"]
        res = sql.clientes_sin_renovaciones(r, {'555': [(date(2025, 1, 2), '777')]})
        self.assertEqual(res, [('555', "El cliente 555 (también viene en el incidente) no cuenta con renovaciones de "
                                       "esos días (la última: 777 del 02/01/2025)")])


class TestVariosClientes(unittest.TestCase):
    def test_el_incidente_guarda_todos_sus_clientes(self):
        items = Parser.extract_disaggregated_renewals(
            "CANCELACION DE RENOVACION", "Cliente 1184282 crédito 10430. Cliente 3386153 crédito 61116", [])
        self.assertTrue(items)
        for x in items:
            self.assertEqual(x.get('clientes_ticket'), ["1184282", "3386153"])


if __name__ == "__main__":
    unittest.main()
