#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Validación contra CR_RENOVACIONES (base simulada) y datos leídos de capturas de pantalla."""

import os
import sys
import unittest
from datetime import datetime, date

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core_renovaciones import completar_con_imagenes  # noqa: E402
from ocr_imagenes import extraer_cliente_credito, extraer_de_versiones  # noqa: E402
import sql_renovaciones as sql  # noqa: E402

# Lo que devolvió la consulta del 23/10/2026 para el cliente 1184282
BASE = [(datetime(2026, 10, 23), '1184282', 10430), (datetime(2025, 10, 22), '1184282', 84846),
        (datetime(2026, 10, 23), '4363419', 17626), (datetime(2026, 10, 22), '7982848', 27179),
        (datetime(2026, 10, 22), '3430084', 24969)]
# Segunda consulta: (cuenta, no_renovacion) -> (sucursal_ejerce, plazo de la renovación, plazo de la disposición, cédula)
DETALLE = {('1184282', '84846'): (12, 18, 18, 1), ('7982848', '27179'): (104, 12, 18, 0),
           ('3430084', '24969'): (104, 18, 18, 0)}


class Cursor:
    def __init__(self, consultas):
        self.consultas = consultas
        self._res = []

    def execute(self, texto, binds):
        self.consultas.append((texto, binds))
        valores = {str(v) for k, v in binds.items() if k[0] in 'cr'}
        if "CR_DISPOSICIONES" in texto:
            self._res = [(c, r, *d) for (c, r), d in DETALLE.items() if c in valores]
        elif "WHERE CUENTA IN" in texto:
            self._res = [x for x in BASE if x[1] in valores]
        else:
            self._res = [x for x in BASE if str(x[2]) in valores and binds['desde'] <= x[0] < binds['hasta']]

    def fetchall(self):
        return list(self._res)

    def close(self):
        pass


class Conexion:
    def __init__(self):
        self.consultas = []

    def cursor(self):
        return Cursor(self.consultas)


def fila(ticket, cli, cred, llegada=datetime(2026, 10, 23, 11, 0)):
    return {'Row_ID': ticket, 'Numero_Ticket': ticket, 'Numero_Cliente': cli, 'Credito_Renovado_Hijo': cred,
            'Fecha_Llegada_DT': llegada, 'Datos_Incompletos': cli == "SIN CLIENTE" or cred == "SIN CRÉDITO",
            'Indice_En_Ticket': 1, 'Total_En_Ticket': 1, 'Es_Multiple': False,
            'Alerta_Renovacion': "🚨 FALTA DATO", 'Nivel_Alerta': "incompleto"}


class TestValidacionSQL(unittest.TestCase):
    def validar(self, *rows, enviadas=()):
        conn = Conexion()
        resumen = sql.validar_contra_base(conn, "TOPAZ.CR_RENOVACIONES", list(rows), list(enviadas))
        return resumen, conn

    def test_renovacion_del_dia(self):
        r = fila("I-1", "1184282", "10430")
        resumen, conn = self.validar(r)
        self.assertEqual(r['SQL_Estado'], sql.SQL_OK)
        self.assertIn("el mismo día", r['SQL_Detalle'])
        self.assertEqual(resumen['problemas'], 0)
        # Solo SELECT con variables (los números no van pegados en el SQL)
        texto, binds = conn.consultas[0]
        self.assertTrue(texto.startswith("SELECT FECHA_ALTA, CUENTA, NO_RENOVACION FROM TOPAZ.CR_RENOVACIONES"))
        self.assertNotIn("1184282", texto)
        self.assertEqual(binds, {'c0': '1184282'})

    def test_renovacion_vieja_no_procede(self):
        r = fila("I-2", "1184282", "84846")
        resumen, conn = self.validar(r)
        self.assertEqual(r['SQL_Estado'], sql.SQL_NO_PROCEDE)
        self.assertTrue(r['SQL_Detalle'].startswith("No procede, la renovación es de otro día"))
        self.assertIn("22/10/2025", r['SQL_Detalle'])
        self.assertIn("no es de la APP", r['SQL_Detalle'])
        self.assertIn("10430 (23/10/2026)", r['SQL_Detalle'])
        self.assertEqual(r['SQL_Fecha_Alta'], date(2025, 10, 22))
        self.assertIn("ha emitido", r['SQL_Cedula'])
        self.assertEqual(resumen['problemas'], 1)
        # La segunda consulta también con variables y solo SELECT
        texto, binds = conn.consultas[-1]
        self.assertTrue(texto.startswith("SELECT c.CUENTA, c.NO_RENOVACION"))
        self.assertIn("TOPAZ.CR_DISPOSICIONES", texto)
        self.assertIn("TOPAZ.CR_CEDULAS_DETALLE", texto)
        self.assertNotIn("1184282", texto)

    def test_de_otro_dia_pero_de_la_app_con_plazo_menor(self):
        r = fila("I-8", "7982848", "27179")
        self.validar(r)
        self.assertEqual(r['SQL_Estado'], sql.SQL_APP_PLAZO)
        self.assertIn("se envía", r['SQL_Detalle'])
        self.assertIn("12 vs. disposición 18", r['SQL_Detalle'])
        self.assertNotIn(r['SQL_Estado'], sql.SQL_PROBLEMAS)

    def test_de_la_app_con_el_mismo_plazo_no_procede(self):
        r = fila("I-9", "3430084", "24969")
        self.validar(r)
        self.assertEqual(r['SQL_Estado'], sql.SQL_NO_PROCEDE)
        self.assertIn("el plazo no es menor", r['SQL_Detalle'])

    def test_si_falla_la_segunda_consulta_queda_otra_fecha(self):
        conn = Conexion()
        original = Cursor.execute

        def execute(cur, texto, binds):
            if "CR_DISPOSICIONES" in texto:
                raise RuntimeError("ORA-00942: table or view does not exist")
            original(cur, texto, binds)
        Cursor.execute = execute
        try:
            r = fila("I-10", "7982848", "27179")
            resumen = sql.validar_contra_base(conn, "TOPAZ.CR_RENOVACIONES", [r])
        finally:
            Cursor.execute = original
        self.assertEqual(r['SQL_Estado'], sql.SQL_OTRA_FECHA)
        self.assertIn("ORA-00942", r['SQL_Detalle'])
        self.assertIn("ORA-00942", resumen['error_detalle'])

    def test_credito_que_el_cliente_no_tiene(self):
        r = fila("I-3", "1184282", "10431")
        resumen, _ = self.validar(r)
        self.assertEqual(r['SQL_Estado'], sql.SQL_NO_EXISTE)
        self.assertIn("¿Será 10430?", r['SQL_Detalle'])
        self.assertEqual(resumen['problemas'], 1)

    def test_cliente_de_otro_incidente(self):
        r = fila("I-4", "1184282", "17626")
        self.validar(r)
        self.assertEqual(r['SQL_Estado'], sql.SQL_NO_EXISTE)
        self.assertIn("es del cliente 4363419", r['SQL_Detalle'])

    def test_sugiere_el_dato_que_falta(self):
        sin_credito, sin_cliente = fila("I-5", "1184282", "SIN CRÉDITO"), fila("I-6", "SIN CLIENTE", "17626")
        self.validar(sin_credito, sin_cliente)
        self.assertEqual(sin_credito['SQL_Estado'], sql.SQL_SUGERENCIA)
        self.assertIn("10430", sin_credito['SQL_Detalle'])
        self.assertIn("cliente 4363419", sin_cliente['SQL_Detalle'])

    def test_cliente_sin_renovaciones(self):
        r = fila("I-7", "9999999", "12345")
        self.validar(r)
        self.assertEqual(r['SQL_Estado'], sql.SQL_SIN_RENOVACIONES)

    def test_renglones_del_ticket_unico(self):
        enviada = {'cliente': '1184282', 'credito': '84846', 'fecha': date(2026, 10, 23)}
        resumen, _ = self.validar(fila("I-1", "1184282", "10430"), enviadas=[enviada])
        self.assertEqual(enviada['sql_estado'], sql.SQL_NO_PROCEDE)
        self.assertEqual(resumen['enviadas_con_problema'], 1)

    def test_nombre_de_tabla_seguro(self):
        self.assertEqual(sql.tabla_valida(" topaz.cr_renovaciones "), "topaz.cr_renovaciones")
        with self.assertRaises(sql.ErrorSQL):
            sql.tabla_valida("TOPAZ.CR_RENOVACIONES; DELETE FROM X")

    def test_servidor_como_en_sql_developer(self):
        self.assertEqual(sql.armar_dsn(" 10.0.0.5:1521/TOPAZ "), "10.0.0.5:1521/TOPAZ")
        # Separado por espacios o comas ya no se busca como alias de tnsnames.ora (DPY-4000 «unable to find "1521"»)
        for texto in ("scanfonacot.fonacot.gob.mx 1521 fonatop", "scanfonacot.fonacot.gob.mx, 1521, fonatop"):
            self.assertEqual(sql.armar_dsn(texto), "scanfonacot.fonacot.gob.mx:1521/fonatop")
        self.assertEqual(sql.armar_dsn("PRODUCCION"), "PRODUCCION")  # un alias sí se deja igual
        self.assertEqual(sql.servidor_de("scanfonacot.fonacot.gob.mx", "", "fonatop"),
                         "scanfonacot.fonacot.gob.mx:1521/fonatop")
        self.assertEqual(sql.separar_servidor("scanfonacot.fonacot.gob.mx:1521/fonatop"),
                         ("scanfonacot.fonacot.gob.mx", "1521", "fonatop"))


def palabra(texto, x, y, ancho=None, alto=20):
    return (texto, x, y, ancho or len(texto) * 10, alto)


class TestLecturaDeImagen(unittest.TestCase):
    def test_pantalla_consulta_de_credito(self):
        lineas = [[palabra("250013,", 0, 0), palabra("Consulta", 80, 0), palabra("de", 170, 0), palabra("Crédito,", 200, 0),
                   palabra("Transacción", 290, 0), palabra("3", 410, 0)],
                  [palabra("No.", 10, 60), palabra("Cliente", 45, 60), palabra("4363419", 200, 60),
                   palabra("GALVAN", 400, 60), palabra("REYNA", 480, 60)],
                  [palabra("Datos", 10, 100), palabra("del", 70, 100), palabra("Credito", 110, 100)],
                  [palabra("No.", 10, 130), palabra("Credito", 45, 130), palabra("17626", 200, 130)],
                  [palabra("Producto", 10, 160), palabra("301", 200, 160)],
                  [palabra("Importe", 500, 190), palabra("Ejercido", 580, 190), palabra("12,158.24", 700, 190)]]
        self.assertEqual({k: v for k, v in extraer_cliente_credito(lineas).items() if k != 'texto'},
                         {'cliente': '4363419', 'credito': '17626'})

    def test_etiquetas_ilegibles_usa_el_nombre_y_el_titulo(self):
        # Texto blanco sobre gris: el OCR no lee «No. Cliente» ni «No. Credito»
        lineas = [[palabra("Datos", 0, 100), palabra("del", 60, 100), palabra("Credito", 100, 100)],
                  [palabra("4363419", 300, 60), palabra("GALVAN", 700, 60), palabra("REYNA", 800, 60)],
                  [palabra("17626", 350, 135)]]
        datos = extraer_cliente_credito(lineas)
        self.assertEqual((datos['cliente'], datos['credito']), ('4363419', '17626'))

    def test_resultado_de_sql(self):
        lineas = [[palabra("WHERE", 0, 0), palabra("CUENTA", 70, 0), palabra("IN", 150, 0), palabra("('1184282')", 180, 0)],
                  [palabra("FECHA_ALTA", 20, 50), palabra("CUENTA", 160, 50), palabra("NO_RENOVACION", 260, 50)],
                  [palabra("23/10/26", 20, 80), palabra("1184282", 160, 80), palabra("10430", 300, 80)],
                  [palabra("22/10/25", 20, 108), palabra("1184282", 160, 108), palabra("84846", 300, 108)]]
        datos = extraer_cliente_credito(lineas)
        self.assertEqual((datos['cliente'], datos['credito']), ('1184282', '10430'))

    def test_lecturas_distintas_quedan_dudosas(self):
        a = [[palabra("Cliente", 0, 0), palabra("4363419", 100, 0)], [palabra("Credito", 0, 40), palabra("17626", 100, 40)]]
        b = [[palabra("Cliente", 0, 0), palabra("4363419", 100, 0)], [palabra("Credito", 0, 40), palabra("17625", 100, 40)]]
        self.assertTrue(extraer_de_versiones([a, b])['dudosa'])
        self.assertFalse(extraer_de_versiones([a, a])['dudosa'])


class TestCompletarConImagenes(unittest.TestCase):
    def test_incidente_solo_con_captura(self):
        r = fila("I-362557", "SIN CLIENTE", "SIN CRÉDITO")
        filas = completar_con_imagenes([r], [{'cliente': '4363419', 'credito': '17626', 'imagen': 'captura.png'}])
        self.assertEqual(len(filas), 1)
        self.assertEqual((r['Numero_Cliente'], r['Credito_Renovado_Hijo']), ('4363419', '17626'))
        self.assertFalse(r['Datos_Incompletos'])
        self.assertIn("leído de la imagen «captura.png»", r['Dato_Imagen'])
        self.assertIn("revísalo", r['Alerta_Renovacion'])

    def test_solo_completa_lo_que_falta_y_si_coincide(self):
        r = fila("I-1", "1184282", "SIN CRÉDITO")
        completar_con_imagenes([r], [{'cliente': '4363419', 'credito': '17626'}])  # otra persona: no se usa
        self.assertEqual(r['Credito_Renovado_Hijo'], "SIN CRÉDITO")
        completar_con_imagenes([r], [{'cliente': '1184282', 'credito': '10430'}])
        self.assertEqual(r['Credito_Renovado_Hijo'], "10430")

    def test_una_renovacion_por_imagen(self):
        r = fila("I-9", "SIN CLIENTE", "SIN CRÉDITO")
        filas = completar_con_imagenes([r], [{'cliente': '4363419', 'credito': '17626', 'imagen': 'a.png'},
                                             {'cliente': '1184282', 'credito': '10430', 'imagen': 'b.png'}])
        self.assertEqual([(f['Numero_Cliente'], f['Credito_Renovado_Hijo'], f['Indice_En_Ticket'], f['Total_En_Ticket'])
                          for f in filas], [('4363419', '17626', 1, 2), ('1184282', '10430', 2, 2)])
        self.assertEqual(len({f['Row_ID'] for f in filas}), 2)

    def test_datos_completos_no_se_tocan(self):
        r = fila("I-1", "1184282", "10430")
        completar_con_imagenes([r], [{'cliente': '4363419', 'credito': '17626'}])
        self.assertEqual((r['Numero_Cliente'], r['Credito_Renovado_Hijo']), ("1184282", "10430"))
        self.assertNotIn('Dato_Imagen', r)


if __name__ == '__main__':
    unittest.main()
