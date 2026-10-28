#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
v3.3.4: correos que ponen el crédito antes que la renovación o el cliente (tablas «Crédito | Renovación | Cliente»,
«No. de crédito … No. de renovación …», «se renovó con el …»), renovaciones que ya están canceladas en la base
(TZ_LOCK distinto de 0) e incidentes resueltos en iTop cuya renovación sigue sin cancelar (TZ_LOCK = 0).
"""

import os
import sys
import unittest
from datetime import date, datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core_renovaciones import RenovacionTicketParser as Parser, LOTE_ENVIADA  # noqa: E402
import reglas_lote as reglas  # noqa: E402
import sql_renovaciones as sql  # noqa: E402
from test_sql_y_ocr import Conexion  # noqa: E402
from test_v30 import renovacion, cerrada  # noqa: E402
import test_sql_y_ocr  # noqa: E402


def leer(desc="", html=""):
    rows = Parser.parse_and_disaggregate_email("CANCELACION DE RENOVACION", html, desc, datetime(2026, 10, 27, 9, 0),
                                               "I-1", "2609", "Nancy")
    return [(r['Numero_Cliente'], r['Credito_Renovado_Hijo'], r['Credito_Origen_Padre']) for r in rows]


def tabla(*filas):
    celdas = ''.join('<tr>' + ''.join(f'<td>{c}</td>' for c in f) + '</tr>' for f in filas)
    return ("<html><body><table><tr><td>Incidente #</td><td>I-1</td></tr>"
            "<tr><td>Resumen</td><td>CANCELACION DE RENOVACION</td></tr>"
            f"<tr><td>Descripción</td><td><div>Cancelar renovación</div><table>{celdas}</table></td></tr>"
            "</table></body></html>")


class TestCreditoAntesQueElCliente(unittest.TestCase):
    def test_tabla_credito_renovacion_cliente(self):
        self.assertEqual(leer(html=tabla(['Crédito', 'Renovación', 'Cliente'], ['65000', '71470', '4919283'])),
                         [('4919283', '71470', '65000')])

    def test_tabla_con_no_de(self):
        # Antes: el cliente salía 65000 y el crédito «SIN CRÉDITO»
        self.assertEqual(leer(html=tabla(['No. de crédito', 'No. de renovación', 'No. de cliente'],
                                         ['65000', '71470', '4919283'])), [('4919283', '71470', '65000')])

    def test_tabla_varias_filas(self):
        self.assertEqual(leer(html=tabla(['CREDITO', 'RENOVACION', 'CLIENTE'], ['65000', '71470', '4919283'],
                                         ['65001', '71471', '5864479'])),
                         [('4919283', '71470', '65000'), ('5864479', '71471', '65001')])

    def test_tabla_cliente_primero_sigue_igual(self):
        self.assertEqual(leer(html=tabla(['CLIENTE', 'CREDITO', 'FECHA'], ['4919283', '71470', '22/10/2026'])),
                         [('4919283', '71470', '-')])

    def test_tabla_nombre_del_cliente_no_es_el_numero(self):
        self.assertEqual(leer(html=tabla(['Nombre del cliente', 'No. cliente', 'Crédito'],
                                         ['JUAN PEREZ', '4919283', '71470'])), [('4919283', '71470', '-')])

    def test_texto_una_fila_credito_cliente(self):
        # Antes: «CLIENTE\n71470» tomaba el crédito como cliente
        self.assertEqual(leer("CRÉDITO\tCLIENTE\n71470\t4919283"), [('4919283', '71470', '-')])
        self.assertEqual(leer("CREDITO CLIENTE\n71470 4919283"), [('4919283', '71470', '-')])
        self.assertEqual(leer("Crédito\nCliente\n71470\n4919283"), [('4919283', '71470', '-')])

    def test_texto_tabla_credito_renovacion_cliente(self):
        self.assertEqual(leer("CREDITO RENOVACION CLIENTE\n65000 71470 4919283"), [('4919283', '71470', '65000')])
        self.assertEqual(leer("CRÉDITO RENOVACIÓN\n65000 71470\nCliente 4919283"), [('4919283', '71470', '65000')])

    def test_etiquetas_credito_y_renovacion(self):
        for desc in ("Crédito: 65000\nRenovación: 71470\nCliente: 4919283",
                     "No. de crédito 65000\nNo. de renovación 71470\nNo. de cliente 4919283",
                     "Cliente: 4919283\nCrédito: 65000\nRenovación: 71470",
                     "Renovación: 71470\nCrédito: 65000\nCliente: 4919283",
                     "credito: 65000 / renovacion: 71470 / cliente: 4919283"):
            self.assertEqual(leer(desc), [('4919283', '71470', '65000')], desc)

    def test_se_renovo_con(self):
        self.assertEqual(leer("Crédito 65000 se renovó con el 71470 del cliente 4919283"),
                         [('4919283', '71470', '65000')])
        self.assertEqual(leer("Crédito 65000 renovado con 71470, cliente 4919283"), [('4919283', '71470', '65000')])

    def test_varios_creditos_no_se_vuelven_padre(self):
        self.assertEqual([(c, r) for c, r, _ in leer("Cliente 4919283 créditos 71470 y 71471")],
                         [('4919283', '71470'), ('4919283', '71471')])
        self.assertEqual([(c, r) for c, r, _ in leer("Cliente 4919283\nCréditos 71470, 71471\nRenovación 71472")],
                         [('4919283', '71470'), ('4919283', '71471'), ('4919283', '71472')])

    def test_renovado_con_que_no_es_la_renovacion(self):
        # «renovado con» seguido de un plazo, un monto o un año no es la renovación
        for desc in ("Cliente 4919283 crédito 71470 renovado con 12 meses de plazo",
                     "Cliente 4919283 crédito 71470 renovado con 1200 de descuento mensual",
                     "Cliente 4919283 crédito 71470 renovado con 5000 pesos",
                     "Cliente 4919283 crédito 71470 se renovó en el 2026 sin autorización",
                     "Cliente 4919283 crédito 71470 renovado con error, favor de cancelar"):
            self.assertEqual(leer(desc), [('4919283', '71470', '-')], desc)

    def test_credito_y_renovacion_juntos_se_cancelan_los_dos(self):
        for desc in ("Cliente 4919283 cancelar el crédito 71470 y la renovación 71471",
                     "Cliente 4919283 cancelar ambas: crédito 71470 y renovación 71471"):
            self.assertEqual([(c, r) for c, r, _ in leer(desc)], [('4919283', '71470'), ('4919283', '71471')], desc)

    def test_encabezado_credito_renovacion_con_varias_filas(self):
        self.assertEqual(leer("CRÉDITO RENOVACIÓN\n65000 71470\n65001 71471\nCliente 4919283"),
                         [('4919283', '71470', '65000'), ('4919283', '71471', '65001')])

    def test_cliente_creditos_en_un_renglon(self):
        for desc in ("CLIENTE CREDITO\n4919283 71470 71471", "Cliente Créditos\n4919283 71470, 71471"):
            self.assertEqual(leer(desc), [('4919283', '71470', '-'), ('4919283', '71471', '-')], desc)

    def test_una_fila_con_el_credito_antes_aunque_el_encabezado_diga_cliente(self):
        self.assertEqual(leer("Cliente Crédito\n71470 4919283 2609"), [('4919283', '71470', '-')])
        self.assertEqual(leer("CREDITO CLIENTE\n71470 4919283\nGafete 2609"), [('4919283', '71470', '-')])

    def test_tabla_html_con_otros_nombres_de_columna(self):
        casos = [
            (['Cliente', 'Crédito Renovado', 'Crédito Nuevo'], [('4919283', '71470', '65000')]),
            (['Cliente', 'Crédito anterior', 'Crédito nuevo'], [('4919283', '71470', '65000')]),
            (['Cliente', 'Crédito', 'Crédito a cancelar'], [('4919283', '71470', '-')]),
            (['Cuenta Credere', 'Crédito'], None),
            (['Cliente', 'Pagaré'], None),
        ]
        for enc, esperado in casos:
            fila = ['4919283', '65000', '71470'][:len(enc)] if len(enc) == 3 else ['4919283', '71470']
            self.assertEqual(leer(html=tabla(enc, fila)), esperado or [('4919283', '71470', '-')], enc)

    def test_celdas_etiqueta_valor_no_son_encabezado(self):
        self.assertEqual(leer(html=tabla(['Cliente', '4919283'], ['Crédito', '71470'], ['Renovación', '71471'])),
                         [('4919283', '71471', '71470')])


class TestYaCancelada(unittest.TestCase):
    def test_tz_lock_distinto_de_cero_no_se_manda(self):
        r = renovacion(SQL_Tz_Lock=312610280362901, SQL_Estado=sql.SQL_MESA_CREDITO)
        conteo = reglas.aplicar_reglas([r])
        self.assertEqual(r['Informativa_Tipo'], reglas.INFO_CANCELADA)
        self.assertIn("ya está bloqueada", r['Informativa'])
        self.assertEqual(conteo[reglas.INFO_CANCELADA], 1)
        self.assertEqual(conteo['informativas'], 1)
        self.assertFalse(reglas.va_en_lote(r))
        ver = reglas.veredicto(r, en_lote=False, validada=True)
        self.assertEqual((ver['nivel'], ver['titulo']), (reglas.NO_SE_MANDA, "No se manda: ya está cancelada"))
        self.assertTrue(ver['checks'][0][1].startswith("Base de datos: la renovación ya está bloqueada"))
        self.assertEqual(reglas.texto_no_se_mandan({reglas.INFO_CANCELADA: 2}), "2 ya están canceladas en la base")

    def test_aunque_se_haya_cerrado_por_error_o_no_haya_salido(self):
        r = cerrada("Es de otro día", SQL_Tz_Lock=5, No_Enviada=True, Dia_Lote=date(2026, 10, 26))
        reglas.aplicar_reglas([r])
        self.assertEqual((r['Informativa_Tipo'], r['Cerrado_Por_Error'], r['No_Enviada']),
                         (reglas.INFO_CANCELADA, '', False))
        self.assertEqual(r['Lote_Resumen'], "No salió el 26/10  ·  ya está cancelada")

    def test_emitida_gana(self):
        r = renovacion(SQL_Tz_Lock=5, Bloqueo_Tipo='emision', Bloqueo_Emision="YA NO SE PUEDE CANCELAR: ya se emitió")
        reglas.aplicar_reglas([r])
        self.assertEqual(r['Informativa_Tipo'], reglas.INFO_EMISION)

    def test_ya_enviada_no_cambia(self):
        r = renovacion(Lote_Estado=LOTE_ENVIADA, Enviado_Fecha=date(2026, 10, 22), SQL_Tz_Lock=5)
        reglas.aplicar_reglas([r])
        self.assertEqual(r['Informativa_Tipo'], '')
        self.assertEqual(reglas.veredicto(r)['nivel'], reglas.ENVIADA)
        self.assertIn((reglas.INFO, "Base de datos: la renovación ya está bloqueada (TZ_LOCK = 5, distinto de 0): "
                                    "ya se canceló."), reglas.comprobaciones(r))

    def test_metida_a_mano_sigue_fuera(self):
        r = renovacion(SQL_Tz_Lock=5)
        reglas.aplicar_reglas([r])
        grupos = reglas.grupos_del_dia([r], incluidas={r['Row_ID']})
        self.assertEqual(grupos['no_se_mandan'], [r])
        self.assertEqual(grupos['por_tipo'], {reglas.INFO_CANCELADA: 1})

    def test_descartada(self):
        r = renovacion(SQL_Tz_Lock=5)
        reglas.aplicar_reglas([r])
        self.assertIn("ya está cancelada en la base", reglas.veredicto(r, descartada=True)['accion'])


class TestResueltoSinCancelar(unittest.TestCase):
    def test_resuelto_por_otro_motivo_y_sin_bloquear(self):
        r = renovacion(Cerrado_iTop=True, Cierre_iTop={'motivo': 'Ya no requiere', 'cerro': '7353 Sammy Avila',
                                                       'fecha_dt': datetime(2026, 10, 27, 13, 27)})
        conteo = reglas.aplicar_reglas([r])
        self.assertTrue(r['Cierre_Sin_Cancelar'])
        self.assertTrue(r['Posible_Cierre_Error'].startswith("RESUELTO SIN CANCELAR"))
        self.assertEqual((r['Informativa'], r['Cerrado_Por_Error']), ('', ''))
        self.assertEqual(conteo['cierre_sin_cancelar'], 1)
        self.assertFalse(reglas.va_en_lote(r))
        ver = reglas.veredicto(r)
        self.assertEqual((ver['nivel'], ver['corto']), (reglas.REVISAR, "¿Resuelto sin cancelar?"))
        self.assertIn("Confirma con 7353 Sammy Avila", ver['accion'])
        self.assertIn("«Restaurar» la manda hoy", ver['accion'])
        self.assertTrue(any(e == reglas.AVISO and "sigue sin cancelar" in t for e, t in reglas.comprobaciones(r)))
        self.assertEqual(reglas.grupos_del_dia([r])['posibles_cierre'], [r])
        # Si se mete a mano, va en el lote
        self.assertEqual(reglas.grupos_del_dia([r], incluidas={r['Row_ID']})['posibles_cierre'], [])

    def test_no_salio_en_su_dia(self):
        r = cerrada("Ya no requiere", No_Enviada=True, Dia_Lote=date(2026, 10, 26))
        reglas.aplicar_reglas([r])
        self.assertTrue(r['Cierre_Sin_Cancelar'])
        self.assertEqual((r['No_Enviada'], r['Cerrada_Sin_Envio']), (False, False))
        self.assertEqual(r['Lote_Resumen'], "¿Resuelto sin cancelar?  ·  no está bloqueada")

    def test_otro_dia_sigue_siendo_cierre_por_error(self):
        r = cerrada("Es de otro día")
        reglas.aplicar_reglas([r])
        self.assertTrue(r['Cerrado_Por_Error'])
        self.assertFalse(r['Cierre_Sin_Cancelar'])
        self.assertTrue(reglas.va_en_lote(r))

    def test_no_aplica(self):
        casos = {
            "sin TZ_LOCK": dict(SQL_Tz_Lock=None),
            "ya bloqueada": dict(SQL_Tz_Lock=9),
            "no procede": dict(SQL_Estado=sql.SQL_NO_PROCEDE),
            "no existe": dict(SQL_Estado=sql.SQL_NO_EXISTE),
            "emitida": dict(Bloqueo_Tipo='emision', Bloqueo_Emision="YA NO SE PUEDE CANCELAR"),
            "ya enviada": dict(Lote_Estado=LOTE_ENVIADA),
        }
        for nombre, extra in casos.items():
            r = renovacion(Cerrado_iTop=True, Cierre_iTop={'motivo': 'Ya no requiere', 'cerro': 'Ana'}, **extra)
            reglas.aplicar_reglas([r])
            self.assertFalse(r['Cierre_Sin_Cancelar'], nombre)

    def test_abierto_no_aplica(self):
        r = renovacion()
        reglas.aplicar_reglas([r])
        self.assertFalse(r['Cierre_Sin_Cancelar'])
        self.assertEqual(reglas.veredicto(r)['nivel'], reglas.LISTA)


class TestTzLockDelPadreCorregido(unittest.TestCase):
    def test_se_lee_el_de_la_renovacion(self):
        test_sql_y_ocr.TZLOCK[('1993891', '52011')] = 4
        try:
            r = test_sql_y_ocr.fila("I-9", "1993891", "51937")
            sql.validar_contra_base(Conexion(), "TOPAZ.CR_RENOVACIONES", [r])
        finally:
            del test_sql_y_ocr.TZLOCK[('1993891', '52011')]
        self.assertEqual((r['SQL_Estado'], r['Credito_Renovado_Hijo'], r['SQL_Tz_Lock']),
                         (sql.SQL_PADRE_CORREGIDO, '52011', 4))
        r['Lote_Estado'] = reglas.LOTE_HOY
        reglas.aplicar_reglas([r])
        self.assertEqual(r['Informativa_Tipo'], reglas.INFO_CANCELADA)


if __name__ == '__main__':
    unittest.main()
