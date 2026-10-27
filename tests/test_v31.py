#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
v3.1: trabajo en equipo por carpeta compartida (cada analista publica lo que validó y sus dudas; el validador lo lee),
bitácora y dudas por renovación, y la comparación incidente | base de datos | ticket único del desglose.
"""

import json
import os
import shutil
import sys
import tempfile
import unittest
from datetime import date, datetime, timedelta

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import equipo_compartido as equipo  # noqa: E402
import reglas_lote as reglas  # noqa: E402
import sql_renovaciones as sql  # noqa: E402
from core_renovaciones import LOTE_ENVIADA, LOTE_HOY, LOTE_SIGUIENTE  # noqa: E402
from lote_utils import EstadoValidaciones, llaves_fila, row_key  # noqa: E402
from test_v30 import renovacion  # noqa: E402

HOY = date(2026, 10, 26)


class Carpeta(unittest.TestCase):
    def setUp(self):
        self.carpeta = tempfile.mkdtemp(prefix="equipo_")

    def tearDown(self):
        shutil.rmtree(self.carpeta, ignore_errors=True)


class TestCarpetaCompartida(Carpeta):
    def publicar(self, gafete, nombre, marcas, dia=HOY, avance=None):
        self.assertEqual(equipo.publicar(self.carpeta, gafete, nombre, dia, marcas, avance), "")

    def test_publica_y_el_validador_lo_lee(self):
        r = renovacion()
        k = row_key(r)
        self.publicar("2609", "Nancy Pérez", {k: equipo.marca(r, equipo.VALIDADA, datetime(2026, 10, 26, 11, 32))},
                      avance={"total": 4, "validadas": 1})
        res = equipo.leer_equipo(self.carpeta, HOY, ahora=datetime(2026, 10, 26, 11, 40))
        self.assertEqual(res["errores"], [])
        self.assertEqual([a["gafete"] for a in res["analistas"]], ["2609"])
        analista = res["analistas"][0]
        self.assertFalse(analista["sin_senal"])
        self.assertEqual(analista["avance"], {"total": 4, "validadas": 1})
        m = res["marcas"][k]
        self.assertEqual((m["estado"], m["gafete"], m["nombre"], m["cliente"]),
                         (equipo.VALIDADA, "2609", "Nancy Pérez", "1184282"))
        # Escritura atómica: no quedan temporales en la carpeta
        self.assertEqual(os.listdir(self.carpeta), ["marcas_2609.json"])

    def test_cada_analista_su_archivo_y_republicar_reemplaza(self):
        r = renovacion()
        self.publicar("2609", "Nancy", {row_key(r): equipo.marca(r, equipo.VALIDADA)})
        self.publicar("3100", "Luis", {})
        self.publicar("2609", "Nancy", {})  # quitó la validación
        res = equipo.leer_equipo(self.carpeta, HOY)
        self.assertEqual(sorted(a["gafete"] for a in res["analistas"]), ["2609", "3100"])
        self.assertEqual(res["marcas"], {})

    def test_ignora_otro_dia_y_avisa_archivo_danado(self):
        self.publicar("2609", "Nancy", {"x": {"estado": equipo.VALIDADA, "hora": "2026-10-23T10:00:00"}},
                      dia=date(2026, 10, 23))
        with open(os.path.join(self.carpeta, "marcas_9999.json"), "w", encoding="utf-8") as f:
            f.write("{no es json")
        with open(os.path.join(self.carpeta, "otro_archivo.json"), "w", encoding="utf-8") as f:
            f.write("{}")
        res = equipo.leer_equipo(self.carpeta, HOY)
        self.assertEqual(res["analistas"], [])
        self.assertEqual(res["marcas"], {})
        self.assertEqual(len(res["errores"]), 1)
        self.assertIn("marcas_9999.json", res["errores"][0])

    def test_gana_la_marca_mas_reciente(self):
        r = renovacion()
        k = row_key(r)
        self.publicar("2609", "Nancy", {k: equipo.marca(r, equipo.VALIDADA, datetime(2026, 10, 26, 9, 0))})
        self.publicar("3100", "Luis", {k: equipo.marca(r, equipo.DUDA, datetime(2026, 10, 26, 10, 0), "¿es del padre?")})
        m = equipo.leer_equipo(self.carpeta, HOY)["marcas"][k]
        self.assertEqual((m["estado"], m["gafete"], m["nota"]), (equipo.DUDA, "3100", "¿es del padre?"))

    def test_sin_senal(self):
        self.publicar("2609", "Nancy", {})
        tarde = datetime.now() + timedelta(minutes=equipo.MINUTOS_SIN_SENAL + 5)
        self.assertTrue(equipo.leer_equipo(self.carpeta, HOY, ahora=tarde)["analistas"][0]["sin_senal"])

    def test_carpeta_valida(self):
        self.assertEqual(equipo.carpeta_valida(self.carpeta), "")
        self.assertEqual(os.listdir(self.carpeta), [])  # la prueba de escritura no deja basura
        self.assertIn("No hay carpeta", equipo.carpeta_valida(""))
        self.assertIn("No se encuentra", equipo.carpeta_valida(os.path.join(self.carpeta, "no_existe")))
        self.assertIn("No se puede leer", equipo.leer_equipo(os.path.join(self.carpeta, "no"), HOY)["errores"][0])

    def test_gafete_raro_no_sale_de_la_carpeta(self):
        ruta = equipo.archivo_de(self.carpeta, "..\\..\\2609")
        self.assertEqual(os.path.dirname(ruta), self.carpeta)
        self.assertEqual(os.path.basename(ruta), "marcas_2609.json")


class TestLlaves(unittest.TestCase):
    def test_padre_corregido_tiene_las_dos_llaves(self):
        r = renovacion(Credito_Renovado_Hijo="52011", Credito_Padre_Enviado="51937")
        self.assertEqual(llaves_fila(r), [row_key(r), row_key(renovacion(Credito_Renovado_Hijo="51937"))])
        self.assertEqual(llaves_fila(renovacion()), [row_key(renovacion())])

    def test_el_validador_reconoce_lo_que_marco_un_analista_sin_base(self):
        # El analista no validó en la base: su renglón trae el padre. El validador sí: lo cambió por la renovación.
        analista = renovacion(Credito_Renovado_Hijo="51937")
        validador = renovacion(Credito_Renovado_Hijo="52011", Credito_Padre_Enviado="51937")
        marcas = {row_key(analista): equipo.marca(analista, equipo.VALIDADA, llaves=llaves_fila(analista))}
        self.assertIsNotNone(equipo.marca_de_fila(llaves_fila(validador), marcas))
        self.assertIsNone(equipo.marca_de_fila(llaves_fila(renovacion(Credito_Renovado_Hijo="1")), marcas))

    def test_al_reves_por_las_llaves_de_la_marca(self):
        validador_sin_base = renovacion(Credito_Renovado_Hijo="51937")
        analista = renovacion(Credito_Renovado_Hijo="52011", Credito_Padre_Enviado="51937")
        marcas = {row_key(analista): equipo.marca(analista, equipo.VALIDADA, llaves=llaves_fila(analista))}
        self.assertIsNotNone(equipo.marca_de_fila(llaves_fila(validador_sin_base), marcas))


class TestBitacoraYDudas(Carpeta):
    def estado(self):
        return EstadoValidaciones(os.path.join(self.carpeta, "estado.json"))

    def test_bitacora_y_duda_se_guardan(self):
        e = self.estado()
        e.registrar(["k1"], "Descartada", "es de otro día")
        e.registrar(["k1"], "Restaurada")
        e.poner_duda("k1", "¿el cliente es el correcto?")
        e2 = self.estado()
        self.assertEqual([a for _, a, _ in e2.historial("k1")], ["Descartada", "Restaurada"])
        self.assertEqual(e2.historial("k1")[0][2], "es de otro día")
        self.assertEqual(e2.duda("k1"), "¿el cliente es el correcto?")
        e2.poner_duda("k1", "")
        self.assertEqual(self.estado().duda("k1"), "")
        self.assertEqual(e2.historial("nada"), [])

    def test_bitacora_guarda_las_ultimas_30(self):
        e = self.estado()
        for i in range(35):
            e.registrar(["k"], f"acción {i}", guardar=False)
        self.assertEqual(len(e.historial("k")), 30)
        self.assertEqual(e.historial("k")[0][1], "acción 5")

    def test_purga_lo_viejo(self):
        viejo = (datetime.now() - timedelta(days=60)).isoformat(timespec="seconds")
        nuevo = datetime.now().isoformat(timespec="seconds")
        with open(os.path.join(self.carpeta, "estado.json"), "w", encoding="utf-8") as f:
            json.dump({"validados": {"a": viejo, "b": nuevo},
                       "dudas": {"a": {"hora": viejo, "nota": "x"}, "b": {"hora": nuevo, "nota": "y"}},
                       "bitacora": {"a": [[viejo, "Validada", ""]], "b": [[viejo, "X", ""], [nuevo, "Validada", ""]]}},
                      f)
        e = self.estado()
        self.assertEqual(list(e.data["validados"]), ["b"])
        self.assertEqual(list(e.data["dudas"]), ["b"])
        self.assertEqual(list(e.data["bitacora"]), ["b"])

    def test_estado_de_la_version_anterior(self):
        with open(os.path.join(self.carpeta, "estado.json"), "w", encoding="utf-8") as f:
            json.dump({"validados": {"a": datetime.now().isoformat()}}, f)
        e = self.estado()
        self.assertTrue(e.has("validados", "a"))
        self.assertEqual(e.duda("a"), "")
        self.assertEqual(e.historial("a"), [])


class TestComparacion(unittest.TestCase):
    def por_dato(self, r, en_lote=None):
        return {f["dato"]: f for f in reglas.comparacion(r, en_lote)}

    def test_todo_cuadra(self):
        r = renovacion(Dia_Lote=HOY)
        c = self.por_dato(r, en_lote=True)
        self.assertEqual((c["Cliente"]["incidente"], c["Cliente"]["base"], c["Cliente"]["ticket"]),
                         ("1184282", "1184282", "1184282"))
        self.assertEqual(c["Crédito"]["estado"], reglas.OK)
        self.assertEqual(c["Fecha"]["base"], "23/10/2026")
        self.assertEqual(c["Fecha"]["ticket"], "hoy (26/10/2026)")
        self.assertTrue(all(f["estado"] == reglas.OK for f in c.values()))

    def test_padre_corregido_avisa(self):
        r = renovacion(Credito_Renovado_Hijo="52011", Credito_Padre_Enviado="51937", SQL_Estado=sql.SQL_PADRE_CORREGIDO)
        c = self.por_dato(r, en_lote=True)["Crédito"]
        self.assertEqual((c["incidente"], c["ticket"]), ("51937", "52011"))
        self.assertIn("renovación del padre 51937", c["base"])
        self.assertEqual(c["estado"], reglas.AVISO)

    def test_no_existe_en_la_base(self):
        c = self.por_dato(renovacion(SQL_Estado=sql.SQL_NO_EXISTE), en_lote=True)
        self.assertEqual(c["Crédito"]["base"], "no existe para ese cliente")
        self.assertEqual(c["Crédito"]["estado"], reglas.MAL)

    def test_es_el_padre(self):
        c = self.por_dato(renovacion(SQL_Estado=sql.SQL_ES_PADRE, SQL_Renovacion_Hija="52011"), en_lote=True)
        self.assertEqual(c["Crédito"]["base"], "10430 es el padre de 52011")
        self.assertEqual(c["Crédito"]["estado"], reglas.MAL)

    def test_otra_fecha(self):
        c = self.por_dato(renovacion(SQL_Estado=sql.SQL_OTRA_FECHA, SQL_Fecha_Alta=date(2026, 10, 2)), en_lote=True)
        self.assertEqual(c["Fecha"]["base"], "02/10/2026 (otro día)")
        self.assertEqual(c["Fecha"]["estado"], reglas.MAL)

    def test_sin_base_es_informativo(self):
        c = self.por_dato(renovacion(SQL_Estado=None), en_lote=True)
        self.assertEqual(c["Cliente"]["base"], "sin validar")
        self.assertTrue(all(f["estado"] == reglas.INFO for f in c.values()))

    def test_falta_dato(self):
        c = self.por_dato(renovacion(Numero_Cliente="SIN CLIENTE", SQL_Estado=None), en_lote=True)
        self.assertEqual(c["Cliente"]["incidente"], "falta")
        self.assertEqual(c["Cliente"]["estado"], reglas.MAL)

    def test_ya_enviada_y_del_siguiente(self):
        c = self.por_dato(renovacion(Lote_Estado=LOTE_ENVIADA, Enviado_Fecha=date(2026, 10, 22)), en_lote=False)
        self.assertEqual(c["Fecha"]["ticket"], "salió el 22/10/2026")
        c = self.por_dato(renovacion(Lote_Estado=LOTE_SIGUIENTE), en_lote=False)
        self.assertEqual(c["Cliente"]["ticket"], "no se manda")
        self.assertEqual(c["Cliente"]["estado"], reglas.OK)

    def test_usa_va_en_lote_si_no_se_dice(self):
        r = renovacion(Lote_Estado=LOTE_HOY, Dia_Lote=HOY)
        self.assertEqual(self.por_dato(r)["Fecha"]["ticket"], "hoy (26/10/2026)")


if __name__ == "__main__":
    unittest.main()
