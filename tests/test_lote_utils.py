# -*- coding: utf-8 -*-
import os
import sys
import tempfile
import unittest
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core_renovaciones import LOTE_REZAGADA  # noqa: E402
from lote_utils import (  # noqa: E402
    EstadoValidaciones, row_key, build_lote_tsv, build_lote_html, make_cf_html, insert_into_html_body, plain_text
)


class TestPlainText(unittest.TestCase):
    def test_quita_iconos(self):
        self.assertEqual(plain_text("⚠️ MÚLTIPLE: Renovación 1 de 2"), "MÚLTIPLE: Renovación 1 de 2")
        self.assertEqual(plain_text("⏰ Rezagada 22/10 18:05 (no está)"), "Rezagada 22/10 18:05 (no está)")
        self.assertEqual(plain_text("A ➔ B"), "A → B")
        self.assertEqual(plain_text("Crédito ñ 56572"), "Crédito ñ 56572")


def fila(ticket, cli, cred, idx=1, total=1, **extra):
    r = {
        'Row_ID': f"{ticket}_{idx}", 'Numero_Ticket': ticket, 'Numero_Cliente': cli, 'Credito_Renovado_Hijo': cred,
        'Credito_Origen_Padre': '-', 'Gafete': '7353', 'Nombre_Asignado': 'SAMMY', 'Motivo': 'Sin disponible',
        'Indice_En_Ticket': idx, 'Total_En_Ticket': total, 'Fecha_Llegada_DT': datetime(2026, 10, 23, 9, 30),
    }
    r.update(extra)
    return r


class TestEstado(unittest.TestCase):
    def test_persiste_entre_instancias(self):
        with tempfile.TemporaryDirectory() as d:
            path = os.path.join(d, "estado.json")
            e = EstadoValidaciones(path)
            k = row_key(fila("I-361001", "5864479", "56572"))
            e.set('validados', [k], True)
            self.assertTrue(EstadoValidaciones(path).has('validados', k))
            e.set('validados', [k], False)
            self.assertFalse(EstadoValidaciones(path).has('validados', k))

    def test_purga_antiguos(self):
        with tempfile.TemporaryDirectory() as d:
            path = os.path.join(d, "estado.json")
            e = EstadoValidaciones(path)
            e.data['validados']['viejo'] = "2000-01-01T00:00:00"
            e.save()
            self.assertFalse(EstadoValidaciones(path).has('validados', 'viejo'))

    def test_llave_estable(self):
        a = fila("I-361001", "5864479", "56572")
        b = dict(a, Row_ID="otro_id", Gafete="2609")
        self.assertEqual(row_key(a), row_key(b))
        self.assertNotEqual(row_key(a), row_key(fila("I-361001", "5864479", "56574")))


class TestLote(unittest.TestCase):
    def test_tsv(self):
        rows = [fila("I-361001", "5864479", "56572"), fila("I-361002", "4530500", "91258", Lote_Estado=LOTE_REZAGADA)]
        tsv = build_lote_tsv(rows, {"I-361001_1"}).splitlines()
        self.assertEqual(len(tsv), 3)
        self.assertIn("5864479\t56572", tsv[1])
        self.assertTrue(tsv[1].endswith("SÍ"))
        self.assertIn("Rezagada", tsv[2])

    def test_html_escapa_y_resume(self):
        rows = [fila("I-361001", "5864479", "56572", Motivo="<b>x</b>", Lote_Estado=LOTE_REZAGADA)]
        html = build_lote_html(rows, set(), titulo="Lote")
        self.assertIn("&lt;b&gt;x&lt;/b&gt;", html)
        self.assertIn("incluye 1 rezagada", html)

    def test_cf_html_offsets(self):
        frag = "<table><tr><td>Crédito ñ</td></tr></table>"
        data = make_cf_html(frag)
        head = data.decode('utf-8').split("<html>")[0]
        vals = dict(line.split(":", 1) for line in head.strip().splitlines())
        self.assertEqual(data[int(vals['StartFragment']):int(vals['EndFragment'])].decode('utf-8'), frag)
        self.assertTrue(data[int(vals['StartHTML']):].startswith(b"<html>"))
        self.assertEqual(int(vals['EndHTML']), len(data))

    def test_insertar_antes_de_firma(self):
        out = insert_into_html_body('<html><body lang=ES><p>Firma</p></body></html>', '<table/>')
        self.assertIn('<body lang=ES><table/><p>Firma</p>', out)
        self.assertEqual(insert_into_html_body('', '<t/>'), '<t/>')


if __name__ == "__main__":
    unittest.main()
