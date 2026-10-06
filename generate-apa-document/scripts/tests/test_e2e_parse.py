# -*- coding: utf-8 -*-
"""
End-to-end test of the parse step against a synthetic document.

It runs the real `md-a-manifiesto.py` (no mocks) on
`fixtures/sintetico.md` and checks the manifest it produces. Only the standard
library and the parser are used, so this test can run in CI on any OS without
LibreOffice, Node or PyMuPDF. The fixture is a made-up document: never a real
one.
"""

import json
import os
import subprocess
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.dirname(HERE)
FIXTURE = os.path.join(HERE, "fixtures", "sintetico.md")
PARSER = os.path.join(SCRIPTS, "md-a-manifiesto.py")


class TestParseSyntheticDocument(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.mkdtemp(prefix="apa7_e2e_")
        cls.out = os.path.join(cls.tmp, "MANIFEST.json")
        env = dict(os.environ)
        env["PYTHONIOENCODING"] = "utf-8"
        cls.proc = subprocess.run(
            [sys.executable, PARSER, "--md", FIXTURE, "--out", cls.out],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            env=env,
        )
        with open(cls.out, encoding="utf-8") as fh:
            cls.manifest = json.load(fh)

    def test_the_parser_exits_zero_and_writes_the_manifest(self):
        self.assertEqual(
            self.proc.returncode,
            0,
            self.proc.stdout.decode("utf-8", "replace"),
        )
        self.assertTrue(os.path.isfile(self.out))

    def test_the_manifest_has_the_expected_shape(self):
        for key in ("secciones", "tablas", "figuras", "referencias", "bloques"):
            self.assertIn(key, self.manifest)
        self.assertEqual(len(self.manifest["secciones"]), 7)
        self.assertEqual(len(self.manifest["tablas"]), 1)
        self.assertEqual(len(self.manifest["figuras"]), 1)
        self.assertEqual(len(self.manifest["referencias"]), 2)

    def test_the_table_and_figure_keep_their_content(self):
        tabla = self.manifest["tablas"][0]
        self.assertEqual(tabla["indice"], 1)
        # The note now comes from the .md ("Nota. Elaboración propia."), with its
        # own punctuation, instead of the generic default.
        self.assertEqual(tabla["nota"], "Elaboración propia.")
        self.assertEqual(tabla["nota_origen"], "md")
        figura = self.manifest["figuras"][0]
        self.assertEqual(figura["indice"], 1)
        self.assertEqual(figura["titulo"], "Gráfico de resultados")
        self.assertEqual(figura["ruta"], "images/figura.png")
        self.assertEqual(figura["nota"], "Elaboración propia.")
        self.assertEqual(figura["nota_origen"], "md")

    def test_the_note_lines_are_consumed_not_duplicated_as_body_text(self):
        # A "Nota. ..." line right under a table/figure becomes nota_tabla /
        # nota_figura. If it leaked as an ordinary paragraph the document would
        # end up with two notes (the line AND the one built from the manifest).
        tipos = [b.get("tipo") for b in self.manifest["bloques"]]
        self.assertIn("nota_tabla", tipos)
        self.assertIn("nota_figura", tipos)
        for b in self.manifest["bloques"]:
            if b.get("tipo") == "p":
                self.assertNotIn("Elaboración propia", str(b.get("texto", "")))

    def test_urls_and_dois_are_never_touched_by_the_parser(self):
        texto = " ".join(str(b.get("texto", "")) for b in self.manifest["bloques"])
        self.assertIn("https://example.org/ruta", texto)
        self.assertIn("10.1000/xyz123", texto)

    def test_the_optional_cover_fields_are_asked_but_never_block(self):
        # vicerrectoria, logo and instructor title are asked about when missing,
        # but a missing one never blocks delivery. materia_nrc is now blocking
        # (it must be answered before delivery).
        preguntas = {q["id"]: q for q in self.manifest["diagnostico"]["preguntas"]}
        pendientes = set(self.manifest["diagnostico"]["pendientes_bloqueantes"])
        # Fields that are still optional (never block): logo, docente título, vicerrectoria.
        opcionales = {"portada_logo", "portada_docente_titulo", "portada_vicerrectoria"}
        # Fields that block if missing: materia_nrc.
        bloqueantes = {"portada_materia_nrc"}
        # Contrast: fields that DO block, so the assertion cannot pass silently
        # on a manifest that has no cover questions at all.
        for pid in opcionales:
            self.assertFalse(preguntas[pid]["bloqueante"], pid)
            self.assertNotIn(pid, pendientes)
        for pid in bloqueantes:
            self.assertTrue(preguntas[pid]["bloqueante"], pid)
            self.assertIn(pid, pendientes)
        # Both appear in campos_faltantes, but only materia_nrc is blocking.
        for campo in ("vicerrectoria", "materia_nrc"):
            self.assertIn(campo, self.manifest["portada"]["campos_faltantes"])


if __name__ == "__main__":
    unittest.main()
