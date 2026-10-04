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
        self.assertEqual(tabla["nota"], "Elaboración propia")
        figura = self.manifest["figuras"][0]
        self.assertEqual(figura["indice"], 1)
        self.assertEqual(figura["titulo"], "Gráfico de resultados")
        self.assertEqual(figura["ruta"], "images/figura.png")

    def test_urls_and_dois_are_never_touched_by_the_parser(self):
        texto = " ".join(str(b.get("texto", "")) for b in self.manifest["bloques"])
        self.assertIn("https://example.org/ruta", texto)
        self.assertIn("10.1000/xyz123", texto)


if __name__ == "__main__":
    unittest.main()
