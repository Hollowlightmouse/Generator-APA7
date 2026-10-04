# -*- coding: utf-8 -*-
"""
Tests for deglue() in md-a-manifiesto.py (STEP 6).

The module name has a hyphen, so it is loaded with importlib from its real
path instead of with a plain import.
"""

import importlib.util
import os
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.dirname(HERE)
sys.path.insert(0, SCRIPTS)


def cargar_modulo():
    ruta = os.path.join(SCRIPTS, "md-a-manifiesto.py")
    spec = importlib.util.spec_from_file_location("md_a_manifiesto", ruta)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


class TestDeglue(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.mod = cargar_modulo()

    def setUp(self):
        self.mod.TERMINOS_PROTEGIDOS = []

    def tearDown(self):
        self.mod.TERMINOS_PROTEGIDOS = []

    def _d(self, texto):
        return self.mod.deglue(texto)

    def test_correccion_clara(self):
        out, cambios = self._d("ACTIVIDAD4")
        self.assertEqual(out, "ACTIVIDAD 4")
        self.assertTrue(cambios)

    def test_no_protegido_rompe_camelcase(self):
        out, cambios = self._d("JavaScript")
        self.assertEqual(out, "Java Script")
        self.assertTrue(cambios[0][3], "camelCase must be flagged as ambiguous")

    def test_no_protegido_rompe_github(self):
        out, _ = self._d("GitHub")
        self.assertEqual(out, "Git Hub")

    def test_no_protegido_rompe_sha256(self):
        out, cambios = self._d("SHA256")
        self.assertEqual(out, "SHA 256")
        self.assertTrue(cambios[0][3], "mixed alphanumeric must be flagged")

    def test_url_intacta(self):
        url = "https://example.com/NodeJS?x=1"
        out, cambios = self._d("Ver %s para detalles" % url)
        self.assertIn(url, out)
        self.assertFalse(cambios)

    def test_doi_intacto(self):
        doi = "10.1038/s41586-021-03819-2"
        out, cambios = self._d("doi: %s" % doi)
        self.assertIn(doi, out)
        self.assertFalse(cambios)

    def test_email_intacto(self):
        correo = "juan.perez@example.com"
        out, _ = self._d("Escribir a %s" % correo)
        self.assertIn(correo, out)

    def test_nombre_archivo_intacto(self):
        archivo = "informe_final_v2.docx"
        out, _ = self._d("Abrir %s" % archivo)
        self.assertIn(archivo, out)

    def test_termino_protegido_sobrevive(self):
        self.mod.TERMINOS_PROTEGIDOS = ["JavaScript", "GitHub", "SHA256"]
        texto = "JavaScript GitHub SHA256 y ACTIVIDAD4"
        out, cambios = self._d(texto)
        self.assertIn("JavaScript", out)
        self.assertIn("GitHub", out)
        self.assertIn("SHA256", out)
        self.assertEqual(out, "JavaScript GitHub SHA256 y ACTIVIDAD 4")
        for _, antes, _, _ in cambios:
            self.assertNotIn(antes, ("JavaScript", "GitHub", "SHA256"))

    def test_termino_protegido_case_sensitive(self):
        self.mod.TERMINOS_PROTEGIDOS = ["NodeJS"]
        out, _ = self._d("nodejs y NodeJS")
        self.assertEqual(out, "nodejs y NodeJS")

    def test_cambios_son_cuatro_tuplas(self):
        _, cambios = self._d("ACTIVIDAD4")
        self.assertEqual(len(cambios[0]), 4)
        self.assertIsInstance(cambios[0][3], bool)


if __name__ == "__main__":
    unittest.main()
