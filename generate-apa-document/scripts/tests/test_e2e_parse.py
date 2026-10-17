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
import shutil
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

    def test_a_line_before_the_first_section_is_kept_as_body_text(self):
        # "Este documento es un fixture ..." sits before the first '##'. The
        # cover only claims the lines it recognizes as fields; this one is prose,
        # so it must be promoted to the body in place, never silently dropped
        # (it used to end up as a portada_dato that build-docx.js discards).
        bloques = self.manifest["bloques"]
        idx_h = next(i for i, b in enumerate(bloques) if b.get("tipo") == "h")
        idx_p = next(i for i, b in enumerate(bloques)
                     if b.get("tipo") == "p" and "fixture" in (b.get("texto") or ""))
        self.assertLess(idx_p, idx_h, "the promoted line must stay before the first section")
        portada_datos = " ".join((b.get("texto") or "") for b in bloques
                                 if b.get("tipo") == "portada_dato")
        self.assertNotIn("fixture", portada_datos)

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


class TestBlockQuote(unittest.TestCase):
    """A Markdown block quote ('>') is an APA 7 quote of more than 40 words."""

    def _parsea(self, md):
        tmp = tempfile.mkdtemp(prefix="apa7_cita_")
        self.addCleanup(shutil.rmtree, tmp, True)
        ruta_md = os.path.join(tmp, "doc.md")
        with open(ruta_md, "w", encoding="utf-8") as fh:
            fh.write(md)
        out = os.path.join(tmp, "MANIFEST.json")
        env = dict(os.environ, PYTHONIOENCODING="utf-8")
        proc = subprocess.run(
            [sys.executable, PARSER, "--md", ruta_md, "--out", out],
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT, env=env)
        self.assertEqual(proc.returncode, 0,
                         proc.stdout.decode("utf-8", "replace"))
        with open(out, encoding="utf-8") as fh:
            return json.load(fh)

    def test_a_block_quote_becomes_one_cita_block_without_the_caret(self):
        man = self._parsea(
            "# Titulo\n\n## Seccion\n\n"
            "> Primera linea de la cita textual extensa\n"
            ">\n"
            "> Segunda linea de la misma cita.\n\n"
            "Texto normal despues.\n")
        citas = [b for b in man["bloques"] if b.get("tipo") == "cita"]
        self.assertEqual(len(citas), 1, man["bloques"])
        self.assertEqual(
            citas[0]["texto"],
            "Primera linea de la cita textual extensa Segunda linea de la misma cita.")
        self.assertNotIn(">", citas[0]["texto"])
        self.assertIn("segmentos", citas[0])
        # The ordinary paragraph that follows is untouched.
        tipos = [b.get("tipo") for b in man["bloques"]]
        self.assertIn("p", tipos)

    def test_the_nbsp_of_the_guide_is_normalised_to_a_space(self):
        man = self._parsea(
            "# Titulo\n\n## Seccion\n\n"
            "> &nbsp;&nbsp;&nbsp;&nbsp;Una cita con la sangria del ejemplo.\n")
        cita = [b for b in man["bloques"] if b.get("tipo") == "cita"][0]
        self.assertEqual(cita["texto"], "Una cita con la sangria del ejemplo.")
        self.assertNotIn("\u00a0", cita["texto"])
        self.assertNotIn("&nbsp;", cita["texto"])


class TestAtribucion(unittest.TestCase):
    """The 'Elaboración propia' default of a note is an assumption, so it is asked."""

    def _parsea(self, md):
        tmp = tempfile.mkdtemp(prefix="apa7_atrib_")
        self.addCleanup(shutil.rmtree, tmp, True)
        ruta_md = os.path.join(tmp, "doc.md")
        with open(ruta_md, "w", encoding="utf-8") as fh:
            fh.write(md)
        out = os.path.join(tmp, "MANIFEST.json")
        env = dict(os.environ, PYTHONIOENCODING="utf-8")
        proc = subprocess.run(
            [sys.executable, PARSER, "--md", ruta_md, "--out", out],
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT, env=env)
        self.assertEqual(proc.returncode, 0,
                         proc.stdout.decode("utf-8", "replace"))
        with open(out, encoding="utf-8") as fh:
            return json.load(fh)

    _TABLA = "| A | B |\n|---|---|\n| 1 | 2 |\n"

    def test_a_table_without_a_note_raises_the_blocking_attribution_question(self):
        man = self._parsea("# Titulo\n\n## Seccion\n\n" + self._TABLA)
        preguntas = {q["id"]: q for q in man["diagnostico"]["preguntas"]}
        self.assertIn("atribucion_fuente", preguntas)
        self.assertTrue(preguntas["atribucion_fuente"]["bloqueante"])
        self.assertIn("atribucion_fuente",
                      man["diagnostico"]["pendientes_bloqueantes"])
        self.assertEqual(man["tablas"][0]["nota"], "Elaboración propia")
        self.assertEqual(man["tablas"][0]["nota_origen"], "default")

    def test_a_table_with_its_own_note_does_not_raise_the_question(self):
        man = self._parsea(
            "# Titulo\n\n## Seccion\n\n" + self._TABLA +
            "\nNota. De \"Censo\" por DANE, 2018. (https://x)\n")
        preguntas = {q["id"]: q for q in man["diagnostico"]["preguntas"]}
        self.assertNotIn("atribucion_fuente", preguntas)
        self.assertEqual(man["tablas"][0]["nota_origen"], "md")


class TestTipografia(unittest.TestCase):
    """The guide allows one typeface; the parser records it in the manifest."""

    def _parsea(self, extra):
        tmp = tempfile.mkdtemp(prefix="apa7_font_")
        self.addCleanup(shutil.rmtree, tmp, True)
        ruta_md = os.path.join(tmp, "doc.md")
        with open(ruta_md, "w", encoding="utf-8") as fh:
            fh.write("# Titulo\n\n## Seccion\n\nTexto.\n")
        out = os.path.join(tmp, "MANIFEST.json")
        env = dict(os.environ, PYTHONIOENCODING="utf-8")
        proc = subprocess.run(
            [sys.executable, PARSER, "--md", ruta_md, "--out", out] + extra,
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT, env=env)
        man = None
        if os.path.isfile(out):
            with open(out, encoding="utf-8") as fh:
                man = json.load(fh)
        return proc.returncode, man

    def test_the_default_typeface_is_times_new_roman_12(self):
        code, man = self._parsea([])
        self.assertEqual(code, 0)
        self.assertEqual(man["tipografia"], {"nombre": "Times New Roman", "pt": 12})

    def test_arial_brings_its_guide_size(self):
        code, man = self._parsea(["--fuente", "Arial"])
        self.assertEqual(code, 0)
        self.assertEqual(man["tipografia"], {"nombre": "Arial", "pt": 11})

    def test_computer_modern_brings_10_pt(self):
        code, man = self._parsea(["--fuente", "Computer Modern"])
        self.assertEqual(code, 0)
        self.assertEqual(man["tipografia"], {"nombre": "Computer Modern", "pt": 10})

    def test_the_size_can_be_overridden(self):
        code, man = self._parsea(["--fuente", "Arial", "--tamano-fuente", "12"])
        self.assertEqual(code, 0)
        self.assertEqual(man["tipografia"], {"nombre": "Arial", "pt": 12})

    def test_an_unknown_font_is_rejected(self):
        code, _ = self._parsea(["--fuente", "Comic Sans MS"])
        self.assertEqual(code, 2)


class TestTitulosMayusculas(unittest.TestCase):
    """APA 7: a heading drops its number and is never written in ALL CAPS."""

    def _parsea(self, md):
        tmp = tempfile.mkdtemp(prefix="apa7_titulo_")
        self.addCleanup(shutil.rmtree, tmp, True)
        ruta_md = os.path.join(tmp, "doc.md")
        with open(ruta_md, "w", encoding="utf-8") as fh:
            fh.write(md)
        out = os.path.join(tmp, "MANIFEST.json")
        env = dict(os.environ, PYTHONIOENCODING="utf-8")
        proc = subprocess.run(
            [sys.executable, PARSER, "--md", ruta_md, "--out", out],
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT, env=env)
        self.assertEqual(proc.returncode, 0,
                         proc.stdout.decode("utf-8", "replace"))
        with open(out, encoding="utf-8") as fh:
            return json.load(fh)

    def _textos(self, man):
        return [s["texto"] for s in man["secciones"]]

    def test_an_all_caps_heading_is_rewritten_to_sentence_case(self):
        man = self._parsea(
            "# Mi trabajo\n\n## INTRODUCCIÓN GENERAL\n\nTexto.\n")
        self.assertIn("Introducción general", self._textos(man))

    def test_the_numbered_heading_drops_its_number(self):
        man = self._parsea(
            "# Mi trabajo\n\n## 1. METODOLOGÍA DEL ESTUDIO\n\nTexto.\n")
        secs = man["secciones"]
        self.assertEqual(secs[0]["texto"], "Metodología del estudio")
        self.assertEqual(secs[0]["etiqueta"], "1")
        self.assertEqual(secs[0]["nivel"], 1)

    def test_acronyms_are_kept_but_spanish_words_are_lowercased(self):
        man = self._parsea(
            "# Mi trabajo\n\n## ANÁLISIS DE OWASP SAMM Y DEL SDL\n\nTexto.\n")
        self.assertIn("Análisis de OWASP SAMM y del SDL", self._textos(man))

    def test_a_word_written_in_lowercase_in_the_prose_is_not_uppercased(self):
        man = self._parsea(
            "# Mi trabajo\n\n## MARCO TEÓRICO\n\nEl marco teórico guía.\n")
        self.assertIn("Marco teórico", self._textos(man))

    def test_a_heading_already_in_sentence_case_is_untouched(self):
        man = self._parsea(
            "# Mi trabajo\n\n## Marco Teórico\n\nTexto.\n")
        self.assertIn("Marco Teórico", self._textos(man))


if __name__ == "__main__":
    unittest.main()
