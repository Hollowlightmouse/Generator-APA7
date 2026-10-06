# -*- coding: utf-8 -*-
"""
Tests for verificar-pdf.py (STEP 7).

The verifier script has a hyphen in its name, so it cannot be imported with a
plain `import`. It is loaded with importlib from its real path.

The PDFs are created with PyMuPDF at test time; if PyMuPDF is not installed the
whole module is skipped instead of failing.
"""

import importlib.util
import json
import os
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.dirname(HERE)
sys.path.insert(0, SCRIPTS)

try:
    import pymupdf
except ImportError:  # pymupdf < 1.24.3 published the same module as 'fitz'
    try:
        import fitz as pymupdf
    except ImportError:
        pymupdf = None


def cargar_verificador():
    ruta = os.path.join(SCRIPTS, "verificar-pdf.py")
    spec = importlib.util.spec_from_file_location("verificar_pdf", ruta)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def manifiesto(n_tablas=1, nota="Elaboracion propia"):
    return {
        "portada": {},
        "secciones": [],
        "tablas": [{"indice": i + 1, "titulo": "T", "filas": [["1", "2"]], "nota": nota}
                   for i in range(n_tablas)],
        "figuras": [],
        "referencias": [],
        "opciones": {},
        "diagnostico": {},
    }


class _Base(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if pymupdf is None:
            raise unittest.SkipTest("pymupdf is not installed")
        cls.mod = cargar_verificador()
        cls.tmp = tempfile.mkdtemp(prefix="apa7_verif_")

    def _escribe(self, nombre, contenido):
        ruta = os.path.join(self.tmp, nombre)
        with open(ruta, "w", encoding="utf-8") as fh:
            json.dump(contenido, fh, ensure_ascii=False)
        return ruta

    def _verifica(self, pdf, man):
        class Args:
            pass
        a = Args()
        a.pdf = pdf
        a.manifiesto = man
        R, _ = self.mod.verificar(a)
        return {it["nombre"]: it for it in R.items}

    def _pdf_con_lineas(self, nombre, lineas):
        ruta = os.path.join(self.tmp, nombre)
        doc = pymupdf.open()
        page = doc.new_page(width=612, height=792)
        y = 72
        for linea in lineas:
            page.insert_text((72, y), linea, fontsize=12)
            y += 20
        doc.save(ruta)
        doc.close()
        return ruta

    def _pdf_un_bloque(self, nombre, texto):
        """Inserts the whole text in a single call, so PyMuPDF groups it into
        one text block (the case that produced the false FAIL)."""
        ruta = os.path.join(self.tmp, nombre)
        doc = pymupdf.open()
        page = doc.new_page(width=612, height=792)
        page.insert_text((72, 72), texto, fontsize=12)
        doc.save(ruta)
        doc.close()
        return ruta


class TestNotaDeTabla(_Base):
    CLAVE = "Tables carry a note below (in the PDF, not only in the manifest)"

    def test_nota_debajo_pasa(self):
        man = self._escribe("m1.json", manifiesto())
        pdf = self._pdf_con_lineas("nota_ok.pdf",
                                   ["Tabla 1", "1 2 3 4 5 6", "Nota. Elaboracion propia"])
        self.assertTrue(self._verifica(pdf, man)[self.CLAVE]["ok"])

    def test_nota_ausente_falla(self):
        man = self._escribe("m2.json", manifiesto())
        pdf = self._pdf_con_lineas("sin_nota.pdf", ["Tabla 1", "1 2 3 4 5 6"])
        res = self._verifica(pdf, man)[self.CLAVE]
        self.assertFalse(res["ok"])
        self.assertIn("no note", res["detalle"])

    def test_bloque_fusionado_pasa(self):
        # Regression: block extraction merges the cells and the note into one
        # block ('Tabla 1\n1 2 3 4 5 6\nNota. Elaboracion propia'). The old
        # check required the block to START with "Nota." and failed; the
        # word-level check must pass.
        man = self._escribe("m3.json", manifiesto())
        pdf = self._pdf_un_bloque("fusionado.pdf",
                                  "Tabla 1\n1 2 3 4 5 6\nNota. Elaboracion propia")
        self.assertTrue(self._verifica(pdf, man)[self.CLAVE]["ok"])

    def test_tabla_que_ocupa_varias_paginas_pasa(self):
        # Regression: a table wider than a page ends on the next one and its
        # note follows there. Looking only at the label's page reported a false
        # critical FAIL for every such table.
        man = self._escribe("m4.json", manifiesto())
        ruta = os.path.join(self.tmp, "multipagina.pdf")
        doc = pymupdf.open()
        p1 = doc.new_page(width=612, height=792)
        p1.insert_text((72, 72), "Tabla 1", fontsize=12)
        p1.insert_text((72, 100), "1 2 3 4 5 6", fontsize=12)
        p2 = doc.new_page(width=612, height=792)
        p2.insert_text((72, 72), "6 5 4 3 2 1", fontsize=12)
        p2.insert_text((72, 100), "Nota. Elaboracion propia", fontsize=12)
        doc.save(ruta)
        doc.close()
        self.assertTrue(self._verifica(ruta, man)[self.CLAVE]["ok"])

    def test_nota_en_la_pagina_siguiente_sobre_la_siguiente_tabla_pasa(self):
        # The note lands at the top of the page where the NEXT caption starts,
        # so the search window ends at that caption instead of at the page end.
        man = self._escribe("m5.json", manifiesto(2))
        ruta = os.path.join(self.tmp, "nota_encima_de_la_siguiente.pdf")
        doc = pymupdf.open()
        p1 = doc.new_page(width=612, height=792)
        p1.insert_text((72, 72), "Tabla 1", fontsize=12)
        p1.insert_text((72, 100), "1 2 3 4 5 6", fontsize=12)
        p2 = doc.new_page(width=612, height=792)
        p2.insert_text((72, 72), "Nota. Elaboracion propia", fontsize=12)
        p2.insert_text((72, 140), "Tabla 2", fontsize=12)
        p2.insert_text((72, 168), "1 2 3 4 5 6", fontsize=12)
        p2.insert_text((72, 196), "Nota. Elaboracion propia", fontsize=12)
        doc.save(ruta)
        doc.close()
        self.assertTrue(self._verifica(ruta, man)[self.CLAVE]["ok"])

    def test_mencion_en_texto_no_cuenta_como_rotulo(self):
        # Regression: "Tabla 1" inside a sentence is not a caption. Taking it
        # for one moved the search window and hid the real note.
        man = self._escribe("m6.json", manifiesto())
        pdf = self._pdf_con_lineas("mencion.pdf",
                                   ["Como se ve en la Tabla 1 el riesgo aumenta",
                                    "Tabla 1",
                                    "1 2 3 4 5 6",
                                    "Nota. Elaboracion propia"])
        self.assertTrue(self._verifica(pdf, man)[self.CLAVE]["ok"])

    def test_nota_de_otra_tabla_no_cuenta(self):
        # The note of the next table must not vouch for the previous one.
        man = self._escribe("m7.json", manifiesto(2))
        ruta = os.path.join(self.tmp, "nota_cruzada.pdf")
        doc = pymupdf.open()
        doc.new_page(width=612, height=792)
        p2 = doc.new_page(width=612, height=792)
        p1 = doc[0]
        p1.insert_text((72, 72), "Tabla 1", fontsize=12)
        p1.insert_text((72, 100), "1 2 3 4 5 6", fontsize=12)
        p2.insert_text((72, 72), "Tabla 2", fontsize=12)
        p2.insert_text((72, 100), "1 2 3 4 5 6", fontsize=12)
        p2.insert_text((72, 128), "Nota. Elaboracion propia", fontsize=12)
        doc.save(ruta)
        doc.close()
        res = self._verifica(ruta, man)[self.CLAVE]
        self.assertFalse(res["ok"])
        self.assertIn("Tabla 1", res["detalle"])


class TestNumeracion(_Base):
    CLAVE = "Every page after the cover shows its number in the top right corner"

    def _pdf(self, nombre, numeros):
        """numeros: dict {page_index: number} to print in the top strip."""
        ruta = os.path.join(self.tmp, nombre)
        doc = pymupdf.open()
        for _ in range(3):
            doc.new_page(width=612, height=792)
        for i, n in numeros.items():
            doc[i].insert_text((560, 40), str(n), fontsize=11)
        for i in range(3):
            doc[i].insert_text((72, 100), "Contenido de la pagina", fontsize=12)
        doc.save(ruta)
        doc.close()
        return ruta

    def test_numeracion_correcta_pasa(self):
        man = self._escribe("mn1.json", manifiesto(0))
        pdf = self._pdf("num_ok.pdf", {1: 2, 2: 3})
        self.assertTrue(self._verifica(pdf, man)[self.CLAVE]["ok"])

    def test_numero_equivocado_falla(self):
        man = self._escribe("mn2.json", manifiesto(0))
        pdf = self._pdf("num_mal.pdf", {1: 2, 2: 9})
        res = self._verifica(pdf, man)[self.CLAVE]
        self.assertFalse(res["ok"])
        self.assertIn("page 3", res["detalle"])


class TestNotaDeFigura(_Base):
    """The figure's note goes BELOW the image (APA 7), like the table's.

    The check is geometric, so these PDFs are built with a real image and the
    note above or below it. The manifest says the figure does not exist on disk:
    this suite is about the PDF's geometry, and check 15 (sizing against the
    manifest) has nothing to say about a picture the builder never inserted.
    """

    CLAVE = "Figure notes go below the image"

    def _manifiesto(self, nombre):
        man = manifiesto(0)
        man["figuras"] = [{"indice": 1, "titulo": "Grafico de resultados",
                           "ruta": "images/figura.png", "nota": "Elaboracion propia",
                           "existe": False}]
        return self._escribe(nombre, man)

    def _pdf(self, nombre, nota_antes):
        """Two pages: the cover, and the body with caption, image (y 300-400)
        and the note above or below it."""
        ruta = os.path.join(self.tmp, nombre)
        doc = pymupdf.open()
        for _ in range(2):
            doc.new_page(width=612, height=792)
        # Page 0 is the cover, so the body starts on page 1 (CUERPO).
        cuerpo = doc[1]
        cuerpo.insert_text((72, 200), "Figura 1", fontsize=12)
        cuerpo.insert_text((72, 220), "Grafico de resultados", fontsize=12)
        pix = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, 20, 20), False)
        pix.set_rect(pix.irect, (200, 30, 30))
        cuerpo.insert_image(pymupdf.Rect(150, 300, 450, 400), pixmap=pix)
        y_nota = 250 if nota_antes else 430
        cuerpo.insert_text((72, y_nota), "Nota. Elaboracion propia", fontsize=12)
        doc.save(ruta)
        doc.close()
        return ruta

    def test_nota_debajo_pasa(self):
        man = self._manifiesto("mf1.json")
        pdf = self._pdf("fig_ok.pdf", nota_antes=False)
        res = self._verifica(pdf, man)[self.CLAVE]
        self.assertTrue(res["ok"], res["detalle"])
        self.assertIn("below", res["detalle"])

    def test_nota_encima_falla(self):
        man = self._manifiesto("mf2.json")
        pdf = self._pdf("fig_arriba.pdf", nota_antes=True)
        res = self._verifica(pdf, man)[self.CLAVE]
        self.assertFalse(res["ok"])
        self.assertIn("no note below", res["detalle"])

    def test_una_nota_que_cayo_a_la_pagina_siguiente_pasa_con_resguardo(self):
        # The image fills the page and the note is pushed over. That is not an
        # error, but it is also not a clean pass, and the detail has to name it.
        man = self._manifiesto("mf3.json")
        ruta = os.path.join(self.tmp, "fig_siguiente.pdf")
        doc = pymupdf.open()
        for _ in range(3):
            doc.new_page(width=612, height=792)
        pix = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, 20, 20), False)
        pix.set_rect(pix.irect, (200, 30, 30))
        doc[1].insert_image(pymupdf.Rect(150, 300, 450, 400), pixmap=pix)
        doc[2].insert_text((72, 90), "Nota. Elaboracion propia", fontsize=12)
        doc.save(ruta)
        doc.close()
        res = self._verifica(ruta, man)[self.CLAVE]
        self.assertTrue(res["ok"], res["detalle"])
        self.assertIn("fell to the next page", res["detalle"])

    def test_un_documento_sin_imagenes_no_se_evalua(self):
        man = self._manifiesto("mf4.json")
        pdf = self._pdf_un_bloque("fig_vacio.pdf", "Cuerpo sin figuras")
        res = self._verifica(pdf, man)[self.CLAVE]
        self.assertTrue(res["ok"])
        self.assertFalse(res["critico"])
        self.assertIn("not applicable", res["detalle"])


class TestPortadaOpcional(_Base):
    """vicerrectoria/materia_nrc absence is a WARNING, never a failure."""

    def _man(self, portada, nombre):
        man = manifiesto(0)
        man["portada"] = portada
        return self._escribe(nombre, man)

    def test_missing_optional_fields_warn_without_blocking(self):
        man = self._man({}, "port1.json")
        pdf = self._pdf_con_lineas("portada_vacia.pdf", ["Titulo del trabajo"])
        res = self._verifica(pdf, man)
        for clave in ("The vice-rector's office line was asked about",
                      "The cover page logo was asked about"):
            self.assertIn(clave, res, clave)
            self.assertFalse(res[clave]["critico"], clave)
            self.assertFalse(res[clave]["ok"], clave)

    def test_present_fields_do_not_emit_the_asked_about_warning(self):
        man = self._man({"materia_nrc": "NRC 91827",
                         "vicerrectoria": "Vicerrectoría Académica"}, "port2.json")
        pdf = self._pdf_con_lineas("portada_con_datos.pdf",
                                   ["Titulo del trabajo",
                                    "Vicerrectoría Académica",
                                    "NRC 91827"])
        res = self._verifica(pdf, man)
        self.assertNotIn("The course/NRC line was asked about", res)
        self.assertNotIn("The vice-rector's office line was asked about", res)


class TestTipografia(_Base):
    """Only the declared typeface, and only where there is something to read.

    Microsoft Word stamps its own default (Aptos) on the blank runs that
    separate the entries of a TOC field it regenerates on export. Those glyphs
    are spaces: invisible, and rejecting them failed every Word-rendered PDF for
    something nobody can see. A font that does draw text is still judged.

    The fonts are embedded from the system files, because PyMuPDF's built-in
    base-14 fonts report names ("Helvetica", "Times-Roman") that the check does
    not accept, which would test the wrong thing entirely.
    """

    CLAVE = "Only the declared typeface is used"
    TIMES = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", "times.ttf")
    ARIAL = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", "arial.ttf")

    def setUp(self):
        if not os.path.exists(self.TIMES):
            self.skipTest("times.ttf is not on this machine")
        if not os.path.exists(self.ARIAL):
            self.skipTest("arial.ttf is not on this machine")

    def _pdf(self, nombre, texto, fontfile):
        """Body in Times New Roman, plus one run in `fontfile` on page 1."""
        ruta = os.path.join(self.tmp, nombre)
        doc = pymupdf.open()
        for _ in range(2):
            doc.new_page(width=612, height=792)
        cuerpo = doc[1]
        cuerpo.insert_font(fontname="tr", fontfile=self.TIMES)
        cuerpo.insert_text((72, 100), "Cuerpo del documento", fontsize=12,
                           fontname="tr")
        if texto is not None:
            nombre_fuente = "ajena"
            cuerpo.insert_font(fontname=nombre_fuente, fontfile=fontfile)
            cuerpo.insert_text((72, 300), texto, fontsize=12,
                               fontname=nombre_fuente)
        doc.save(ruta)
        doc.close()
        return ruta

    def test_una_fuente_ajena_que_solo_espacios_no_falla(self):
        man = self._escribe("mt1.json", manifiesto(0))
        pdf = self._pdf("tipo_espacios.pdf", " ", self.ARIAL)
        res = self._verifica(pdf, man)[self.CLAVE]
        self.assertTrue(res["ok"], res["detalle"])
        self.assertIn("no visible text", res["detalle"])
        self.assertIn("Arial", res["detalle"])

    def test_una_fuente_ajena_que_escribe_texto_si_falla(self):
        man = self._escribe("mt2.json", manifiesto(0))
        pdf = self._pdf("tipo_texto.pdf", "Texto que se lee", self.ARIAL)
        res = self._verifica(pdf, man)[self.CLAVE]
        self.assertFalse(res["ok"])
        self.assertIn("NOT accepted", res["detalle"])
        self.assertIn("Arial", res["detalle"])

    def test_la_fuente_declarada_siempre_pasa(self):
        man = self._escribe("mt3.json", manifiesto(0))
        pdf = self._pdf("tipo_ok.pdf", "Times New Roman", self.TIMES)
        res = self._verifica(pdf, man)[self.CLAVE]
        self.assertTrue(res["ok"], res["detalle"])
        self.assertNotIn("NOT accepted", res["detalle"])


if __name__ == "__main__":
    unittest.main()
