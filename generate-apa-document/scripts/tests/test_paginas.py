# -*- coding: utf-8 -*-
"""
Tests for paginas-de-pdf.py (the measuring half of the two-pass index build).

The script has a hyphen in its name, so it is loaded with importlib from its
real path. The PDFs are created with PyMuPDF at test time; if PyMuPDF is not
installed the whole module is skipped instead of failing.
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


def cargar():
    ruta = os.path.join(SCRIPTS, "paginas-de-pdf.py")
    spec = importlib.util.spec_from_file_location("paginas_de_pdf", ruta)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


class _Base(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if pymupdf is None:
            raise unittest.SkipTest("pymupdf is not installed")
        cls.mod = cargar()
        cls.tmp = tempfile.mkdtemp(prefix="apa7_paginas_")

    def _pdf(self, nombre, paginas, toc=None):
        """`paginas` is a list of page texts (one block per line). `toc` is a
        PyMuPDF table of contents: [[nivel, titulo, pagina_1based], ...]."""
        ruta = os.path.join(self.tmp, nombre)
        doc = pymupdf.open()
        for texto in paginas:
            page = doc.new_page(width=612, height=792)
            y = 72
            for linea in texto.splitlines():
                page.insert_text((72, y), linea, fontsize=12)
                y += 20
        if toc:
            doc.set_toc(toc)
        doc.save(ruta)
        doc.close()
        return ruta

    def _manifiesto(self, nombre, secciones):
        ruta = os.path.join(self.tmp, nombre)
        with open(ruta, "w", encoding="utf-8") as fh:
            json.dump({"secciones": secciones, "tablas": [], "figuras": []}, fh,
                      ensure_ascii=False)
        return ruta


class TestMapaDeSecciones(_Base):
    def test_las_secciones_siguen_el_orden_del_outline(self):
        pdf = self._pdf("secciones.pdf", [
            "Portada",
            "Tabla de contenido\nCapitulo Uno..........4",
            "Indice de tablas",
            "Capitulo Uno\nSeccion A",
            "Capitulo Dos",
        ], toc=[[1, "Capitulo Uno", 4], [2, "Seccion A", 4], [1, "Capitulo Dos", 5]])
        doc = pymupdf.open(pdf)
        mapa, cuerpo = self.mod.paginas_de_secciones(doc, [
            {"hid": 1, "texto": "Capitulo Uno"},
            {"hid": 2, "texto": "Seccion A"},
            {"hid": 3, "texto": "Capitulo Dos"},
        ])
        doc.close()
        self.assertEqual(mapa, {"1": 4, "2": 4, "3": 5})
        self.assertEqual(cuerpo, 4)

    def test_un_titulo_que_no_aparece_no_desplaza_a_los_demas(self):
        # A heading can be missing from the outline (e.g. an empty one). The
        # title check keeps the following sections on their real pages instead
        # of shifting them by one.
        pdf = self._pdf("hueco.pdf", [
            "Portada", "Tabla de contenido", "Indice de tablas", "Uno", "Dos",
        ], toc=[[1, "Dos", 5]])
        doc = pymupdf.open(pdf)
        mapa, _ = self.mod.paginas_de_secciones(doc, [
            {"hid": 1, "texto": "Uno"},
            {"hid": 2, "texto": "Dos"},
        ])
        doc.close()
        self.assertEqual(mapa, {"2": 5})


class TestMapaDeIndices(_Base):
    def test_detecta_las_paginas_de_indice(self):
        pdf = self._pdf("indices.pdf", [
            "Portada", "Tabla de contenido", "Indice de tablas",
            "Indice de figuras", "Cuerpo",
        ])
        doc = pymupdf.open(pdf)
        indices = self.mod.paginas_de_indices(doc)
        doc.close()
        self.assertEqual(indices, [1, 2, 3])

    def test_sin_indices_no_hay_paginas_de_indice(self):
        pdf = self._pdf("sin_indices.pdf", ["Portada", "Cuerpo"])
        doc = pymupdf.open(pdf)
        indices = self.mod.paginas_de_indices(doc)
        doc.close()
        self.assertEqual(indices, [])


class TestMapaDeEtiquetas(_Base):
    def test_primera_aparicion_por_numero(self):
        pdf = self._pdf("etiquetas.pdf", [
            "Portada",
            "Indice de tablas\nTabla 1. algo",
            "Indice de figuras",
            "Cuerpo\nTabla 1\nFigura 1",       # page index 3 -> page 4
            "Tabla 2",                          # page index 4 -> page 5
        ])
        doc = pymupdf.open(pdf)
        tablas = self.mod.paginas_de_etiquetas(doc, 3, "Tabla")
        figuras = self.mod.paginas_de_etiquetas(doc, 3, "Figura")
        doc.close()
        self.assertEqual(tablas, {"1": 4, "2": 5})
        self.assertEqual(figuras, {"1": 4})


class TestMain(_Base):
    def test_main_escribe_el_json_con_paginas_1based(self):
        pdf = self._pdf("todo.pdf", [
            "Portada",
            "Tabla de contenido\nCapitulo Uno..........4",
            "Indice de tablas\nTabla 1. algo",
            "Capitulo Uno\nTabla 1",
            "Capitulo Dos",
        ], toc=[[1, "Capitulo Uno", 4], [1, "Capitulo Dos", 5]])
        man = self._manifiesto("m.json", [
            {"hid": 1, "texto": "Capitulo Uno"},
            {"hid": 2, "texto": "Capitulo Dos"},
        ])
        salida = os.path.join(self.tmp, "paginas.json")
        codigo = self.mod.main(["--pdf", pdf, "--manifiesto", man, "--out", salida])
        self.assertEqual(codigo, 0)
        with open(salida, encoding="utf-8") as fh:
            datos = json.load(fh)
        self.assertEqual(datos["secciones"], {"1": 4, "2": 5})
        self.assertEqual(datos["tablas"], {"1": 4})

    def test_main_sin_outline_deja_las_secciones_vacias_pero_mide_tablas(self):
        # No headings at all: the outline is empty, so the body start is found
        # from the index pages instead. Tables must still be measured.
        pdf = self._pdf("sin_toc.pdf", [
            "Portada", "Tabla de contenido", "Indice de tablas", "Tabla 1",
        ])
        man = self._manifiesto("m2.json", [])
        salida = os.path.join(self.tmp, "paginas2.json")
        codigo = self.mod.main(["--pdf", pdf, "--manifiesto", man, "--out", salida])
        self.assertEqual(codigo, 0)
        with open(salida, encoding="utf-8") as fh:
            datos = json.load(fh)
        self.assertEqual(datos["secciones"], {})
        self.assertEqual(datos["tablas"], {"1": 4})


if __name__ == "__main__":
    unittest.main()
