"""Prueba real de build-docx.js con node.

El resto de la suite no ejecuta node: usa dobles para todo lo que cuesta. Aqui
si hace falta, porque las dos cosas que se comprueban solo existen dentro del
.js: el orden de los parrafos de una figura (imagen y nota) y la etiqueta
"Nota." que el script escribe delante del texto de la nota. Una prueba con
dobles pasaria aunque el .js estuviera mal.

Se salta cuando node no esta en PATH o cuando docx no esta instalado en el
workdir real, de modo que en un equipo sin la instalacion la suite sigue verde
sin mentir.
"""

import importlib.util
import json
import os
import shutil
import struct
import subprocess
import sys
import tempfile
import unittest
import zipfile
import zlib
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parents[1]
BUILD = SCRIPTS / "build-docx.js"
WORKDIR = SCRIPTS.parent / ".work"


def _cargar_apa7():
    spec = importlib.util.spec_from_file_location(
        "apa7", SCRIPTS / "apa7.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def _png(ruta, ancho, alto, color=(70, 110, 160)):
    """A real PNG, so dimensionesImagen() reads a header like the real thing."""
    def trozo(tipo, datos):
        cuerpo = tipo + datos
        return (struct.pack(">I", len(datos)) + cuerpo
                + struct.pack(">I", zlib.crc32(cuerpo) & 0xFFFFFFFF))

    crudo = b"".join(b"\x00" + bytes(color) * ancho for _ in range(alto))
    with open(ruta, "wb") as fh:
        fh.write(b"\x89PNG\r\n\x1a\n")
        fh.write(trozo(b"IHDR", struct.pack(">IIBBBBB", ancho, alto, 8, 2, 0, 0, 0)))
        fh.write(trozo(b"IDAT", zlib.compress(crudo)))
        fh.write(trozo(b"IEND", b""))


class _BuildReal(unittest.TestCase):
    maxDiff = None

    @classmethod
    def setUpClass(cls):
        cls.node = shutil.which("node")
        if not cls.node:
            raise unittest.SkipTest("node is not on PATH")
        if not (WORKDIR / "node_modules" / "docx" / "package.json").exists():
            raise unittest.SkipTest("docx is not installed in %s" % WORKDIR)

    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.tmp, True)

    def _construye(self, manifiesto):
        ruta_man = self.tmp / "MANIFEST.json"
        ruta_man.write_text(json.dumps(manifiesto), encoding="utf-8")
        salida = self.tmp / "salida.docx"
        entorno = dict(os.environ, APA7_WORKDIR=str(WORKDIR))
        proc = subprocess.run(
            [self.node, str(BUILD), "--manifiesto", str(ruta_man),
             "--out", str(salida)],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
            env=entorno, timeout=180)
        self.assertEqual(proc.returncode, 0,
                         "build-docx.js failed:\n%s\n%s" % (proc.stdout, proc.stderr))
        with zipfile.ZipFile(salida) as z:
            return z.read("word/document.xml").decode("utf-8")

    def _manifiesto(self, nota_tabla, nota_figura):
        img = self.tmp / "grafico.png"
        _png(img, 240, 160)
        return {
            "version": 1,
            "generado": "2026-10-05T00:00:00",
            "fuente": "prueba automatica",
            "opciones": {"indice_tablas": False, "indice_figuras": False},
            "portada": {"titulo": "Nota de figura", "autor": "Prueba",
                        "institucion": "Prueba", "fecha": "2026"},
            "docling": None,
            "secciones": [],
            "bloques": [
                {"tipo": "tabla", "indice": 0},
                {"tipo": "p", "texto": "separador",
                 "marcador": None,
                 "segmentos": [{"text": "separador", "bold": False, "italics": False}]},
                {"tipo": "figura", "indice": 0},
            ],
            "tablas": [{
                "indice": 1, "titulo": "Tabla de prueba",
                "filas": [["Columna A", "Columna B"], ["1", "2"]],
                "nota": nota_tabla, "origen": "prueba",
            }],
            "figuras": [{
                "indice": 1, "titulo": "Figura de prueba",
                "nota": nota_figura, "ruta": img.name,
                "ruta_absoluta": str(img), "ancho_in": 3.0, "alto_in": 2.0,
                "existe": True, "origen": "prueba",
            }],
            "referencias": [],
            "diagnostico": {"preguntas": [], "pendientes_bloqueantes": []},
        }

    @staticmethod
    def _parrafos(xml):
        """Every <w:p> as its concatenated text, in document order."""
        import re
        salida = []
        for p in re.findall(r"<w:p[ >].*?</w:p>", xml, re.S):
            salida.append("".join(re.findall(r"<w:t[^>]*>(.*?)</w:t>", p, re.S)))
        return salida

    # -- the label ---------------------------------------------------------

    def test_la_etiqueta_nota_no_se_repite_en_la_tabla(self):
        xml = self._construye(self._manifiesto(
            nota_tabla="Nota. Fuente de la tabla.", nota_figura=None))
        notas = [t for t in self._parrafos(xml) if "Fuente de la tabla" in t]
        self.assertEqual(len(notas), 1, notas)
        self.assertEqual(notas[0].count("Nota."), 1, notas[0])
        self.assertTrue(notas[0].startswith("Nota."), notas[0])

    def test_la_etiqueta_nota_no_se_repite_en_la_figura(self):
        xml = self._construye(self._manifiesto(
            nota_tabla=None, nota_figura="Nota. Fuente de la figura."))
        notas = [t for t in self._parrafos(xml) if "Fuente de la figura" in t]
        self.assertEqual(len(notas), 1, notas)
        self.assertEqual(notas[0].count("Nota."), 1, notas[0])

    def test_una_nota_sin_etiqueta_no_gana_una(self):
        xml = self._construye(self._manifiesto(
            nota_tabla="Elaboracion propia", nota_figura="Note. From the author."))
        parrafos = self._parrafos(xml)
        tabla = [t for t in parrafos if "Elaboracion propia" in t][0]
        figura = [t for t in parrafos if "From the author" in t][0]
        self.assertEqual(tabla, "Nota. Elaboracion propia")
        self.assertEqual(figura, "Nota. From the author.")

    def test_una_nota_que_solo_menciona_la_palabra_no_se_toca(self):
        xml = self._construye(self._manifiesto(
            nota_tabla="Segun la nota del autor, el valor es 3.", nota_figura=None))
        notas = [t for t in self._parrafos(xml) if "nota del autor" in t]
        self.assertEqual(len(notas), 1, notas)
        self.assertTrue(notas[0].startswith("Nota. Segun la nota del autor"),
                        notas[0])

    # -- the order ---------------------------------------------------------

    def test_la_nota_de_figura_viene_despues_de_la_imagen(self):
        xml = self._construye(self._manifiesto(
            nota_tabla=None, nota_figura="Fuente de la figura."))
        import re
        # The drawing element is inside the image's paragraph.
        p_imagen = next(p for p in re.findall(r"<w:p[ >].*?</w:p>", xml, re.S)
                        if "<w:drawing>" in p)
        p_nota = next(p for p in re.findall(r"<w:p[ >].*?</w:p>", xml, re.S)
                      if "Fuente de la figura" in p)
        self.assertLess(xml.index(p_imagen), xml.index(p_nota),
                        "the note paragraph comes BEFORE the image")
        # keepNext belongs to the image, so the pair moves to the next page
        # together instead of the note stranding the figure on the previous one.
        self.assertIn("<w:keepNext/>", p_imagen)
        self.assertNotIn("<w:keepNext/>", p_nota)

    # -- the caption legend (real SEQ, all bold) ---------------------------
    # T4b: the legend ("Tabla 1.", "Figura 2.") is a real SEQ field, which is
    # what "TOC \c" collects, and the WHOLE legend is bold. CampoSecuencia
    # exists because SimpleField wrote its cached number with a DEFAULT run,
    # so the number always came out plain while only the word was bold.

    def _caption_de(self, xml, etiqueta):
        import re
        return next(p for p in re.findall(r"<w:p[ >].*?</w:p>", xml, re.S)
                    if "<w:fldSimple" in p and etiqueta in p)

    def _cada_run_visible_es_negrita(self, parrafo):
        import re
        vacios = []
        for run in re.findall(r"<w:r>.*?</w:r>", parrafo, re.S):
            if "<w:t" in run and "<w:b/>" not in run:
                vacios.append(run)
        return vacios

    def test_la_leyenda_de_la_tabla_es_una_secuencia_real_toda_en_negrita(self):
        import re
        xml = self._construye(self._manifiesto(
            nota_tabla="Fuente.", nota_figura="Fuente."))
        p = self._caption_de(xml, "Tabla ")
        self.assertIn('w:instr=" SEQ Tabla \\* ARABIC "', p)
        self.assertEqual(self._cada_run_visible_es_negrita(p), [], p)
        # The legend reads the cached number, not a blank field.
        self.assertEqual("".join(re.findall(r"<w:t[^>]*>(.*?)</w:t>", p, re.S)),
                         "Tabla 1", p)

    def test_la_leyenda_de_la_figura_es_una_secuencia_real_toda_en_negrita(self):
        import re
        xml = self._construye(self._manifiesto(
            nota_tabla="Fuente.", nota_figura="Fuente."))
        p = self._caption_de(xml, "Figura ")
        self.assertIn('w:instr=" SEQ Figura \\* ARABIC "', p)
        self.assertEqual(self._cada_run_visible_es_negrita(p), [], p)
        self.assertEqual("".join(re.findall(r"<w:t[^>]*>(.*?)</w:t>", p, re.S)),
                         "Figura 1", p)

    # -- the TOC entries are not bold --------------------------------------
    # T4c: the cached entries of the indexes come out in the TOC style, never
    # bold. (The index TITLES are headings and are bold; the entries are not.)

    def test_las_entradas_del_toc_no_son_negritas(self):
        import re
        man = self._manifiesto(nota_tabla="Fuente.", nota_figura="Fuente.")
        man["opciones"]["indice_tablas"] = True
        man["opciones"]["indice_figuras"] = True
        man["secciones"] = [{"texto": "Introduccion", "hid": "introduccion",
                             "nivel": 1}]
        xml = self._construye(man)
        toc = [p for p in re.findall(r"<w:p[ >].*?</w:p>", xml, re.S)
               if "<w:instrText" in p and "TOC " in p and "<w:hyperlink" in p]
        self.assertTrue(len(toc) >= 2, len(toc))  # content + tables + figures
        for p in toc:
            self.assertIn('w:pStyle w:val="TOC1"', p, p)
            self.assertNotIn("<w:b/>", p, p)
            self.assertIn('w:rStyle w:val="IndexLink"', p, p)
        # Every index entry points at its caption's bookmark, proving the
        # cached entry list was really built (not an empty placeholder field).
        self.assertIn('w:anchor="apa_tbl_1"', xml)
        self.assertIn('w:anchor="apa_fig_1"', xml)
        self.assertIn('w:anchor="apa_sec_', xml)


if __name__ == "__main__":
    unittest.main()
