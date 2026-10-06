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


if __name__ == "__main__":
    unittest.main()
