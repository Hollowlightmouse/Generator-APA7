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

    def _componentes(self, manifiesto):
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
            return (z.read("word/document.xml").decode("utf-8"),
                    z.read("word/styles.xml").decode("utf-8"))

    def _construye(self, manifiesto):
        return self._componentes(manifiesto)[0]

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

    # -- the APA block quote (>40 words) -----------------------------------
    # APA 7: a textual quote of more than 40 words is set in its own paragraph,
    # with a 1.27 cm (0.5 in) left indent, no first-line indent, no quotation
    # marks and double spacing.

    def test_la_cita_en_bloque_va_indentada_y_sin_comillas(self):
        import re
        man = self._manifiesto(nota_tabla=None, nota_figura=None)
        texto = ("Las universidades hoy por hoy se preocupan por desarrollar y "
                 "mejorar su actividad investigativa por varias razones.")
        man["bloques"] = [
            {"tipo": "cita", "texto": texto,
             "segmentos": [{"text": texto, "bold": False, "italics": False}]},
        ]
        xml = self._construye(man)
        p = next(p for p in re.findall(r"<w:p[ >].*?</w:p>", xml, re.S)
                 if "actividad investigativa" in p)
        self.assertIn('w:left="720"', p, p)
        self.assertNotIn("<w:firstLine", p, p)
        self.assertIn('w:line="480"', p, p)
        # No quotation marks are added around the quote.
        parrafos = self._parrafos(xml)
        cita = [t for t in parrafos if "actividad investigativa" in t][0]
        self.assertEqual(cita, texto)

    # -- the typeface comes from the manifest ------------------------------

    def test_sin_tipografia_el_documento_usa_times_new_roman_12(self):
        xml = self._construye(self._manifiesto(
            nota_tabla="Fuente.", nota_figura=None))
        self.assertIn('w:ascii="Times New Roman"', xml)
        self.assertIn('w:val="24"', xml)  # 12 pt = 24 half-points

    def test_la_tipografia_del_manifiesto_se_aplica_en_todo_el_documento(self):
        man = self._manifiesto(nota_tabla="Fuente.", nota_figura="Fuente.")
        man["tipografia"] = {"nombre": "Arial", "pt": 11}
        xml = self._construye(man)
        self.assertIn('w:ascii="Arial"', xml)
        self.assertIn('w:hAnsi="Arial"', xml)
        self.assertIn('w:val="22"', xml)  # 11 pt = 22 half-points
        self.assertNotIn('w:ascii="Times New Roman"', xml)

    # -- the table content is 2 pt smaller than the body (APA 7) -----------
    # The guide's table anatomy: element "Contenido" may drop 2 points, while
    # the number, the title and the note keep the body size.

    def _tamano_del_run(self, xml, texto):
        import re
        for run in re.findall(r"<w:r>.*?</w:r>", xml, re.S):
            if re.search(r"<w:t[^>]*>" + re.escape(texto) + r"</w:t>", run):
                m = re.search(r'<w:sz w:val="(\d+)"', run)
                if m:
                    return int(m.group(1))
        return None

    def test_el_contenido_de_la_tabla_va_dos_puntos_menor_que_el_cuerpo(self):
        xml = self._construye(self._manifiesto(
            nota_tabla="Fuente de la tabla.", nota_figura=None))
        # Body is Times New Roman 12 pt (24 half-points); the table content
        # drops 2 pt to 10 pt (20 half-points), header row included.
        self.assertEqual(self._tamano_del_run(xml, "Columna A"), 20)
        self.assertEqual(self._tamano_del_run(xml, "Columna B"), 20)
        # The caption number and the note keep the body size (24).
        self.assertEqual(self._tamano_del_run(xml, "Fuente de la tabla."), 24)

    def test_el_contenido_de_la_tabla_baja_dos_puntos_tambien_con_otra_fuente(self):
        man = self._manifiesto(nota_tabla=None, nota_figura=None)
        man["tipografia"] = {"nombre": "Arial", "pt": 11}
        xml = self._construye(man)
        # Arial 11 pt (22 half-points) -> table content 9 pt (18 half-points).
        self.assertEqual(self._tamano_del_run(xml, "Columna A"), 18)

    # -- the headings are REAL APA styles (Heading1..Heading5) --------------
    # APA 7: a heading is a real Heading 1..5 style, not bold text on a plain
    # paragraph. The library's built-in styles are blue and oversized, so each
    # level has to be redefined black and in the body font; otherwise Word's
    # navigation pane and a `TOC \o` field see a flat, blue document.

    def _manifiesto_con_titulos(self):
        man = self._manifiesto(nota_tabla=None, nota_figura=None)
        man["secciones"] = [
            {"texto": "Introduccion", "hid": 1, "nivel": 1},
            {"texto": "Detalle", "hid": 2, "nivel": 3},
        ]
        man["bloques"] = [
            {"tipo": "h", "hid": 1, "nivel": 1, "texto": "Introduccion",
             "segmentos": [{"text": "Introduccion", "bold": False, "italics": False}]},
            {"tipo": "p", "texto": "cuerpo", "marcador": None,
             "segmentos": [{"text": "cuerpo", "bold": False, "italics": False}]},
            {"tipo": "h", "hid": 2, "nivel": 3, "texto": "Detalle",
             "segmentos": [{"text": "Detalle", "bold": False, "italics": False}]},
        ]
        return man

    def test_un_titulo_lleva_el_estilo_real_Heading_n(self):
        import re
        xml, _ = self._componentes(self._manifiesto_con_titulos())
        p1 = next(p for p in re.findall(r"<w:p[ >].*?</w:p>", xml, re.S)
                  if "Introduccion" in p and "<w:bookmarkStart" in p)
        self.assertIn('w:pStyle w:val="Heading1"', p1, p1)
        p3 = next(p for p in re.findall(r"<w:p[ >].*?</w:p>", xml, re.S)
                  if "Detalle" in p and "<w:bookmarkStart" in p)
        self.assertIn('w:pStyle w:val="Heading3"', p3, p3)

    def test_los_estilos_de_titulo_son_negros_y_no_los_azules_de_la_libreria(self):
        import re
        _, estilos = self._componentes(self._manifiesto_con_titulos())
        for nivel in range(1, 6):
            m = re.search(r'<w:style [^>]*w:styleId="Heading%d">.*?</w:style>' % nivel,
                          estilos, re.S)
            self.assertIsNotNone(m, "Heading%d is not defined" % nivel)
            bloque = m.group(0)
            self.assertIn('<w:color w:val="000000"/>', bloque)
            # The default heading colours must not survive the override.
            self.assertNotIn("2E74B5", bloque)
            self.assertNotIn("1F4D78", bloque)
            # Every level carries the outline level a `TOC \o` field collects.
            self.assertIn('<w:outlineLvl w:val="%d"/>' % (nivel - 1), bloque)
        h3 = re.search(r'<w:style [^>]*w:styleId="Heading3">.*?</w:style>',
                       estilos, re.S).group(0)
        self.assertIn("<w:b/>", h3)
        self.assertIn("<w:i/>", h3)


if __name__ == "__main__":
    unittest.main()
