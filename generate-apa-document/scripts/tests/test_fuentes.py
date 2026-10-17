"""Tests for the font acceptance rule.

The failure this prevents is the reason the rule exists: Times New Roman does
not ship with Linux or macOS, LibreOffice substitutes a metrically compatible
font, and the check used to call that a failure. So a correct document built
outside Windows could never pass.

The other half matters just as much: the substitution must not become a blanket
"accept anything". A font with different metrics still fails, because the line
advance measured in the PDF is then not comparable with the Times New Roman
figures the check assumes.
"""

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from lib import fuentes  # noqa: E402


class TestNormaliza(unittest.TestCase):
    def test_the_subset_prefix_is_removed(self):
        self.assertEqual(fuentes.normaliza("BAAAAA+TimesNewRomanPSMT"),
                         "timesnewromanpsmt")

    def test_separators_and_case_are_ignored(self):
        self.assertEqual(fuentes.normaliza("Liberation Serif"), "liberationserif")
        self.assertEqual(fuentes.normaliza("LiberationSerif"), "liberationserif")
        # The style stays in the normalised name, which is why matching is by
        # prefix and not equality.
        self.assertEqual(fuentes.normaliza("LiberationSerif-Bold"),
                         "liberationserifbold")

    def test_an_empty_name_does_not_explode(self):
        self.assertEqual(fuentes.normaliza(None), "")
        self.assertEqual(fuentes.normaliza(""), "")


class TestClasifica(unittest.TestCase):
    def test_times_new_roman_itself_is_accepted_and_not_a_substitute(self):
        ok, sustituto, _ = fuentes.clasifica("BAAAAA+TimesNewRomanPS-BoldMT")
        self.assertTrue(ok)
        self.assertFalse(sustituto)

    def test_the_metric_clones_are_accepted_as_substitutes(self):
        for name in ("CAAAAA+LiberationSerif", "NimbusRoman-Regular",
                     "TeXGyreTermes-Regular", "FreeSerif", "Tinos-Regular"):
            ok, sustituto, _ = fuentes.clasifica(name)
            self.assertTrue(ok, name)
            self.assertTrue(sustituto, "%s should be flagged as a substitute" % name)

    def test_the_style_suffix_does_not_hide_the_family(self):
        ok, _, _ = fuentes.clasifica("AAAAAA+LiberationSerif-Bold")
        self.assertTrue(ok)

    def test_a_family_with_different_metrics_is_refused(self):
        # DejaVu Serif is wider than Times New Roman: accepting it would make
        # the line-advance check meaningless.
        ok, _, legible = fuentes.clasifica("DejaVuSerif")
        self.assertFalse(ok)
        self.assertEqual(legible, "DejaVu Serif")

    def test_arial_is_refused_even_with_the_mt_suffix(self):
        # Regression: "ArialMT" used to be listed as a metric clone, so a
        # document in Arial (a different family) passed the typography check.
        ok, _, legible = fuentes.clasifica("ArialMT")
        self.assertFalse(ok)
        self.assertEqual(legible, "Arial")

    def test_an_unknown_font_is_refused(self):
        ok, _, _ = fuentes.clasifica("ComicPapyrusStd")
        self.assertFalse(ok)

    def test_a_name_merely_containing_a_family_is_not_accepted(self):
        # "NotTimesNewRomanLookalike" must not slip through a substring test.
        ok, _, _ = fuentes.clasifica("NotTimesNewRomanLookalike")
        self.assertFalse(ok)


class TestParticiona(unittest.TestCase):
    def test_it_separates_the_three_groups(self):
        aceptadas, sustituciones, rechazadas = fuentes.particiona([
            "BAAAAA+TimesNewRomanPSMT",
            "CAAAAA+LiberationSerif",
            "ComicPapyrusStd",
        ])
        self.assertEqual(len(aceptadas), 1)
        self.assertEqual(len(sustituciones), 1)
        self.assertEqual(len(rechazadas), 1)

    def test_the_rejection_reason_distinguishes_known_from_unknown(self):
        _, _, rechazadas = fuentes.particiona(["DejaVuSerif", "ComicPapyrusStd"])
        motivos = dict((nombre, motivo) for nombre, motivo in rechazadas)
        self.assertIn("another family", motivos["DejaVu Serif"])
        self.assertIn("not a declared", motivos["ComicPapyrusStd"])

    def test_nothing_is_lost_or_counted_twice(self):
        nombres = ["BAAAAA+TimesNewRomanPSMT", "CAAAAA+LiberationSerif", "XyZ"]
        aceptadas, sustituciones, rechazadas = fuentes.particiona(nombres)
        total = len(aceptadas) + len(sustituciones) + len(rechazadas)
        self.assertEqual(total, len(nombres))

    def test_an_empty_document_rejects_nothing(self):
        self.assertEqual(fuentes.particiona([]), ([], [], []))

    def test_only_a_substitution_leaves_the_check_passing(self):
        aceptadas, sustituciones, rechazadas = fuentes.particiona(
            ["AAAAAA+LiberationSerif"])
        self.assertFalse(rechazadas, "a metric substitute must not fail the check")
        self.assertEqual(len(aceptadas) + len(sustituciones), 1)


class TestFamilias(unittest.TestCase):
    def test_the_guide_families_are_all_known(self):
        for nombre in ("Times New Roman", "Arial", "Calibri",
                       "Lucida Sans Unicode", "Georgia", "Computer Modern"):
            self.assertTrue(fuentes.es_familia_valida(nombre), nombre)

    def test_the_guide_sizes_are_the_ones_the_guide_fixes(self):
        esperado = {
            "Times New Roman": 12, "Arial": 11, "Calibri": 11,
            "Lucida Sans Unicode": 10, "Georgia": 11, "Computer Modern": 10,
        }
        for nombre, pt in esperado.items():
            self.assertEqual(fuentes.tamano_por_defecto(nombre), pt, nombre)

    def test_an_unknown_family_is_not_valid(self):
        self.assertFalse(fuentes.es_familia_valida("Comic Sans MS"))
        self.assertIsNone(fuentes.tamano_por_defecto("Comic Sans MS"))


class TestClasificaPorFamilia(unittest.TestCase):
    def test_arial_accepts_arial_and_its_substitutes(self):
        for nombre in ("ArialMT", "AAAAAA+LiberationSans", "Arimo-Regular"):
            ok, _, _ = fuentes.clasifica(nombre, familia="Arial")
            self.assertTrue(ok, nombre)

    def test_arial_marks_the_metric_clone_as_a_substitute_but_not_arial(self):
        _, es_sub, _ = fuentes.clasifica("ArialMT", familia="Arial")
        self.assertFalse(es_sub)
        _, es_sub, _ = fuentes.clasifica("LiberationSans", familia="Arial")
        self.assertTrue(es_sub)

    def test_calibri_accepts_carlito(self):
        ok, es_sub, _ = fuentes.clasifica("Carlito-Regular", familia="Calibri")
        self.assertTrue(ok)
        self.assertTrue(es_sub)

    def test_georgia_accepts_gelasio(self):
        ok, _, _ = fuentes.clasifica("Gelasio", familia="Georgia")
        self.assertTrue(ok)

    def test_computer_modern_accepts_latin_modern(self):
        for nombre in ("LatinModernRoman", "LMRoman10-Regular", "CMUSerif-Roman"):
            ok, _, _ = fuentes.clasifica(nombre, familia="Computer Modern")
            self.assertTrue(ok, nombre)

    def test_a_serif_is_refused_when_the_family_is_sans(self):
        # Liberation Serif is Times-metric, not Arial-metric: it must not pass
        # a document declared as Arial.
        ok, _, _ = fuentes.clasifica("LiberationSerif", familia="Arial")
        self.assertFalse(ok)


class TestMetricas(unittest.TestCase):
    def test_times_new_roman_keeps_its_historical_threshold(self):
        # 1.15 x 12 = 13.8 pt single; the old hardcoded double threshold was 20.0.
        self.assertAlmostEqual(fuentes.avance_simple_pt("Times New Roman"), 13.8,
                               places=1)
        self.assertAlmostEqual(fuentes.umbral_doble("Times New Roman"), 20.0,
                               places=1)

    def test_a_different_family_moves_the_threshold(self):
        # Arial 11 single is ~12.65 pt, so its double threshold is lower than the
        # Times one; keeping the Times number would false-fail the check.
        self.assertLess(fuentes.umbral_doble("Arial"), fuentes.umbral_doble("Times New Roman"))
        self.assertGreater(fuentes.avance_simple_pt("Calibri"),
                           fuentes.avance_simple_pt("Arial"))

    def test_an_unknown_family_falls_back_to_times_metrics(self):
        self.assertAlmostEqual(fuentes.avance_simple_pt("Comic Sans MS"),
                               fuentes.avance_simple_pt("Times New Roman"), places=3)


if __name__ == "__main__":
    unittest.main()