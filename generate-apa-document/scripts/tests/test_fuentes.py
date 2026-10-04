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
        self.assertIn("metrics differ", motivos["DejaVu Serif"])
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


if __name__ == "__main__":
    unittest.main()