"""Tests for the PDF engine choice: lib/motores.py and lib/word.py.

Standard library only, and nothing here ever launches Word: every export is a
fake `_correr`. The only test that talks to real Word is a manual one, because
`check` on a machine that has Word must be provable by hand and never by a
surprise in the suite.

Run them with:
    python -m unittest discover -s scripts/tests -v
"""
import base64
import contextlib
import io
import os
import sys
import tempfile
import unittest
import zipfile
from contextlib import contextmanager
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import lib.rutas as rutas  # noqa: E402
from lib import motores as motores_mod  # noqa: E402
from lib import word as word_mod  # noqa: E402


class PlatformMixin:
    """Pins the platform so one test machine can test all three."""

    @contextmanager
    def as_platform(self, system):
        with mock.patch.object(rutas, "is_windows", return_value=system == "Windows"), \
                mock.patch.object(rutas, "is_macos", return_value=system == "Darwin"), \
                mock.patch.object(rutas, "is_linux", return_value=system == "Linux"):
            yield

    def as_windows(self):
        return self.as_platform("Windows")

    def as_macos(self):
        return self.as_platform("Darwin")

    def as_linux(self):
        return self.as_platform("Linux")


def _motor(nombre, ok=True, version="", **extra):
    return motores_mod.Motor(nombre=nombre, ok=ok, version=version, **extra)


class TestNormaliza(unittest.TestCase):
    """The spelling is forgiving; the MEANING is not.

    A typo must be an error, never a quiet switch to a different renderer than
    the one that was asked for.
    """

    def test_canonical_values(self):
        self.assertEqual(motores_mod.normaliza("auto"), "auto")
        self.assertEqual(motores_mod.normaliza("word"), "word")
        self.assertEqual(motores_mod.normaliza("libreoffice"), "libreoffice")

    def test_aliases_and_case_and_dashes(self):
        for valor in ("Word", "WORD", " winword ", "Microsoft Word", "office", "--word"):
            self.assertEqual(motores_mod.normaliza(valor), "word", valor)
        for valor in ("LibreOffice", "LO", "soffice", "headless", "_libreoffice_"):
            self.assertEqual(motores_mod.normaliza(valor), "libreoffice", valor)

    def test_empty_is_auto(self):
        self.assertEqual(motores_mod.normaliza(""), "auto")
        self.assertEqual(motores_mod.normaliza(None), "auto")

    def test_an_unknown_value_is_an_error_listing_the_valid_ones(self):
        with self.assertRaises(ValueError) as ctx:
            motores_mod.normaliza("wordp")
        self.assertIn("unknown PDF engine", str(ctx.exception))
        for nombre in ("auto", "word", "libreoffice"):
            self.assertIn(nombre, str(ctx.exception))

    def test_preference_order_is_word_then_libreoffice(self):
        self.assertEqual(motores_mod.MOTORES, ("word", "libreoffice"))


class TestElegir(unittest.TestCase):
    def test_auto_prefers_word_when_it_is_viable(self):
        word = _motor(motores_mod.WORD, version="16.0")
        with mock.patch.object(motores_mod, "estado_word", return_value=word), \
                mock.patch.object(motores_mod, "estado_libreoffice") as libreoffice:
            elegido, estados = motores_mod.elegir("auto")
        self.assertIs(elegido, word)
        self.assertEqual(estados, [word])
        # Word answered it; LibreOffice was never launched or even looked up.
        libreoffice.assert_not_called()

    def test_auto_falls_back_to_libreoffice_when_word_cannot_convert(self):
        word = _motor(motores_mod.WORD, ok=False, motivo="not installed")
        libre = _motor(motores_mod.LIBREOFFICE, version="7.0")
        with mock.patch.object(motores_mod, "estado_word", return_value=word), \
                mock.patch.object(motores_mod, "estado_libreoffice", return_value=libre):
            elegido, estados = motores_mod.elegir("auto")
        self.assertIs(elegido, libre)
        self.assertEqual(estados, [word, libre])

    def test_auto_returns_none_and_explains_when_nothing_can_convert(self):
        word = _motor(motores_mod.WORD, ok=False, motivo="not installed")
        libre = _motor(motores_mod.LIBREOFFICE, ok=False, motivo="not found")
        with mock.patch.object(motores_mod, "estado_word", return_value=word), \
                mock.patch.object(motores_mod, "estado_libreoffice", return_value=libre):
            elegido, estados = motores_mod.elegir("auto")
        self.assertIsNone(elegido)
        self.assertEqual([e.nombre for e in estados], ["word", "libreoffice"])
        for estado in estados:
            self.assertTrue(estado.motivo)

    def test_a_forced_engine_that_cannot_convert_is_never_replaced(self):
        # The whole reason `elegir` takes a preference: `--motor word` must not
        # quietly hand back a PDF rendered by LibreOffice.
        with mock.patch.object(motores_mod, "estado_word",
                               return_value=_motor(motores_mod.WORD, ok=False, motivo="no")) as w, \
                mock.patch.object(motores_mod, "estado_libreoffice") as libre:
            elegido, estados = motores_mod.elegir("word")
        self.assertIsNone(elegido)
        self.assertEqual(estados[0].nombre, "word")
        libre.assert_not_called()

    def test_a_forced_libreoffice_never_asks_about_word(self):
        libre = _motor(motores_mod.LIBREOFFICE, version="7.0")
        with mock.patch.object(motores_mod, "estado_libreoffice", return_value=libre), \
                mock.patch.object(motores_mod, "estado_word") as w:
            elegido, _ = motores_mod.elegir("lo")
        self.assertIs(elegido, libre)
        w.assert_not_called()

    def test_probar_word_false_reaches_the_word_lookup(self):
        with mock.patch.object(motores_mod, "estado_word",
                               return_value=_motor(motores_mod.WORD, version="16.0")) as w:
            motores_mod.elegir("auto", probar_word=False)
        self.assertFalse(w.call_args.kwargs["probar"])

    def test_reevaluar_reaches_the_word_lookup(self):
        with mock.patch.object(motores_mod, "estado_word",
                               return_value=_motor(motores_mod.WORD, version="16.0")) as w:
            motores_mod.elegir("word", reevaluar=True)
        self.assertTrue(w.call_args.kwargs["reevaluar"])

    def test_an_unknown_engine_is_rejected_before_anything_is_launched(self):
        with mock.patch.object(motores_mod, "estado_word") as w, \
                mock.patch.object(motores_mod, "estado_libreoffice") as libre:
            with self.assertRaises(ValueError):
                motores_mod.elegir("openoffice")
        w.assert_not_called()
        libre.assert_not_called()


class TestMotor(unittest.TestCase):
    def test_the_label_names_the_program(self):
        self.assertEqual(_motor(motores_mod.WORD, version="16.0").etiqueta(), "Microsoft Word 16.0")
        self.assertEqual(_motor(motores_mod.LIBREOFFICE, version="7.4").etiqueta(),
                         "headless LibreOffice 7.4")

    def test_a_probed_engine_says_so_and_an_unchecked_one_says_not_verified(self):
        probed = motores_mod.estado_word.__wrapped__ if False else None
        del probed
        with mock.patch.object(word_mod, "disponible",
                               return_value=word_mod.Estado(ok=True, version="16.0",
                                                            ruta=r"C:\P\WINWORD.EXE",
                                                            sondeado=True)):
            detalle = motores_mod.estado_word().detalle
        self.assertIn("probe: minimal .docx -> PDF", detalle)
        self.assertIn("16.0", detalle)

        with mock.patch.object(word_mod, "disponible",
                               return_value=word_mod.Estado(ok=True, ruta=r"C:\P\WINWORD.EXE",
                                                            sondeado=False)):
            detalle = motores_mod.estado_word(probar=False).detalle
        self.assertIn("not verified", detalle)

    def test_a_failed_word_keeps_its_reason_as_detail_and_motivo(self):
        with mock.patch.object(word_mod, "disponible",
                               return_value=word_mod.Estado(ok=False, motivo="not installed. Searched: x")):
            motor = motores_mod.estado_word()
        self.assertFalse(motor.ok)
        self.assertEqual(motor.motivo, motor.detalle)
        self.assertIn("not installed", motor.motivo)


class TestEstadoLibreOffice(unittest.TestCase):
    def _probed(self, first_line, exit_code=0):
        stdout = (first_line + "\n") if first_line else ""
        return mock.Mock(stdout=stdout, exit_code=exit_code, first_line=first_line)

    def test_missing_soffice_lists_where_it_looked(self):
        with mock.patch.object(rutas, "soffice_path", return_value=None), \
                mock.patch.object(rutas, "soffice_candidates",
                                  return_value=["/usr/bin/soffice", "/snap/bin/libreoffice"]):
            motor = motores_mod.estado_libreoffice()
        self.assertFalse(motor.ok)
        self.assertIn("/usr/bin/soffice", motor.motivo)
        self.assertIn("/snap/bin/libreoffice", motor.motivo)

    def test_the_version_comes_from_the_launcher_through_run_soffice(self):
        # soffice.exe detaches: asking the launcher directly hangs the preflight.
        with mock.patch.object(rutas, "soffice_path", return_value="/usr/bin/soffice"), \
                mock.patch.object(rutas, "run_soffice",
                                  return_value=self._probed("LibreOffice 7.4.7.2 2023-01-01")) as run_soffice:
            motor = motores_mod.estado_libreoffice()
        self.assertTrue(motor.ok)
        self.assertEqual(motor.version, "7.4.7.2 2023-01-01")
        run_soffice.assert_called_once_with(["--version"], timeout=60)

    def test_a_launcher_that_fails_is_not_an_engine(self):
        with mock.patch.object(rutas, "soffice_path", return_value="/usr/bin/soffice"), \
                mock.patch.object(rutas, "run_soffice", return_value=self._probed("", exit_code=1)):
            motor = motores_mod.estado_libreoffice()
        self.assertFalse(motor.ok)
        self.assertIn("exit code 1", motor.motivo)

    def test_a_launcher_that_cannot_be_started_is_reported_with_its_reason(self):
        with mock.patch.object(rutas, "soffice_path", return_value="/usr/bin/soffice"), \
                mock.patch.object(rutas, "run_soffice",
                                  side_effect=rutas.SofficeNotFound("soffice is not executable")):
            motor = motores_mod.estado_libreoffice()
        self.assertFalse(motor.ok)
        self.assertIn("not executable", motor.motivo)


class TestPlataforma(PlatformMixin, unittest.TestCase):
    def test_linux_has_no_word_driver(self):
        with self.as_linux():
            self.assertIn("cannot be automated", word_mod.plataforma_soportada())

    def test_windows_needs_powershell(self):
        with self.as_windows(), mock.patch.object(word_mod, "powershell_bin", return_value=None):
            self.assertIn("PowerShell", word_mod.plataforma_soportada())
        with self.as_windows(), mock.patch.object(word_mod, "powershell_bin", return_value="powershell.exe"):
            self.assertIsNone(word_mod.plataforma_soportada())

    def test_macos_needs_osascript(self):
        with self.as_macos(), mock.patch.object(word_mod.shutil, "which", return_value=None):
            self.assertIn("osascript", word_mod.plataforma_soportada())
        with self.as_macos(), mock.patch.object(word_mod.shutil, "which", return_value="/usr/bin/osascript"):
            self.assertIsNone(word_mod.plataforma_soportada())

    def test_powershell_is_preferred_over_pwsh_because_it_is_what_windows_ships(self):
        with mock.patch.object(word_mod.shutil, "which",
                               side_effect=lambda name: {"powershell": r"C:\P\powershell.exe"}.get(name)):
            self.assertEqual(word_mod.powershell_bin(), r"C:\P\powershell.exe")
        with mock.patch.object(word_mod.shutil, "which",
                               side_effect=lambda name: {"pwsh": "/usr/bin/pwsh"}.get(name)):
            self.assertEqual(word_mod.powershell_bin(), "/usr/bin/pwsh")


class TestDocxMinimo(unittest.TestCase):
    """The probe's document is built here, so no .docx fixture can rot."""

    def test_it_is_a_valid_package_with_the_three_required_parts(self):
        with tempfile.TemporaryDirectory() as tmp:
            destino = word_mod.docx_minimo(Path(tmp) / "sub" / "sondeo.docx")
            self.assertTrue(destino.is_file())
            with zipfile.ZipFile(str(destino)) as paquete:
                self.assertIsNone(paquete.testzip())
                nombres = paquete.namelist()
                self.assertEqual(nombres[0], "[Content_Types].xml")
                for parte in ("_rels/.rels", "word/document.xml"):
                    self.assertIn(parte, nombres)
                cuerpo = paquete.read("word/document.xml").decode("utf-8")
        # The text is what proves the document really opened.
        self.assertIn("APA7", cuerpo)
        self.assertIn("wordprocessingml", cuerpo)


class TestScriptPowerShell(unittest.TestCase):
    def setUp(self):
        self.script = word_mod.script_powershell(r"C:\in\a.docx", r"C:\out\b.pdf")

    def test_close_and_quit_pass_the_zero_by_reference(self):
        # Without [ref] the exception is swallowed by the surrounding catch and
        # WINWORD.EXE survives the export, which then blocks every later run.
        self.assertEqual(self.script.count("[ref]$cero"), 2)
        self.assertIn("Close([ref]$cero)", self.script)
        self.assertIn("Quit([ref]$cero)", self.script)

    def test_it_is_invisible_and_never_writes_back(self):
        self.assertIn("$palabra.Visible = $false", self.script)
        self.assertIn("$palabra.DisplayAlerts = 0", self.script)
        self.assertIn("$cero = 0", self.script)

    def test_the_document_is_opened_read_only(self):
        # Documents.Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, ...)
        self.assertIn("Documents.Open('C:\\in\\a.docx', $false, $true, $false,", self.script)

    def test_the_pdf_format_constant_is_explicit(self):
        self.assertIn("ExportAsFixedFormat('C:\\out\\b.pdf', 17)", self.script)

    def test_the_sentinels_both_appear(self):
        self.assertIn("APA7-VERSION:", self.script)
        self.assertIn("APA7-PDF-OK", self.script)
        self.assertIn("APA7-PDF-ERROR", self.script)

    def test_fields_are_updated_only_when_asked(self):
        self.assertNotIn("Fields.Update", self.script)
        con_campos = word_mod.script_powershell(r"C:\in\a.docx", r"C:\out\b.pdf", True)
        self.assertIn("Fields.Update()", con_campos)

    def test_a_quote_in_a_path_cannot_break_out_of_the_literal(self):
        script = word_mod.script_powershell("C:\\in\\it's.docx", "C:\\out\\b.pdf")
        self.assertIn("it''s.docx", script)
        self.assertNotIn("'C:\\in\\it's.docx'", script)


class TestScriptAppleScript(unittest.TestCase):
    """macOS: built and asserted here, verified on real hardware by hand."""

    def setUp(self):
        self.script = word_mod.script_applescript("/tmp/in/a.docx", "/tmp/out/b.pdf")

    def test_the_same_sentinels_as_powershell(self):
        self.assertIn("APA7-VERSION:", self.script)
        self.assertIn("APA7-PDF-OK", self.script)
        self.assertIn("APA7-PDF-ERROR", self.script)

    def test_it_reads_and_closes_without_saving(self):
        self.assertIn("read only", self.script)
        self.assertIn("close elDocumento saving no", self.script)
        self.assertIn("quit saving no", self.script)

    def test_an_error_is_reported_not_swallowed(self):
        self.assertIn("on error", self.script)
        self.assertIn("APA7-PDF-ERROR: ", self.script)


class TestComando(PlatformMixin, unittest.TestCase):
    def test_windows_ships_the_script_encoded_so_no_path_is_ever_quoted_into_it(self):
        with self.as_windows(), \
                mock.patch.object(word_mod, "powershell_bin", return_value="powershell.exe"):
            argv = word_mod.comando(r"C:\in\it's a.docx", r"C:\out\b.pdf")
        self.assertEqual(argv[:4], ["powershell.exe", "-NoProfile", "-NonInteractive",
                                    "-EncodedCommand"])
        script = base64.b64decode(argv[4]).decode("utf-16-le")
        self.assertEqual(script, word_mod.script_powershell(r"C:\in\it's a.docx", r"C:\out\b.pdf"))

    def test_macos_uses_osascript(self):
        with self.as_macos(), \
                mock.patch.object(word_mod.shutil, "which",
                                  side_effect=lambda name: "/usr/bin/osascript" if name == "osascript" else None):
            argv = word_mod.comando("/tmp/a.docx", "/tmp/b.pdf")
        self.assertEqual(argv[0], "/usr/bin/osascript")
        self.assertIn("-e", argv)

    def test_a_platform_without_a_driver_gets_none(self):
        with self.as_linux():
            self.assertIsNone(word_mod.comando("/tmp/a.docx", "/tmp/b.pdf"))


class TestInterpretar(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.pdf = Path(self.tmp.name) / "out.pdf"

    def _escribir_pdf(self):
        self.pdf.write_bytes(b"%PDF-1.7\n%\xe2\xe3\xcf\xd3\n")

    def test_success_needs_the_sentinel_and_a_real_pdf(self):
        self._escribir_pdf()
        ok, version, motivo = word_mod._interpretar(
            "APA7-VERSION:16.0\r\nAPA7-PDF-OK\r\n", 0, self.pdf, False)
        self.assertTrue(ok)
        self.assertEqual(version, "16.0")
        self.assertEqual(motivo, "")

    def test_the_error_sentinel_wins_over_everything(self):
        ok, _, motivo = word_mod._interpretar(
            "APA7-PDF-ERROR: the file is locked", 1, self.pdf, True)
        self.assertFalse(ok)
        self.assertEqual(motivo, "the file is locked")

    def test_a_timeout_says_timeout(self):
        ok, _, motivo = word_mod._interpretar("", 124, self.pdf, True)
        self.assertFalse(ok)
        self.assertIn("within the timeout", motivo)

    def test_a_claimed_success_without_a_file_is_not_a_success(self):
        ok, _, motivo = word_mod._interpretar("APA7-PDF-OK", 0, self.pdf, False)
        self.assertFalse(ok)
        self.assertIn("no readable PDF", motivo)

    def test_silence_is_not_a_success(self):
        ok, _, motivo = word_mod._interpretar("", 0, self.pdf, False)
        self.assertFalse(ok)
        self.assertIn("no PDF", motivo)

    def test_the_version_is_read_through_the_carriage_returns(self):
        # Word prints \r\n; a parser that splits on \n alone keeps a stray \r and
        # reports an empty version, which is what the log showed before.
        ok, version, _ = word_mod._interpretar(
            "APA7-VERSION:16.0\r\nAPA7-PDF-OK\r\n", 0, self._escribir_pdf() or self.pdf, False)
        self.assertTrue(ok)
        self.assertEqual(version, "16.0")


class TestExportar(PlatformMixin, unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        self.docx = word_mod.docx_minimo(self.root / "in.docx")
        self.pdf = self.root / "out.pdf"
        self.winword = mock.patch.object(rutas, "word_path", return_value=r"C:\P\WINWORD.EXE")
        self.winword.start()
        self.addCleanup(self.winword.stop)
        # No sleeping in the leftover poll.
        self.reloj = mock.patch.object(word_mod, "time")
        self.reloj.start()
        self.addCleanup(self.reloj.stop)

    def _driver(self, stdout, exit_code=0, timed_out=False, escribir_pdf=True):
        """A fake `_correr` for the argv `comando` builds as ["fake", docx, destino]."""

        def _correr(argv, timeout, out_file, err_file):
            if escribir_pdf:
                Path(argv[2]).write_bytes(b"%PDF-1.7\ntrailing\n")
            return exit_code, stdout, "", timed_out, None

        return mock.patch.object(word_mod, "_correr", side_effect=_correr)

    @contextmanager
    def _con_driver(self, stdout, **kwargs):
        """Patch the driver away: argv is ["fake", docx, destino] so the fake
        `_correr` knows which file the export was supposed to write."""
        with mock.patch.object(word_mod, "comando",
                               side_effect=lambda docx, pdf, campos: ["fake", str(docx), str(pdf)]), \
                self._driver(stdout, **kwargs):
            yield

    def test_a_missing_document_fails_without_launching_anything(self):
        with mock.patch.object(word_mod, "comando") as comando:
            resultado = word_mod.exportar(self.root / "nope.docx", self.pdf)
        self.assertEqual(resultado.exit_code, 1)
        self.assertIn("does not exist", resultado.motivo)
        comando.assert_not_called()

    def test_it_refuses_while_the_user_has_word_open(self):
        # Word is single-instance COM: attaching would mean quitting the user's
        # own Word at the end, with their unsaved documents on it.
        with mock.patch.object(rutas, "word_pids", return_value=[4321]), \
                mock.patch.object(word_mod, "comando") as comando:
            resultado = word_mod.exportar(self.docx, self.pdf)
        self.assertEqual(resultado.exit_code, 1)
        self.assertIn("already open", resultado.motivo)
        self.assertIn("4321", resultado.motivo)
        self.assertTrue(resultado.transitorio)
        comando.assert_not_called()

    def test_a_refusal_for_an_open_word_is_never_a_failure_of_word(self):
        with mock.patch.object(rutas, "word_pids", return_value=[4321]):
            resultado = word_mod.exportar(self.docx, self.pdf)
        self.assertEqual(resultado.exit_code, 1)
        # Nothing was proven and nothing was disproven: the caller must probe
        # again instead of trusting this.
        self.assertTrue(resultado.transitorio)

    def test_a_real_export_leaves_the_pdf_in_place_and_no_temporary_behind(self):
        with mock.patch.object(rutas, "word_pids", return_value=[]), \
                mock.patch.object(word_mod, "_esperar_que_word_salga", return_value=[]), \
                self._con_driver("APA7-VERSION:16.0\nAPA7-PDF-OK\n"):
            resultado = word_mod.exportar(self.docx, self.pdf)
        self.assertEqual(resultado.exit_code, 0)
        self.assertTrue(resultado.pdf_creado)
        self.assertTrue(self.pdf.read_bytes().startswith(b"%PDF-"))
        self.assertEqual(resultado.version, "16.0")
        self.assertFalse((self.root / "out.word.tmp.pdf").exists())

    def test_a_failed_export_does_not_destroy_a_previous_pdf(self):
        self.pdf.write_bytes(b"%PDF-1.7 the good one\n")
        with mock.patch.object(rutas, "word_pids", return_value=[]), \
                mock.patch.object(word_mod, "_esperar_que_word_salga", return_value=[]), \
                self._con_driver("APA7-PDF-ERROR: locked\n", exit_code=1, escribir_pdf=False):
            resultado = word_mod.exportar(self.docx, self.pdf)
        self.assertEqual(resultado.exit_code, 1)
        self.assertEqual(resultado.motivo, "locked")
        self.assertIn(b"the good one", self.pdf.read_bytes())
        self.assertFalse((self.root / "out.word.tmp.pdf").exists())

    def test_a_leftover_word_of_ours_is_stopped(self):
        with mock.patch.object(rutas, "word_pids", return_value=[]), \
                mock.patch.object(word_mod, "_esperar_que_word_salga", return_value=[777]), \
                mock.patch.object(word_mod, "kill_word") as matar, \
                self._con_driver("APA7-PDF-OK\n"):
            resultado = word_mod.exportar(self.docx, self.pdf)
        self.assertEqual(resultado.exit_code, 0)
        matar.assert_called_once_with([777])
        self.assertEqual(resultado.pids_restantes, [777])

    def test_a_timeout_is_reported_as_a_timeout(self):
        with mock.patch.object(rutas, "word_pids", return_value=[]), \
                mock.patch.object(word_mod, "_esperar_que_word_salga", return_value=[]), \
                self._con_driver("", exit_code=124, timed_out=True, escribir_pdf=False):
            resultado = word_mod.exportar(self.docx, self.pdf, timeout=3)
        self.assertTrue(resultado.timed_out)
        self.assertEqual(resultado.exit_code, 1)
        self.assertIn("within the timeout", resultado.motivo)

    def test_the_logs_are_written_where_the_caller_asked(self):
        logs = self.root / "logs"
        with mock.patch.object(rutas, "word_pids", return_value=[]), \
                mock.patch.object(word_mod, "_esperar_que_word_salga", return_value=[]), \
                self._con_driver("APA7-PDF-OK\n"):
            resultado = word_mod.exportar(self.docx, self.pdf, log_dir=logs)
        self.assertEqual(Path(resultado.log_dir), logs.resolve())
        self.assertEqual(Path(resultado.stdout_path).name, "word.out.log")

    def test_kill_word_never_raises(self):
        with mock.patch.object(rutas, "word_pids", return_value=[1]), \
                mock.patch.object(rutas, "kill_process_tree", side_effect=OSError("gone")):
            word_mod.kill_word()


class TestDisponible(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.work = Path(self.tmp.name)
        self.exe = self.work / "WINWORD.EXE"
        self.exe.write_bytes(b"MZ")
        # workdir() normally creates this; here it is created explicitly so the
        # tests exercise the cache writing, not mkdir.
        (self.work / ".work").mkdir()
        self.winword = mock.patch.object(rutas, "word_path", return_value=str(self.exe))
        self.winword.start()
        self.addCleanup(self.winword.stop)
        self.workdir = mock.patch.object(rutas, "workdir", return_value=self.work / ".work")
        self.workdir.start()
        self.addCleanup(self.workdir.stop)

    def _sondeo(self, ok=True, version="16.0", motivo="", transitorio=False):
        resultado = word_mod.Resultado(
            exit_code=0 if ok else 1, pdf_creado=ok, version=version,
            motivo=motivo, transitorio=transitorio)
        return mock.patch.object(word_mod, "sondear", return_value=resultado)

    def test_a_platform_without_a_driver_is_never_probed(self):
        with mock.patch.object(rutas, "is_windows", return_value=False), \
                mock.patch.object(rutas, "is_macos", return_value=False), \
                mock.patch.object(word_mod, "sondear") as sondear:
            estado = word_mod.disponible()
        self.assertFalse(estado.ok)
        self.assertIn("cannot be automated", estado.motivo)
        sondear.assert_not_called()

    def test_a_missing_executable_says_where_it_looked(self):
        with mock.patch.object(rutas, "word_path", return_value=None), \
                mock.patch.object(rutas, "word_candidates", return_value=["/a/WINWORD.EXE", "/b/WINWORD.EXE"]):
            estado = word_mod.disponible()
        self.assertFalse(estado.ok)
        self.assertIn("/a/WINWORD.EXE", estado.motivo)
        self.assertIn("/b/WINWORD.EXE", estado.motivo)

    def test_installed_without_the_com_class_is_not_an_engine(self):
        with mock.patch.object(rutas, "word_clase_registrada", return_value=False), \
                mock.patch.object(word_mod, "sondear") as sondear:
            estado = word_mod.disponible()
        self.assertFalse(estado.ok)
        self.assertIn("not registered", estado.motivo)
        sondear.assert_not_called()

    def test_the_cheap_answer_does_not_launch_word(self):
        with mock.patch.object(rutas, "word_clase_registrada", return_value=True), \
                mock.patch.object(word_mod, "sondear") as sondear:
            estado = word_mod.disponible(probar=False)
        self.assertTrue(estado.ok)
        self.assertFalse(estado.sondeado)
        self.assertEqual(estado.ruta, str(self.exe))
        sondear.assert_not_called()

    def test_the_probe_is_the_only_thing_that_makes_an_engine(self):
        with mock.patch.object(rutas, "word_clase_registrada", return_value=True), self._sondeo():
            estado = word_mod.disponible()
        self.assertTrue(estado.ok)
        self.assertTrue(estado.sondeado)
        self.assertFalse(estado.desde_cache)
        self.assertEqual(estado.version, "16.0")

    def test_a_failed_probe_reports_its_reason(self):
        with mock.patch.object(rutas, "word_clase_registrada", return_value=True), \
                self._sondeo(ok=False, motivo="the driver produced no PDF"):
            estado = word_mod.disponible()
        self.assertFalse(estado.ok)
        self.assertEqual(estado.motivo, "the driver produced no PDF")

    def test_the_second_run_answers_from_the_cache(self):
        with mock.patch.object(rutas, "word_clase_registrada", return_value=True), self._sondeo():
            word_mod.disponible()
        with mock.patch.object(rutas, "word_clase_registrada", return_value=True), \
                mock.patch.object(word_mod, "sondear") as sondear:
            estado = word_mod.disponible()
        sondear.assert_not_called()
        self.assertTrue(estado.ok)
        self.assertTrue(estado.desde_cache)
        self.assertTrue(estado.sondeado)

    def test_reevaluar_ignores_the_cache(self):
        with mock.patch.object(rutas, "word_clase_registrada", return_value=True), self._sondeo():
            word_mod.disponible()
        with mock.patch.object(rutas, "word_clase_registrada", return_value=True), \
                self._sondeo() as sondear:
            estado = word_mod.disponible(reevaluar=True)
        self.assertEqual(sondear.call_count, 1)
        self.assertFalse(estado.desde_cache)

    def test_a_word_upgrade_invalidates_the_cache(self):
        with mock.patch.object(rutas, "word_clase_registrada", return_value=True), self._sondeo():
            word_mod.disponible()
        # The build changes: same path, different bytes.
        self.exe.write_bytes(b"MZ plus a whole new Word")
        with mock.patch.object(rutas, "word_clase_registrada", return_value=True), \
                mock.patch.object(word_mod, "sondear") as sondear:
            word_mod.disponible()
        self.assertEqual(sondear.call_count, 1)

    def test_a_word_the_user_had_open_is_not_cached_as_a_failure(self):
        # Otherwise the next run reports "Word cannot convert here" forever,
        # after the user closed it.
        with mock.patch.object(rutas, "word_clase_registrada", return_value=True), \
                self._sondeo(ok=False, motivo="already open", transitorio=True):
            estado = word_mod.disponible()
        self.assertFalse(estado.sondeado)
        self.assertFalse((self.work / ".work" / word_mod._CACHE_NAME).exists())

    def test_a_corrupt_cache_is_ignored(self):
        with mock.patch.object(rutas, "word_clase_registrada", return_value=True), self._sondeo():
            word_mod.disponible()
        (self.work / ".work" / word_mod._CACHE_NAME).write_text("{not json", encoding="utf-8")
        with mock.patch.object(rutas, "word_clase_registrada", return_value=True), \
                mock.patch.object(word_mod, "sondear") as sondear:
            word_mod.disponible()
        self.assertEqual(sondear.call_count, 1)

    def test_dry_run_launches_nothing_and_proves_nothing(self):
        rutas.set_dry_run(True)
        try:
            with mock.patch.object(word_mod, "exportar") as exportar:
                resultado = word_mod.sondear()
            self.assertEqual(resultado.exit_code, 1)
            self.assertIn("--dry-run", resultado.motivo)
            exportar.assert_not_called()

            # The lookup stops before the probe: no answer, no cache, no Word.
            with mock.patch.object(rutas, "word_clase_registrada", return_value=True), \
                    mock.patch.object(word_mod, "exportar") as exportar:
                estado = word_mod.disponible()
            exportar.assert_not_called()
            self.assertFalse(estado.ok)
            self.assertFalse(estado.sondeado)
            self.assertFalse((self.work / ".work" / word_mod._CACHE_NAME).exists())
        finally:
            rutas.set_dry_run(False)


class TestSondear(unittest.TestCase):
    def test_it_proves_viability_by_really_exporting_a_minimal_document(self):
        visto = {}

        def exportar(docx, pdf, **kwargs):
            # Asserted HERE, inside the call: the scratch dir is gone by the
            # time sondear returns, which is the point of the test below.
            visto["docx_existia"] = Path(docx).is_file()
            visto["docx"] = Path(docx).name
            visto["pdf"] = Path(pdf).name
            visto["timeout"] = kwargs.get("timeout")
            return word_mod.Resultado(exit_code=0, pdf_creado=True, version="16.0")

        with mock.patch.object(word_mod, "exportar", side_effect=exportar), \
                mock.patch.object(rutas, "is_windows", return_value=True), \
                mock.patch.object(rutas, "word_pids", return_value=[]), \
                mock.patch.object(word_mod, "powershell_bin", return_value="powershell.exe"), \
                mock.patch.object(rutas, "word_path", return_value=r"C:\P\WINWORD.EXE"):
            resultado = word_mod.sondear(timeout=17)
        self.assertEqual(resultado.exit_code, 0)
        self.assertTrue(visto["docx_existia"])
        self.assertEqual(visto["docx"], "sondeo.docx")
        self.assertEqual(visto["pdf"], "sondeo.pdf")
        self.assertEqual(visto["timeout"], 17)

    def test_the_scratch_directory_does_not_survive(self):
        antes = set(Path(tempfile.gettempdir()).glob("apa7-sondeo-word-*"))
        with mock.patch.object(word_mod, "exportar",
                               return_value=word_mod.Resultado(exit_code=0, pdf_creado=True)), \
                mock.patch.object(rutas, "is_windows", return_value=True), \
                mock.patch.object(rutas, "word_pids", return_value=[]), \
                mock.patch.object(word_mod, "powershell_bin", return_value="powershell.exe"), \
                mock.patch.object(rutas, "word_path", return_value=r"C:\P\WINWORD.EXE"):
            word_mod.sondear()
        # A probe that leaks its scratch directory would leave a directory with a
        # PDF in %TEMP% on every single run.
        self.assertEqual(set(Path(tempfile.gettempdir()).glob("apa7-sondeo-word-*")) - antes, set())

    def test_the_scratch_directory_is_deleted_even_when_the_export_raises(self):
        antes = set(Path(tempfile.gettempdir()).glob("apa7-sondeo-word-*"))
        with mock.patch.object(word_mod, "exportar", side_effect=OSError("boom")), \
                mock.patch.object(rutas, "is_windows", return_value=True), \
                mock.patch.object(rutas, "word_pids", return_value=[]), \
                mock.patch.object(word_mod, "powershell_bin", return_value="powershell.exe"), \
                mock.patch.object(rutas, "word_path", return_value=r"C:\P\WINWORD.EXE"):
            with self.assertRaises(OSError):
                word_mod.sondear()
        self.assertEqual(set(Path(tempfile.gettempdir()).glob("apa7-sondeo-word-*")) - antes, set())

    def test_a_directory_that_survives_is_reported_instead_of_being_forgotten(self):
        # rmtree(ignore_errors=True) can leave it when Word still holds the PDF.
        with mock.patch.object(rutas, "remove_tree", return_value=False), \
                mock.patch.object(rutas, "log") as log:
            self.assertFalse(word_mod._limpiar_scratch(Path(tempfile.gettempdir()) / "x"))
        log.assert_called_once()
        self.assertIn("safe to remove it by hand", log.call_args[0][0])


if __name__ == "__main__":
    unittest.main()
