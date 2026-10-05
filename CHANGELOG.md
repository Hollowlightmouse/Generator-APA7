# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

Output organization: a converted document now lands as a predictable set of
files next to its `.md`, instead of the mixed pile that three phases each chose
for themselves.

### Added

- `scripts/lib/rutas.py` owns **where a document's files go**:
  `rutas_documento()` (the `.md`, `.docx`, `.pdf`, working folder, `datos/`,
  `logs/` and the default `--out` / `--log` / `--json` / `--outdir` per phase),
  `sanea_nombre()` for the working-folder name (NFKD without diacritics,
  invalid and control characters and spaces to `_`, Windows reserved names with
  a `_` suffix, capped at 60 characters), and `anota_fuente()` for the
  `datos/fuente.json` anchor that lets a phase resolve its document without
  being told again.
- `--carpeta-trabajo <dir>` on `parse`, `build`, `export` and `verify`, for the
  documents that must keep their intermediates somewhere other than next to
  them.
- Warning on stderr when a working folder that already holds an anchor is
  pointed at a different `.md`.
- `export` without `--docx` resolves the document from the working folder, and
  says so instead of guessing when the folder holds more than one `.docx`.

### Changed

- **Defaults, not requirements.** Each phase now derives its own paths, so
  `--out`, `--log`, `--outdir` and `--json` became optional. The deliverables
  keep the original name of the `.md`: `informe.md` → `informe.docx`,
  `informe.pdf` and `informe_apa/`.
- The working folder is **reused, never versioned**: a rerun replaces the files
  in place instead of leaving `informe (2).docx` behind.
- **Explicit flags still win.** `export --outdir X` without
  `--carpeta-trabajo` keeps its historical meaning, PDF *and*
  `X/_logs/03-export.log`, so existing callers are unaffected; that is the only
  case that still writes `_logs/`.
- The double pass keeps its temporaries in `logs/` instead of the folder of the
  first-pass `.docx`, `paginas-de-pdf.py` is invoked without `--out` so the page
  map lands in `datos/`'s sibling `logs/paginas.json`, and the throwaway PDF is
  **deleted once the second pass succeeds**. It is kept when the second pass
  fails, because it is then the only record of what was measured.
- A failed `verify` keeps its PDF and its whole working folder, as the failure
  report intends.
- Documentation: an "Output layout" section in `SKILL.md`, a "Where the files
  land" section in `README.md`, and the per-phase path table in
  `references/system-requirements.md`.
- `.gitignore` ignores `*_apa/` so a conversion run inside the repository cannot
  be committed by accident.

### Fixed

- **A false critical failure on every table wider than a page.** "Tables carry a
  note below" only looked for the note on the page holding the `Tabla N` label,
  but a table that fills a page ends on the next one and its note follows it
  there. The check now searches from the label up to the next table's label, so
  the note is accepted wherever the table ends while still belonging to it, and
  an in-text cross-reference (`... se observa en la Tabla 6 ...`) is no longer
  mistaken for a caption.
- The build log is `logs/02-build.log`, the name the documentation had been
  promising since 1.0.0; the code wrote `02-docx.log`.
- `export` without `--docx` now works at all: the fallback was documented but
  the argument was still required, so it could not be reached.

### Note on the acceptance criterion

The original criterion for this change was "only the PDF and the working folder
appear next to the `.md`". That was **amended by decision**: the `.docx` is a
deliverable in its own right and is not hidden inside the working folder, so
next to `informe.md` there are now `informe.docx`, `informe.pdf` and
`informe_apa/` — and nothing else.

## [1.0.0] - 2026-10-03

First release of the `generate-apa-document` skill. There is no earlier tag, so
this section records the state that is being published, including the hardening
that was done before the first release.

### Added

- Initial release of the `generate-apa-document` skill.
- Interactive-free pipeline: `check` -> `install` -> `parse` -> `build` ->
  `export` -> `verify`, driven by `scripts/apa7.py`.
- Markdown to `MANIFEST.json` parser with automatic correction of conversion
  artifacts (glued words, duplicated list markers, glued numbering) and a
  report split into corrections and items to review.
- Word `.docx` builder with institutional 3-zone cover, real `TOC` fields for
  the table of contents and the lists of tables and figures, wide tables on an
  occasional landscape page, and APA 7 references with hanging indent.
- Headless LibreOffice export to PDF (Microsoft Word is not required).
- PDF verification with `pymupdf`: cover, indices, captions, landscape pages,
  notes, complete figure images and references.
- `scripts/paginas-de-pdf.py`: reads a throwaway PDF and reports the real page
  of every heading, table and figure, used by the TOC double pass.
- `scripts/tests/fixtures/sintetico.md`: synthetic source document for the
  end-to-end parse test (never a real document).

### Changed

- `install --dry-run` creates nothing at all, not even the working directory or
  the npm anchor, and distinguishes "simulated" from "not confirmed" in its
  messages.
- Every command suggested by `apa7.py` prints the real, copyable command with
  the interpreter that is running (`sys.executable`), instead of assuming a
  `python` on `PATH`.
- The `soffice` candidate list uses the launcher names Snap and Flatpak
  document (`/snap/bin/libreoffice`, `org.libreoffice.LibreOffice`), marked as
  best effort and unverified.
- `scripts/lib/fuentes.py`: `ArialMT` is no longer accepted as a Times New
  Roman substitute, and the comments about the accepted metric clones were
  corrected.

### Fixed

- Arial PDFs are now rejected by the typography check instead of passing it.
- The test suite is independent of the account it runs under: the elevation
  tests state the privilege they assume, so they pass both as root and as an
  ordinary user.
- The TOC fields ship the measured page numbers as their cached result, so the
  exported PDF shows the indices correctly without Microsoft Word.
