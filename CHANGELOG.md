# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
