#!/usr/bin/env python3
"""paginas-de-pdf.py - real page of every index entry of an APA document.

Part of the two-pass index build. The first pass exports the .docx with the TOC
fields unresolved; this script reads that PDF and returns, as JSON, the page
where each section, table and figure really landed:

    {"secciones": {"<hid>": 2}, "tablas": {"1": 5}, "figuras": {"1": 6}}

apa7.py feeds that file back to build-docx.js (--paginas-json), so the second
pass ships those numbers as the cached result of the TOC fields.

Runs on the venv interpreter, because pymupdf is installed there.
"""

import argparse
import json
import re
import sys
from pathlib import Path


def norm(texto):
    return re.sub(r"\s+", " ", texto or "").strip().lower()


def paginas_de_secciones(doc, secciones):
    """hid -> 1-based page, from the PDF outline.

    LibreOffice writes one outline entry per paragraph that carries an applied
    outline level, in document order. The entries are matched against the
    manifest's sections by order AND title: the title check keeps one skipped
    heading from shifting every page that follows it.
    """
    mapa = {}
    toc = doc.get_toc()          # [[nivel, titulo, pagina], ...], pagina 1-based
    if not toc:
        return mapa, 0
    cuerpo = min(p for _n, _t, p in toc)
    idx = 0
    for _nivel, titulo, pagina in toc:
        nt = norm(titulo)
        elegido = None
        if idx < len(secciones) and norm(secciones[idx].get("texto", "")) == nt:
            elegido = idx
            idx += 1
        else:
            for j in range(idx, len(secciones)):
                if norm(secciones[j].get("texto", "")) == nt:
                    elegido = j
                    idx = j + 1
                    break
        if elegido is None:
            continue
        mapa[str(secciones[elegido].get("hid"))] = pagina
    return mapa, cuerpo


def paginas_de_indices(doc):
    """0-based pages that belong to the cover or to an index.

    The body always starts on a fresh page (build-docx.js closes the indexes
    with a page break), so the page after the last index page is where the
    captions begin. This is what lets tables and figures be found even when the
    document has no headings at all and the outline is empty.
    """
    marcadores = ("tabla de contenido", "indice de tablas", "indice de figuras")
    paginas = []
    for p in range(doc.page_count):
        texto = norm(doc[p].get_text())
        if any(m in texto for m in marcadores):
            paginas.append(p)
    return paginas


def paginas_de_etiquetas(doc, desde, etiqueta):
    """index (as string) -> 1-based page of the first '<Etiqueta> <n>' line at
    or after `desde` (0-based). The caption is the first occurrence in the
    body: starting after the first outline page leaves the indexes behind."""
    paginas = {}
    for p in range(desde, doc.page_count):
        for linea in doc[p].get_text().splitlines():
            m = re.match(r"^\s*(%s)\s+(\d+)\b" % etiqueta, linea, re.I)
            if not m:
                continue
            n = int(m.group(2))
            if n not in paginas:
                paginas[n] = p + 1
    return {str(k): v for k, v in sorted(paginas.items())}


def escribir(destino, datos):
    Path(destino).parent.mkdir(parents=True, exist_ok=True)
    with open(destino, "w", encoding="utf-8") as fh:
        json.dump(datos, fh, ensure_ascii=False)


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--pdf", required=True)
    ap.add_argument("--manifiesto", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--log")
    args = ap.parse_args(argv)

    vacio = {"secciones": {}, "tablas": {}, "figuras": {}}

    try:
        with open(args.manifiesto, encoding="utf-8") as fh:
            M = json.load(fh)
    except (OSError, ValueError) as exc:
        sys.stderr.write("paginas-de-pdf: could not read the manifest (%s)\n" % exc)
        escribir(args.out, vacio)
        return 0

    try:
        import fitz
    except ImportError as exc:
        sys.stderr.write("paginas-de-pdf: pymupdf is not available (%s)\n" % exc)
        escribir(args.out, vacio)
        return 0

    if not Path(args.pdf).is_file():
        sys.stderr.write("paginas-de-pdf: the PDF does not exist: %s\n" % args.pdf)
        return 1

    salida = dict(vacio)
    doc = fitz.open(args.pdf)
    try:
        secciones = M.get("secciones", []) or []
        mapa, cuerpo = paginas_de_secciones(doc, secciones)
        salida["secciones"] = mapa
        indices = paginas_de_indices(doc)
        if indices:
            inicio = max(indices) + 1
        elif cuerpo:
            inicio = cuerpo - 1
        else:
            inicio = 1
        salida["tablas"] = paginas_de_etiquetas(doc, inicio, "Tabla")
        salida["figuras"] = paginas_de_etiquetas(doc, inicio, "Figura")
    finally:
        doc.close()

    escribir(args.out, salida)
    sys.stderr.write(
        "paginas-de-pdf: %d section(s), %d table(s), %d figure(s)\n"
        % (len(salida["secciones"]), len(salida["tablas"]), len(salida["figuras"])))
    if args.log:
        Path(args.log).parent.mkdir(parents=True, exist_ok=True)
        with open(args.log, "w", encoding="utf-8") as fh:
            json.dump(salida, fh, ensure_ascii=False, indent=2)
    return 0


if __name__ == "__main__":
    sys.exit(main())
