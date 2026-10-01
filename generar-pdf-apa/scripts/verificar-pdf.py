#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
verificar-pdf.py - Comprueba el PDF generado contra su MANIFEST.json.

Este script es el PASO 7 del pipeline. No juzga la calidad academica del
documento (eso lo hace el agente leyendo el texto), comprueba lo que se puede
comprobar de forma automatica y REPRODUCIBLE:

  1. El PDF se abre y tiene paginas.
  2. El tamano de pagina es Carta (612 x 792 pt) con margenes de 1 pulgada.
  3. La portada NO lleva numero de pagina.
  4. La pagina 2 SI lleva el numero, y en la esquina superior derecha.
  5. Toda seccion del manifiesto aparece como titulo en el PDF.
  6. Toda tabla aparece, con su numero y su contenido.
  7. Toda figura aparece.
  8. Toda referencia aparece en la seccion de referencias.
  9. Los numeros de la tabla de contenido COINCIDEN con la pagina real de cada
     titulo. Este es el control que de verdad importa: un PAGEREF sin resolver
     deja un 0 o un guion, y un desajuste de paginas invalida la TOC.
 10. No hay marcadores rotos ("Error! Bookmark not defined", "0", "??").
 11. Solo se usa la fuente tipografica declarada.
 12. Interlineado doble en los parrafos de cuerpo.

Uso:
    python verificar-pdf.py --pdf salida.pdf --manifiesto MANIFIEST.json
                            [--log salida/_logs/04-verificacion.log]
                            [--json salida/_logs/04-verificacion.json]

Devuelve 0 si no hay FALLAS, 1 si hay alguna.
"""

import argparse
import json
import os
import re
import sys
import unicodedata
from datetime import datetime

try:
    import pymupdf
except ImportError:  # python <3.13 expone el mismo paquete como 'fitz'
    import fitz as pymupdf

PT_PULGADA = 72.0
PAGINA_CARTA = (612.0, 792.0)

# Bandas verticales de la portada, en puntos desde el borde superior de la hoja.
# Debe coincidir con las constantes ZONA_* de build-docx.js.
PORTADA_BANDA_ALTA = (72.0, 288.0)
PORTADA_BANDA_CENTRO = (288.0, 504.0)
PORTADA_BANDA_BAJA = (504.0, 720.0)


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------

def norm(text):
    """Minusculas, sin acentos, sin puntuacion, espacios colapsados."""
    t = unicodedata.normalize("NFD", (text or "").lower())
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    t = re.sub(r"[^a-z0-9]+", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def y_de_texto(bloques_pagina, aguja, prefijo=24):
    """y del primer bloque que contiene el comienzo de `aguja`, o None.

    Se usa para anclar la comprobacion de zonas a los textos REALES del
    manifiesto. Contar bloques por posicion no sirve: un titulo que ocupa dos
    lineas produce dos bloques y desplaza el guessed "el segundo bloque son los
    integrantes".
    """
    a = norm(aguja)[:prefijo]
    if len(a) < 6:
        return None
    for b in sorted(bloques_pagina, key=lambda x: x[1]):
        if a and a in norm(b[4]):
            return b[1]
    return None


def cargar_texto(pdf):
    """Devuelve (lista_de_textos_por_pagina, lista_de_bloques_por_pagina)."""
    doc = pymupdf.open(pdf)
    paginas = []
    bloques = []
    for i in range(doc.page_count):
        paginas.append(doc[i].get_text())
        bloques.append(doc[i].get_text("blocks"))
    return doc, paginas, bloques


def paginas_con(paginas, aguja_norm, ignorar_portada=True, desde=0):
    """Indices (0-based) de las paginas cuyo texto normalizado contiene la aguja."""
    out = []
    for i, t in enumerate(paginas):
        if i < desde:
            continue
        if ignorar_portada and i == 0:
            continue
        if aguja_norm and aguja_norm in norm(t):
            out.append(i)
    return out


def lineas_toc(paginas):
    """
    Extrae las lineas de la tabla de contenido: 'Titulo ......... 12'.
    Se apoya en el leader de puntos, que es lo que genera el tab con puntos.
    """
    filas = []
    for i, t in enumerate(paginas[:4]):     # la TOC vive al principio
        for linea in t.splitlines():
            l = linea.strip()
            m = re.match(r"^(.*?)\s*\.{4,}\s*(\d{1,4})\s*$", l)
            if m:
                filas.append({"titulo": m.group(1).strip(), "pagina": int(m.group(2)), "pagina_pdf": i})
    return filas


RE_ETIQUETA_INDICE = re.compile(r"^(tabla|figura)\s+\d+")


def pagina_fin_de_indices(paginas, lineas):
    """
    Ultima pagina ocupada por los indices (TOC, indice de tablas, de figuras).

    Es imprescindible conocerla: el titulo de cada seccion aparece TAMbien
    dentro de la propia tabla de contenido. Buscar "la pagina donde esta
    Introduccion" sin excluir la TOC devuelve la pagina 2 (la de la TOC) en
    lugar de la pagina real, y produce 29 desajustes falsos.
    """
    if not lineas:
        return 0
    return max(f["pagina_pdf"] for f in lineas)


def texto_en_orden_visual(doc, desde, hasta):
    """
    Texto de las paginas [desde, hasta) en ORDEN DE LECTURA real.

    `get_text()` devuelve los bloques en el orden del flujo interno del PDF, no
    en el visual: al haber fragmentos en cursiva dentro de cada referencia, el
    orden se descoloca y una comprobacion de "orden alfabetico" da un falso
    negativo. Ordenar por (y, x) reconstruye la lectura correcta.
    """
    partes = []
    for i in range(desde, min(hasta, doc.page_count)):
        bloques = [b for b in doc[i].get_text("blocks") if len(b) >= 5 and isinstance(b[4], str)]
        bloques.sort(key=lambda b: (round(b[1], 1), round(b[0], 1)))
        partes.append("\n".join(b[4] for b in bloques))
    return "\n".join(partes)


# ---------------------------------------------------------------------------
# Verificaciones
# ---------------------------------------------------------------------------

class Resultado:
    def __init__(self):
        self.items = []

    def anota(self, nombre, ok, detalle="", critico=True):
        self.items.append({
            "nombre": nombre,
            "ok": bool(ok),
            "detalle": detalle,
            "critico": critico,
        })
        return ok

    @property
    def fallas(self):
        return [i for i in self.items if not i["ok"] and i["critico"]]

    @property
    def avisos(self):
        return [i for i in self.items if not i["ok"] and not i["critico"]]


def verificar(args):
    R = Resultado()
    with open(args.manifiesto, encoding="utf-8") as fh:
        M = json.load(fh)

    doc, paginas, bloques = cargar_texto(args.pdf)
    total = doc.page_count
    texto_total = "\n".join(paginas)
    norm_total = norm(texto_total)

    # Pagina a partir de la cual empieza el cuerpo: los indices se repiten a si
    # mismos, asi que buscar un titulo antes de aqui da un falso positivo.
    _toc = lineas_toc(paginas)
    fin_indices = pagina_fin_de_indices(paginas, _toc)
    CUERPO = fin_indices + 1

    # --- 1. Apertura -----------------------------------------------------
    R.anota("El PDF se abre y tiene paginas", total > 0,
            "%d paginas" % total)

    # --- 2. Tamano de pagina y margenes ----------------------------------
    r0 = doc[0].rect if total else None
    if r0:
        w, h = round(r0.width, 1), round(r0.height, 1)
        R.anota("Pagina Carta (612x792 pt)", abs(w - PAGINA_CARTA[0]) < 2 and abs(h - PAGINA_CARTA[1]) < 2,
                "medida real %.1f x %.1f pt" % (w, h))
    else:
        R.anota("Pagina Carta (612x792 pt)", False, "sin paginas", critico=False)

    # Margen izquierdo: el texto no debe empezar antes de 1 pulgada (con margen
    # de tolerancia de 2 pt por el redondeo del motor de texto).
    xs = []
    for pg in bloques[1:] if len(bloques) > 1 else bloques:
        for b in pg:
            if len(b) >= 5 and isinstance(b[4], str) and b[4].strip():
                xs.append(b[0])
    if xs:
        minx = min(xs)
        R.anota("Margen izquierdo >= 1 pulgada", minx >= PT_PULGADA - 2,
                "x minimo del texto = %.1f pt (1 pulgada = 72 pt)" % minx)
    else:
        R.anota("Margen izquierdo >= 1 pulgada", False, "no se localizaron bloques de texto", critico=False)

    # --- 3 y 4. Numeracion de la portada ---------------------------------
    # La portada no debe mostrar ningun numero suelto en la esquina superior
    # derecha; la pagina 2 si, y debe coincidir con el 2 real.
    def numero_en_esquina(i):
        if i >= total:
            return None
        zona = doc[i].get_text("text", clip=pymupdf.Rect(0, 0, r0.width, 60))
        m = re.search(r"\b(\d{1,4})\b", zona)
        return int(m.group(1)) if m else None

    n1 = numero_en_esquina(0)
    R.anota("La portada NO lleva numero de pagina", n1 is None,
            "numero detectado en la esquina superior derecha: %s" % (n1 if n1 is not None else "ninguno"))

    if total > 1:
        n2 = numero_en_esquina(1)
        R.anota("La pagina 2 muestra el numero 2", n2 == 2,
                "numero detectado: %s" % (n2 if n2 is not None else "ninguno"))
    else:
        R.anota("La pagina 2 muestra el numero 2", False, "el PDF solo tiene 1 pagina", critico=False)

    # --- 3b. Portada en 3 zonas (AVISOS, nunca bloquean) -------------------
    # Estas cuatro NO son fallas: la portada se construye con los datos que haya.
    # Si falta el logo o el titulo del docente se avisa para que el agente lo
    # pregunte, pero el documento sale igual. Bloquear aqui seria contraproducente.
    portada = M.get("portada") or {}

    if not total:
        R.anota("La portada respeta las 3 zonas", False, "el PDF no tiene paginas", critico=False)
    else:
        pg0 = doc[0]
        # Se ignoran los bloques de la franja de la cabecera (y < 72 pt), donde
        # solo puede aparecer el numero de pagina.
        pb = [b for b in bloques[0]
              if len(b) >= 5 and isinstance(b[4], str) and b[4].strip() and b[1] >= PT_PULGADA - 2]
        pb.sort(key=lambda b: b[1])
        alto = PORTADA_BANDA_ALTA[1]        # 288: fin de la banda alta
        bajo = PORTADA_BANDA_BAJA[0]        # 504: fin de la banda central
        centro = PORTADA_BANDA_CENTRO[0]    # 288: inicio de la banda central

        # Cada zona se ancla a su propio texto del manifiesto, no al orden de
        # los bloques: asi un titulo de dos lineas no falsea la medicion.
        autores_txt = [a for a in (portada.get("autores") or []) if str(a).strip()]
        campos_bajos = [c for c in ("docente", "materia_nrc", "facultad",
                                    "vicerrectoria", "sede", "fecha")
                        if portada.get(c)]
        y_titulo = y_de_texto(pb, portada.get("titulo") or "")
        y_autores = y_de_texto(pb, autores_txt[0]) if autores_txt else None

        problemas = []
        if not pb:
            R.anota("La portada respeta las 3 zonas", False,
                    "no se localizo texto en la portada", critico=False)
        else:
            if y_titulo is not None and y_titulo >= alto:
                problemas.append("el titulo arranca en y=%.0f pt, deberia estar sobre y=%.0f" % (y_titulo, alto))
            if autores_txt and y_autores is None:
                problemas.append("no se localizaron los integrantes en la portada")
            elif y_autores is not None and not (centro <= y_autores < bajo):
                problemas.append("los integrantes estan en y=%.0f pt, deberian caer entre %.0f y %.0f"
                                 % (y_autores, centro, bajo))
            # La banda baja solo se comprueba si el manifiesto trae datos que
            # anclar. En una portada reduced a un titulo no hay bloque que medir,
            # y avisar de eso seria un falso positivo.
            y_bajo = None
            if campos_bajos:
                y_bajo = y_de_texto(pb, portada.get(campos_bajos[0]))
                if y_bajo is None:
                    # El campo existe en el manifiesto pero no aparece en la
                    # pagina 1: lo mas probable es que la portada se haya
                    # desbordado y el bloque institucional este en la pagina 2.
                    problemas.append("'%s' no aparece en la portada: la portada se desbordo a otra pagina"
                                     % campos_bajos[0])
                elif y_bajo <= bajo:
                    problemas.append("el bloque institucional arranca en y=%.0f pt, deberia estar bajo y=%.0f"
                                     % (y_bajo, bajo))
            fin = pb[-1][3]
            if campos_bajos and fin <= bajo:
                problemas.append("el bloque institucional termina en y=%.0f pt, deberia llegar a la banda baja" % fin)
            if campos_bajos:
                zona_baja_txt = ("y=%.0f..%.0f" % (y_bajo, fin)) if y_bajo is not None else "no encontrado en la pagina 1"
            else:
                zona_baja_txt = "no aplica (sin datos)"
            detalle = "titulo y=%s | integrantes y=%s | bloque institucional %s" % (
                "%.0f" % y_titulo if y_titulo is not None else "n/d",
                "%.0f" % y_autores if y_autores is not None else "n/d",
                zona_baja_txt)
            R.anota("La portada respeta las 3 zonas", not problemas,
                    "; ".join(problemas) if problemas else detalle,
                    critico=False)

        # Los integrantes van en UN SOLO parrafo: si hay dos o mas, ningun
        # nombre puede empezar una linea nueva, salvo el ultimo si el anterior
        # acaba justo en el corte de linea.
        if len(autores_txt) > 1:
            lineas_p0 = [" ".join(l.split()) for l in pg0.get_text().splitlines() if l.strip()]
            i_ult = next((i for i, l in enumerate(lineas_p0) if l.endswith(" ".join(str(autores_txt[-1]).split()))), None)
            i_pre = next((i for i, l in enumerate(lineas_p0) if l.endswith(" ".join(str(autores_txt[-2]).split()))), None)
            if i_ult is None:
                R.anota("Los integrantes van en un solo parrafo", False,
                        "no se encontro el ultimo nombre (%s) en la portada" % autores_txt[-1], critico=False)
            elif i_pre is not None and i_pre - i_ult > 1:
                R.anota("Los integrantes van en un solo parrafo", False,
                        "hay una linea entre el penultimo y el ultimo integrante: parece un parrafo por persona",
                        critico=False)
            else:
                R.anota("Los integrantes van en un solo parrafo", True,
                        "%d integrantes en un solo parrafo" % len(autores_txt))

        # Logo: coherencia entre lo que pide el manifiesto y lo que se dibujo.
        logo = portada.get("logo")
        if logo:
            if not os.path.exists(logo):
                R.anota("El logo de la portada existe", False,
                        "el manifiesto apunta a un archivo inexistente: %s" % logo, critico=False)
            else:
                n_img = len(pg0.get_images())
                R.anota("El logo de la portada existe", n_img > 0,
                        "archivo presente" if n_img else "el archivo existe pero la portada no muestra imagenes",
                        critico=False)
        else:
            R.anota("Se pregunto por el logo de la portada", False,
                    "no hay logo en el manifiesto; es opcional, pero hay que preguntar siempre antes de omitirlo",
                    critico=False)

        # El titulo o profesion del docente se pregunta siempre.
        if portada.get("docente") and not portada.get("docente_titulo"):
            R.anota("Se pregunto por el titulo del docente", False,
                    "docente_titulo vacio: hay que preguntar su titulo o profesion", critico=False)

    # --- 5. Secciones ----------------------------------------------------
    faltantes = []
    for s in M.get("secciones", []):
        if not paginas_con(paginas, norm(s["texto"]), desde=CUERPO):
            faltantes.append(s["texto"][:60])
    R.anota("Todas las secciones del manifiesto aparecen en el PDF", not faltantes,
            "faltan %d de %d (buscando desde la pagina %d)%s"
            % (len(faltantes), len(M.get("secciones", [])), CUERPO + 1,
               (": " + "; ".join(faltantes[:5])) if faltantes else ""))

    # --- 6. Tablas -------------------------------------------------------
    problemas_tablas = []
    for t in M.get("tablas", []):
        if ("tabla %d" % t["indice"]) not in norm_total:
            problemas_tablas.append("no aparece 'Tabla %d'" % t["indice"])
            continue
        # contenido: al menos la primera celda de la primera fila de datos
        celdas = [c for fila in t["filas"] for c in fila if c.strip()]
        if celdas:
            muestra = norm(celdas[0])[:40]
            if muestra and muestra not in norm_total:
                problemas_tablas.append("Tabla %d: no aparece el contenido '%s'" % (t["indice"], muestra))
    R.anota("Todas las tablas aparecen con su numero y contenido", not problemas_tablas,
            "%d tabla(s); %s" % (len(M.get("tablas", [])),
                                 "; ".join(problemas_tablas[:4]) if problemas_tablas else "todo presente"))

    # --- 7. Figuras ------------------------------------------------------
    problemas_fig = []
    figs_omitidas = []
    figs_esperadas = []
    for f in M.get("figuras", []):
        if not f.get("existe"):
            figs_omitidas.append(f["indice"])
            continue
        figs_esperadas.append(f)
        if ("figura %d" % f["indice"]) not in norm_total:
            problemas_fig.append("no aparece 'Figura %d'" % f["indice"])
        if f.get("titulo") and norm(f["titulo"])[:40] not in norm_total:
            problemas_fig.append("Figura %d: no aparece la leyenda" % f["indice"])
    R.anota("Todas las figuras aparecen con su leyenda",
            not problemas_fig if figs_esperadas else True,
            "%d figura(s) insertada(s); %s" % (
                len(figs_esperadas),
                "; ".join(problemas_fig[:4]) if problemas_fig else "todo presente"),
            critico=bool(figs_esperadas))
    # Una figura sin archivo no es un fallo del PDF: el pipeline la omite y lo
    # avisa, asi que se informa aparte y sin tumbar la comprobacion.
    if figs_omitidas:
        R.anota("Figuras sin archivo en el origen", False,
                "%d: %s" % (len(figs_omitidas),
                            ", ".join("Figura %d" % i for i in figs_omitidas)
                            + " no se insertaron; el .md las apunta a un archivo "
                              "inexistente"),
                critico=False)

    # --- 8. Referencias --------------------------------------------------
    refs = M.get("referencias", [])
    faltan_refs = []
    for r in refs:
        # comparar por un fragmento distintivo: los 60 primeros caracteres
        frag = norm(r["texto"])[:60]
        if frag and frag not in norm_total:
            faltan_refs.append(r["texto"][:55])
    R.anota("Todas las referencias del .md estan en el PDF", not faltan_refs,
            "%d referencia(s); faltan %d%s" % (len(refs), len(faltan_refs),
                                               (": " + "; ".join(faltan_refs[:3])) if faltan_refs else ""))

    # indice de referencias ordenado alfabeticamente
    if len(refs) > 1:
        # La pagina de la seccion de referencias es la del CUERPO, no la de la
        # TOC (que tambien contiene la palabra "Referencias").
        pag_ref = None
        for i in range(CUERPO, total):
            if "referencias" in norm(paginas[i]):
                pag_ref = i
                break
        if pag_ref is not None:
            pag = norm(texto_en_orden_visual(doc, pag_ref, total))
            pos = [pag.find(norm(r["texto"])[:30]) for r in refs]
            # Ojo: -1 significa "no encontrado", y -1 es verdadero en Python.
            todos = all(p >= 0 for p in pos)
            if todos:
                # El orden hay que juzgarlo sobre el ORDEN VISUAL del PDF, no
                # sobre el del manifiesto: comparar `pos` contra `sorted(pos)`
                # en el orden del manifiesto daria siempre DESORDENADAS cuando
                # el .md no venia ordenado, que es el caso normal.
                visual = [refs[k]["texto"] for k in sorted(range(len(pos)), key=lambda i: pos[i])]
                desorden = [visual[i][:40] for i in range(1, len(visual))
                            if visual[i].lower() < visual[i - 1].lower()]
                R.anota("Las referencias estan en orden alfabetico", not desorden,
                        ("orden visual correcto: %s" % " | ".join(t[:22] for t in visual))
                        if not desorden
                        else "%d fuera de orden: %s" % (len(desorden), "; ".join(desorden[:3])),
                        critico=False)
            else:
                R.anota("Las referencias estan en orden alfabetico", False,
                        "no verificable: %d de %d referencias no se localizaron en la seccion"
                        % (sum(1 for p in pos if p < 0), len(refs)),
                        critico=False)
        else:
            R.anota("Las referencias estan en orden alfabetico", False,
                    "no se encontro la pagina de la seccion de referencias", critico=False)

    # --- 9. Coherencia de la tabla de contenido --------------------------
    # Las entradas de titulo vacio son continuaciones de una linea de la TOC que
    # se partio en dos (los titulos largos se envuelven). No son entradas.
    toc = [f for f in _toc if f["titulo"].strip()]
    if not toc and (M.get("opciones", {}).get("indice_tablas") or M.get("secciones")):
        R.anota("La tabla de contenido tiene entradas", False,
                "no se encontro ninguna linea con leader de puntos", critico=False)
    else:
        R.anota("La tabla de contenido tiene entradas", bool(toc),
                "%d entradas con numero de pagina" % len(toc), critico=False)

    # Los titulos REPETIDOS (p. ej. "Fortalezas, Limitaciones y Aplicabilidad"
    # cuatro veces) no se pueden comprobar emparejando uno a uno: hay que
    # emparejar la k-esima aparicion de la TOC con la k-esima del cuerpo.
    reclamadas = {}
    for fila in toc:
        reclamadas.setdefault(norm(fila["titulo"]), []).append(fila["pagina"])

    desajustes = []
    for titulo_norm, pedidas in reclamadas.items():
        # Una entrada de indice de tablas o figuras reza "Tabla 1. (sin titulo)",
        # pero en el cuerpo solo se imprime "Tabla 1": el titulo completo no
        # existe ahi y buscarlo daria un falso negativo. Para esas entradas se
        # comprueba solo la etiqueta.
        etiqueta = RE_ETIQUETA_INDICE.match(titulo_norm)
        busca = etiqueta.group(0).strip() if etiqueta else titulo_norm
        # Si la entrada fue generada con un placeholder de titulo ausente
        # (por ejemplo "(sin titulo)"), no debe tolerarse: el proceso debe haber
        # abortado antes de generar el DOCX. Se marca como fallo critico para
        # detectar el caso.
        if "(sin titulo)" in titulo_norm or "(sin leyenda)" in titulo_norm:
            desajustes.append("'%s' contiene '(sin titulo/leyenda)' en la TOC" % titulo_norm[:40])
            continue
        encontradas = [p + 1 for p in paginas_con(paginas, busca, desde=CUERPO)]
        if not encontradas:
            desajustes.append("'%s': no se encuentra en el cuerpo" % titulo_norm[:40])
            continue
        if len(encontradas) < len(pedidas):
            desajustes.append("'%s': la TOC lista %d veces y el cuerpo la tiene %d"
                              % (titulo_norm[:40], len(pedidas), len(encontradas)))
            continue
        for k, pedida in enumerate(pedidas):
            if encontradas[k] == pedida:
                continue
            # Un titulo puede aparecer de paso en el cuerpo (una mencion dentro
            # de un parrafo) ANTES de que empiece la seccion real. En ese caso la
            # pagina esperada sigue siendo correcta y el emparejamiento por
            # posicion no debe marcar fallo: basta con que la pagina que dice la
            # TOC este entre las paginas donde aparece el titulo.
            if pedida in encontradas[k:]:
                continue
            desajustes.append("'%s' (aparicion %d): la TOC dice %d pero esta en la pagina %d"
                              % (titulo_norm[:40], k + 1, pedida, encontradas[k]))
    R.anota("Los numeros de la TOC coinciden con las paginas reales", not desajustes,
            ("%d desajuste(s): %s" % (len(desajustes), "; ".join(desajustes[:4])))
            if desajustes else "%d entradas comprobadas (%d titulos distintos)"
            % (len(toc), len(reclamadas)))

    # --- 10. Marcadores rotos -------------------------------------------
    rotas = []
    for patron in (r"error!\s*bookmark", r"error!\s*reference", r"\?\?\s*$"):
        m = re.search(patron, texto_total, re.I | re.M)
        if m:
            rotas.append(patron)
    # un 0 o un guion suelto tras el leader de puntos tambien es un PAGEREF roto
    for linea in texto_total.splitlines():
        if re.search(r"\.{4,}\s*(0|-|\?)\s*$", linea.strip()):
            rotas.append("TOC con numero de pagina vacio/0")
            break
    R.anota("No hay marcadores ni referencias de pagina rotas", not rotas,
            "; ".join(sorted(set(rotas))) if rotas else "ninguno", critico=False)

    # --- 11. Tipografia --------------------------------------------------
    # Los PDF incrustan las fuentes con prefijo de subconjunto ("BAAAAA+"), asi
    # que hay que quitarlo antes de comparar el nombre.
    fuentes = set()
    for i in range(total):
        for f in doc[i].get_fonts(full=False):
            fuentes.add(f[3] if len(f) > 3 else str(f))

    def nombre_limpio(n):
        n = re.sub(r"^[A-Z]{6}\+", "", n or "")     # prefijo de subconjunto
        return re.sub(r"[^a-z]", "", n.lower())

    raras = [f for f in fuentes if "timesnewroman" not in nombre_limpio(f)]
    R.anota("Solo se usa la fuente tipografica declarada", not raras,
            "fuentes: %s%s" % (", ".join(sorted(fuentes)) if fuentes else "ninguna",
                               ("  -> NO declaradas: " + ", ".join(raras)) if raras else ""))

    # --- 12. Interlineado ------------------------------------------------
    # pymupdf devuelve cada LINEA como un bloque separado, no el parrafo, asi
    # que no se puede medir "alto de caja / numero de lineas".
    #
    # Se mide el AVANCE de linea: la diferencia entre el borde superior de dos
    # lineas consecutivas de la misma columna. Times New Roman 12 avanza ~13,8 pt
    # a espacio simple y ~27,6 pt a doble espacio.
    #   Ojo: no sirve medir el hueco entre cajas (y0_siguiente - y1_anterior):
    #   a doble espaiado ese hueco vale tambien ~13,8 pt (el doble espaciado
    #   empuja la caja, no la hace mas alta), y daria un falso negativo.
    UMBRAL_DOUBLE = 20.0
    avances = []
    for i in range(CUERPO, min(total, CUERPO + 15)):
        bs = [b for b in bloques[i] if len(b) >= 5 and isinstance(b[4], str) and b[4].strip()]
        bs.sort(key=lambda b: b[1])
        for a, b in zip(bs, bs[1:]):
            if abs(a[0] - b[0]) < 3 and 0 < (b[1] - a[1]) < 45:
                avances.append(b[1] - a[1])
    if avances:
        avances.sort()
        mediana = avances[len(avances) // 2]
        R.anota("Interlineado doble en el cuerpo", mediana >= UMBRAL_DOUBLE,
                "avance mediano entre lineas = %.1f pt (doble ~27,6 | simple ~13,8)"
                % mediana, critico=False)
    else:
        R.anota("Interlineado doble en el cuerpo", False,
                "no se pudieron emparejar lineas consecutivas para medir", critico=False)

    # --- 13. Notas de tablas y figuras -----------------------------------
    # Dos niveles de comprobacion, porque no es lo mismo:
    #   a) que el manifiesto declare una nota, y
    #   b) que esa nota aparezca de verdad en el PDF.
    # Solo (a) seria trivial (el parser la rellena por defecto) y dejaria pasar
    # un documento entregado sin notas.
    Mobj = {}
    try:
        with open(args.manifiesto, "r", encoding="utf-8") as fh:
            Mobj = json.load(fh)
    except Exception:
        Mobj = {}

    # Flujo de lectura global excluyendo las paginas de indice: en el indice de
    # tablas cada entrada reza "Tabla N. titulo" y generaria falsos positivos.
    paginas_indice = set()
    for i in range(doc.page_count):
        t = paginas[i]
        if re.search(r"\bIndice de (tablas|figuras)\b", t, re.I) and re.search(r"\.{4,}", t):
            paginas_indice.add(i)

    flujo = []
    for i in range(doc.page_count):
        if i in paginas_indice:
            continue
        for b in bloques[i]:
            if len(b) >= 5 and isinstance(b[4], str) and b[4].strip():
                flujo.append((i + 1, round(b[1], 1), b[4].strip()))
    flujo.sort(key=lambda x: (x[0], x[1]))

    sin_nota = []
    sin_nota_en_pdf = []
    pos = 0
    for t in Mobj.get("tablas", []):
        etiqueta = "Tabla %d" % t["indice"]
        nota = t.get("nota")
        if not nota:
            sin_nota.append(etiqueta)
            continue
        # La nota debe aparecer DESPUES de la etiqueta de su tabla en el flujo y
        # ANTES de la siguiente tabla: asi queda debajo de la tabla (APA 7).
        ini = None
        for k in range(pos, len(flujo)):
            if re.match(r"^%s\b" % re.escape(etiqueta), flujo[k][2]):
                ini = k
                break
        if ini is None:
            sin_nota_en_pdf.append("%s: no se encontro la etiqueta" % etiqueta)
            continue
        needle = norm(nota)[:40]
        enc = None
        for k in range(ini, min(len(flujo), ini + 60)):
            if re.match(r"^Nota\.\s", flujo[k][2]) and norm(flujo[k][2]).find(needle) >= 0:
                enc = k
                break
        if enc is None:
            sin_nota_en_pdf.append("%s: la nota no aparece bajo la tabla" % etiqueta)
        else:
            pos = enc + 1
    R.anota("Las tablas llevan nota debajo (en el PDF, no solo en el manifiesto)",
            not (sin_nota or sin_nota_en_pdf),
            "; ".join(sin_nota + sin_nota_en_pdf) if (sin_nota or sin_nota_en_pdf)
            else "%d tabla(s) con nota verificada bajo la tabla" % len(Mobj.get("tablas", [])))

    # Las figuras se comprueban por geometria mas abajo (nota encima de la
    # imagen). Aqui solo se avisa si el manifiesto no declara nota.
    fig_sin_nota = [f["indice"] for f in Mobj.get("figuras", []) if not f.get("nota")]
    if fig_sin_nota:
        R.anota("El manifiesto declara nota para cada figura", False,
                "sin nota: Figura %s" % ", ".join(str(i) for i in fig_sin_nota))
    else:
        R.anota("El manifiesto declara nota para cada figura", True,
                "%d figura(s)" % len(Mobj.get("figuras", [])))

    # --- 14. Posicion de la nota de figura (encima de la imagen) ----------
    # Buscar la nota por texto NO sirve: varias figuras suelen compartir la
    # misma nota por defecto ("Elaboracion propia"), asi que una busqueda
    # textual devuelve siempre la primera y marca como error a todas las
    # demas. Hay que comparar la GEOMETRIA: la nota general de la figura es el
    # bloque de texto "Nota." que queda por ENCIMA (menor y) de la imagen, en la
    # misma pagina.
    sin_nota_encima = []
    paginas_con_foto = 0
    for i in range(CUERPO, total):
        rects_imagen = []
        for img in doc[i].get_images(full=True):
            rects_imagen.extend(doc[i].get_image_rects(img[0]))
        if not rects_imagen:
            continue
        paginas_con_foto += 1
        # Bloques de texto de la pagina que empiezan por "Nota.". Ojo: hay que
        # mirar el texto CRUDO, no norm(), porque norm() quita la puntuacion y
        # "Nota." se volveria "nota" (y el prefijo con punto no casaria nunca).
        bs = [b for b in bloques[i] if len(b) >= 5 and isinstance(b[4], str) and b[4].strip()]
        notas = [b for b in bs if re.match(r"^Nota\.\s", b[4].strip())]
        for r in rects_imagen:
            # El borde inferior del bloque de nota debe quedar por encima (o
            # justo en el borde superior) de la imagen.
            if not [b for b in notas if b[3] <= r.y0 + 4]:
                sin_nota_encima.append("pag. %d: imagen en y0=%.0f sin nota encima"
                                       % (i + 1, r.y0))

    if paginas_con_foto == 0:
        R.anota("Las notas de figura van encima de la imagen", True,
                "el documento no tiene imagenes: no aplica", critico=False)
    elif sin_nota_encima:
        R.anota("Las notas de figura van encima de la imagen", False,
                "%d imagen(es) sin nota encima: %s"
                % (len(sin_nota_encima), "; ".join(sin_nota_encima[:3])), critico=False)
    else:
        R.anota("Las notas de figura van encima de la imagen", True,
                "%d pagina(s) con imagen, todas con nota encima" % paginas_con_foto)

    return R, total


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description="Verifica el PDF generado contra su manifiesto")
    ap.add_argument("--pdf", required=True)
    ap.add_argument("--manifiesto", required=True)
    ap.add_argument("--log")
    ap.add_argument("--json")
    args = ap.parse_args()

    lineas = []

    def w(s=""):
        print(s)
        lineas.append(s)

    if not os.path.isfile(args.pdf):
        w("FALLO: no existe el PDF: %s" % args.pdf)
        return 1

    w("=== FASE 4: verificacion del PDF ===")
    w("[%s] PDF        : %s" % (datetime.now().strftime("%H:%M:%S"), args.pdf))
    w("[%s] Manifiesto : %s" % (datetime.now().strftime("%H:%M:%S"), args.manifiesto))
    w("")

    R, total = verificar(args)

    for it in R.items:
        w("  %-6s %-52s %s" % ("OK" if it["ok"] else ("FALLA" if it["critico"] else "AVISO"),
                                it["nombre"], it["detalle"]))
    w("")

    criticas = len(R.fallas)
    no_criticas = len(R.avisos)
    w("Comprobaciones: %d   OK: %d   FALLAS: %d   AVISOS: %d"
      % (len(R.items), len(R.items) - criticas - no_criticas, criticas, no_criticas))
    w("Paginas: %d" % total)
    w("")
    if criticas:
        w("RESULTADO: FALLA (%d comprobacion(es) critica(s))" % criticas)
    else:
        w("RESULTADO: OK")

    if args.log:
        d = os.path.dirname(os.path.abspath(args.log))
        if d and not os.path.isdir(d):
            os.makedirs(d, exist_ok=True)
        with open(args.log, "w", encoding="utf-8") as fh:
            fh.write("\n".join(lineas) + "\n")
        w("Log escrito: %s" % args.log)

    if args.json:
        d = os.path.dirname(os.path.abspath(args.json))
        if d and not os.path.isdir(d):
            os.makedirs(d, exist_ok=True)
        with open(args.json, "w", encoding="utf-8") as fh:
            json.dump({
                "pdf": os.path.abspath(args.pdf),
                "paginas": total,
                "resultado": "FALLA" if criticas else "OK",
                "fallas_criticas": criticas,
                "avisos": no_criticas,
                "comprobaciones": R.items,
            }, fh, ensure_ascii=False, indent=2)
        w("Informe JSON: %s" % args.json)

    return 1 if criticas else 0


if __name__ == "__main__":
    sys.exit(main())
