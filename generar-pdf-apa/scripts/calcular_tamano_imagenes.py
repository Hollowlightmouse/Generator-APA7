#!/usr/bin/env python3
"""
Calcula el tamaño (en pulgadas) que debe tener cada imagen al insertarla en el
.docx, preservando su proporción real, a partir del JSON de layout generado
junto con el .md (bbox de cada imagen/tabla en la página).

Uso:
    python3 calcular_tamano_imagenes.py <archivo.json> [--ancho-contenido 6.5]

Acepta tanto el formato "*_model.json" (bbox normalizado 0-1 por página, un
item por página con "type": "image") como el formato "*_content_list.json"
(bbox en píxeles absolutos, con "page_idx").

--ancho-contenido es el ancho disponible de contenido en pulgadas (página
carta 8.5in - 1in de margen a cada lado = 6.5in por defecto, según
references/normas-apa7.md).

Salida: una línea JSON por imagen con:
    img_path, page, ancho_in, alto_in
El ancho siempre queda acotado a --ancho-contenido; el alto se deriva de la
proporción real de la imagen en la página (no de su tamaño en píxeles del
archivo, que puede no coincidir con cómo se veía en el documento original).
"""
import json
import sys
import argparse

PAGE_W_IN = 8.5
PAGE_H_IN = 11.0


def from_model_json(data, ancho_contenido):
    resultados = []
    for page_idx, page_items in enumerate(data):
        for item in page_items:
            if item.get("type") != "image":
                continue
            bbox = item.get("bbox")
            if not bbox or len(bbox) != 4:
                continue
            x0, y0, x1, y1 = bbox
            frac_w = x1 - x0
            frac_h = y1 - y0
            if frac_w <= 0 or frac_h <= 0:
                continue
            real_w_in = frac_w * PAGE_W_IN
            real_h_in = frac_h * PAGE_H_IN
            aspecto = real_h_in / real_w_in
            ancho_final = min(ancho_contenido, real_w_in)
            alto_final = ancho_final * aspecto
            resultados.append({
                "img_path": item.get("content") or item.get("img_path") or f"(sin path, pagina {page_idx})",
                "page": page_idx,
                "ancho_in": round(ancho_final, 2),
                "alto_in": round(alto_final, 2),
            })
    return resultados


def from_content_list(data, ancho_contenido):
    # bbox en píxeles absolutos; no conocemos el tamaño de página en píxeles
    # con certeza para todos los casos, así que usamos la proporción
    # ancho/alto del bbox, que es independiente de la unidad.
    resultados = []
    for item in data:
        if item.get("type") != "image":
            continue
        bbox = item.get("bbox")
        img_path = item.get("img_path")
        if not bbox or len(bbox) != 4 or not img_path:
            continue
        x0, y0, x1, y1 = bbox
        w = x1 - x0
        h = y1 - y0
        if w <= 0 or h <= 0:
            continue
        aspecto = h / w
        ancho_final = ancho_contenido
        alto_final = ancho_final * aspecto
        resultados.append({
            "img_path": img_path,
            "page": item.get("page_idx"),
            "ancho_in": round(ancho_final, 2),
            "alto_in": round(alto_final, 2),
        })
    return resultados


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("json_path")
    parser.add_argument("--ancho-contenido", type=float, default=6.5)
    args = parser.parse_args()

    # utf-8-sig: en Windows Set-Content -Encoding UTF8 y la mayoria de editores
    # guardan con BOM, y json.load aborta con 'Unexpected UTF-8 BOM' sin el.
    with open(args.json_path, "r", encoding="utf-8-sig") as f:
        data = json.load(f)

    # Heurística simple para distinguir el formato:
    # model.json es una lista de listas (una lista de items por página).
    if isinstance(data, list) and data and isinstance(data[0], list):
        resultados = from_model_json(data, args.ancho_contenido)
    elif isinstance(data, list):
        resultados = from_content_list(data, args.ancho_contenido)
    else:
        print("Formato de JSON no reconocido para este script.", file=sys.stderr)
        sys.exit(1)

    for r in resultados:
        print(json.dumps(r, ensure_ascii=False))


if __name__ == "__main__":
    main()
