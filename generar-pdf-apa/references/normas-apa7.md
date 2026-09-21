# Reglas de formato — Normas APA, 7ª edición

Nota: estas son las reglas generales y ampliamente documentadas de la 7ª edición del manual APA (formato para trabajos de estudiante). El usuario indicó que Uniminuto no pide nada adicional a APA 7, solo tamaño carta y los márgenes de la normativa. Si en algún momento aparece un requisito institucional puntual que no esté aquí, este archivo debe actualizarse con lo que el usuario confirme — no se debe asumir nada más allá de lo que sigue.

## Página

- Tamaño de papel: **Carta** (Letter, 21.59 cm × 27.94 cm / 8.5 in × 11 in).
- Márgenes: **2.54 cm (1 pulgada)** en los cuatro lados.
- Numeración de página en la esquina **superior derecha**, en todas las páginas (incluida la portada).

## Tipografía y espaciado

- Fuente recomendada (elegir una y mantenerla en todo el documento): Times New Roman 12 pt (la más común en trabajos de Uniminuto), Arial 11 pt, Calibri 11 pt o Georgia 11 pt.
- Interlineado **doble** en todo el documento, incluidas portada, citas en bloque, tablas/figuras (leyendas), la **tabla de contenido** y referencias. En la práctica, el TOC generado por Word nace con interlineado sencillo: se aplica doble a sus entradas por COM tras el `Update()` (ver `campos-word-toc.md`).
- **Entradas de la Tabla de contenido (estándar del proyecto):** interlineado doble, **Times New Roman 12 pt y SIN negrilla** (los títulos de sección sí van en negrita, pero sus entradas en el TOC no). Word copia el negrita del heading al resultado del campo; se quita por COM antes de guardar.
- Texto del cuerpo alineado a la **izquierda** (no justificado), con sangría de primera línea de 1.27 cm (0.5 in) en cada párrafo.

## Niveles de encabezado

| Nivel | Formato |
|---|---|
| 1 | Centrado, negrita, con mayúsculas y minúsculas normales |
| 2 | Alineado a la izquierda, negrita |
| 3 | Alineado a la izquierda, negrita y cursiva |
| 4 | Con sangría de primera línea, negrita, termina en punto. El texto sigue en la misma línea |
| 5 | Con sangría de primera línea, negrita y cursiva, termina en punto. El texto sigue en la misma línea |

**Estándar de títulos (definido en el proyecto):** todos los títulos (niveles 1 y 2) en **color negro, negrilla y Times New Roman 12 pt**, además de su alineación según nivel. Aplícalo a nivel de run (`bold: true, color: '000000', font: 'Times New Roman', size: 24`) porque los estilos de encabezado de la librería `docx` traen color/negrita propios que sobreescriben el documento.

Estos encabezados deben aplicarse siempre como **estilos reales de Word** (Heading 1/2/3, etc.), nunca solo como texto en negrita — de eso depende que la tabla de contenido funcione (ver `campos-word-toc.md`).

## Tablas o figuras muy anchas → página horizontal

Si una tabla o figura es demasiado ancha para verse bien en vertical, está permitido por APA 7 poner **esa página puntual en horizontal** dentro de un documento que en todo lo demás va en vertical (fuente: apastyle.apa.org, sección "Table setup"). Al hacerlo:

- Los márgenes de 2.54 cm se mantienen igual en la página horizontal.
- El número de página se sigue mostrando, reubicado para la nueva orientación (no se omite).
- Usar un salto de **sección** (no solo de página) antes y después de la página horizontal, para poder cambiar su orientación sin afectar el resto del documento.
- La tabla/figura debe seguir dentro de los márgenes (no puede desbordarse del ancho horizontal disponible).

## Tablas y figuras

- Numeración independiente: "Tabla 1", "Tabla 2"... y "Figura 1", "Figura 2"...
- El número y el título van en líneas separadas, en negrita el número, cursiva el título.
- Nota debajo (fuente/autoría): en cursiva la palabra "Nota.", seguida del texto normal. Por defecto: "Nota. Elaboración propia." salvo que el `.md` indique otra fuente.
- **Bordes de tablas (estándar del proyecto/APA):** solo líneas horizontales — superior, inferior y las internas que separan filas (`insideHorizontal`). Sin líneas verticales ni caja lateral. La fila de encabezado puede llevar sombreado claro.

## Listas

- Cuando una línea de la fuente inicie con dos o más marcadores de lista seguidos ("•  •", "– •", "• ••") o con numeración duplicada ("1. 1.", "4. 1."), normalizarla a **un solo marcador** ("• ") o un solo número, sin duplicados — artefactos de conversión que no cumplen APA.

## Palabras pegadas (artefactos de extracción)

La conversión del PDF a texto suele perder espacios internos y dejar palabras pegadas. Casos reales vistos:

- `EJERCICIO 2:Diferenciaciónentre Bugs yFlaws` → debe quedar `EJERCICIO 2: Diferenciación entre Bugs y Flaws`.
- `ACTIVIDAD1:MapeodeControlesISO27002` → `ACTIVIDAD 1: Mapeo de Controles ISO 27002`.
- `SoftwareyDatos`, `MatrizArgumentativade Controles`, `Aplicación (AppSec)estructurados` → separar las palabras.

**Regla:** al leer cada encabezado y cada párrafo del `.md`, detectar y corregir:

1. **Minúscula→mayúscula pegadas** (`[a-zñáéíóú][A-ZÁÉÍÓÚ]`): insertar un espacio en el límite. Ej.: `Diferenciaciónentre` → `Diferenciación entre`, `yFlaws` → `y Flaws`.
2. **Siguiente palabra en mayúscula tras dos puntos** (`:[A-ZÁÉÍÓÚ]` sin espacio): insertar espacio tras `:`. Ej.: `2:Diferenciación` → `2: Diferenciación`.
3. **Números pegados a palabras** (`ACTIVIDAD1:` → `ACTIVIDAD 1:`): separar el número de la palabra.

**Cuidado — no separar nombres propios o términos técnicos legítimos** (el mismo límite camelCase es válido en estos): `AppSec`, `DevSecOps`, `FinTechX`, `MediCloud`, `EcoPay`, `ISO27002`/`ISO/IEC 27002`, `SSG`, `SDLC`, `BSIMM`, `OWASP`. Usar una lista blanca de términos conocidos del documento antes de aplicar la separación automática, o corregir a mano los casos no ambiguos. Ejemplo de normalización (JS, aplicable a `bodyParagraph` y a los textos de encabezado):

```js
// Heurística con lista blanca: si el contexto de la palabra pegada contiene un
// término técnico conocido, NO la separa; en caso contrario inserta el espacio.
const GLUE_RE = /([a-zñáéíóú])([A-ZÁÉÍÓÚ][a-zñáéíóúÁÉÍÓÚ]*)/g;
function deglue(text, whitelist = ['AppSec', 'DevSecOps', 'FinTechX', 'MediCloud', 'EcoPay', 'ISO27002', 'BSIMM']) {
  return text.replace(GLUE_RE, (m, a, b) => {
    const i = text.indexOf(m);
    const around = text.slice(Math.max(0, i - 12), i + 12);
    return whitelist.some(w => around.includes(w)) ? m : a + ' ' + b;
  });
}
```

(Si la heurística automática se complica por nombres propios, la opción segura es corregir a mano los pocos casos que aparecen en el documento antes de construir el `.docx`.)

## Referencias

- Encabezado de la sección: "Referencias" (nivel 1, centrado, negrita).
- Orden alfabético por apellido del primer autor.
- **Sangría francesa**: primera línea sin sangría, líneas siguientes con 1.27 cm (0.5 in) de sangría.
- Interlineado doble, sin espacio extra entre referencias.
- Títulos de libros/informes en cursiva; títulos de artículos sin cursiva (la revista sí va en cursiva).
- Si una referencia del `.md` no sigue este formato (falta cursiva, falta año entre paréntesis, orden incorrecto, etc.), corregirla antes de incluirla.
