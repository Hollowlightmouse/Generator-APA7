# Reglas de formato - Normas APA, 7a edicion

Reglas generales de la 7a edicion del manual APA, en el formato de trabajo de
estudiante. Este archivo no anade requisitos institucionales: si la institucion
donde se entrega pide algo puntual que no este aqui, se actualiza el archivo con
lo que confirme el usuario y no se asume nada mas.

## Pagina

- Tamano de papel: **Carta** (Letter, 21,59 x 27,94 cm / 8,5 x 11 in).
- Margenes: **2,54 cm (1 pulgada)** en los cuatro lados.
- Numero de pagina en la esquina **superior derecha**, en el encabezado.
- **La portada no lleva numero visible.** La numeracion arranca en la pagina 2,
  que es lo que exige APA 7. La portada cuenta como pagina 1 para el resto de
  calculos, pero no se imprime su numero.

## Tipografia y espaciado

- Fuente: Times New Roman 12 pt en todo el documento, incluida la portada, la
  tabla de contenido, los indices, las leyendas de tablas y figuras y las
  referencias.
- Interlineado **doble** en todo el documento.
- El cuerpo va alineado a la **izquierda**, no justificado, con sangria de
  primera linea de 1,27 cm (0,5 in) en cada parrafo.
- Las entradas de la tabla de contenido van en Times New Roman 12 pt y **sin
  negrita**, aunque el titulo de seccion si la lleve.

## Niveles de encabezado

| Nivel | Formato |
|---|---|
| 1 | Centrado, negrita, con mayusculas y minusculas normales |
| 2 | Alineado a la izquierda, negrita |
| 3 | Alineado a la izquierda, negrita y cursiva |
| 4 | Con sangria de primera linea, negrita, termina en punto. El texto sigue en la misma linea |
| 5 | Con sangria de primera linea, negrita y cursiva, termina en punto. El texto sigue en la misma linea |

Los titulos de los niveles 1 y 2 van en **negro, negrita y Times New Roman
12 pt**. Se aplica a nivel de run (`bold: true, color: '000000', font: 'Times
New Roman', size: 24`) porque los estilos de encabezado de la libreria `docx`
traen color y negrita propios que pisan los del documento.

Los encabezados deben aplicarse siempre como **estilos reales** (Heading 1/2/3,
etc.), nunca solo como texto en negrita: de eso depende que la tabla de
contenido se construya. Ver `campos-word-toc.md`.

### Como se mapea desde el `.md`

Este documento usa `##` para los dos primeros niveles, como el resto de la
documentacion de la skill:

| En el `.md` | Nivel APA |
|---|---|
| `#` | (portada, no es encabezado de cuerpo) |
| `## Numero. Texto` | 1 (centrado, negrita, pagina nueva) |
| `## Texto` | 2 (izquierda, negrita) |
| `###` | 3 (izquierda, negrita + cursiva) |
| `####` | 4 (sangria, negrita, punto final, texto en la misma linea) |
| `#####` | 5 (sangria, negrita + cursiva, punto final, texto en la misma linea) |

Los niveles 3, 4 y 5 **no se aplanan**: cada hash conserva su nivel. Markdown
no puede expresar "titulo en linea con el cuerpo continuando", asi que un
`####` o `#####` sale como parrafo propio. Para el caso autentico de APA
(titulo en linea con sangria) el parser trae `--detectar-niveles-en-linea`,
que reconoce la convencion:

- `**Titulo.** El texto sigue en la misma linea.` → nivel 4
- `***Titulo.*** El texto sigue en la misma linea.` → nivel 5

Va **opt-in a proposito**: la negrita al abrir un parrafo tambien es texto
normal corriente, y activarlo sin control llenaria la tabla de contenido de
titulos falsos. El parseo se hace sobre el texto crudo, antes de
`strip_leading_markers()`, que se come los asteriscos iniciales.

## Tablas o figuras muy anchas: pagina horizontal

APA 7 permite poner esa pagina puntual en horizontal dentro de un documento que
en lo demas va en vertical. Al hacerlo:

- Los margenes de 2,54 cm se mantienen.
- El numero de pagina se sigue mostrando, reubicado para la nueva orientacion.
- Se usa un salto de **seccion**, no solo de pagina, antes y despues, para poder
  cambiar la orientacion sin afectar al resto del documento.
- La tabla o figura debe caber dentro de los margenes horizontales.
- La tabla ancha, su titulo y su nota van **en la misma pagina horizontal**, y el
  texto que continua despues vuelve a vertical en otra seccion.
- `build-docx.js` decide esto solo, por legibilidad y no por numero de filas: una
  tabla de 200 filas y 3 columnas puede seguir en vertical sin problema. Se pasa a
  horizontal cuando el ancho no basta:
  - **muchas columnas**: 6 o mas (`UMBRAL_COLUMNAS_ANCHAS`), o
  - **celdas tan largas que no caben**, siempre que girar la pagina las alivie de
    verdad (la celda mas larga pasa de 8 lineas o mas en vertical a menos de 8 en
    horizontal, `UMBRAL_LINEAS_APRETADAS`).

  Los dos casos se miden sobre el ancho util de cada pagina (6,5" en vertical y 9"
  en horizontal) y no sobre una regla de filas. Una tabla de 3 columnas con celdas
  larguisimas **no** se gira: pasar de 24 a 34 caracteres por linea no la hace
  legible, y conviene partirla o mover el contenido a texto.
- Aviso tecnico: al declarar la seccion horizontal hay que dar a `docx` las
  medidas **en vertical** (12240 x 15840) y marcar `orientation: LANDSCAPE`. La
  libreria las intercambia por su cuenta; darlas ya invertidas produce un
  `w:pgSz` con `w` menor que `h`, que Word y LibreOffice leen como vertical y
  deja la tabla apretada.

## Tablas y figuras

- Numeracion independiente: "Tabla 1", "Tabla 2"... y "Figura 1", "Figura 2"...
- El numero y el titulo van en lineas separadas, en negrita el numero y en
  cursiva el titulo.
- **El titulo y la leyenda no son opcionales.** Si el `.md` no los trae, el
  parser deja una pregunta bloqueante por cada uno y `build-docx.js` aborta
  con codigo 4. No existe un `(sin titulo)` que se imprima en su lugar.
- Nota en cursiva solo la palabra **"Nota."**, seguida del texto normal. Por
  defecto **"Nota. Elaboración propia."**, salvo que el `.md` indique otra
  fuente; el parser la rellena sola (`nota_origen: "default"`) y con
  `--notas-tabla-json` / `--notas-figura-json` se sustituye una en concreto.
- **La posicion si distingue el caso, y no es negociable:**
  - **Nota de FIGURA: encima de la imagen.**
  - **Nota de TABLA: debajo de la tabla.**
- La nota de figura lleva `keepNext` para que no se quede sola al final de una
  pagina mientras su imagen salta a la siguiente. Sin eso el documento queda
  incoherente aunque el orden en el flujo sea el correcto.
- Bordes de tabla: **solo tres lineas horizontales** - el borde superior de la
  tabla, la linea que separa la fila de encabezado del cuerpo y el borde inferior.
  **Sin** lineas verticales, sin caja lateral y **sin** lineas entre las filas de
  datos: esas separaciones las hace el interlineado, no un borde. La fila de
  encabezado puede llevar sombreado claro.

## Listas

Cuando una linea empiece con dos o mas marcadores ("•  •", "– •", "• ••") o con
numeracion duplicada ("1. 1.", "4. 1."), se normaliza a un solo marcador o un
solo numero. Son artefactos de la conversion y no cumplen APA.

## Palabras pegadas (artefactos de extraccion)

Convertir un PDF a texto suele perder espacios internos. El parser los corrige
automaticamente, pero conviene revisarlos al leer el `.md`.

**Casos que se separan:**

1. Minuscula seguida de mayuscula (`[a-z][A-Z]`): `Diferenciacionentre` ->
   `Diferenciacion entre`, `yFlaws` -> `y Flaws`.
2. Dos puntos pegados a una mayuscula (`:[A-Z]`): `2:Diferenciacion` ->
   `2: Diferenciacion`.
3. Numero pegado a palabra: `ACTIVIDAD1` -> `ACTIVIDAD 1`.

**Casos que NO se separan:**

- Siglas en mayusculas puras (`OWASP`, `CVSS`, `BPMN`): la regla exige minuscula
  seguida de mayuscula, asi que ya son seguras.
- Nombres con camelCase (`NodeJS`, `OpenID`, `MySQL`) y los que mezclan letra y
  digito (`IPv6`, `SHA256`). Estos **no** se pueden distinguir por patron, por
  eso van en `terminos-whitelist.txt`, que se pasa con `--whitelist`.
- Apellidos con prefijo (`McDonald`, `MacArthur`): anadir el apellido concreto a
  `terminos-whitelist.txt` si se rompe.

El deglue nunca es silencioso: **todo lo que cambia queda registrado** y se
puede revisar antes de construir el `.docx`. Si se partio un termino que no
debia, se anade al fichero de lista blanca y se vuelve a ejecutar.

## Referencias

- Encabezado de la seccion: "Referencias" (nivel 1, centrado, negrita).
- Orden alfabetico por apellido del primer autor.
- **Sangria francesa**: primera linea sin sangria, lineas siguientes con 1,27 cm
  (0,5 in) de sangria.
- Interlineado doble, sin espacio extra entre referencias.
- Titulos de libros e informes en cursiva; titulos de articulos sin cursiva
  (la revista si va en cursiva).
- Si una referencia del `.md` no sigue el formato (falta la cursiva, falta el ano
  entre parentesis, el orden es incorrecto), se corrige antes de incluirla.
