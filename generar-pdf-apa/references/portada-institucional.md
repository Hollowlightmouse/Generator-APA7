# Portada institucional

Disposicion de la primera pagina de un documento academico. La portada es una
**plantilla con marcadores**, no un texto fijo: se rellena con lo que traiga el
documento y con lo que responda el usuario. No lleva el nombre de ninguna
institucion, asignatura ni persona.

## Organizacion posicional en 3 zonas

Pagina carta (612 x 792 pt) con margenes de 72 pt (1 pulgada). El contenido NO
se apila consecutivamente: se reparte en tres zonas verticales. Las posiciones
se miden desde el borde superior de la hoja y las verifica `verificar-pdf.py`
como **aviso**, nunca como falla.

El alto util (792 - 72 - 72 = 648 pt) se parte en tres bandas iguales de 216 pt:

```
   y= 36   Numero de pagina (header, esquina superior derecha)
   y= 72  +----------------------------------------------------+  <- margen
           |  [ZONA ALTA]   72 .. 288 pt                       |
   y=108   |     Logo opcional (160x160 px) encima del titulo   |
           |     Titulo del trabajo (negrita, centrado)        |
   y=288  +----------------------------------------------------+
           |  [ZONA CENTRO]  288 .. 504 pt                     |
   y=383   |     Integrantes: UN parrafo centrado, vertical      |
           |     centrado en la banda (y ~383 con 1 linea)      |
   y=504  +----------------------------------------------------+
           |  [ZONA BAJA]    504 .. 720 pt                     |
   y=555   |     Docente (+ titulo/profesion)                   |
           |     Materia y NRC                                 |
   y=583   |     Facultad / universidad                        |
   y=610   |     Vicerrectoria                                 |
           |     Sede                                         |
   y=693   |     Fecha          <- ultima linea, siempre a ~706 |
   y=720  +----------------------------------------------------+
```

**Regla de 3 zonas (no negociable):**

1. **Arriba** — logo opcional y titulo del trabajo. El logo va SIEMPRE encima
   del titulo, nunca al lado.
2. **Centro vertical** — integrantes, unico elemento en el medio de la pagina.
3. **Abajo** — el resto de los datos institucionales, con un pitch de unos
   27,6 pt entre lineas y **anclados por abajo**: la ultima linea termina siempre
   cerca de y=706, en vez de empezar en una cota fija. Asi el bloque no se mueve
   segun quantos campos traiga el manifiesto.

El orden del bloque inferior es el de siempre y depende de la informacion
disponible: docente, materia/NRC, facultad, vicerrectoria, sede, fecha. Un campo
que no exista simplemente no se imprime: no deja huecos en blanco.

### Como se implementa

En `build-docx.js`, `construirPortada()` devuelve **una tabla de 1 columna x 3
filas sin bordes visibles**, con `HeightRule.ATLEAST` y `verticalAlign`
`TOP` / `CENTER` / `BOTTOM` respectivamente. La tabla es solo un mecanismo de
posicionamiento: no lleva bordes, no se imprime y el verificador la ignora.

Dos detalles que no son negociables:

- El aire de la zona alta (720 twips) se **descuenta** de la altura de su fila.
  Las tres filas tienen que sumar justo 12960 twips; si suman mas, la ultima fila
  se desborda a una pagina nueva y el documento gana una pagina.
- El parrafo del logo va a interlineado **simple**. Con doble, la caja de la
  imagen se infla y empuja el resto de la portada.

Antes esto se hacia apilando parrafos con lineas vacias de relleno, lo que
dejaba a los integrantes pegados al titulo y el bloque institucional a media
pagina. No volver a esa forma: las zonas no se sostienen si el contenido cambia.


## Campos

| Campo | Marcador | Regla |
|---|---|---|
| Logo | `logo` | Opcional. Solo se inserta si el usuario aporta el archivo. Si no, no se inserta nada: sin placeholder, sin marco. Ancho maximo ~2 in, altura proporcional. **Preguntar siempre.** |
| Numero de pagina | — | Superior derecha, igual que el resto del documento. La portada **no lleva** numero visible. |
| Titulo | `titulo` | Negrita, centrado. Puede ocupar varias lineas. |
| Autores | `autores` | **Un solo parrafo centrado**, sin negrita. Nombres completos separados por coma y con "y" antes del ultimo, sin importar cuantos sean. No uno por linea. Acepta `"Nombre Apellido y Nombre Apellido"` o `["Nombre Apellido", "Nombre Apellido"]`: se respeta el texto tal cual, sin partirlo ni suponer autores. |
| Facultad y universidad | `facultad` | Una linea centrada, tal como aparezca en el documento fuente. |
| Sede o unidad | `vicerrectoria` | Centrada, opcional. |
| Materia y codigo | `materia_nrc` | Centrada, tal como aparezca en el documento. El campo admite un codigo de asignatura cualquiera (por ejemplo un NRC), no solo ese formato. |
| Docente | `docente` | Centrado. Se pregunta siempre por su titulo o profesion (`docente_titulo`), salvo que el usuario diga que esa vez se omite. |
| Fecha | `fecha` | Centrada, en el formato del documento. |

## Que hacer cuando falta un dato

Regla general de la skill, no solo de la portada:

- **Ante cualquier valor que no aparezca en los archivos de entrada, preguntar.**
  Nunca inferir, asumir ni inventar. Aplica a todos los campos de la portada
  y a cualquier otro dato del documento (fuente de una tabla o figura, titulo
  de seccion, autores, fecha, etc.).
- Si el usuario pide omitir un campo, **no** dejar linea en blanco ni
  placeholder: simplemente no se incluye esa linea. Para el logo, no se
  inserta nada.
- El titulo o profesion del docente es un caso aparte: **se pregunta siempre**,
  aunque el resto de la portada este completa.

## Como se rellena en la practica

El parser detecta por patron los campos que suele traer el `.md` de un trabajo
academico (facultad, sede, materia con codigo, docente, autores, fecha) y los
pone en `portada` dentro de `MANIFEST.json`. Lo que no detecte aparece en
`campos_faltantes` y hay que preguntarlo.

Las respuestas del usuario se pasan en un `portada.json` que se entrega con
`--portada`, y tienen prioridad sobre lo deducido. El generador nunca inventa
un nombre de institucion: si `facultad` no viene, no imprime ninguna linea.
