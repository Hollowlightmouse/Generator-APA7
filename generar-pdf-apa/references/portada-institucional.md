# Estructura de portada — estándar de organización (plantilla del proyecto)

**Fuente canónica de la organización:** página 1 del PDF `ARQUITECTURA_X_SEGMED_APA7_corregido.pdf` ("Dossier de Arquitectura de Software: SGEMD"). La portada de TODO documento debe respetar esta **disposición posicional** (en 3 zonas verticales) y este formato exactos. No inventar campos ni reordenar.

## Organización posicional (3 zonas, página carta 612×792 pts, márgenes 72 pts = 1 in)

El diseño NO apila los campos de arriba hacia abajo consecutivamente: distribuye el contenido en **tres zonas verticales** con posiciones objetivo (medidas desde el borde superior de la página, y = pts):

```
                                    Nº de página (superior derecha, header, y≈36)          ← header

        [ZONA ARRIBA]
        Logo opcional (centrado, solo si el usuario lo aporta)                             ← debajo del header


                                          Título del trabajo (negrita, centrado)           ← y≈100
                                          (2 líneas en el canónico; puede ocupar varias)


        [ZONA CENTRO]
                                          Autores (UN solo párrafo centrado)               ← y≈293

        [ZONA ABAJO]
               Facultad de Ingeniería de Sistemas, Corporación Universitaria Minuto de Dios  ← y≈514
               Vicerrectoría Regional Centro Sur, Sede Ibagué, Tolima                       ← y≈542
               Materia NRC (Ej.: Arquitectura de Software NRC 60-95404)                     ← y≈569
               Nombre del docente                                                           ← y≈597
               Fecha (Ej.: 16 de septiembre de 2026)                                        ← y≈625

                                                                                             (pitch del bloque
                                                                                              inferior ≈27.6 pts)
```

**Regla de 3 zonas (no negociable):**
1. **Arriba** → logo opcional (solo si se aporta) y **título** del trabajo.
2. **Centro vertical** → **autores** (único elemento en el medio de la página).
3. **Abajo** → todo el resto de la información institucional (facultad, vicerrectoría, materia-NRC, docente, fecha), consecutivo con pitch entre líneas ≈27.6 pts.

## Reglas por elemento

1. **Logo (opcional)** — centrado, arriba del título, **solo si el usuario lo aporta** (el PDF canónico no lleva logo). Si no se aporta, **no se inserta ninguno** (sin placeholder). Si se aporta: ancho máximo ≈2 in (≈192 px a 96 dpi), altura proporcional a la proporción real de la imagen. **Siempre preguntar** si quiere aportarse un logo.
2. **Número de página** — superior derecha, igual al resto del documento (header del docx).
3. **Título del trabajo** — negrita, centrado; puede ocupar varias líneas.
4. **Autores** — en **un solo párrafo centrado**, sin negrita. Formato: nombres completos separados por coma y **"y" antes del último**, sin importar la cantidad. Ejemplo canónico (3 autores): "Nombre Uno Apellido, Nombre Dos Apellido y Nombre Tres Apellido". NO usar uno por línea.
5. **Facultad + Universidad** — una sola línea, centrado: "Facultad de Ingeniería de Sistemas, Corporación Universitaria Minuto de Dios".
6. **Vicerrectoría + Sede** — centrado: "Vicerrectoría Regional Centro Sur, Sede Ibagué, Tolima".
7. **Materia + NRC** — centrado, tal como aparece en el documento fuente (puede traer prefijo de facultad, p.ej. "Arquitectura de Software NRC 60-95404").
8. **Nombre del docente** — centrado. **Siempre preguntar** al usuario por el título/estudio profesional del docente para incluirlo, salvo que indique explícitamente que se omita esa vez.
9. **Fecha** — centrado, en el formato del documento (día de mes de año), según datos del `.md` o lo que confirme el usuario.

## Qué hacer cuando falta un dato (regla general de la skill)

- **Ante cualquier valor faltante que no se encuentre en el análisis del `.md`, preguntar** — nunca inferir, asumir ni inventar. Esto aplica a TODOS los campos de la portada (incluido el logo) **y a cualquier dato faltante del documento en general** (fuente de una tabla/figura, NRC, título de una sección, autores, fecha, etc.).
- Si el usuario pide omitir un campo, no dejar un espacio vacío ni un placeholder — simplemente no incluir esa línea (y si es el logo, no insertar nada).
- El título/profesión del docente es un caso aparte: **siempre se pregunta**, incluso si el resto de la portada está completa.