# Estándar de Presentaciones PowerPoint — REDCO

**Especificación para agentes generadores de presentaciones**
Versión 1.0 · Junio 2026
*Complementa el Estándar de Documentos — REDCO (v1.0). Esta guía está pensada específicamente para el desarrollo de presentaciones en formato PowerPoint (.pptx) — no aplica a Word, PDF ni otro tipo de documento.*

---

## 1. Propósito

Este documento define las condiciones de formato, jerarquía visual y estructura de contenido que debe respetar todo agente o sistema que genere presentaciones en PowerPoint en nombre de REDCO. A diferencia de un informe o documento corrido, una presentación ejecutiva comunica a través de bloques visuales discretos. Por eso, el principio que rige tanto la forma como el contenido de cada diapositiva es la estructura **MECE**, descrita en la sección 2.

---

## 2. Principio rector: estructura MECE

Es la condición fundamental de este estándar. Toda diapositiva que organice información en categorías — frameworks, columnas, tarjetas numeradas, alternativas de decisión, niveles de un estado — debe cumplir dos condiciones simultáneamente:

- **Mutuamente excluyente (ME):** cada categoría cubre un aspecto que ninguna otra categoría cubre. Si dos tarjetas o columnas podrían contener el mismo dato, hay que fusionarlas o redefinir el corte.
- **Colectivamente exhaustivo (CE):** la suma de las categorías cubre el 100% del universo relevante del tema tratado. Si existe un caso, estado o responsable que no calza en ninguna categoría, falta una categoría.

Esta verificación se aplica tanto al **contenido** (qué dice cada categoría) como a la **forma visual** (cuántas tarjetas o columnas se muestran y cómo se nombran). El agente generador debe revisar explícitamente ambas condiciones antes de cerrar cualquier slide con framework — ver checklist en la sección 9.

**Ejemplos de aplicación:**

| Framework | Categorías ME/CE |
| --- | --- |
| Leyenda de estado | Exactamente 3 estados que cubren el 100% de los casos posibles: *en línea con lo esperado / requiere atención puntual / crítico*. No se admite un cuarto estado ambiguo que se superponga con los anteriores (ej. "en revisión"). |
| Plan o cronograma de fases | Fases consecutivas y no superpuestas que cubren el periodo completo de inicio a fin, sin vacíos entre una fase y la siguiente. |
| Alternativas de decisión | Opciones que no se solapan entre sí y que, entre todas, agotan los caminos posibles a tomar — incluyendo, si corresponde, la opción explícita de no actuar. |
| Desglose de un tema en partes | Cada parte cubre una dimensión distinta del tema; ninguna palabra clave del tema queda sin asignar a una parte. |

---

## 3. Formato de slide

| Parámetro | Valor |
| --- | --- |
| Relación de aspecto | 16:9 |
| Dimensiones | 13.333 in × 7.5 in (33.87 × 19.05 cm) |
| Margen de seguridad | 0.5 in en los cuatro bordes |
| Separación mínima entre bloques de contenido | 0.3 in |

---

## 4. Tipografía

### 4.1 Familia tipográfica

**Lexend Deca** — confirmada para presentaciones REDCO, consistente con la tipografía ya definida en el Estándar de Documentos REDCO.

> ⚠️ **Nota técnica de implementación:** Lexend Deca no viene preinstalada en PowerPoint ni en LibreOffice. Si el archivo `.ttf` no puede incrustarse (embed) dentro del `.pptx`, el agente debe declarar igualmente Lexend Deca en el tema, pero dejar configurado un fallback explícito a una fuente segura (Calibri o Arial) para evitar sustituciones erráticas en equipos donde la fuente no esté instalada.
> 

### 4.2 Escala tipográfica

| Elemento | Tamaño | Peso |
| --- | --- | --- |
| Título de portada / cierre | 40–44 pt | Bold |
| Título de slide de contenido | 26–28 pt | Bold |
| Frase de síntesis (bajo el título, ver 7.3) | 13–14 pt | Regular, itálica |
| Encabezado de categoría dentro de slide | 13–14 pt | Bold |
| Cuerpo / bullets | 11–13 pt | Regular |
| Etiquetas de campo (labels en mayúscula) | 8–9 pt | Bold |
| Notas, pies de página, fuente | 8–9 pt | Regular |

---

## 5. Paleta de colores

Paleta ejecutiva de alto nivel, dominada por azul marino profundo (regla de dominancia: un color cubre 60–70% del peso visual; el resto son tonos de apoyo y un acento puntual).

| Nombre | HEX | Uso |
| --- | --- | --- |
| **Azul medianoche** (primario, dominante) | `#1E2761` | Fondos de portada y cierre, títulos de sección, fondo de tarjetas numeradas, franjas de cabecera |
| **Azul acero** (secundario) | `#2C4A7C` | Sub-bloques dentro de tarjetas, numeración de categorías, íconos |
| **Celeste hielo** (acento claro) | `#CADCFC` | Texto sobre fondo azul medianoche, líneas divisorias sutiles, chips informativos |
| Blanco | `#FFFFFF` | Fondo de slides de contenido, texto sobre azul |
| **Tinta** (texto cuerpo) | `#232733` | Texto principal sobre fondo blanco |
| Gris medio | `#6B7280` | Metadatos, labels, notas |
| Gris línea | `#E2E5EA` | Bordes de tarjetas y tablas |
| Gris fondo | `#F4F6F9` | Fondo de tarjetas sobre slide blanco |

### Colores de estado (uso exclusivo en leyendas de estado MECE)

Estos tres colores son la única excepción a la dominancia del azul. Su uso está reservado **únicamente** a indicadores de estado dentro de tablas de seguimiento — nunca como acento decorativo general de una slide:

| Estado | HEX | Significado |
| --- | --- | --- |
| Verde institucional | `#2E7D5B` | En línea con lo esperado |
| Ámbar | `#C8862D` | Requiere atención puntual |
| **Rojo REDCO** | `#C8102E` | Crítico / requiere intervención |

> El rojo es el mismo definido como color de marca en el Estándar de Documentos REDCO. Aquí se usa exclusivamente para el estado más urgente — no como acento general — de modo que conserve su peso de alerta y no compita con el azul como color dominante de la presentación.
> 

---

## 6. Logo

La presentación debe **reservar el espacio** para el logo de REDCO en la esquina superior izquierda de cada slide de contenido, y centrado (o en posición de cierre de marca) en portada y cierre. Esta guía fija únicamente la **sección reservada**; el tamaño, archivo y versión (clara/oscura) del isotipo se resuelven al momento de generar la presentación real, cuando se cuente con el activo gráfico definitivo.

---

## 7. Tipos de slide

Todo slide de una presentación REDCO corresponde a exactamente uno de los siguientes seis tipos (conjunto ME/CE: ningún slide pertenece a más de un tipo, y no hay necesidad de comunicación que quede fuera de este conjunto).

1. **Portada** — fondo azul medianoche a sangre completa, título y subtítulo en blanco/celeste hielo, espacio de logo.
2. **Agenda** — lista numerada de los temas a tratar. Debe ser MECE respecto al contenido real del deck: ningún tema del cuerpo queda fuera de la agenda, y ningún ítem de la agenda está de más.
3. **Framework MECE** (grid de categorías) — *n* tarjetas numeradas, cada una con título y lista de criterios; se usa para desglosar un tema en sus partes constituyentes.
4. **Comparación / alternativas de decisión** — columnas paralelas (2 o más), mismo set de atributos por columna, una alternativa real y distinta por columna.
5. **Tabla de seguimiento** — filas por tema, columnas de atributos fijos (objetivo, actividades, próximos pasos, estado), cerrada con la leyenda de estado de la sección 5.
6. **Cierre** — replica visualmente la portada (mismo fondo, mismo logo) para enmarcar el documento.

### 7.1 Encabezado de slide de contenido

Cada slide de contenido (tipos 2 a 5) abre con:

- Una **viñeta circular** azul acero sólida con punto blanco interior — mismo motivo de marca que el Estándar de Documentos REDCO, para mantener consistencia visual entre formatos.
- Título en Bold 26–28 pt, Tinta.

### 7.2 Frase de síntesis

Inmediatamente bajo el título, una línea en itálica (13–14 pt, Gris medio) que resume en una oración el mensaje central de la slide — qué debe concluir el lector antes de leer el detalle. Esta frase debe ser coherente con las categorías que siguen: si la frase promete "tres factores", la slide debe mostrar exactamente tres, ni más ni menos.

---

## 8. Componentes visuales

### Tarjeta numerada (categoría)

Contenedor con fondo `#F4F6F9`, borde `#E2E5EA` de 0.75 pt, sin esquinas redondeadas pronunciadas (radio mínimo). Estructura interna:

- Número de orden en círculo sólido azul acero, texto blanco Bold.
- Título de la categoría en Bold, Tinta.
- Lista de criterios o bullets en Regular, Tinta — máximo 4 ítems por tarjeta para evitar sobrecarga visual.
- Línea de cierre opcional con label "Responsable" en mayúscula Bold gris medio + valor en Bold Tinta.

A diferencia de esquemas donde cada tarjeta numerada lleva un color distinto, en REDCO **todas las tarjetas usan el mismo azul acero** para la numeración. Esto preserva la regla de dominancia de color (sección 5): un esquema multicolor por categoría diluye el azul como color dominante de marca.

### Leyenda de estado

Fila horizontal de 3 indicadores (punto de color + texto), siempre en la misma posición — pie de la slide — en toda tabla de seguimiento. Nunca se muestran menos ni más de 3 estados (ver sección 2).

### Columnas de comparación

Ancho igual entre columnas, separadas por al menos 0.3 in de espacio en blanco (no se usa línea divisoria vertical ni franja de color como separador — ver restricciones de diseño en la sección 10). Cada columna lleva el mismo set de atributos en el mismo orden, para que la comparación sea visualmente legible de un vistazo.

### Tabla de datos

Mismo patrón que el Estándar de Documentos REDCO: label en mayúscula Bold (gris medio) en la línea superior, valor en Regular (Tinta) en la línea inferior. Los campos que son el objeto principal de la fila (proyecto, responsable, fecha límite) se muestran en Bold.

---

## 9. Checklist MECE antes de exportar

- [ ] ¿Alguna categoría de una slide se solapa en contenido con otra categoría de la misma slide? → fusionar.
- [ ] ¿Existe algún caso, estado o responsable real del tema que no calce en ninguna categoría mostrada? → agregar categoría o ítem explícito.
- [ ] ¿La leyenda de estado usa exactamente 3 estados, sin un cuarto ambiguo?
- [ ] ¿El slide de Agenda coincide 1:1 con las secciones reales del deck — ni de más ni de menos?
- [ ] ¿La frase de síntesis de cada slide promete un número de elementos que coincide con el número de tarjetas/columnas mostradas?

---

## 10. Restricciones de diseño (heredadas de buenas prácticas de generación de slides)

- No usar franjas o barras de color decorativas como separador de columnas o cabecera/pie — leer como genérico de slide generado automáticamente.
- No subrayar títulos como recurso decorativo — usar espacio en blanco o el color de fondo para crear jerarquía.
- No centrar texto de cuerpo o bullets — alinear a la izquierda; solo los títulos de portada/cierre se centran.
- No dejar una slide con solo texto: todo framework MECE debe representarse con las tarjetas, columnas o tabla descritas en la sección 8 — nunca como una lista plana de viñetas.

---

## 11. Pendientes de validación

- [ ] Archivo vectorial del logo REDCO (versión clara y versión oscura, para uso sobre fondo azul medianoche y sobre fondo blanco)
- [ ] Confirmar licencia de distribución/embedding de Lexend Deca para equipos de cliente que no tengan la fuente instalada
