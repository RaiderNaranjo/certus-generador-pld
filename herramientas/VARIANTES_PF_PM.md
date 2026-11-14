# Variantes persona física / persona moral (kit profesionistas)

Las plantillas de `plantillas/profesionistas` están redactadas para persona física (PF) y marcan lo que cambia para persona moral (PM):

| Marca | Estilo de carácter | Se ve |
|---|---|---|
| `[PF: …]` | Variante PF | sombreado verde |
| `[PM: …]` | Variante PM | sombreado amarillo |
| Nota de edición | Nota Certus | rojo, cursiva |

El texto sin marca aplica a ambos.

## En el generador
`generador.js` lee `TipoPersona` (PF/PM). Si viene, elimina la variante que no aplica y la nota, y quita etiquetas y sombreado de la que se conserva. Si no viene, entrega los documentos con ambas variantes marcadas para editarlos a mano.

## En Word
Importar `CertusVariantes.bas` (Alt+F11 > Archivo > Importar) y ejecutar `ResolverComoPF` o `ResolverComoPM`.

## Reglas para editar plantillas
- Cada variante abre y cierra dentro del mismo párrafo y no lleva corchetes dentro (no meter `[Nombre]` u otros campos dentro de una variante).
- Párrafo completo de una sola variante: la marca de párrafo también lleva el estilo, para que el párrafo desaparezca al resolver. No usar párrafos completos de variante dentro de celdas de tabla (usar variantes dentro del texto de la celda).
