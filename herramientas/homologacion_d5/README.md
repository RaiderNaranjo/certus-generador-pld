# Homologación del D5

`build_d5.py <dir_fuentes> <dir_salida>` reconstruye los cuatro D5 a partir de estas fuentes:
- `PROF_pr12.docx`: D5-PROF del 12-nov-2026 (rama del PR #12). Es la referencia y aporta el encabezado y pie con los logotipos.
- `PROF_pr2`, `INM_pr2`, `NOT_pr2` y `COR_pr2`: contenido del Capítulo X RCG (PR #2).
- `INM_pr13`: variantes PF/PM del inmobiliario (PR #13).

Las fuentes se extraen con `git show <rama>:<ruta>`. Las reglas están en `../GUIA_ESTILO_D5.md`.
