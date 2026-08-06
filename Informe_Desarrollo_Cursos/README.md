# Dashboard de Desarrollo de Cursos

## Actualización automática

1. Cargue o reemplace un archivo `.xlsx` dentro de esta carpeta.
2. Para un proceso predecible, use el nombre `reporte_desarrollo_cursos.xlsx`.
3. GitHub Actions ejecutará `generate_dashboard.py`.
4. `ucan_course_development_dashboard.html` se regenerará y se publicará automáticamente.

También puede ejecutarse manualmente desde **Actions → Actualizar Dashboard de Cursos → Run workflow**.

El generador busca columnas `Recurso`, `Curso`, `Estatus` y `Total Comprometido`. También reconoce la hoja `Faltan Year 3`.
