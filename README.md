# matiasdiaz.cl

Sitio web personal de Matías Díaz Díaz, Profesor Asociado del Departamento de Ingeniería Eléctrica de la Universidad de Santiago de Chile (DIE USACH).

Sitio estático sin framework ni proceso de build: HTML, CSS y JS planos, servido con GitHub Pages. Diseño híbrido: portada visual y páginas interiores sobrias tipo al-folio.

## Estructura

- `index.html`: portada (investigación, en terreno, ecosistema, prensa, trayectoria).
- `publicaciones.html`: 24 revistas WoS-JCR y 50 conferencias del período 2021-2026, con DOI. Generada con `tools/gen_publicaciones.py`.
- `publications.bib`: BibTeX generado desde el índice de publicaciones.
- `prensa.html`: selección curada y registro completo de apariciones en prensa.
- `videos.html`: clases, seminarios y entrevistas del canal de YouTube.
- `investigacion.html`, `docencia.html`, `equipo.html`, `postula.html`.
- `assets/`: estilos y JS compartidos.

## Actualización de contenido

- Publicaciones: actualizar el Excel fuente (repositorio de jerarquización, carpeta `04.-Envío_Titular_MDD`) y correr `python3 tools/gen_publicaciones.py <listado.xlsx> <indice.xlsx>`.
- Prensa y videos: editar las listas en sus HTML.
- Los archivos fuente (Excel) no se versionan en este repositorio.

## Pendientes

- Fotografías reales (retrato, laboratorio, cargador V2G, Lollapalooza, equipo).
- Nombres de equipo y temas de tesis vigentes.
- Versión en inglés (i18n) y CNAME al comprar el dominio.
