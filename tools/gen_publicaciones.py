#!/usr/bin/env python3
"""Genera publicaciones.html y publications.bib desde los Excel de la jerarquización.

Uso: python3 tools/gen_publicaciones.py <Listado_Publicaciones.xlsx> <Indice_Pubs_Nuevas.xlsx>

Fuentes (repositorio de jerarquización, carpeta 04.-Envío_Titular_MDD):
- Listado consolidado: hojas "Revistas WoS-JCR" y "Conferencias-Proceedings".
- Índice de publicaciones nuevas: referencias completas para el BibTeX.
"""
import sys, html, re, unicodedata
from openpyxl import load_workbook

HEAD = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Publicaciones · Matías Díaz</title>
<meta name="description" content="Publicaciones de Matías Díaz (DIE USACH): revistas WoS-JCR y conferencias, período 2021-2026, con enlace DOI.">
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Instrument+Serif:ital@0;1&family=Archivo:wght@400;500;600;700&display=swap">
<link rel="stylesheet" href="assets/site.css">
</head>
<body>
<nav><div class="wrap">
  <a class="logo" href="index.html">Matías <b>Díaz</b></a>
  <a href="investigacion.html">Investigación</a>
  <a class="on" href="publicaciones.html">Publicaciones</a>
  <a href="videos.html">Videos</a>
  <a href="prensa.html">Prensa</a>
  <a href="docencia.html">Docencia</a>
  <a href="equipo.html">Equipo</a>
  <a href="postula.html">Postula</a>
</div></nav>
<header class="pagehead wrap">
  <h1>Publicaciones</h1>
  <p>{n_j} artículos en revistas WoS-JCR y {n_c} en conferencias en el período 2021-2026, según el índice consolidado de la postulación a Profesor Titular (junio de 2026). Perfil completo en <a href="https://scholar.google.com/citations?user=-43YaJIAAAAJ" target="_blank" rel="noopener">Google Scholar</a> · <a href="publications.bib">BibTeX</a>.</p>
</header>
"""

FOOT = """
<footer><div class="wrap">
  <span>© 2026 Matías Díaz · DIE USACH</span>
  <span>Generado desde el índice consolidado de publicaciones · <a href="publications.bib">publications.bib</a></span>
</div></footer>
<script src="assets/site.js"></script>
</body>
</html>
"""

def esc(x):
    return html.escape(str(x).strip()) if x else ""

def link(url, label):
    if url and str(url).startswith("http"):
        return f'<a href="{esc(url)}" target="_blank" rel="noopener">{label}</a>'
    return ""

def bibkey(title, year):
    t = unicodedata.normalize("NFKD", title).encode("ascii", "ignore").decode()
    words = [w.lower() for w in re.findall(r"[A-Za-z]{4,}", t)][:3]
    return f"diaz{year}" + "".join(w[:4] for w in words)

def main(listado, indice):
    wb = load_workbook(listado, data_only=True)
    def _yr(v):
        try: return int(v)
        except (TypeError, ValueError): return None
    js = [r for r in wb["Revistas WoS-JCR"].iter_rows(min_row=2, values_only=True) if _yr(r[1]) and r[2]]
    cs = [r for r in wb["Conferencias-Proceedings"].iter_rows(min_row=2, values_only=True) if _yr(r[1]) and r[2]]

    out = [HEAD.format(n_j=len(js), n_c=len(cs))]

    out.append('<section style="padding-top:40px"><div class="wrap"><div class="head"><h2>Revistas WoS-JCR</h2></div>')
    for year in sorted({int(r[1]) for r in js}, reverse=True):
        out.append(f'<div class="pubyear">{year}</div>')
        for r in [x for x in js if int(x[1]) == year]:
            _, y, title, journal, cita, _, url = (list(r) + [None]*7)[:7]
            rev = esc(journal) if journal and "verificar" not in str(journal) else "Revista"
            cita_txt = f", {esc(cita)}" if cita else ""
            out.append(
                f'<div class="pub"><div class="thumb">{rev.split()[0][:10]}</div><div>'
                f'<p>{esc(title)}. <em>{rev}{cita_txt}.</em></p>'
                f'<div class="links">{link(url, "DOI")}<span>{year}</span></div></div></div>')
    out.append("</div></section>")

    out.append('<section class="alt"><div class="wrap"><div class="head"><h2>Conferencias y proceedings</h2></div>')
    for year in sorted({int(r[1]) for r in cs}, reverse=True):
        out.append(f'<div class="pubyear">{year}</div>')
        for r in cs:
            if int(r[1]) != year:
                continue
            _, y, title, fuente, url = (list(r) + [None]*5)[:5]
            out.append(
                f'<div class="pub"><div class="thumb">Conf.</div><div>'
                f'<p>{esc(title)}. <em>{esc(fuente) or "Proceedings"}.</em></p>'
                f'<div class="links">{link(url, "DOI")}<span>{year}</span></div></div></div>')
    out.append("</div></section>")
    out.append(FOOT)

    with open("publicaciones.html", "w", encoding="utf-8") as f:
        f.write("\n".join(out))

    # BibTeX desde el índice de publicaciones nuevas (referencia completa + DOI)
    wb2 = load_workbook(indice, data_only=True)
    ws = wb2.active
    bibs, seen = [], set()
    for r in ws.iter_rows(min_row=2, values_only=True):
        n, year, tipo, title, url, ref = (list(r) + [None]*8)[:6]
        if not (year and title):
            continue
        key = bibkey(str(title), year)
        while key in seen:
            key += "x"
        seen.add(key)
        entry = "article" if tipo and "WOS" in str(tipo).upper() else ("article" if tipo and "revista" in str(tipo).lower() else "inproceedings")
        doi = ""
        if url and "doi.org/" in str(url):
            doi = str(url).split("doi.org/")[-1].strip()
        bibs.append(
            f"@{entry}{{{key},\n"
            f"  title = {{{str(title).strip()}}},\n"
            f"  year = {{{year}}},\n"
            + (f"  doi = {{{doi}}},\n" if doi else "")
            + (f"  note = {{{str(ref).strip()[:300]}}},\n" if ref else "")
            + "}\n")
    with open("publications.bib", "w", encoding="utf-8") as f:
        f.write("% Publicaciones de Matías Díaz, período 2021-2026.\n"
                "% Generado desde 00_INDICE_Publicaciones_Nuevas_2021-2026.xlsx.\n\n" + "\n".join(bibs))
    print(f"OK: {len(js)} revistas, {len(cs)} conferencias, {len(bibs)} entradas BibTeX")

if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
