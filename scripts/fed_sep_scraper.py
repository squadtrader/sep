#!/usr/bin/env python3
"""
fed_sep_scraper.py
-------------------
Collecte les projections économiques du FOMC (Summary of Economic Projections - SEP)
publiées sur federalreserve.gov, à chaque réunion trimestrielle qui en comporte
(mars, juin, septembre, décembre).

Étapes :
1. Récupère la page calendrier des réunions FOMC pour trouver tous les liens vers
   les pages de projections ("fomcprojtabl<AAAAMMJJ>.htm").
2. Pour chaque réunion trouvée, télécharge la page "accessible version" et en extrait
   le Tableau 1 (Median / Central Tendency / Range pour GDP, chômage, inflation PCE,
   inflation PCE core, taux des fed funds), ainsi que la ligne de comparaison avec la
   projection précédente quand elle est présente.
3. Sauvegarde le tout en JSON (un enregistrement par réunion) et, en option, exporte
   un CSV "à plat" (pour un dashboard) et/ou un classeur Excel (.xlsx) déjà mis en
   forme et lisible directement (feuille résumé + une feuille détaillée par variable).

Usage :
    python fed_sep_scraper.py                       # toutes les réunions SEP trouvées
    python fed_sep_scraper.py --since 2023           # seulement à partir de 2023
    python fed_sep_scraper.py --latest               # seulement la réunion la plus récente
    python fed_sep_scraper.py --out data/fed_sep.json --xlsx data/fed_sep.xlsx
    python fed_sep_scraper.py --cache-dir .cache     # évite de re-télécharger les pages déjà vues

Dépendances : requests, beautifulsoup4, openpyxl
    pip install requests beautifulsoup4 openpyxl
"""

from __future__ import annotations

import argparse
import csv
import json
import re
import sys
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

CALENDAR_URL = "https://www.federalreserve.gov/monetarypolicy/fomccalendars.htm"
PROJ_URL_TMPL = "https://www.federalreserve.gov/monetarypolicy/fomcprojtabl{date}.htm"

HEADERS = {
    # Un User-Agent "normal" évite certains blocages basiques côté serveur.
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    )
}

# Variables que l'on cherche à extraire du Tableau 1, dans l'ordre où elles
# apparaissent sur la page. On matche sur un préfixe car le libellé exact
# varie légèrement d'une année à l'autre ("Change in real GDP", etc.)
VARIABLE_PATTERNS = [
    ("gdp_growth", re.compile(r"change in real gdp", re.I)),
    ("unemployment_rate", re.compile(r"unemployment rate", re.I)),
    ("pce_inflation", re.compile(r"^pce inflation", re.I)),
    ("core_pce_inflation", re.compile(r"core pce inflation", re.I)),
    ("fed_funds_rate", re.compile(r"federal funds rate", re.I)),
]

STAT_TYPES = ["median", "central_tendency", "range"]


@dataclass
class SepMeeting:
    date: str  # AAAAMMJJ
    url: str
    variables: dict = field(default_factory=dict)
    # Comparaison avec la projection précédente (ex: "March projection")
    prior_projection_label: Optional[str] = None
    prior_variables: dict = field(default_factory=dict)


def fetch(url: str, cache_dir: Optional[Path] = None) -> str:
    """Télécharge une URL (avec cache disque optionnel pour ne pas re-frapper le serveur)."""
    if cache_dir:
        cache_dir.mkdir(parents=True, exist_ok=True)
        cache_file = cache_dir / (re.sub(r"[^a-zA-Z0-9]+", "_", url) + ".html")
        if cache_file.exists():
            return cache_file.read_text(encoding="utf-8")

    resp = requests.get(url, headers=HEADERS, timeout=30)
    resp.raise_for_status()
    # federalreserve.gov ne déclare pas toujours son charset dans l'en-tête HTTP ;
    # sans ça, requests suppose ISO-8859-1 par défaut et corrompt les tirets "–"
    # ("2.0â2.3" au lieu de "2.0–2.3"). On force UTF-8, qui est le charset réel du site.
    resp.encoding = "utf-8"
    html = resp.text

    if cache_dir:
        cache_file.write_text(html, encoding="utf-8")

    return html


def find_sep_meeting_dates(calendar_html: str) -> list[str]:
    """Extrait toutes les dates AAAAMMJJ présentes dans des liens fomcprojtabl*.htm."""
    dates = sorted(set(re.findall(r"fomcprojtabl(\d{8})\.htm", calendar_html)))
    return dates


def _clean_cell(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def _looks_like_table1(table) -> bool:
    """Heuristique : la bonne table contient 'Median', 'Central Tendency' et 'Range'
    dans ses en-têtes, sans dépendre d'un nom de classe CSS précis (qui peut changer)."""
    header_text = " ".join(_clean_cell(c.get_text(" ")) for c in table.find_all(["th", "td"])[:20])
    return (
        "median" in header_text.lower()
        and "central tendency" in header_text.lower()
        and "range" in header_text.lower()
    )


def parse_table1(html: str) -> tuple[dict, Optional[str], dict]:
    """Parse le Tableau 1 (projections économiques) d'une page fomcprojtabl*.htm.

    Retourne (variables, prior_label, prior_variables) où :
      - variables[var_key][stat_type][year] = valeur (str, ex "2.2" ou "2.0-2.3")
      - prior_label = ex "March projection" (None si absent)
      - prior_variables = même structure que variables, pour la projection précédente
    """
    soup = BeautifulSoup(html, "html.parser")

    table1 = None
    for table in soup.find_all("table"):
        if _looks_like_table1(table):
            table1 = table
            break

    if table1 is None:
        raise ValueError("Impossible de localiser le Tableau 1 sur la page.")

    rows = table1.find_all("tr")

    # --- Repérer les colonnes années dans la ligne d'en-tête ---
    # La structure typique a 2 lignes d'en-tête : une avec Median/Central Tendency/Range
    # (groupes de 4 colonnes chacun) et une avec les années répétées 3 fois.
    year_header_row = None
    for row in rows[:4]:
        cells = [_clean_cell(c.get_text(" ")) for c in row.find_all(["th", "td"])]
        if sum(1 for c in cells if re.fullmatch(r"(19|20)\d{2}", c)) >= 3:
            year_header_row = cells
            break

    if year_header_row is None:
        raise ValueError("Impossible de localiser la ligne des années dans le Tableau 1.")

    years_in_order = [c for c in year_header_row if re.fullmatch(r"(19|20)\d{2}", c) or c.lower() == "longer run"]
    # Normalement 4 années/colonnes (an1, an2, an3, longer run) répétées 3 fois
    # (Median, Central Tendency, Range). On reconstruit les 3 blocs.
    block_size = 4
    blocks = [years_in_order[i:i + block_size] for i in range(0, len(years_in_order), block_size)]
    # Sécurité : si jamais on n'a pas exactement 3 blocs pleins, on garde ce qu'on a.
    blocks = [b for b in blocks if b]

    variables: dict = {}
    prior_label: Optional[str] = None
    prior_variables: dict = {}

    current_var_key: Optional[str] = None

    for row in rows:
        cells = row.find_all(["th", "td"])
        if not cells:
            continue
        texts = [_clean_cell(c.get_text(" ")) for c in cells]
        if not texts or not texts[0]:
            continue

        label = texts[0]
        data_cells = texts[1:]

        # Ligne de comparaison avec la projection précédente
        prior_match = re.match(r"^(January|February|March|April|May|June|July|August|"
                                r"September|October|November|December)\s+projection$", label, re.I)
        if prior_match:
            prior_label = label
            if current_var_key and data_cells:
                prior_variables.setdefault(current_var_key, {})
                _fill_stats(prior_variables[current_var_key], data_cells, blocks)
            continue

        # Ligne de variable principale
        matched_key = None
        for key, pattern in VARIABLE_PATTERNS:
            if pattern.search(label):
                matched_key = key
                break

        if matched_key:
            current_var_key = matched_key
            variables.setdefault(matched_key, {"label": label})
            if data_cells:
                _fill_stats(variables[matched_key], data_cells, blocks)

    return variables, prior_label, prior_variables


def _fill_stats(target: dict, data_cells: list[str], blocks: list[list[str]]) -> None:
    """Répartit les valeurs d'une ligne de données dans Median / Central Tendency / Range,
    par année, en s'appuyant sur les blocs d'années détectés dans l'en-tête."""
    idx = 0
    for stat_type, block in zip(STAT_TYPES, blocks):
        stat_dict = target.setdefault(stat_type, {})
        for year in block:
            if idx >= len(data_cells):
                break
            value = data_cells[idx]
            if value and value != "-":
                stat_dict[year] = value
            idx += 1


def scrape_meeting(date: str, cache_dir: Optional[Path] = None, pause: float = 1.0) -> Optional[SepMeeting]:
    url = PROJ_URL_TMPL.format(date=date)
    try:
        html = fetch(url, cache_dir=cache_dir)
    except requests.HTTPError as exc:
        print(f"  ! {date}: page introuvable ou erreur HTTP ({exc})", file=sys.stderr)
        return None

    try:
        variables, prior_label, prior_variables = parse_table1(html)
    except ValueError as exc:
        print(f"  ! {date}: échec du parsing ({exc})", file=sys.stderr)
        return None

    time.sleep(pause)  # politesse envers le serveur de la Fed
    return SepMeeting(
        date=date,
        url=url,
        variables=variables,
        prior_projection_label=prior_label,
        prior_variables=prior_variables,
    )


def meeting_to_dict(m: SepMeeting) -> dict:
    return {
        "date": m.date,
        "date_iso": f"{m.date[0:4]}-{m.date[4:6]}-{m.date[6:8]}",
        "url": m.url,
        "prior_projection_label": m.prior_projection_label,
        "variables": m.variables,
        "prior_variables": m.prior_variables,
    }


def write_json(meetings: list[SepMeeting], out_path: Path) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    payload = [meeting_to_dict(m) for m in meetings]
    out_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def write_csv(meetings: list[SepMeeting], out_path: Path) -> None:
    """Export 'à plat' : une ligne par (réunion, variable, statistique, année)."""
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["meeting_date", "variable", "label", "stat_type", "year", "value", "is_prior"])
        for m in meetings:
            for var_key, var_data in m.variables.items():
                label = var_data.get("label", var_key)
                for stat_type in STAT_TYPES:
                    for year, value in var_data.get(stat_type, {}).items():
                        writer.writerow([m.date, var_key, label, stat_type, year, value, False])
            for var_key, var_data in m.prior_variables.items():
                label = var_data.get("label", var_key) if isinstance(var_data, dict) else var_key
                for stat_type in STAT_TYPES:
                    for year, value in var_data.get(stat_type, {}).items():
                        writer.writerow([m.date, var_key, label, stat_type, year, value, True])


# --- Export Excel -----------------------------------------------------------

EXCEL_VARIABLES = [
    ("gdp_growth", "GDP"),
    ("unemployment_rate", "Chômage"),
    ("pce_inflation", "Inflation PCE"),
    ("core_pce_inflation", "Inflation PCE core"),
    ("fed_funds_rate", "Taux Fed Funds"),
]
EXCEL_STAT_LABELS = [("median", "Médiane"), ("central_tendency", "Central Tendency"), ("range", "Range")]

_XLSX_FONT_NAME = "Arial"
_XLSX_HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
_XLSX_HEADER_FONT = Font(name=_XLSX_FONT_NAME, size=10, bold=True, color="FFFFFF")
_XLSX_SUBHEADER_FILL = PatternFill("solid", fgColor="D9E1F2")
_XLSX_SUBHEADER_FONT = Font(name=_XLSX_FONT_NAME, size=10, bold=True)
_XLSX_CELL_FONT = Font(name=_XLSX_FONT_NAME, size=10)
_XLSX_THIN = Side(style="thin", color="B7B7B7")
_XLSX_BORDER = Border(left=_XLSX_THIN, right=_XLSX_THIN, top=_XLSX_THIN, bottom=_XLSX_THIN)
_XLSX_CENTER = Alignment(horizontal="center", vertical="center")


def _xlsx_year_order(years: set) -> list:
    numeric = sorted((y for y in years if y != "Longer run"), key=int)
    return numeric + (["Longer run"] if "Longer run" in years else [])


def _xlsx_style_header(cell, fill, font) -> None:
    cell.fill = fill
    cell.font = font
    cell.alignment = _XLSX_CENTER
    cell.border = _XLSX_BORDER


def _xlsx_autofit(ws, min_width: int = 10, max_width: int = 22) -> None:
    for col_cells in ws.columns:
        length = max((len(str(c.value)) for c in col_cells if c.value is not None), default=0)
        col_letter = get_column_letter(col_cells[0].column)
        ws.column_dimensions[col_letter].width = min(max(length + 2, min_width), max_width)


def _xlsx_build_summary_sheet(wb: Workbook, meetings: list[SepMeeting], years: list) -> None:
    ws = wb.active
    ws.title = "Résumé (Médianes)"

    ws.cell(row=1, column=1, value="Date réunion")
    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)
    _xlsx_style_header(ws.cell(row=1, column=1), _XLSX_HEADER_FILL, _XLSX_HEADER_FONT)

    col = 2
    for var_key, var_label in EXCEL_VARIABLES:
        start_col = col
        for year in years:
            c = ws.cell(row=2, column=col, value=year)
            _xlsx_style_header(c, _XLSX_SUBHEADER_FILL, _XLSX_SUBHEADER_FONT)
            col += 1
        end_col = col - 1
        ws.merge_cells(start_row=1, start_column=start_col, end_row=1, end_column=end_col)
        _xlsx_style_header(ws.cell(row=1, column=start_col, value=var_label), _XLSX_HEADER_FILL, _XLSX_HEADER_FONT)

    row = 3
    for m in sorted(meetings, key=lambda x: x.date):
        date_iso = f"{m.date[0:4]}-{m.date[4:6]}-{m.date[6:8]}"
        cell = ws.cell(row=row, column=1, value=date_iso)
        cell.font, cell.border, cell.alignment = _XLSX_CELL_FONT, _XLSX_BORDER, _XLSX_CENTER
        col = 2
        for var_key, _ in EXCEL_VARIABLES:
            medians = m.variables.get(var_key, {}).get("median", {})
            for year in years:
                c = ws.cell(row=row, column=col, value=medians.get(year, ""))
                c.font, c.border, c.alignment = _XLSX_CELL_FONT, _XLSX_BORDER, _XLSX_CENTER
                col += 1
        row += 1

    ws.freeze_panes = "B3"
    _xlsx_autofit(ws)


def _xlsx_build_variable_sheet(wb: Workbook, meetings: list[SepMeeting], var_key: str, var_label: str, years: list) -> None:
    ws = wb.create_sheet(var_label[:31])

    ws.cell(row=1, column=1, value="Date réunion")
    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)
    _xlsx_style_header(ws.cell(row=1, column=1), _XLSX_HEADER_FILL, _XLSX_HEADER_FONT)

    col = 2
    for year in years:
        start_col = col
        for _, stat_label in EXCEL_STAT_LABELS:
            c = ws.cell(row=2, column=col, value=stat_label)
            _xlsx_style_header(c, _XLSX_SUBHEADER_FILL, _XLSX_SUBHEADER_FONT)
            col += 1
        end_col = col - 1
        ws.merge_cells(start_row=1, start_column=start_col, end_row=1, end_column=end_col)
        _xlsx_style_header(ws.cell(row=1, column=start_col, value=year), _XLSX_HEADER_FILL, _XLSX_HEADER_FONT)

    row = 3
    for m in sorted(meetings, key=lambda x: x.date):
        date_iso = f"{m.date[0:4]}-{m.date[4:6]}-{m.date[6:8]}"
        var_data = m.variables.get(var_key, {})
        cell = ws.cell(row=row, column=1, value=date_iso)
        cell.font, cell.border, cell.alignment = _XLSX_CELL_FONT, _XLSX_BORDER, _XLSX_CENTER
        col = 2
        for year in years:
            for stat_key, _ in EXCEL_STAT_LABELS:
                value = var_data.get(stat_key, {}).get(year, "")
                c = ws.cell(row=row, column=col, value=value)
                c.font, c.border, c.alignment = _XLSX_CELL_FONT, _XLSX_BORDER, _XLSX_CENTER
                col += 1
        row += 1

    ws.freeze_panes = "B3"
    _xlsx_autofit(ws)


def write_xlsx(meetings: list[SepMeeting], out_path: Path) -> None:
    """Écrit un classeur Excel lisible : une feuille résumé (médianes) + une feuille
    détaillée (Médiane / Central Tendency / Range par année) pour chaque variable."""
    if not meetings:
        return

    years: set = set()
    for m in meetings:
        for var_data in m.variables.values():
            years.update(var_data.get("median", {}).keys())
    years = _xlsx_year_order(years)

    wb = Workbook()
    _xlsx_build_summary_sheet(wb, meetings, years)
    for var_key, var_label in EXCEL_VARIABLES:
        _xlsx_build_variable_sheet(wb, meetings, var_key, var_label, years)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


def main() -> None:
    parser = argparse.ArgumentParser(description="Collecte les projections économiques (SEP) de la Fed.")
    parser.add_argument("--out", default="fed_sep.json", help="Chemin du fichier JSON de sortie.")
    parser.add_argument("--csv", default=None, help="Chemin optionnel d'export CSV à plat.")
    parser.add_argument("--xlsx", default=None, help="Chemin optionnel d'export Excel (.xlsx) lisible.")
    parser.add_argument("--since", type=int, default=None, help="Année minimale à inclure (ex: 2023).")
    parser.add_argument("--latest", action="store_true", help="Ne récupérer que la dernière réunion SEP.")
    parser.add_argument("--cache-dir", default=None, help="Dossier de cache des pages HTML téléchargées.")
    parser.add_argument("--pause", type=float, default=1.0, help="Pause en secondes entre deux requêtes.")
    args = parser.parse_args()

    cache_dir = Path(args.cache_dir) if args.cache_dir else None

    print(f"Récupération du calendrier FOMC : {CALENDAR_URL}")
    calendar_html = fetch(CALENDAR_URL, cache_dir=cache_dir)
    dates = find_sep_meeting_dates(calendar_html)

    if args.since:
        dates = [d for d in dates if int(d[:4]) >= args.since]

    if args.latest and dates:
        dates = [dates[-1]]

    if not dates:
        print("Aucune réunion SEP trouvée avec ces critères.", file=sys.stderr)
        sys.exit(1)

    print(f"{len(dates)} réunion(s) SEP à traiter : {', '.join(dates)}")

    meetings: list[SepMeeting] = []
    for date in dates:
        print(f"-> {date} ...")
        meeting = scrape_meeting(date, cache_dir=cache_dir, pause=args.pause)
        if meeting:
            meetings.append(meeting)

    out_path = Path(args.out)
    write_json(meetings, out_path)
    print(f"JSON écrit : {out_path} ({len(meetings)} réunions)")

    if args.csv:
        csv_path = Path(args.csv)
        write_csv(meetings, csv_path)
        print(f"CSV écrit : {csv_path}")

    if args.xlsx:
        xlsx_path = Path(args.xlsx)
        write_xlsx(meetings, xlsx_path)
        print(f"Excel écrit : {xlsx_path}")


if __name__ == "__main__":
    main()
