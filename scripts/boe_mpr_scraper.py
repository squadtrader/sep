#!/usr/bin/env python3
"""
boe_mpr_scraper.py
--------------------
Collecte les projections économiques du Monetary Policy Report (MPR) de la
Banque d'Angleterre (BoE), publié 4 fois par an (février, mai, août, novembre).

Particularités de la BoE par rapport à la Fed / BOC / BCE :
  - Les projections sont présentées sous forme de "fan charts" (probabilistes),
    mais un tableau récapitulatif ("Table 1.A: Forecast summary") donne les
    valeurs modales/centrales par année (année 1 à année 4), avec la valeur du
    rapport précédent entre parenthèses -- comme la Fed.
  - Ce tableau inclut le taux directeur ("Bank Rate") : contrairement à la BOC
    et à la BCE, la BoE affiche bien un chemin de taux dans son tableau
    récapitulatif -- mais il s'agit du chemin IMPLICITE PAR LES MARCHÉS sur
    lequel les projections sont conditionnées, pas d'une prévision propre de la
    BoE (la Banque ne prévoit pas sa propre trajectoire de taux).
  - Certaines valeurs sont exprimées en fractions unicode (¼, ½, ¾) plutôt qu'en
    décimales -- le script les convertit en décimal pour l'export.
  - La page HTML du rapport ne semble pas utiliser de <table> HTML standard
    pour ce tableau (contrairement aux 3 autres banques) : le texte du tableau
    apparaît "collé" sans séparateur. Le script travaille donc directement sur
    le texte brut de la page plutôt que sur la structure DOM d'une <table>.
    C'est la partie la plus fragile de ce script -- si elle échoue sur le vrai
    site, un fichier de diagnostic est sauvegardé (voir plus bas).

Étapes :
1. Parcourt le plan du site des MPR (bankofengland.co.uk/sitemap/monetary-policy-report)
   pour trouver l'URL de chaque rapport (remonte jusqu'en 2019).
2. Pour chaque rapport, télécharge sa page et en extrait la Table 1.A (GDP, CPI
   inflation, Unemployment rate, Excess supply/demand, Bank Rate) par année.
3. Sauvegarde le tout en JSON, avec exports CSV et Excel automatiques.

Usage :
    python boe_mpr_scraper.py
    python boe_mpr_scraper.py --since 2023
    python boe_mpr_scraper.py --latest
    python boe_mpr_scraper.py --out boe_mpr.json --no-xlsx
    python boe_mpr_scraper.py --cache-dir .cache_boe

Dépendances : requests, beautifulsoup4, openpyxl
    pip install requests beautifulsoup4 openpyxl
"""

from __future__ import annotations

import argparse
import csv
import html as html_module
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

SITEMAP_URL = "https://www.bankofengland.co.uk/sitemap/monetary-policy-report"

# Les rapports antérieurs à mai 2022 sont bloqués par la protection anti-bot du
# site (Akamai) : la page reçue ne contient pas le contenu réel de l'article,
# quelle que soit la structure recherchée. Testé et confirmé sur plusieurs
# rapports 2019-2022. On se limite donc par défaut à la période qui fonctionne
# de façon fiable ; quelques rapports isolés dans cette période peuvent encore
# échouer ponctuellement (même protection, de façon intermittente) -- le
# script les ignore proprement et continue avec les suivants.
MIN_SUPPORTED_PERIOD = "2022-05"
BASE_URL = "https://www.bankofengland.co.uk"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    )
}

# Lien vers un rapport, ex: href="https://www.bankofengland.co.uk/monetary-policy-report/2025/february-2025"
# (URL absolue sur cette page -- on exclut les sous-pages comme .../february-2025/annex-...
# en exigeant que le guillemet fermant suive immédiatement le mois-année).
REPORT_LINK_RE = re.compile(
    r'href="(?:https?://www\.bankofengland\.co\.uk)?(/monetary-policy-report/(\d{4})/([a-z]+-\d{4}))"'
)

ROW_LABELS = [
    ("gdp", "GDP"),
    ("cpi_inflation", "CPI inflation"),
    ("unemployment_rate", "Unemployment rate"),
    ("output_gap", "Excess supply/Excess demand"),
    ("bank_rate", "Bank Rate"),
]

FRAC_SUFFIX = {"¼": "25", "½": "5", "¾": "75"}

# Nombre : décimal classique, entier+fraction accolée ("1¼" = 1.25), entier
# seul, ou fraction seule ("¼"), avec signe + ou - optionnel.
_NUM_CORE = r"(?:\d+\.\d+|\d+[¼½¾]|\d+|[¼½¾])"
NUM_RE = re.compile(rf"([-+]?{_NUM_CORE})(?:\s*\(([-+]?{_NUM_CORE})\))?")
YEAR_Q_RE = re.compile(r"\d{4}\s*Q[1-4]")


@dataclass
class BoeReport:
    date: str   # AAAA-MM
    url: str
    table: dict = field(default_factory=dict)  # variable -> year -> {"value":.., "prior":..}


def fetch(url: str, cache_dir: Optional[Path] = None) -> str:
    if cache_dir:
        cache_dir.mkdir(parents=True, exist_ok=True)
        cache_file = cache_dir / (re.sub(r"[^a-zA-Z0-9]+", "_", url) + ".html")
        if cache_file.exists():
            return cache_file.read_text(encoding="utf-8")

    resp = requests.get(url, headers=HEADERS, timeout=30)
    resp.raise_for_status()
    resp.encoding = "utf-8"
    html = resp.text

    if cache_dir:
        cache_file.write_text(html, encoding="utf-8")

    return html


def discover_reports(cache_dir: Optional[Path] = None, min_period: str = MIN_SUPPORTED_PERIOD) -> list[tuple[str, str]]:
    """Retourne une liste de (date AAAA-MM, url) pour chaque rapport trouvé,
    à partir de min_period (voir MIN_SUPPORTED_PERIOD)."""
    html = fetch(SITEMAP_URL, cache_dir=cache_dir)
    seen = {}
    for m in REPORT_LINK_RE.finditer(html):
        path, year, month_year = m.group(1), m.group(2), m.group(3)
        month_name = month_year.rsplit("-", 1)[0]
        month_num = {
            "january": "01", "february": "02", "march": "03", "april": "04",
            "may": "05", "june": "06", "july": "07", "august": "08",
            "september": "09", "october": "10", "november": "11", "december": "12",
        }.get(month_name)
        if not month_num:
            continue
        date = f"{year}-{month_num}"
        if min_period and date < min_period:
            continue
        seen[date] = (date, BASE_URL + path)

    if not seen:
        debug_path = Path("debug_boe_sitemap.html")
        debug_path.write_text(html, encoding="utf-8")
        print(f"  ! Aucun rapport trouvé. HTML sauvegardé dans {debug_path.resolve()} "
              f"pour diagnostic.", file=sys.stderr)

    return sorted(seen.values())


def _to_decimal(token: str) -> str:
    """Convertit un nombre éventuellement signé (+/-) contenant une fraction
    unicode (¼, ½, ¾), seule ou accolée à un entier ("1¼" = 1.25), en décimal."""
    sign = ""
    if token and token[0] in "+-":
        sign = "-" if token[0] == "-" else ""
        token = token[1:]

    m = re.match(r"^(\d+)([¼½¾])$", token)
    if m:
        integer_part, frac = m.groups()
        return f"{sign}{integer_part}.{FRAC_SUFFIX[frac]}"
    if token in FRAC_SUFFIX:
        return f"{sign}0.{FRAC_SUFFIX[token]}"
    return f"{sign}{token}"


def parse_table_1a(html: str) -> tuple[dict, list[str]]:
    """Extrait la Table 1.A (Forecast summary) du texte brut de la page. Ce
    tableau n'apparaît pas comme une <table> HTML standard sur le site de la
    BoE (contrairement aux 3 autres banques) : on travaille donc sur le texte
    de la page avec les balises retirées, dans l'ordre où le contenu apparaît."""
    # Texte brut, balises retirées, entités HTML décodées -- au plus proche de
    # ce qu'affiche la page, sans réintroduire d'espaces que le site n'a pas.
    plain = re.sub(r"<[^>]+>", "", html)
    plain = html_module.unescape(plain)

    # Le numéro du tableau ("Table 1.A", "Table 3.A"...) varie selon l'édition
    # (nombre de sections différent en amont) -- on cherche son titre, stable
    # d'une édition à l'autre, plutôt que son numéro.
    start_match = re.search(r"Table\s+\d+\.[A-Z]\s*:\s*Forecast summary", plain)
    if start_match is None:
        raise ValueError("Table 'Forecast summary' introuvable sur la page.")
    start = start_match.start()

    # La section utile se termine avant les notes de bas de page (qui commencent
    # par "Footnotes") ou, à défaut, avant le prochain grand tableau ("Table 1.B").
    next_table_match = re.search(r"Table\s+\d+\.[A-Z]\s*:", plain[start_match.end():])
    end_candidates = [plain.find("Footnotes", start)]
    if next_table_match:
        end_candidates.append(start_match.end() + next_table_match.start())
    end_candidates = [c for c in end_candidates if c != -1]
    end = min(end_candidates) if end_candidates else start + 2000
    section = plain[start:end]

    years = [re.sub(r"\s+", " ", y) for y in YEAR_Q_RE.findall(section)]
    years = list(dict.fromkeys(years))  # dédoublonne en gardant l'ordre
    if not years:
        raise ValueError("Années introuvables dans la Table 1.A.")
    n_years = len(years)

    # Position de chaque libellé de ligne dans la section (pour découper les blocs de données)
    label_positions = []
    for key, label in ROW_LABELS:
        # Insensible à la casse, et "unemployment rate" tolère un préfixe "LFS "
        # utilisé dans les rapports plus anciens ("LFS unemployment rate").
        label_pattern_str = r"(?:LFS\s+)?" + re.escape(label) if label == "Unemployment rate" else re.escape(label)
        pattern = re.compile(label_pattern_str + r"\s*(?:\([a-z]\))?", re.I)
        m = pattern.search(section)
        if m:
            label_positions.append((key, m.start(), m.end()))

    if not label_positions:
        raise ValueError("Aucune ligne de la Table 1.A reconnue.")

    label_positions.sort(key=lambda x: x[1])

    result: dict = {}
    for i, (key, _, data_start) in enumerate(label_positions):
        data_end = label_positions[i + 1][1] if i + 1 < len(label_positions) else len(section)
        chunk = section[data_start:data_end]
        tokens = NUM_RE.findall(chunk)
        if not tokens:
            continue
        tokens = tokens[:n_years]
        entry = {}
        for year, (value, prior) in zip(years, tokens):
            entry[year] = {
                "value": _to_decimal(value),
                "prior": _to_decimal(prior) if prior else None,
            }
        result[key] = entry

    return result, years


def scrape_report(date: str, url: str, cache_dir: Optional[Path] = None, pause: float = 1.0) -> Optional[BoeReport]:
    try:
        html = fetch(url, cache_dir=cache_dir)
    except requests.HTTPError as exc:
        print(f"  ! {date}: page introuvable ({exc})", file=sys.stderr)
        return None

    report = BoeReport(date=date, url=url)
    try:
        report.table, _ = parse_table_1a(html)
    except ValueError as exc:
        print(f"  ! {date}: Table 1.A non extraite ({exc})", file=sys.stderr)
        debug_path = Path(f"debug_boe_report_{date}.html")
        debug_path.write_text(html, encoding="utf-8")
        print(f"    -> HTML brut sauvegardé dans {debug_path.resolve()} pour diagnostic.", file=sys.stderr)
        return None

    time.sleep(pause)
    return report


def report_to_dict(r: BoeReport) -> dict:
    return {"date": r.date, "url": r.url, "table_1a_forecast_summary": r.table}


def write_json(reports: list[BoeReport], out_path: Path) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    payload = [report_to_dict(r) for r in reports]
    out_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def write_csv(reports: list[BoeReport], out_path: Path) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["report_date", "variable", "period", "value", "prior_report_value"])
        for r in reports:
            for var_key, var_data in r.table.items():
                for year, vals in var_data.items():
                    writer.writerow([r.date, var_key, year, vals["value"], vals.get("prior") or ""])


# --- Export Excel -----------------------------------------------------------

_XLSX_FONT_NAME = "Arial"
_XLSX_HEADER_FILL = PatternFill("solid", fgColor="7A1E3A")  # bordeaux BoE
_XLSX_HEADER_FONT = Font(name=_XLSX_FONT_NAME, size=10, bold=True, color="FFFFFF")
_XLSX_SUBHEADER_FILL = PatternFill("solid", fgColor="EAD3DA")
_XLSX_SUBHEADER_FONT = Font(name=_XLSX_FONT_NAME, size=10, bold=True)
_XLSX_CELL_FONT = Font(name=_XLSX_FONT_NAME, size=10)
_XLSX_THIN = Side(style="thin", color="B7B7B7")
_XLSX_BORDER = Border(left=_XLSX_THIN, right=_XLSX_THIN, top=_XLSX_THIN, bottom=_XLSX_THIN)
_XLSX_CENTER = Alignment(horizontal="center", vertical="center")

VAR_LABELS = {
    "gdp": "PIB (glissement annuel)",
    "cpi_inflation": "Inflation IPC",
    "unemployment_rate": "Taux de chômage",
    "output_gap": "Écart de production",
    "bank_rate": "Bank Rate (implicite marchés)*",
}


def _xlsx_style_header(cell, fill, font) -> None:
    cell.fill, cell.font, cell.alignment, cell.border = fill, font, _XLSX_CENTER, _XLSX_BORDER


def _xlsx_autofit(ws, min_width=10, max_width=24) -> None:
    for col_cells in ws.columns:
        length = max((len(str(c.value)) for c in col_cells if c.value is not None), default=0)
        col_letter = get_column_letter(col_cells[0].column)
        ws.column_dimensions[col_letter].width = min(max(length + 2, min_width), max_width)


def _fmt_cell(var_data: dict, year: str) -> str:
    vals = var_data.get(year)
    if not vals:
        return ""
    value, prior = vals.get("value"), vals.get("prior")
    return f"{value} ({prior})" if prior else (value or "")


def write_xlsx(reports: list[BoeReport], out_path: Path) -> None:
    if not reports:
        return
    reports = sorted(reports, key=lambda r: r.date)

    years: set = set()
    for r in reports:
        for var_data in r.table.values():
            years.update(var_data.keys())
    def _period_sort_key(period: str) -> tuple[int, int]:
        y, q = period.split(" Q")
        return int(y), int(q)

    years_sorted = sorted(years, key=_period_sort_key)

    var_keys = [k for k, _ in ROW_LABELS]

    wb = Workbook()
    ws = wb.active
    ws.title = "Résumé (Table 1.A)"
    ws.cell(row=1, column=1, value="Date du MPR")
    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)
    _xlsx_style_header(ws.cell(row=1, column=1), _XLSX_HEADER_FILL, _XLSX_HEADER_FONT)

    col = 2
    for var_key in var_keys:
        start_col = col
        for year in years_sorted:
            c = ws.cell(row=2, column=col, value=year)
            _xlsx_style_header(c, _XLSX_SUBHEADER_FILL, _XLSX_SUBHEADER_FONT)
            col += 1
        if col == start_col:
            c = ws.cell(row=2, column=col, value="N/D")
            _xlsx_style_header(c, _XLSX_SUBHEADER_FILL, _XLSX_SUBHEADER_FONT)
            col += 1
        if col - 1 > start_col:
            ws.merge_cells(start_row=1, start_column=start_col, end_row=1, end_column=col - 1)
        _xlsx_style_header(ws.cell(row=1, column=start_col, value=VAR_LABELS[var_key]),
                            _XLSX_HEADER_FILL, _XLSX_HEADER_FONT)

    row = 3
    for r in reports:
        cell = ws.cell(row=row, column=1, value=r.date)
        cell.font, cell.border, cell.alignment = _XLSX_CELL_FONT, _XLSX_BORDER, _XLSX_CENTER
        col = 2
        for var_key in var_keys:
            var_data = r.table.get(var_key, {})
            for year in years_sorted:
                c = ws.cell(row=row, column=col, value=_fmt_cell(var_data, year))
                c.font, c.border, c.alignment = _XLSX_CELL_FONT, _XLSX_BORDER, _XLSX_CENTER
                col += 1
        row += 1

    ws.freeze_panes = "B3"
    _xlsx_autofit(ws)
    note_row = row + 1
    ws.cell(row=note_row, column=1,
            value="* Bank Rate : chemin implicite par les marchés sur lequel les projections sont conditionnées "
                  "(la BoE ne prévoit pas sa propre trajectoire de taux).")
    ws.cell(row=note_row, column=1).font = Font(name=_XLSX_FONT_NAME, size=9, italic=True)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


def main() -> None:
    parser = argparse.ArgumentParser(description="Collecte les projections économiques (MPR) de la Banque d'Angleterre.")
    parser.add_argument("--out", default="boe_mpr.json", help="Chemin du fichier JSON de sortie.")
    parser.add_argument("--csv", default=None, help="Chemin optionnel d'export CSV à plat.")
    parser.add_argument("--xlsx", default=None,
                         help="Chemin de l'export Excel (.xlsx). Par défaut : même nom que --out, en .xlsx.")
    parser.add_argument("--no-xlsx", action="store_true", help="Ne pas générer le fichier Excel.")
    parser.add_argument("--since", type=int, default=None, help="Année minimale à inclure (ex: 2023).")
    parser.add_argument("--latest", action="store_true", help="Ne récupérer que le rapport le plus récent.")
    parser.add_argument("--cache-dir", default=None, help="Dossier de cache des pages HTML téléchargées.")
    parser.add_argument("--pause", type=float, default=1.0, help="Pause en secondes entre deux requêtes.")
    args = parser.parse_args()

    cache_dir = Path(args.cache_dir) if args.cache_dir else None

    print("Recherche des rapports MPR disponibles...")
    entries = discover_reports(cache_dir=cache_dir)

    if args.since:
        entries = [e for e in entries if int(e[0][:4]) >= args.since]
    if args.latest and entries:
        entries = [entries[-1]]

    if not entries:
        print("Aucun rapport trouvé avec ces critères.", file=sys.stderr)
        sys.exit(1)

    print(f"{len(entries)} rapport(s) à traiter : {', '.join(e[0] for e in entries)}")

    reports: list[BoeReport] = []
    for date, url in entries:
        print(f"-> {date} ...")
        report = scrape_report(date, url, cache_dir=cache_dir, pause=args.pause)
        if report:
            reports.append(report)

    out_path = Path(args.out)
    write_json(reports, out_path)
    print(f"JSON écrit : {out_path} ({len(reports)} rapports)")

    if args.csv:
        write_csv(reports, Path(args.csv))
        print(f"CSV écrit : {args.csv}")

    if not args.no_xlsx:
        xlsx_path = Path(args.xlsx) if args.xlsx else out_path.with_suffix(".xlsx")
        write_xlsx(reports, xlsx_path)
        print(f"Excel écrit : {xlsx_path}")


if __name__ == "__main__":
    main()
