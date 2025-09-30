#!/usr/bin/env python
"""
Extrahiert alle Weblinks aus einer DOCX-Datei, ruft die Zielseiten ab und
schreibt den Seitentitel plus Link und Word-Seite in eine Excel-Arbeitsmappe.
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Iterable, List, Tuple

import requests
from bs4 import BeautifulSoup
from docx import Document
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from openpyxl import Workbook

USER_AGENT = "Mozilla/5.0 (compatible; word_links_to_excel.py; +https://github.com/BaukoDoz)"


def extract_web_links(doc_path: Path) -> List[Tuple[str, int]]:
    """Liest alle externen Hyperlinks und merkt sich die zugehoerige Word-Seite."""
    document = Document(doc_path)
    rels = document.part.rels
    hyperlinks: List[Tuple[str, int]] = []

    page = 1
    for element in document.part.element.iter():
        tag = element.tag
        if tag == qn("w:br") and element.get(qn("w:type")) == "page":
            page += 1
            continue
        if tag == qn("w:lastRenderedPageBreak"):
            page += 1
            continue
        if tag == qn("w:pageBreakBefore"):
            page += 1
            continue
        if tag == qn("w:sectPr"):
            sect_type = element.find(qn("w:type"))
            if sect_type is not None and sect_type.get(qn("w:val")) in {"nextPage", "oddPage", "evenPage"}:
                page += 1
            continue
        if tag != qn("w:hyperlink"):
            continue

        rel_id = element.get(qn("r:id"))
        if not rel_id:
            continue
        rel = rels.get(rel_id)
        if not rel or rel.reltype != RT.HYPERLINK:
            continue
        hyperlinks.append((rel.target_ref, page))

    return hyperlinks


def fetch_page_title(url: str, timeout: float = 10.0) -> str:
    """Laedt eine URL und liefert den Text des title-Tags (oder eine leere Zeichenkette)."""
    headers = {"User-Agent": USER_AGENT}
    try:
        response = requests.get(url, headers=headers, timeout=timeout)
        response.raise_for_status()
    except requests.RequestException:
        return ""

    soup = BeautifulSoup(response.text, "html.parser")
    title_tag = soup.find("title")
    if title_tag is None:
        return ""
    return title_tag.get_text(" ", strip=True)


def build_rows(links: Iterable[Tuple[str, int]]) -> List[Tuple[str, str, int]]:
    """Erzeugt (Seitentitel, URL, Seite)-Tupel fuer alle Links."""
    rows: List[Tuple[str, str, int]] = []
    for url, page in links:
        title_text = fetch_page_title(url)
        rows.append((title_text, url, page))
    return rows


def write_excel(rows: Iterable[Tuple[str, str, int]], output_path: Path) -> None:
    """Schreibt die Zeilen in eine neue Excel-Datei mit drei Spalten."""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Links"
    sheet.append(["Seitentitel", "Link", "Seite"])

    for title_text, url, page in rows:
        sheet.append([title_text, url, page])

    workbook.save(output_path)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Extrahiert Weblinks aus einer DOCX-Datei und schreibt Seitentitel, Links und Seiten in eine XLSX-Datei."
    )
    parser.add_argument("input_docx", type=Path, help="Pfad zur Word-Datei (*.docx)")
    parser.add_argument("output_xlsx", type=Path, nargs="?", default=None, help="Optionaler Pfad zur Zieldatei (*.xlsx, Standard: DOCX-Name mit .xlsx)")
    parser.add_argument(
        "--timeout",
        type=float,
        default=10.0,
        help="Timeout fuer HTTP-Anfragen (Sekunden, Standard 10.0).",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    links = extract_web_links(args.input_docx)
    rows = build_rows(links)
    output_path = args.output_xlsx if args.output_xlsx is not None else args.input_docx.with_suffix(".docx.xlsx")
    write_excel(rows, output_path)


if __name__ == "__main__":
    main()

