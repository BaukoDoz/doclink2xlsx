#!/usr/bin/env python
"""
Extrahiert alle Weblinks aus einer PDF-Datei, ruft die Zielseiten ab und
schreibt den Seitentitel plus Link und PDF-Seite in eine Excel-Arbeitsmappe.
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Iterable, List, Tuple

import requests
from PyPDF2 import PdfReader
from PyPDF2.generic import ArrayObject, IndirectObject
from bs4 import BeautifulSoup
from openpyxl import Workbook

USER_AGENT = "Mozilla/5.0 (compatible; word_links_to_excel.py; +https://github.com/BaukoDoz)"


def extract_web_links(pdf_path: Path) -> List[Tuple[str, int]]:
    """Sucht alle externen Hyperlinks in einer PDF-Datei und merkt sich die Seite."""
    reader = PdfReader(str(pdf_path))
    hyperlinks: List[Tuple[str, int]] = []

    for page_number, page in enumerate(reader.pages, start=1):
        annotations = page.get("/Annots")
        if not annotations:
            continue

        if isinstance(annotations, IndirectObject):
            annotations = annotations.get_object()

        if isinstance(annotations, ArrayObject):
            annotation_refs = list(annotations)
        elif isinstance(annotations, (list, tuple)):
            annotation_refs = list(annotations)
        else:
            annotation_refs = [annotations]

        for annotation_ref in annotation_refs:
            if isinstance(annotation_ref, IndirectObject):
                annotation = annotation_ref.get_object()
            else:
                annotation = annotation_ref
            if not hasattr(annotation, "get"):
                continue
            if annotation.get("/Subtype") != "/Link":
                continue
            action = annotation.get("/A")
            if isinstance(action, IndirectObject):
                action = action.get_object()
            if not action or action.get("/S") != "/URI":
                continue
            uri = action.get("/URI")
            if not uri:
                continue
            hyperlinks.append((str(uri), page_number))

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
        description="Extrahiert Weblinks aus einer PDF-Datei und schreibt Seitentitel, Links und Seiten in eine XLSX-Datei."
    )
    parser.add_argument("input_pdf", type=Path, help="Pfad zur PDF-Datei (*.pdf)")
    parser.add_argument(
        "output_xlsx",
        type=Path,
        nargs="?",
        default=None,
        help="Optionaler Pfad zur Zieldatei (*.xlsx, Standard: PDF-Name mit .xlsx)",
    )
    parser.add_argument(
        "--timeout",
        type=float,
        default=10.0,
        help="Timeout fuer HTTP-Anfragen (Sekunden, Standard 10.0).",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    links = extract_web_links(args.input_pdf)
    rows = build_rows(links)
    output_path = args.output_xlsx if args.output_xlsx is not None else args.input_pdf.with_suffix(".pdf.xlsx")
    write_excel(rows, output_path)


if __name__ == "__main__":
    main()
