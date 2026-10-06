#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import argparse
import csv
import json
import os
import re
import shutil
import time
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Set, Tuple
from urllib.parse import quote_plus

import fitz  # PyMuPDF
import pandas as pd
import requests
from dotenv import load_dotenv


load_dotenv()

# -----------------------------------------------------------------------------
# PDF image filtering thresholds
# -----------------------------------------------------------------------------

MIN_WIDTH = 150
MIN_HEIGHT = 150
MIN_AREA = 40_000
MAX_ASPECT_RATIO = 8.0

# -----------------------------------------------------------------------------
# XML tag/regex constants
# -----------------------------------------------------------------------------

XML_ID = "{http://www.w3.org/XML/1998/namespace}id"
XLINK_HREF = "{http://www.w3.org/1999/xlink}href"

TABLE_LABEL_RE = re.compile(r"\b(?:Table|Tbl\.?)\s*([A-Za-z]?\d+[A-Za-z]?)\b", re.IGNORECASE)
TABLE_ID_RE = re.compile(r"\b(?:tbl|table|tb|t)[-_]?(0*\d+)\b", re.IGNORECASE)
TRAILING_NUMBER_RE = re.compile(r"(\d+)(?!.*\d)")
FIGURE_NUMBER_RE = re.compile(r"\b(?:Fig(?:ure)?\.?)\s*(\d+)\b", re.IGNORECASE)

SKIP_TEXT_TAGS = {
    "annotation",
    "annotation-xml",
    "display",
    "object-id",
}

BREAK_LIKE_TAGS = {
    "break",
    "br",
    "linebreak",
}

BODY_TEXT_EXCLUDED_TAGS = {
    "fig",
    "figure",
    "graphic",
    "inline-graphic",
    "caption",
    "table-wrap",
    "table-wrap-group",
    "table",
    "tgroup",
    "thead",
    "tbody",
    "tfoot",
    "tr",
    "th",
    "td",
    "row",
    "entry",
    "ref-list",
    "bibliography",
    "references",
    "disp-quote",
    "boxed-text",
    "app-group",
    "app",
    "appendix",
    "supplementary-material",
    "suppl",
}

TABLE_CONTAINER_TAGS = {"table-wrap", "table"}
TABLE_ROW_TAGS = {"tr", "row"}
TABLE_CELL_TAGS = {"th", "td", "entry"}
TABLE_SECTION_TAGS = {"thead", "tbody", "tfoot"}
TABLE_NOTE_TAGS = {
    "table-wrap-foot",
    "legend",
    "note",
    "notes",
    "fn",
    "footnote",
    "table-footnote",
    "table-fn",
}

SECTION_CONTAINER_TAGS = {"section", "sec", "subsection"}
SECTION_TITLE_TAGS = {"section-title", "sectitle"}
PARAGRAPH_TAGS = {"para", "p", "simple-para"}
ABSTRACT_TAGS = {"abstract", "abstract-sec"}


# -----------------------------------------------------------------------------
# General helpers
# -----------------------------------------------------------------------------


def _safe_str(v: Any) -> str:
    return "" if v is None else str(v)


def normalize_space(text: str) -> str:
    return re.sub(r"\s+", " ", _safe_str(text)).strip()


def sanitize_doi_filename(doi: str) -> str:
    return _safe_str(doi).replace("/", "_").replace(":", "_").strip()


def safe_paper_slug(paper: str) -> str:
    s = _safe_str(paper).strip()
    s = re.sub(r"[^\w\-.]+", "_", s)
    return s or "paper"


def _local_name(tag: str) -> str:
    if "}" in tag:
        tag = tag.split("}", 1)[1]
    if ":" in tag:
        tag = tag.split(":", 1)[1]
    return tag.lower()


def _elem_text(elem: ET.Element) -> str:
    """Return normalized text from an XML element while respecting simple breaks."""
    parts: List[str] = []

    def walk(node: ET.Element) -> None:
        lname = _local_name(node.tag)
        if lname in SKIP_TEXT_TAGS:
            return
        if node.text:
            parts.append(node.text)
        for child in list(node):
            child_lname = _local_name(child.tag)
            if child_lname in BREAK_LIKE_TAGS:
                parts.append(" ")
            walk(child)
            if child.tail:
                parts.append(child.tail)

    walk(elem)
    return normalize_space("".join(parts))


def _parent_map(root: ET.Element) -> Dict[ET.Element, ET.Element]:
    return {child: parent for parent in root.iter() for child in list(parent)}


def _parent_local_name(elem: ET.Element, parents: Dict[ET.Element, ET.Element]) -> str:
    parent = parents.get(elem)
    return _local_name(parent.tag) if parent is not None else ""


def _has_ancestor_named(
    elem: ET.Element,
    parents: Dict[ET.Element, ET.Element],
    names: Set[str],
) -> bool:
    parent = parents.get(elem)
    while parent is not None:
        if _local_name(parent.tag) in names:
            return True
        parent = parents.get(parent)
    return False


def _has_descendant_named(elem: ET.Element, names: Set[str]) -> bool:
    for child in elem.iter():
        if child is elem:
            continue
        if _local_name(child.tag) in names:
            return True
    return False


def _iter_descendants_named(elem: ET.Element, names: Set[str]) -> Iterable[ET.Element]:
    for child in elem.iter():
        if child is elem:
            continue
        if _local_name(child.tag) in names:
            yield child


def _direct_child_named(elem: ET.Element, names: Set[str]) -> Optional[ET.Element]:
    for child in list(elem):
        if _local_name(child.tag) in names:
            return child
    return None


def _direct_children_named(elem: ET.Element, names: Set[str]) -> List[ET.Element]:
    return [child for child in list(elem) if _local_name(child.tag) in names]


def _element_id(elem: ET.Element) -> str:
    return normalize_space(elem.get("id", "") or elem.get(XML_ID, ""))


def _dedupe_keep_order(items: Sequence[str]) -> List[str]:
    seen: Set[str] = set()
    out: List[str] = []
    for item in items:
        text = normalize_space(item)
        if not text or text in seen:
            continue
        seen.add(text)
        out.append(text)
    return out


# -----------------------------------------------------------------------------
# DOI mapping and Elsevier XML download
# -----------------------------------------------------------------------------


def read_paper_mapping(mapping_xlsx: str) -> List[Dict[str, str]]:
    if not os.path.exists(mapping_xlsx):
        raise FileNotFoundError(f"Mapping workbook not found: {mapping_xlsx}")

    df = pd.read_excel(mapping_xlsx)
    expected = {"Paper", "DOIs"}
    missing = expected.difference(df.columns)
    if missing:
        raise ValueError(
            f"Mapping workbook must contain columns {sorted(expected)}. Missing: {sorted(missing)}"
        )

    records: List[Dict[str, str]] = []
    for _, row in df.iterrows():
        paper = _safe_str(row.get("Paper", "")).strip()
        doi = _safe_str(row.get("DOIs", "")).strip()
        if not paper or paper.lower() == "nan":
            continue
        if doi.lower() == "nan":
            doi = ""
        records.append({"Paper": paper, "DOI": doi})
    return records


def get_elsevier_xml(
    mapping_records: List[Dict[str, str]],
    key: str,
    delay: int = 5,
    download_path: str = "XMLs/Elsevier",
    insttoken: Optional[str] = None,
    retries: int = 1,
    timeout: int = 30,
) -> None:
    headers = {
        "X-ELS-APIKey": key,
        "Accept": "text/xml",
    }
    if insttoken:
        headers["X-ELS-Insttoken"] = insttoken

    os.makedirs(download_path, exist_ok=True)
    base = "https://api.elsevier.com/content/article/doi/"

    print(f"Beginning to save XML files for {len(mapping_records)} record(s)...")
    for i, rec in enumerate(mapping_records):
        doi = _safe_str(rec.get("DOI", "")).strip()
        if not doi:
            print(f"[{i}] Skipping record with empty DOI for paper {rec.get('Paper', 'NR')}")
            continue

        safe_name = sanitize_doi_filename(doi)
        out_path = os.path.join(download_path, f"{safe_name}.xml")
        if os.path.exists(out_path):
            print(f"[{i}] Skipping {doi} (exists)")
            continue

        url_doc = base + quote_plus(doi)
        params = {"httpAccept": "text/xml"}

        attempt = 0
        while attempt <= retries:
            attempt += 1
            try:
                sr = requests.get(url=url_doc, headers=headers, params=params, timeout=timeout)
            except Exception as e:
                print(f"[{i}] Request error for {doi} (attempt {attempt}): {e}")
                if attempt > retries:
                    break
                time.sleep(1 + attempt)
                continue

            print(f"[{i}] {doi} -> HTTP {sr.status_code} (attempt {attempt})")

            if sr.status_code == 200:
                try:
                    with open(out_path, "w", encoding="utf-8") as f:
                        f.write(sr.text)
                    print(f"[{i}] Saved XML to {out_path} ({len(sr.text)} bytes)")
                except Exception as e:
                    print(f"[{i}] Error saving file {out_path}: {e}")
                break

            if sr.status_code in (401, 403):
                print(f"[{i}] AUTH ERROR {sr.status_code} for DOI {doi}.")
                print(sr.text[:1000].replace("\n", " "))
                break

            if sr.status_code == 404:
                print(f"[{i}] Not found (404) for DOI {doi}.")
                break

            print(f"[{i}] HTTP {sr.status_code}. Preview:")
            print(sr.text[:500])
            if attempt > retries:
                break
            wait = delay * attempt
            print(f"[{i}] Retrying in {wait}s...")
            time.sleep(wait)

        time.sleep(delay)

    print("Done downloading XML files.")


# -----------------------------------------------------------------------------
# XML title, DOI, figures, body text
# -----------------------------------------------------------------------------


def _looks_like_journal_line(s: str) -> bool:
    t = _safe_str(s).lower()
    return any(
        tok in t
        for tok in [
            "journal",
            "volume",
            "issue",
            "issn",
            "available online",
            "received",
            "accepted",
            "published by",
            "copyright",
        ]
    )


def _extract_xml_title(root: ET.Element) -> str:
    candidates: List[str] = []

    # Prefer more article-specific title tags when present.
    for elem in root.iter():
        lname = _local_name(elem.tag)
        if lname == "article-title":
            txt = _elem_text(elem)
            if txt:
                candidates.append(txt)

    for elem in root.iter():
        lname = _local_name(elem.tag)
        if lname == "title":
            txt = _elem_text(elem)
            if txt:
                candidates.append(txt)

    if not candidates:
        return "NR"

    candidates = _dedupe_keep_order(candidates)
    filtered = [c for c in candidates if not _looks_like_journal_line(c)]
    if filtered:
        candidates = filtered

    longish = [c for c in candidates if len(c.split()) >= 4]
    if longish:
        candidates = longish

    return candidates[0] if candidates else "NR"


def _extract_xml_doi(root: ET.Element, fallback: str = "") -> str:
    candidates: List[str] = []
    for elem in root.iter():
        lname = _local_name(elem.tag)
        if "doi" not in lname:
            continue
        txt = _elem_text(elem)
        if txt and "10." in txt:
            candidates.append(txt)
    return candidates[0] if candidates else fallback or "NR"


def _parse_figure_number(label: str, fig_id: str = "") -> Optional[int]:
    text = f"{label} {fig_id}".strip()
    match = FIGURE_NUMBER_RE.search(text)
    if match:
        try:
            return int(match.group(1))
        except ValueError:
            return None
    match = TRAILING_NUMBER_RE.search(fig_id)
    if match:
        try:
            return int(match.group(1))
        except ValueError:
            return None
    return None


def _extract_caption_text(container: ET.Element) -> str:
    caption = _direct_child_named(container, {"caption"})
    if caption is None:
        for child in _iter_descendants_named(container, {"caption"}):
            caption = child
            break
    if caption is None:
        return ""

    parts: List[str] = []
    for sub in caption.iter():
        lname = _local_name(sub.tag)
        if lname in {"simple-para", "para", "p", "title"}:
            text = _elem_text(sub)
            if text:
                parts.append(text)

    if parts:
        return normalize_space(" ".join(_dedupe_keep_order(parts)))
    return _elem_text(caption)


def _extract_xml_figures(root: ET.Element) -> List[Dict[str, Any]]:
    figures: List[Dict[str, Any]] = []
    order = 0

    for elem in root.iter():
        lname = _local_name(elem.tag)
        if lname not in {"figure", "fig"}:
            continue

        order += 1
        fig_id = _element_id(elem)
        label = ""
        graphic_href = ""

        label_elem = _direct_child_named(elem, {"label"})
        if label_elem is not None:
            label = _elem_text(label_elem)

        caption = _extract_caption_text(elem)

        for child in elem.iter():
            if _local_name(child.tag) in {"graphic", "inline-graphic"}:
                graphic_href = normalize_space(
                    child.get(XLINK_HREF, "")
                    or child.get("href", "")
                    or child.get("locator", "")
                    or child.get("src", "")
                )
                if graphic_href:
                    break

        figure_number = _parse_figure_number(label, fig_id)
        display_label = label or (f"Figure {figure_number}" if figure_number is not None else f"Figure {order}")
        full_caption = normalize_space(f"{display_label}. {caption}" if caption else display_label)
        raw_text = _elem_text(elem)

        figures.append(
            {
                "order": order,
                "figure_number": figure_number if figure_number is not None else "NR",
                "label": display_label or "NR",
                "caption": caption or "NR",
                "full_caption": full_caption or "NR",
                "graphic_href": graphic_href,
                "fig_id": fig_id,
                "raw_text": raw_text or "NR",
            }
        )

    return figures


# -----------------------------------------------------------------------------
# XML table extraction
# -----------------------------------------------------------------------------


def parse_table_number(raw_label: str, table_id: str = "") -> Optional[int]:
    label_norm = normalize_space(raw_label)
    if label_norm:
        match = TABLE_LABEL_RE.search(label_norm)
        if match:
            raw = match.group(1)
            digits = re.search(r"\d+", raw)
            if digits:
                try:
                    return int(digits.group(0))
                except ValueError:
                    return None
        match = re.search(r"\b(\d+)\b", label_norm)
        if match:
            try:
                return int(match.group(1))
            except ValueError:
                return None

    table_id_norm = normalize_space(table_id)
    if table_id_norm:
        match = TABLE_ID_RE.search(table_id_norm)
        if match:
            try:
                return int(match.group(1))
            except ValueError:
                return None
        match = TRAILING_NUMBER_RE.search(table_id_norm)
        if match:
            try:
                return int(match.group(1))
            except ValueError:
                return None

    return None


def _is_top_level_table_container(
    elem: ET.Element,
    parents: Dict[ET.Element, ET.Element],
) -> bool:
    lname = _local_name(elem.tag)
    if lname == "table-wrap":
        return True
    if lname != "table":
        return False

    # In JATS-like XML, <table> is commonly nested inside <table-wrap>. The
    # wrapper is the correct extraction unit, so skip the nested child table.
    if _has_ancestor_named(elem, parents, {"table-wrap", "table"}):
        return False

    # Elsevier often uses ce:table as the actual top-level table container.
    # Keep it only if it looks like a real data table.
    return _has_descendant_named(elem, TABLE_ROW_TAGS | TABLE_CELL_TAGS | {"caption", "label"})


def _extract_table_label(container: ET.Element) -> str:
    label_elem = _direct_child_named(container, {"label"})
    if label_elem is not None:
        return _elem_text(label_elem)

    # Fallback: first descendant label outside nested table rows/cells.
    for child in _iter_descendants_named(container, {"label"}):
        text = _elem_text(child)
        if text:
            return text
    return ""


def _extract_table_caption(container: ET.Element) -> str:
    caption = _direct_child_named(container, {"caption"})
    if caption is None:
        for child in _iter_descendants_named(container, {"caption"}):
            caption = child
            break
    if caption is None:
        return ""

    parts: List[str] = []
    for sub in caption.iter():
        lname = _local_name(sub.tag)
        if lname in {"simple-para", "para", "p", "title"}:
            text = _elem_text(sub)
            if text:
                parts.append(text)

    if parts:
        return normalize_space(" ".join(_dedupe_keep_order(parts)))
    return _elem_text(caption)


def _find_table_body_element(container: ET.Element) -> Optional[ET.Element]:
    if _local_name(container.tag) == "table":
        return container

    direct_table = _direct_child_named(container, {"table"})
    if direct_table is not None:
        return direct_table

    for child in _iter_descendants_named(container, {"table"}):
        return child
    return None


def _cell_text(cell: ET.Element) -> str:
    return _elem_text(cell)


def _extract_rows_from_section(section_elem: Optional[ET.Element]) -> List[List[str]]:
    rows: List[List[str]] = []
    if section_elem is None:
        return rows

    for row in section_elem.iter():
        if _local_name(row.tag) not in TABLE_ROW_TAGS:
            continue
        cells = [_cell_text(cell) for cell in list(row) if _local_name(cell.tag) in TABLE_CELL_TAGS]
        if cells and any(normalize_space(cell) for cell in cells):
            rows.append(cells)
    return rows


def _extract_all_rows(table_elem: ET.Element) -> List[List[str]]:
    rows: List[List[str]] = []
    for row in table_elem.iter():
        if _local_name(row.tag) not in TABLE_ROW_TAGS:
            continue
        cells = [_cell_text(cell) for cell in list(row) if _local_name(cell.tag) in TABLE_CELL_TAGS]
        if cells and any(normalize_space(cell) for cell in cells):
            rows.append(cells)
    return rows


def _extract_rows_not_in_sections(
    table_elem: ET.Element,
    parents: Dict[ET.Element, ET.Element],
    excluded_sections: Set[str],
) -> List[List[str]]:
    rows: List[List[str]] = []
    for row in table_elem.iter():
        if _local_name(row.tag) not in TABLE_ROW_TAGS:
            continue
        if _has_ancestor_named(row, parents, excluded_sections):
            continue
        cells = [_cell_text(cell) for cell in list(row) if _local_name(cell.tag) in TABLE_CELL_TAGS]
        if cells and any(normalize_space(cell) for cell in cells):
            rows.append(cells)
    return rows


def _extract_table_rows(
    table_elem: Optional[ET.Element],
    parents: Dict[ET.Element, ET.Element],
    infer_first_row_header: bool = True,
) -> Tuple[List[List[str]], List[List[str]], List[List[str]]]:
    if table_elem is None:
        return [], [], []

    thead_sections = list(_iter_descendants_named(table_elem, {"thead"}))
    tbody_sections = list(_iter_descendants_named(table_elem, {"tbody"}))
    tfoot_sections = list(_iter_descendants_named(table_elem, {"tfoot"}))

    header_rows: List[List[str]] = []
    body_rows: List[List[str]] = []
    footer_rows: List[List[str]] = []

    for section in thead_sections:
        header_rows.extend(_extract_rows_from_section(section))
    for section in tbody_sections:
        body_rows.extend(_extract_rows_from_section(section))
    for section in tfoot_sections:
        footer_rows.extend(_extract_rows_from_section(section))

    if header_rows or body_rows or footer_rows:
        if not body_rows:
            body_rows.extend(_extract_rows_not_in_sections(table_elem, parents, {"thead", "tfoot"}))
        return header_rows, body_rows, footer_rows

    # Fallback for tables without explicit thead/tbody/tfoot.
    all_rows = _extract_all_rows(table_elem)
    if infer_first_row_header and len(all_rows) > 1:
        return [all_rows[0]], all_rows[1:], []
    return [], all_rows, []


def _extract_table_notes(container: ET.Element) -> List[str]:
    notes: List[str] = []

    for note_elem in _iter_descendants_named(container, TABLE_NOTE_TAGS):
        # Do not treat table data cells as notes even if they contain note-like tags.
        text = _elem_text(note_elem)
        if not text:
            continue

        # For <fn>, keep label + note body if split into children.
        if _local_name(note_elem.tag) in {"fn", "footnote", "table-footnote", "table-fn"}:
            label = ""
            body_parts: List[str] = []
            for child in list(note_elem):
                cname = _local_name(child.tag)
                if cname == "label":
                    label = _elem_text(child)
                else:
                    child_text = _elem_text(child)
                    if child_text:
                        body_parts.append(child_text)
            note_text = normalize_space(" ".join(body_parts)) if body_parts else text
            if label and note_text.startswith(label):
                note_text = normalize_space(note_text[len(label):])
            if label and note_text:
                notes.append(f"{label}. {note_text}")
            elif note_text:
                notes.append(note_text)
        else:
            # Skip very large note containers that merely duplicate all nested fn text.
            child_note_count = sum(1 for _ in _iter_descendants_named(note_elem, {"fn", "footnote", "table-footnote", "table-fn"}))
            if child_note_count > 0 and _local_name(note_elem.tag) in {"table-wrap-foot", "notes"}:
                continue
            notes.append(text)

    return _dedupe_keep_order(notes)


def _format_table_text(table: Dict[str, Any]) -> str:
    lines: List[str] = []
    if table.get("full_caption"):
        lines.append(str(table["full_caption"]))

    header_rows = table.get("header_rows", []) or []
    body_rows = table.get("body_rows", []) or []
    footer_rows = table.get("footer_rows", []) or []
    notes = table.get("notes", []) or []

    for idx, row in enumerate(header_rows, start=1):
        row_text = " | ".join(normalize_space(cell) for cell in row if normalize_space(cell))
        if row_text:
            prefix = "Header" if len(header_rows) == 1 else f"Header {idx}"
            lines.append(f"{prefix}: {row_text}")

    for idx, row in enumerate(body_rows, start=1):
        row_text = " | ".join(normalize_space(cell) for cell in row if normalize_space(cell))
        if row_text:
            lines.append(f"Row {idx}: {row_text}")

    for idx, row in enumerate(footer_rows, start=1):
        row_text = " | ".join(normalize_space(cell) for cell in row if normalize_space(cell))
        if row_text:
            prefix = "Footer" if len(footer_rows) == 1 else f"Footer {idx}"
            lines.append(f"{prefix}: {row_text}")

    for note in notes:
        if normalize_space(note):
            lines.append(f"Note: {normalize_space(note)}")

    return "\n".join(lines).strip()


def extract_elsevier_tables(
    root: ET.Element,
    infer_first_row_header: bool = True,
) -> List[Dict[str, Any]]:
    tables: List[Dict[str, Any]] = []
    parents = _parent_map(root)
    seen_numbers: Set[int] = set()
    next_number = 1

    for elem in root.iter():
        if not _is_top_level_table_container(elem, parents):
            continue

        table_id = _element_id(elem)
        raw_label = _extract_table_label(elem)
        table_number = parse_table_number(raw_label, table_id)

        if table_number is None:
            while next_number in seen_numbers:
                next_number += 1
            table_number = next_number

        seen_numbers.add(table_number)
        next_number = max(next_number, table_number + 1)

        label = raw_label or f"Table {table_number}"
        title = _extract_table_caption(elem)

        if title and label and title.lower().startswith(label.lower()):
            full_caption = title
        else:
            full_caption = normalize_space(f"{label}. {title}" if title else label)

        table_elem = _find_table_body_element(elem)
        header_rows, body_rows, footer_rows = _extract_table_rows(
            table_elem,
            parents,
            infer_first_row_header=infer_first_row_header,
        )
        notes = _extract_table_notes(elem)

        row_records: List[Dict[str, Any]] = []
        row_index = 1
        for row in header_rows:
            row_records.append({"row_type": "header", "row_index": row_index, "cells": row})
            row_index += 1
        for row in body_rows:
            row_records.append({"row_type": "body", "row_index": row_index, "cells": row})
            row_index += 1
        for row in footer_rows:
            row_records.append({"row_type": "footer", "row_index": row_index, "cells": row})
            row_index += 1

        table: Dict[str, Any] = {
            "sort_index": len(tables) + 1,
            "table_number": table_number,
            "label": label,
            "title": title,
            "full_caption": full_caption,
            "table_id": table_id,
            "raw_label": raw_label,
            "header_rows": header_rows,
            "body_rows": body_rows,
            "footer_rows": footer_rows,
            "notes": notes,
            "rows": row_records,
        }
        table["table_text"] = _format_table_text(table)
        tables.append(table)

    tables.sort(key=lambda row: (row.get("table_number", 999999), row.get("sort_index", 999999)))
    return tables


# -----------------------------------------------------------------------------
# XML body text construction
# -----------------------------------------------------------------------------


def _extract_xml_body_blocks(root: ET.Element) -> List[str]:
    blocks: List[str] = []
    seen: Set[str] = set()
    parents = _parent_map(root)

    for elem in root.iter():
        lname = _local_name(elem.tag)
        if _has_ancestor_named(elem, parents, BODY_TEXT_EXCLUDED_TAGS):
            continue

        text = ""
        if lname in SECTION_TITLE_TAGS:
            text = _elem_text(elem)
        elif lname == "title" and _parent_local_name(elem, parents) in SECTION_CONTAINER_TAGS:
            text = _elem_text(elem)
        elif lname in PARAGRAPH_TAGS:
            text = _elem_text(elem)
        elif lname in ABSTRACT_TAGS:
            # Use abstract text only if it does not have paragraph children that will be
            # collected separately.
            has_para_child = any(_local_name(child.tag) in PARAGRAPH_TAGS for child in elem.iter() if child is not elem)
            if not has_para_child:
                text = _elem_text(elem)

        text = normalize_space(text)
        if not text:
            continue

        # Avoid obvious metadata/title noise and duplicate blocks.
        if text in seen:
            continue
        if lname in SECTION_TITLE_TAGS or (lname == "title" and len(text) < 200):
            blocks.append(text)
            seen.add(text)
        elif len(text) > 20:
            blocks.append(text)
            seen.add(text)

    return blocks


def _extract_xml_body_text(
    root: ET.Element,
    figures: List[Dict[str, Any]],
    tables: List[Dict[str, Any]],
) -> Tuple[List[str], str]:
    paragraphs: List[str] = []
    seen: Set[str] = set()

    title = _extract_xml_title(root)
    if title != "NR":
        paragraphs.append(title)
        seen.add(title)

    for block in _extract_xml_body_blocks(root):
        if block not in seen:
            paragraphs.append(block)
            seen.add(block)

    for fig in figures:
        caption = normalize_space(_safe_str(fig.get("caption", "")))
        full_caption = normalize_space(_safe_str(fig.get("full_caption", "")))
        fig_line = full_caption if full_caption and full_caption != "NR" else caption
        if fig_line and fig_line != "NR" and fig_line not in seen:
            paragraphs.append(fig_line)
            seen.add(fig_line)

    table_text_blocks = [t.get("table_text", "") for t in tables if normalize_space(t.get("table_text", ""))]
    if table_text_blocks:
        paragraphs.append("Tables")
        paragraphs.extend(table_text_blocks)

    full_text = "\n\n".join(paragraphs)
    return paragraphs, full_text


def parse_elsevier_xml(
    xml_path: str,
    paper_name: str,
    doi_fallback: str = "",
    infer_first_row_header: bool = True,
) -> Dict[str, Any]:
    try:
        tree = ET.parse(xml_path)
    except ET.ParseError as e:
        raise RuntimeError(f"Could not parse XML {xml_path}: {e}") from e

    root = tree.getroot()
    title = _extract_xml_title(root)
    doi = _extract_xml_doi(root, fallback=doi_fallback)
    figures = _extract_xml_figures(root)
    tables = extract_elsevier_tables(root, infer_first_row_header=infer_first_row_header)
    paragraphs, full_text = _extract_xml_body_text(root, figures, tables)

    return {
        "paper": paper_name,
        "xml_path": xml_path,
        "title": title,
        "doi": doi,
        "paragraphs": paragraphs,
        "full_text": full_text,
        "figures": figures,
        "tables": tables,
    }


# -----------------------------------------------------------------------------
# Output writers for XML artifacts
# -----------------------------------------------------------------------------


def _write_json(path: str, data: Any) -> None:
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def _write_text(path: str, text: str) -> None:
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)


def write_figure_reference_csv(path: str, figures: List[Dict[str, Any]]) -> None:
    with open(path, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=[
                "order",
                "figure_number",
                "label",
                "caption",
                "full_caption",
                "graphic_href",
                "fig_id",
                "raw_text",
            ],
        )
        writer.writeheader()
        for fig in figures:
            writer.writerow(
                {
                    "order": fig.get("order", "NR"),
                    "figure_number": fig.get("figure_number", "NR"),
                    "label": fig.get("label", "NR"),
                    "caption": fig.get("caption", "NR"),
                    "full_caption": fig.get("full_caption", "NR"),
                    "graphic_href": fig.get("graphic_href", ""),
                    "fig_id": fig.get("fig_id", ""),
                    "raw_text": fig.get("raw_text", "NR"),
                }
            )


def write_table_caption_files(prefix_dir: str, tables: List[Dict[str, Any]], basename: str = "xml_table_captions") -> None:
    txt_path = os.path.join(prefix_dir, f"{basename}.txt")
    csv_path = os.path.join(prefix_dir, f"{basename}.csv")

    with open(txt_path, "w", encoding="utf-8") as f:
        for row in tables:
            f.write(_safe_str(row.get("full_caption", "")) + "\n")

    with open(csv_path, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=["sort_index", "table_number", "label", "title", "full_caption", "table_id", "raw_label"],
        )
        writer.writeheader()
        for row in tables:
            writer.writerow(
                {
                    "sort_index": row.get("sort_index", ""),
                    "table_number": row.get("table_number", ""),
                    "label": row.get("label", ""),
                    "title": row.get("title", ""),
                    "full_caption": row.get("full_caption", ""),
                    "table_id": row.get("table_id", ""),
                    "raw_label": row.get("raw_label", ""),
                }
            )


def write_tables_text(path: str, tables: List[Dict[str, Any]]) -> None:
    with open(path, "w", encoding="utf-8") as f:
        for idx, row in enumerate(tables, start=1):
            if idx > 1:
                f.write("\n\n")
            f.write(_safe_str(row.get("table_text", "")))


def write_tables_rows_csv(path: str, tables: List[Dict[str, Any]]) -> None:
    max_cols = 0
    for table in tables:
        for row in table.get("rows", []) or []:
            max_cols = max(max_cols, len(row.get("cells", []) or []))

    fieldnames = [
        "table_number",
        "label",
        "title",
        "full_caption",
        "table_id",
        "row_type",
        "row_index",
        "cells_joined",
        "cells_json",
    ] + [f"cell_{i}" for i in range(1, max_cols + 1)]

    with open(path, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for table in tables:
            for row in table.get("rows", []) or []:
                cells = row.get("cells", []) or []
                record = {
                    "table_number": table.get("table_number", ""),
                    "label": table.get("label", ""),
                    "title": table.get("title", ""),
                    "full_caption": table.get("full_caption", ""),
                    "table_id": table.get("table_id", ""),
                    "row_type": row.get("row_type", ""),
                    "row_index": row.get("row_index", ""),
                    "cells_joined": " | ".join(_safe_str(cell) for cell in cells),
                    "cells_json": json.dumps(cells, ensure_ascii=False),
                }
                for idx, cell in enumerate(cells, start=1):
                    record[f"cell_{idx}"] = cell
                writer.writerow(record)


def save_parsed_xml_artifacts(parsed: Dict[str, Any], parsed_dir: str, paper_dir: str) -> None:
    os.makedirs(parsed_dir, exist_ok=True)
    os.makedirs(paper_dir, exist_ok=True)

    slug = safe_paper_slug(parsed.get("paper", "paper"))
    json_path = os.path.join(parsed_dir, f"{slug}.json")
    txt_path = os.path.join(parsed_dir, f"{slug}.txt")
    parsed_table_json = os.path.join(parsed_dir, f"{slug}_tables.json")
    parsed_table_rows_csv = os.path.join(parsed_dir, f"{slug}_tables_rows.csv")
    parsed_table_captions_csv = os.path.join(parsed_dir, f"{slug}_table_captions.csv")

    per_paper_text_path = os.path.join(paper_dir, "xml_article_text.txt")
    per_paper_fig_csv = os.path.join(paper_dir, "xml_figure_reference.csv")
    per_paper_tables_json = os.path.join(paper_dir, "xml_tables.json")
    per_paper_tables_text = os.path.join(paper_dir, "xml_tables_text.txt")
    per_paper_tables_rows_csv = os.path.join(paper_dir, "xml_tables_rows.csv")

    _write_json(json_path, parsed)
    _write_text(txt_path, parsed.get("full_text", ""))
    _write_text(per_paper_text_path, parsed.get("full_text", ""))

    figures = parsed.get("figures", []) or []
    tables = parsed.get("tables", []) or []

    write_figure_reference_csv(per_paper_fig_csv, figures)

    _write_json(parsed_table_json, tables)
    _write_json(per_paper_tables_json, tables)
    write_tables_text(per_paper_tables_text, tables)
    write_tables_rows_csv(parsed_table_rows_csv, tables)
    write_tables_rows_csv(per_paper_tables_rows_csv, tables)
    write_table_caption_files(paper_dir, tables, basename="xml_table_captions")

    # Also write corpus-level table captions CSV for convenience.
    with open(parsed_table_captions_csv, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=["sort_index", "table_number", "label", "title", "full_caption", "table_id", "raw_label"],
        )
        writer.writeheader()
        for row in tables:
            writer.writerow(
                {
                    "sort_index": row.get("sort_index", ""),
                    "table_number": row.get("table_number", ""),
                    "label": row.get("label", ""),
                    "title": row.get("title", ""),
                    "full_caption": row.get("full_caption", ""),
                    "table_id": row.get("table_id", ""),
                    "raw_label": row.get("raw_label", ""),
                }
            )


def build_xml_corpus(
    mapping_records: List[Dict[str, str]],
    xml_dir: str,
    parsed_dir: str,
    pdf_base_dir: str,
    infer_first_row_header: bool = True,
) -> Dict[str, Dict[str, Any]]:
    corpus: Dict[str, Dict[str, Any]] = {}
    for rec in mapping_records:
        paper = rec["Paper"]
        doi = rec.get("DOI", "")
        if not doi:
            print(f"  [warning] DOI missing for {paper}; skipping XML parse")
            continue

        xml_path = os.path.join(xml_dir, f"{sanitize_doi_filename(doi)}.xml")
        if not os.path.exists(xml_path):
            print(f"  [warning] XML missing for {paper}: {xml_path}")
            continue

        try:
            parsed = parse_elsevier_xml(
                xml_path,
                paper,
                doi_fallback=doi,
                infer_first_row_header=infer_first_row_header,
            )
            paper_dir = os.path.join(pdf_base_dir, paper)
            save_parsed_xml_artifacts(parsed, parsed_dir, paper_dir)
            corpus[paper] = parsed
            print(
                f"  [xml] Parsed {paper}: "
                f"{len(parsed.get('paragraphs', []))} text block(s), "
                f"{len(parsed.get('figures', []))} XML figure reference(s), "
                f"{len(parsed.get('tables', []))} XML table(s)"
            )
        except Exception as e:
            print(f"  [warning] Failed to parse XML for {paper}: {e}")
    return corpus


# -----------------------------------------------------------------------------
# PDF image extraction
# -----------------------------------------------------------------------------


def is_figure(base_image: Dict[str, Any]) -> Tuple[bool, str]:
    w = base_image.get("width", 0)
    h = base_image.get("height", 0)

    if w < MIN_WIDTH:
        return False, f"too narrow ({w} px wide, min {MIN_WIDTH})"
    if h < MIN_HEIGHT:
        return False, f"too short ({h} px tall, min {MIN_HEIGHT})"
    if w * h < MIN_AREA:
        return False, f"area too small ({w * h} px2, min {MIN_AREA})"

    aspect = max(w, h) / max(min(w, h), 1)
    if aspect > MAX_ASPECT_RATIO:
        return False, f"aspect ratio too extreme ({aspect:.1f}:1, max {MAX_ASPECT_RATIO})"

    return True, ""


def _clear_existing_images(images_dir: str) -> None:
    if not os.path.isdir(images_dir):
        return
    for name in os.listdir(images_dir):
        path = os.path.join(images_dir, name)
        if os.path.isfile(path):
            os.remove(path)
        elif os.path.isdir(path):
            shutil.rmtree(path)


def extract_images_only_from_pdf(
    pdf_path: str,
    images_dir: str,
    clear_existing: bool = False,
) -> List[Dict[str, Any]]:
    if clear_existing:
        _clear_existing_images(images_dir)
    os.makedirs(images_dir, exist_ok=True)

    doc = fitz.open(pdf_path)
    figure_num = 0
    skipped = 0
    seen_xrefs: Set[int] = set()
    manifest: List[Dict[str, Any]] = []

    for page_num, page in enumerate(doc, start=1):
        for img in page.get_images(full=True):
            xref = img[0]
            if xref in seen_xrefs:
                continue
            seen_xrefs.add(xref)

            try:
                base_image = doc.extract_image(xref)
            except Exception as e:
                print(f"    [warning] Could not extract image xref={xref} (page {page_num}): {e}")
                continue

            keep, reason = is_figure(base_image)
            if not keep:
                skipped += 1
                print(f"    [skip] page {page_num}, xref {xref} -- {reason}")
                continue

            img_bytes = base_image["image"]
            img_ext = base_image.get("ext", "png")
            figure_num += 1
            img_filename = f"figure_{figure_num}.{img_ext}"
            img_path = os.path.join(images_dir, img_filename)
            with open(img_path, "wb") as f:
                f.write(img_bytes)

            rec = {
                "figure_index": figure_num,
                "filename": img_filename,
                "page": page_num,
                "xref": xref,
                "width": base_image.get("width", 0),
                "height": base_image.get("height", 0),
                "path": img_path,
            }
            manifest.append(rec)
            print(f"    [saved] {img_filename} ({rec['width']} x {rec['height']} px, page {page_num})")

    doc.close()

    manifest_path = os.path.join(images_dir, "image_manifest.json")
    with open(manifest_path, "w", encoding="utf-8") as f:
        json.dump(manifest, f, indent=2)

    print(f"  -> {figure_num} figure(s) saved | {skipped} small/logo image(s) skipped")
    return manifest


def prepare_pdf_images(
    mapping_records: List[Dict[str, str]],
    pdf_base_dir: str,
    clear_images: bool = False,
) -> None:
    for rec in mapping_records:
        paper = rec["Paper"]
        folder_path = os.path.join(pdf_base_dir, paper)
        pdf_path = os.path.join(folder_path, f"{paper}.pdf")
        if not os.path.exists(pdf_path):
            print(f"  [warning] PDF missing for {paper}: {pdf_path}")
            continue
        images_dir = os.path.join(folder_path, "images")
        print(f"  [pdf] Extracting images for {paper}")
        extract_images_only_from_pdf(pdf_path, images_dir, clear_existing=clear_images)


# -----------------------------------------------------------------------------
# CLI
# -----------------------------------------------------------------------------


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Download Elsevier XML if requested, parse XML text/tables, and extract figure-like images from PDFs. "
            "The XML download behavior uses DOI values from an Excel workbook and Elsevier API credentials from .env."
        )
    )
    parser.add_argument(
        "--mapping-xlsx",
        default="Elsevier dois.xlsx",
        help="Excel workbook with columns 'Paper' and 'DOIs'.",
    )
    parser.add_argument(
        "--xml-dir",
        default="./XMLs/Elsevier",
        help="Directory containing downloaded Elsevier XML files.",
    )
    parser.add_argument(
        "--parsed-xml-dir",
        default="./XMLs/parsed",
        help="Directory for parsed XML JSON/TXT/table artifacts.",
    )
    parser.add_argument(
        "--pdf-base-dir",
        default="./PDFs",
        help="Base directory containing one folder per paper with the PDF inside.",
    )
    parser.add_argument(
        "--download-xml",
        action="store_true",
        help="Download Elsevier XML files before parsing.",
    )
    parser.add_argument(
        "--clear-images",
        action="store_true",
        help="Delete existing files inside each images folder before extracting fresh PDF images.",
    )
    parser.add_argument(
        "--skip-pdf-images",
        action="store_true",
        help="Skip PDF image extraction and only parse/download XML.",
    )
    parser.add_argument(
        "--no-infer-table-header",
        action="store_true",
        help="Do not treat the first row as a header when XML tables lack explicit thead/th tags.",
    )
    parser.add_argument(
        "--download-delay",
        type=int,
        default=5,
        help="Delay between Elsevier XML download requests in seconds.",
    )
    parser.add_argument(
        "--download-retries",
        type=int,
        default=1,
        help="Number of retries for Elsevier XML download.",
    )
    parser.add_argument(
        "--download-timeout",
        type=int,
        default=30,
        help="HTTP timeout for Elsevier XML download.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    mapping_records = read_paper_mapping(args.mapping_xlsx)

    if args.download_xml:
        api_key = os.getenv("ELSEVIER_API_KEY", "")
        inst_token = os.getenv("ELSEVIER_INSTTOKEN", "") or os.getenv("ELSEVIER_INSTOTKEN", "")
        if not api_key:
            raise RuntimeError("ELSEVIER_API_KEY is required when --download-xml is set.")
        get_elsevier_xml(
            mapping_records=mapping_records,
            key=api_key,
            delay=args.download_delay,
            download_path=args.xml_dir,
            insttoken=inst_token or None,
            retries=args.download_retries,
            timeout=args.download_timeout,
        )

    print("\nPreparing XML text and table artifacts...")
    build_xml_corpus(
        mapping_records=mapping_records,
        xml_dir=args.xml_dir,
        parsed_dir=args.parsed_xml_dir,
        pdf_base_dir=args.pdf_base_dir,
        infer_first_row_header=not args.no_infer_table_header,
    )

    if args.skip_pdf_images:
        print("\nSkipping PDF image extraction because --skip-pdf-images was set.")
    else:
        print("\nExtracting figure-like images from PDFs...")
        prepare_pdf_images(mapping_records, pdf_base_dir=args.pdf_base_dir, clear_images=args.clear_images)

    print("\nPreparation complete.")
    print("Review extracted PDF images manually and rename/delete files if needed so they match the article figure numbering.")


if __name__ == "__main__":
    main()
