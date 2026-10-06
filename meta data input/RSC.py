# -*- coding: utf-8 -*-
"""
Created on Tue Sep 15 15:52:34 2026

@author: RDASH
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import re
import shutil
import tempfile
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

# Optional Windows workaround for duplicate OpenMP runtimes.
os.environ.setdefault("KMP_DUPLICATE_LIB_OK", "TRUE")

try:
    from bs4 import BeautifulSoup, NavigableString, Tag
except ImportError as e:
    raise ImportError(
        "This script requires beautifulsoup4. Install it with: pip install beautifulsoup4"
    ) from e

try:
    from PIL import Image
except ImportError as e:
    raise ImportError(
        "This script requires Pillow. Install it with: pip install pillow"
    ) from e

try:
    from chemdataextractor import Document
    from chemdataextractor.reader import RscHtmlReader
except ImportError as e:
    raise ImportError(
        "This script requires chemdataextractor2 / chemdataextractor with RscHtmlReader. "
        "Install it first, for example: pip install chemdataextractor2"
    ) from e


FIGURE_NUMBER_RE = re.compile(r"\b(?:Fig(?:ure)?\.?)\s*(\d+)\b", re.IGNORECASE)
LOCAL_FIGURE_FILE_RE = re.compile(r"(?i)-f(\d+)\.(gif|png|jpe?g|tiff?|bmp|webp)$")
TABLE_NUMBER_RE = re.compile(r"\bTable\s*(\d+)\b", re.IGNORECASE)
TABLE_ID_RE = re.compile(r"(?:^|\b)tab(\d+)(?:\b|$)", re.IGNORECASE)
GENERATED_HTML_FILES = [
    "sentences.txt",
    "fulltext.txt",
    "figure_captions.txt",
    "figure_captions.csv",
    "table_captions.txt",
    "table_captions.csv",
    "tables.json",
    "tables_text.txt",
    "tables_rows.csv",
]


# -----------------------------------------------------------------------------
# General helpers
# -----------------------------------------------------------------------------

def normalize_space(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def iter_paper_html_files(html_root: Path) -> Iterable[Tuple[str, Path]]:
    """
    Yield (paper_name, html_path) for folders like:
      HTMLs/Paper 1/Paper 1.html
    """
    for paper_dir in sorted((p for p in html_root.iterdir() if p.is_dir()), key=lambda p: p.name):
        paper_name = paper_dir.name
        html_path = paper_dir / f"{paper_name}.html"
        if html_path.exists():
            yield paper_name, html_path
        else:
            htmls = sorted(paper_dir.glob("*.html"))
            if htmls:
                yield paper_name, htmls[0]


def cleanup_existing_html_outputs(output_dir: Path) -> None:
    """
    Delete only HTML-side generated text artifacts. Do not touch PDF images.
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    for filename in GENERATED_HTML_FILES:
        path = output_dir / filename
        if path.exists() and path.is_file():
            path.unlink(missing_ok=True)


def find_saved_resource_dir(paper_name: str, html_path: Path) -> Optional[Path]:
    """Return the browser-downloaded sidecar folder, normally Paper X_files."""
    preferred = html_path.parent / f"{paper_name}_files"
    if preferred.is_dir():
        return preferred
    alternatives = sorted(
        p for p in html_path.parent.iterdir()
        if p.is_dir() and p.name.lower().endswith("_files")
    )
    return alternatives[0] if alternatives else None


def reset_images_dir(pdf_root: Path, paper_name: str) -> Path:
    """Recreate PDFs/Paper X/images. This intentionally removes old images."""
    images_dir = pdf_root / paper_name / "images"
    if images_dir.exists():
        shutil.rmtree(images_dir, ignore_errors=True)
    images_dir.mkdir(parents=True, exist_ok=True)
    return images_dir


# -----------------------------------------------------------------------------
# Text extraction using ChemDataExtractor
# -----------------------------------------------------------------------------

def extract_title(doc: Document) -> str:
    for element in doc.elements:
        if element.__class__.__name__ == "Title":
            title = normalize_space(str(element))
            if title:
                return title
    return "NR"


def find_body_bounds(doc: Document) -> Tuple[int, int]:
    """
    Start after Abstract when possible.
    End just before Citation when possible.
    """
    start = 0
    end = len(doc.elements) - 1

    abstract_idx = next(
        (i for i, element in enumerate(doc.elements) if normalize_space(str(element)) == "Abstract"),
        None,
    )
    if abstract_idx is not None:
        start = abstract_idx + 1

    citation_idx = next(
        (i for i, element in enumerate(doc.elements) if element.__class__.__name__ == "Citation"),
        None,
    )
    if citation_idx is not None and citation_idx > start:
        end = citation_idx - 1

    return start, end


def get_fulltext_from_elements(
    elements: Sequence[object],
    start: int,
    end: int,
    include_figures: bool = False,
    include_tables: bool = False,
) -> str:
    parts: List[str] = []

    for i in range(start, end + 1):
        element = elements[i]
        cls = element.__class__.__name__

        if cls == "Figure" and not include_figures:
            continue
        if cls == "Table" and not include_tables:
            continue

        text = normalize_space(str(element))
        if text:
            parts.append(text)

    return " ".join(parts)


def sentence_split_with_cde(text: str) -> List[str]:
    """
    Use CDE on a temporary plain-text file, matching the previous workflow.
    """
    if not text.strip():
        return []

    with tempfile.NamedTemporaryFile("w", suffix=".txt", encoding="utf-8", delete=False) as tmp:
        tmp.write(text)
        temp_path = Path(tmp.name)

    try:
        doc = Document.from_file(str(temp_path))
        if not doc.elements:
            return []

        sentences: List[str] = []
        for element in doc.elements:
            if hasattr(element, "sentences"):
                for sent in element.sentences:
                    s = normalize_space(str(sent))
                    if s:
                        sentences.append(s)

        if sentences:
            return sentences

        return [normalize_space(text)]
    finally:
        try:
            temp_path.unlink(missing_ok=True)
        except Exception:
            pass


def merge_very_short_sentences(sentences: List[str], min_len: int = 10) -> List[str]:
    """
    CatMiner-style cleanup: join tiny fragments with the next sentence.
    """
    if not sentences:
        return []

    merged: List[str] = []
    skip_next = False

    for i, sent in enumerate(sentences):
        if skip_next:
            skip_next = False
            continue

        s = normalize_space(sent)
        if not s:
            continue

        if len(s) < min_len and i + 1 < len(sentences):
            nxt = normalize_space(sentences[i + 1])
            if nxt:
                merged.append(f"{s} {nxt}")
                skip_next = True
                continue

        merged.append(s)

    return merged


def parse_rsc_html_text(html_path: Path) -> Tuple[str, str, List[str]]:
    """
    Returns: (title, base_full_text, base_sentences)

    Tables are intentionally excluded here because the saved RSC HTML tables are
    handled separately and appended in a controlled way.
    """
    doc = Document.from_file(str(html_path), readers=[RscHtmlReader()])
    title = extract_title(doc)
    start, end = find_body_bounds(doc)
    full_text = get_fulltext_from_elements(doc.elements, start, end)
    sentences = sentence_split_with_cde(full_text)
    sentences = merge_very_short_sentences(sentences)
    return title, full_text, sentences


# -----------------------------------------------------------------------------
# Figure caption extraction from saved RSC HTML
# -----------------------------------------------------------------------------

def parse_main_figure_number(text: str) -> Optional[int]:
    if not text:
        return None
    match = FIGURE_NUMBER_RE.search(text)
    if match:
        try:
            return int(match.group(1))
        except ValueError:
            return None
    return None


def extract_figure_captions_from_html(html_path: Path) -> List[Dict[str, Any]]:
    """
    Extract main-article figure captions from saved RSC HTML.

    RSC saved HTML commonly stores captions like:
        <td class="image_title"><b>Fig. 1 </b><span id="fig1">
            <span class="graphic_title">Caption text...</span>
        </span></td>
    """
    html_text = html_path.read_text(encoding="utf-8", errors="ignore")
    soup = BeautifulSoup(html_text, "html.parser")

    results: List[Dict[str, Any]] = []
    seen_keys: set[Tuple[Any, str]] = set()

    for container in soup.select("div.image_table"):
        title_cell = container.select_one("td.image_title")
        if title_cell is None:
            continue

        label_node = title_cell.find("b")
        label = normalize_space(label_node.get_text(" ", strip=True)) if label_node else ""

        caption_parts: List[str] = []
        for node in title_cell.select("span.graphic_title"):
            text = normalize_space(node.get_text(" ", strip=True))
            if text:
                caption_parts.append(text)

        if not caption_parts:
            full_cell_text = normalize_space(title_cell.get_text(" ", strip=True))
            if label and full_cell_text.startswith(label):
                fallback_caption = normalize_space(full_cell_text[len(label):])
            else:
                fallback_caption = full_cell_text
            if fallback_caption:
                caption_parts.append(fallback_caption)

        caption_text = normalize_space(" ".join(caption_parts))
        if not label and not caption_text:
            continue

        figure_number = parse_main_figure_number(label)
        if figure_number is None:
            id_texts: List[str] = []
            for node in container.select("[id]"):
                node_id = normalize_space(str(node.get("id", "")))
                if node_id:
                    id_texts.append(node_id)
            for node_id in id_texts:
                match = re.search(r"(?:imgfig|fig)(\d+)", node_id, flags=re.IGNORECASE)
                if match:
                    figure_number = int(match.group(1))
                    break

        if figure_number is None:
            continue

        full_caption = normalize_space(f"{label} {caption_text}") if label else caption_text
        key = (figure_number, full_caption)
        if key in seen_keys:
            continue
        seen_keys.add(key)

        results.append(
            {
                "figure_number": figure_number,
                "label": label or f"Fig. {figure_number}",
                "caption": caption_text,
                "full_caption": full_caption,
            }
        )

    if not results:
        for fig in soup.select("figure"):
            text = normalize_space(fig.get_text(" ", strip=True))
            if not text:
                continue
            figure_number = parse_main_figure_number(text)
            if figure_number is None:
                continue
            label_match = re.search(r"\b(?:Fig(?:ure)?\.?)\s*\d+\b", text, re.IGNORECASE)
            label = label_match.group(0) if label_match else f"Fig. {figure_number}"
            caption = normalize_space(text.replace(label, "", 1))
            full_caption = normalize_space(f"{label} {caption}")
            key = (figure_number, full_caption)
            if key in seen_keys:
                continue
            seen_keys.add(key)
            results.append(
                {
                    "figure_number": figure_number,
                    "label": label,
                    "caption": caption,
                    "full_caption": full_caption,
                }
            )

    results.sort(key=lambda row: row["figure_number"])
    return results


def write_figure_caption_files(output_dir: Path, captions: List[Dict[str, Any]]) -> None:
    txt_path = output_dir / "figure_captions.txt"
    with open(txt_path, "w", encoding="utf-8") as f:
        for row in captions:
            f.write(row.get("full_caption", "") + "\n")

    csv_path = output_dir / "figure_captions.csv"
    with open(csv_path, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["figure_number", "label", "caption", "full_caption"])
        writer.writeheader()
        for row in captions:
            writer.writerow(
                {
                    "figure_number": row.get("figure_number", ""),
                    "label": row.get("label", ""),
                    "caption": row.get("caption", ""),
                    "full_caption": row.get("full_caption", ""),
                }
            )


# -----------------------------------------------------------------------------
# Table extraction from saved RSC HTML
# -----------------------------------------------------------------------------

def parse_main_table_number(text: str) -> Optional[int]:
    if not text:
        return None
    match = TABLE_NUMBER_RE.search(text)
    if match:
        try:
            return int(match.group(1))
        except ValueError:
            return None
    return None


def _tr_cells(tr: Tag) -> List[Tag]:
    cells = tr.find_all(["th", "td"], recursive=False)
    return cells if cells else tr.find_all(["th", "td"])


def _cell_text(cell: Tag) -> str:
    return normalize_space(cell.get_text(" ", strip=True))


def _parse_span_value(value: Any, default: int = 1) -> int:
    try:
        n = int(str(value))
        return n if n > 0 else default
    except Exception:
        return default


def _expand_rows(tr_tags: List[Tag]) -> List[List[str]]:
    """
    Expand rowspan/colspan into a rectangular text grid.
    """
    grid: List[List[str]] = []
    pending_rowspans: Dict[int, Tuple[str, int]] = {}
    max_cols = 0

    for tr in tr_tags:
        row: List[str] = []
        col_idx = 0

        def flush_pending() -> None:
            nonlocal col_idx
            while col_idx in pending_rowspans:
                value, remaining = pending_rowspans[col_idx]
                row.append(value)
                if remaining > 1:
                    pending_rowspans[col_idx] = (value, remaining - 1)
                else:
                    del pending_rowspans[col_idx]
                col_idx += 1

        flush_pending()

        for cell in _tr_cells(tr):
            flush_pending()
            text = _cell_text(cell)
            rowspan = _parse_span_value(cell.get("rowspan"), 1)
            colspan = _parse_span_value(cell.get("colspan"), 1)

            for offset in range(colspan):
                row.append(text)
                if rowspan > 1:
                    pending_rowspans[col_idx + offset] = (text, rowspan - 1)
            col_idx += colspan

        flush_pending()

        max_cols = max(max_cols, len(row))
        grid.append(row)

    for row in grid:
        if len(row) < max_cols:
            row.extend([""] * (max_cols - len(row)))

    return grid


def _flatten_header_rows(header_rows: List[List[str]], ncols: int) -> List[str]:
    if not header_rows:
        return [f"Column {i + 1}" for i in range(ncols)]

    flattened: List[str] = []
    for col_idx in range(ncols):
        parts: List[str] = []
        for row in header_rows:
            if col_idx < len(row):
                text = normalize_space(row[col_idx])
                if text and (not parts or text != parts[-1]):
                    parts.append(text)
        flattened.append(" / ".join(parts) if parts else f"Column {col_idx + 1}")
    return flattened


def _find_associated_data_table(caption_div: Tag) -> Optional[Tag]:
    """
    Starting from a <div class="table_caption"> block, find the next real data
    table. Skip equation tables, figure tables, and other boilerplate.
    """
    sibling: Any = caption_div
    steps = 0

    while steps < 10:
        sibling = sibling.find_next_sibling()
        if sibling is None:
            break
        if isinstance(sibling, NavigableString):
            continue
        if not isinstance(sibling, Tag):
            continue

        # If we hit a new section or a figure before finding a data table, stop.
        classes = set(sibling.get("class", []))
        if sibling.name in {"h1", "h2", "h3", "h4", "h5"}:
            break
        if "image_table" in classes:
            break
        if "table_caption" in classes and sibling is not caption_div:
            break

        table_tag: Optional[Tag] = None
        if sibling.name == "table":
            table_tag = sibling
        else:
            table_tag = sibling.select_one("table.tgroup.rtable, table.rtable, table")

        if table_tag is not None:
            text = normalize_space(table_tag.get_text(" ", strip=True))
            table_classes = set(table_tag.get("class", []))
            if (
                "rtable" in table_classes
                or "tgroup" in table_classes
                or table_tag.find_parent("div", class_="rtable__wrapper") is not None
            ):
                if text:
                    return table_tag

        steps += 1

    return None


def _extract_table_grid(table_tag: Tag) -> Tuple[List[str], List[List[str]], List[List[str]]]:
    """
    Returns (column_names, header_rows, body_rows).
    """
    thead = table_tag.find("thead")
    tbody = table_tag.find("tbody")

    header_trs: List[Tag] = []
    body_trs: List[Tag] = []

    if thead is not None:
        header_trs = thead.find_all("tr")
    if tbody is not None:
        body_trs = tbody.find_all("tr")

    if not header_trs and not body_trs:
        all_trs = table_tag.find_all("tr")
        body_started = False
        for tr in all_trs:
            cells = _tr_cells(tr)
            if not cells:
                continue
            cell_names = [cell.name for cell in cells]
            if not body_started and all(name == "th" for name in cell_names):
                header_trs.append(tr)
            else:
                body_started = True
                body_trs.append(tr)

    header_rows = _expand_rows(header_trs) if header_trs else []
    body_rows = _expand_rows(body_trs) if body_trs else []

    ncols = 0
    if header_rows:
        ncols = max(ncols, max((len(r) for r in header_rows), default=0))
    if body_rows:
        ncols = max(ncols, max((len(r) for r in body_rows), default=0))

    if ncols == 0:
        return [], [], []

    for rows in (header_rows, body_rows):
        for row in rows:
            if len(row) < ncols:
                row.extend([""] * (ncols - len(row)))

    column_names = _flatten_header_rows(header_rows, ncols)
    return column_names, header_rows, body_rows


def extract_tables_from_html(html_path: Path) -> List[Dict[str, Any]]:
    """
    Extract main-article numbered tables from saved RSC HTML.

    Real data tables in the uploaded RSC HTML appear as:
      <div class="table_caption"><b>Table 1</b> <span id="tab1">...</span></div>
      <div class="rtable__wrapper"><div class="rtable__inner">...<table class="tgroup rtable">...

    Equation images are also stored inside bare <table> tags, so we rely on the
    numbered table-caption blocks and then pick the next matching rtable.
    """
    html_text = html_path.read_text(encoding="utf-8", errors="ignore")
    soup = BeautifulSoup(html_text, "html.parser")

    results: List[Dict[str, Any]] = []
    seen_numbers: set[int] = set()

    for caption_div in soup.select("div.table_caption"):
        label_node = caption_div.find("b")
        label = normalize_space(label_node.get_text(" ", strip=True)) if label_node else ""

        table_number = parse_main_table_number(label)
        if table_number is None:
            for node in caption_div.select("[id]"):
                node_id = normalize_space(str(node.get("id", "")))
                match = TABLE_ID_RE.search(node_id)
                if match:
                    table_number = int(match.group(1))
                    break

        if table_number is None:
            continue
        if table_number in seen_numbers:
            continue

        full_caption_text = normalize_space(caption_div.get_text(" ", strip=True))
        if label and full_caption_text.startswith(label):
            caption_text = normalize_space(full_caption_text[len(label):])
        else:
            caption_text = full_caption_text
        full_caption = normalize_space(f"{label} {caption_text}") if label else caption_text

        table_tag = _find_associated_data_table(caption_div)
        if table_tag is None:
            continue

        column_names, header_rows, body_rows = _extract_table_grid(table_tag)
        if not column_names and not body_rows:
            continue

        seen_numbers.add(table_number)
        results.append(
            {
                "table_number": table_number,
                "label": label or f"Table {table_number}",
                "caption": caption_text,
                "full_caption": full_caption,
                "column_names": column_names,
                "header_rows": header_rows,
                "rows": body_rows,
            }
        )

    results.sort(key=lambda row: row["table_number"])
    return results


def table_to_text_block(table: Dict[str, Any]) -> str:
    lines: List[str] = [table.get("full_caption", "").strip()]

    column_names = [normalize_space(str(x)) for x in table.get("column_names", []) if normalize_space(str(x))]
    if column_names:
        lines.append(" | ".join(column_names))

    for row in table.get("rows", []):
        clean_row = [normalize_space(str(cell)) for cell in row]
        if any(clean_row):
            lines.append(" | ".join(clean_row))

    return "\n".join(line for line in lines if line)


def build_fulltext_with_tables(base_full_text: str, tables: List[Dict[str, Any]]) -> str:
    base = normalize_space(base_full_text)
    table_blocks = [table_to_text_block(table) for table in tables if table_to_text_block(table).strip()]
    if not table_blocks:
        return base
    if base:
        return base + "\n\n" + "\n\n".join(table_blocks)
    return "\n\n".join(table_blocks)


def build_sentences_with_tables(base_sentences: List[str], tables: List[Dict[str, Any]]) -> List[str]:
    lines: List[str] = list(base_sentences)

    for table in tables:
        full_caption = normalize_space(table.get("full_caption", ""))
        if full_caption:
            lines.append(full_caption)

        column_names = [normalize_space(str(x)) for x in table.get("column_names", []) if normalize_space(str(x))]
        if column_names:
            lines.append(f"{table.get('label', 'Table')} columns: " + " | ".join(column_names))

        for idx, row in enumerate(table.get("rows", []), start=1):
            clean_row = [normalize_space(str(cell)) for cell in row]
            if any(clean_row):
                lines.append(f"{table.get('label', 'Table')} row {idx}: " + " | ".join(clean_row))

    return lines


def write_table_caption_files(output_dir: Path, tables: List[Dict[str, Any]]) -> None:
    txt_path = output_dir / "table_captions.txt"
    with open(txt_path, "w", encoding="utf-8") as f:
        for table in tables:
            f.write(table.get("full_caption", "") + "\n")

    csv_path = output_dir / "table_captions.csv"
    with open(csv_path, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["table_number", "label", "caption", "full_caption"])
        writer.writeheader()
        for table in tables:
            writer.writerow(
                {
                    "table_number": table.get("table_number", ""),
                    "label": table.get("label", ""),
                    "caption": table.get("caption", ""),
                    "full_caption": table.get("full_caption", ""),
                }
            )


def write_table_files(output_dir: Path, tables: List[Dict[str, Any]]) -> None:
    json_path = output_dir / "tables.json"
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(tables, f, ensure_ascii=False, indent=2)

    text_path = output_dir / "tables_text.txt"
    with open(text_path, "w", encoding="utf-8") as f:
        for i, table in enumerate(tables):
            if i:
                f.write("\n\n")
            f.write(table_to_text_block(table))

    rows_csv_path = output_dir / "tables_rows.csv"
    with open(rows_csv_path, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=[
                "table_number",
                "label",
                "caption",
                "row_index",
                "row_text",
                "cells_json",
            ],
        )
        writer.writeheader()
        for table in tables:
            for row_index, row in enumerate(table.get("rows", []), start=1):
                clean_row = [normalize_space(str(cell)) for cell in row]
                if not any(clean_row):
                    continue
                writer.writerow(
                    {
                        "table_number": table.get("table_number", ""),
                        "label": table.get("label", ""),
                        "caption": table.get("caption", ""),
                        "row_index": row_index,
                        "row_text": " | ".join(clean_row),
                        "cells_json": json.dumps(clean_row, ensure_ascii=False),
                    }
                )


# -----------------------------------------------------------------------------
# Figure-image extraction from local HTML resources
# -----------------------------------------------------------------------------

def collect_local_figure_files(resources_dir: Path) -> List[Tuple[int, Path]]:
    """Find *-fN image files; for duplicate N values keep the largest file."""
    by_number: Dict[int, Path] = {}
    for path in resources_dir.iterdir():
        if not path.is_file():
            continue
        match = LOCAL_FIGURE_FILE_RE.search(path.name)
        if not match:
            continue
        figure_number = int(match.group(1))
        current = by_number.get(figure_number)
        if current is None or path.stat().st_size > current.stat().st_size:
            by_number[figure_number] = path
    return [(number, by_number[number]) for number in sorted(by_number)]


def image_file_to_jpeg(input_path: Path, output_path: Path) -> Tuple[int, int]:
    """Convert an image to JPEG, compositing transparent pixels onto white."""
    with Image.open(input_path) as image:
        image.load()
        transparent = image.mode in ("RGBA", "LA") or (
            image.mode == "P" and "transparency" in image.info
        )
        if transparent:
            rgba = image.convert("RGBA")
            result = Image.new("RGB", rgba.size, "white")
            result.paste(rgba, mask=rgba.getchannel("A"))
        else:
            result = image.convert("RGB")
        result.save(output_path, format="JPEG", quality=95)
        return result.size


def write_image_manifest(images_dir: Path, manifest: List[Dict[str, Any]]) -> None:
    with open(images_dir / "image_manifest.json", "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    fields = [
        "figure_index", "filename", "source", "original_file", "original_extension",
        "width", "height", "caption", "caption_match_method", "status", "error",
    ]
    with open(images_dir / "image_manifest.csv", "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        for row in manifest:
            writer.writerow({field: row.get(field, "") for field in fields})


def extract_figures_from_saved_html_resources(
    resources_dir: Path, images_dir: Path, captions: List[Dict[str, Any]]
) -> List[Dict[str, Any]]:
    """Convert locally saved figure files and link captions by figure number."""
    captions_by_number = {row["figure_number"]: row for row in captions}
    manifest: List[Dict[str, Any]] = []
    for figure_number, source_path in collect_local_figure_files(resources_dir):
        filename = f"figure_{figure_number}.jpeg"
        caption_row = captions_by_number.get(figure_number)
        caption = caption_row["full_caption"] if caption_row else ""
        row: Dict[str, Any] = {
            "figure_index": figure_number,
            "filename": filename,
            "source": "html_saved_resource",
            "original_file": source_path.name,
            "original_extension": source_path.suffix,
            "caption": caption,
            "caption_match_method": "figure_number" if caption else "",
            "status": "saved",
            "error": "",
        }
        try:
            row["width"], row["height"] = image_file_to_jpeg(
                source_path, images_dir / filename
            )
        except Exception as error:
            row.update({"width": "", "height": "", "status": "failed", "error": str(error)})
        manifest.append(row)
    write_image_manifest(images_dir, manifest)
    return manifest


# -----------------------------------------------------------------------------
# Writers
# -----------------------------------------------------------------------------

def write_sentences(output_path: Path, title: str, sentences: List[str]) -> None:
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(f"{title}\n")
        for sent in sentences:
            f.write(sent + "\n")


# -----------------------------------------------------------------------------
# Pipeline
# -----------------------------------------------------------------------------

def process_all_rsc_htmls(
    html_root: Path,
    pdf_root: Path,
    output_root: Optional[Path] = None,
) -> None:
    if output_root is None:
        output_root = html_root

    paper_count = 0
    fail_count = 0

    for paper_name, html_path in iter_paper_html_files(html_root):
        paper_count += 1
        try:
            title, base_full_text, base_sentences = parse_rsc_html_text(html_path)
            figure_captions = extract_figure_captions_from_html(html_path)
            tables = extract_tables_from_html(html_path)

            paper_dir = html_path.parent if output_root == html_root else output_root / paper_name
            paper_dir.mkdir(parents=True, exist_ok=True)

            cleanup_existing_html_outputs(paper_dir)

            full_text_with_tables = build_fulltext_with_tables(base_full_text, tables)
            sentences_with_tables = build_sentences_with_tables(base_sentences, tables)

            write_sentences(paper_dir / "sentences.txt", title, sentences_with_tables)
            with open(paper_dir / "fulltext.txt", "w", encoding="utf-8") as f:
                f.write(full_text_with_tables)

            write_figure_caption_files(paper_dir, figure_captions)
            write_table_caption_files(paper_dir, tables)
            write_table_files(paper_dir, tables)

            # Rebuild image output only after the HTML parsing/output succeeded.
            images_dir = reset_images_dir(pdf_root, paper_name)
            write_figure_caption_files(images_dir, figure_captions)
            resources_dir = find_saved_resource_dir(paper_name, html_path)
            if resources_dir is None:
                image_manifest: List[Dict[str, Any]] = []
                write_image_manifest(images_dir, image_manifest)
                image_note = "no Paper X_files folder found"
            else:
                image_manifest = extract_figures_from_saved_html_resources(
                    resources_dir, images_dir, figure_captions
                )
                saved_images = sum(row["status"] == "saved" for row in image_manifest)
                image_note = f"{saved_images} figure image(s) -> {images_dir}"

            print(
                f"[OK] {paper_name}: {len(sentences_with_tables)} line(s) -> {paper_dir / 'sentences.txt'} | "
                f"{len(figure_captions)} figure caption(s) -> {paper_dir / 'figure_captions.txt'} | "
                f"{len(tables)} table(s) -> {paper_dir / 'tables.json'} | {image_note}"
            )

        except Exception as e:
            fail_count += 1
            print(f"[FAIL] {paper_name} ({html_path}): {e}")

    print(f"\nDone. Processed {paper_count} paper(s), failed {fail_count}.")


# -----------------------------------------------------------------------------
# CLI
# -----------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Convert downloaded RSC HTML files into sentence/fulltext/caption files "
            "and extract main-article tables plus locally saved figure images."
        )
    )
    parser.add_argument(
        "--pdf-root",
        type=Path,
        default=Path("PDFs"),
        help="Root where PDFs/Paper X/images outputs are created or replaced.",
    )
    parser.add_argument(
        "--html-root",
        type=Path,
        default=Path("HTMLs"),
        help="Root folder containing Paper 1, Paper 2, ... HTML subfolders.",
    )
    parser.add_argument(
        "--output-root",
        type=Path,
        default=None,
        help=(
            "Optional separate output folder for sentences/fulltext/captions/tables. "
            "If omitted, files are written inside each HTML paper folder."
        ),
    )
    args = parser.parse_args()

    if not args.html_root.exists():
        raise FileNotFoundError(f"HTML root not found: {args.html_root}")

    process_all_rsc_htmls(
        html_root=args.html_root,
        pdf_root=args.pdf_root,
        output_root=args.output_root,
    )


if __name__ == "__main__":
    main()
