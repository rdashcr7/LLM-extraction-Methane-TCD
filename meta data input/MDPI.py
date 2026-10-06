# -*- coding: utf-8 -*-

from __future__ import annotations

import argparse
import csv
import io
import json
import os
import re
import shutil
import tempfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Set, Tuple

# Optional Windows workaround for duplicate OpenMP runtimes.
os.environ.setdefault("KMP_DUPLICATE_LIB_OK", "TRUE")

try:
    import fitz  # PyMuPDF
except ImportError as e:
    raise ImportError("This script requires PyMuPDF. Install it with: pip install pymupdf") from e

try:
    from PIL import Image
except ImportError as e:
    raise ImportError("This script requires Pillow. Install it with: pip install pillow") from e

try:
    from chemdataextractor import Document as CDEDocument

    CDE_AVAILABLE = True
except Exception:
    CDE_AVAILABLE = False
    CDEDocument = None  # type: ignore[assignment]


XLINK_HREF = "{http://www.w3.org/1999/xlink}href"
FIGURE_NUMBER_RE = re.compile(r"\b(?:Fig(?:ure)?\.?)\s*(\d+)\b", re.IGNORECASE)
IMAGE_EXTENSIONS = {".tif", ".tiff", ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"}
FIG_ID_RE = re.compile(r"-f0*(\d+)\b", re.IGNORECASE)
TABLE_LABEL_RE = re.compile(r"\bTable\s*(\d+)\b", re.IGNORECASE)
TABLE_ID_RE = re.compile(r"-t0*(\d+)\b", re.IGNORECASE)
TRAILING_NUMBER_RE = re.compile(r"(\d+)(?!.*\d)")
EXCLUDED_BODY_TAGS = {
    "fig",
    "graphic",
    "caption",
    "table-wrap",
    "table-wrap-group",
    "table",
    "thead",
    "tbody",
    "tfoot",
    "tr",
    "th",
    "td",
    "ref-list",
    "disp-quote",
    "boxed-text",
    "app-group",
    "app",
    "supplementary-material",
}
SKIP_TEXT_TAGS = {"annotation", "annotation-xml"}
BREAK_LIKE_TAGS = {"break", "br"}


# -----------------------------------------------------------------------------
# General helpers
# -----------------------------------------------------------------------------


def normalize_space(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()



def _local_name(tag: str) -> str:
    if "}" in tag:
        tag = tag.split("}", 1)[1]
    if ":" in tag:
        tag = tag.split(":", 1)[1]
    return tag.lower()



def _elem_text(elem: ET.Element) -> str:
    parts: List[str] = []

    def walk(node: ET.Element) -> None:
        lname = _local_name(node.tag)
        if lname in SKIP_TEXT_TAGS:
            return
        if node.text:
            parts.append(node.text)
        for child in list(node):
            clname = _local_name(child.tag)
            if clname in BREAK_LIKE_TAGS:
                parts.append(" ")
            walk(child)
            if child.tail:
                parts.append(child.tail)

    walk(elem)
    return normalize_space("".join(parts))



def iter_paper_xml_files(xml_root: Path) -> Iterable[Tuple[str, Path]]:
    """
    Yield (paper_name, xml_path) for folders like:
      XMLs/Paper 1/Paper 1.xml
    """
    for paper_dir in sorted((p for p in xml_root.iterdir() if p.is_dir()), key=lambda p: p.name):
        paper_name = paper_dir.name
        xml_path = paper_dir / f"{paper_name}.xml"
        if xml_path.exists():
            yield paper_name, xml_path
        else:
            xmls = sorted(paper_dir.glob("*.xml"))
            if xmls:
                yield paper_name, xmls[0]


# -----------------------------------------------------------------------------
# Sentence splitting
# -----------------------------------------------------------------------------


def sentence_split_with_cde(text: str) -> List[str]:
    if not text.strip() or not CDE_AVAILABLE:
        return []

    with tempfile.NamedTemporaryFile("w", suffix=".txt", encoding="utf-8", delete=False) as tmp:
        tmp.write(text)
        temp_path = Path(tmp.name)

    try:
        doc = CDEDocument.from_file(str(temp_path))  # type: ignore[union-attr]
        if not getattr(doc, "elements", None):
            return []

        sentences: List[str] = []
        for element in doc.elements:
            if hasattr(element, "sentences"):
                for sent in element.sentences:
                    s = normalize_space(str(sent))
                    if s:
                        sentences.append(s)

        return sentences or []
    finally:
        try:
            temp_path.unlink(missing_ok=True)
        except Exception:
            pass



def sentence_split_fallback(text: str) -> List[str]:
    if not text.strip():
        return []

    protected = text
    protected = re.sub(
        r"\b(Fig|Figure|Table|Eq|Ref|Dr|Mr|Ms|Prof|vs|No)\.",
        lambda m: m.group(1) + "<DOT>",
        protected,
    )
    protected = re.sub(r"\b([A-Z])\.", lambda m: m.group(1) + "<DOT>", protected)
    chunks = re.split(r"(?<=[.!?])\s+(?=[A-Z0-9(\[])", protected)
    sentences = [normalize_space(chunk.replace("<DOT>", ".")) for chunk in chunks if normalize_space(chunk)]
    return sentences



def split_sentences(text: str) -> List[str]:
    sentences = sentence_split_with_cde(text) if CDE_AVAILABLE else []
    if not sentences:
        sentences = sentence_split_fallback(text)
    return merge_very_short_sentences(sentences)



def merge_very_short_sentences(sentences: List[str], min_len: int = 10) -> List[str]:
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


# -----------------------------------------------------------------------------
# MDPI XML body / figure parsing
# -----------------------------------------------------------------------------


def parse_figure_number(label: str, fig_id: str = "") -> Optional[int]:
    text = f"{label} {fig_id}".strip()
    match = FIGURE_NUMBER_RE.search(text)
    if match:
        try:
            return int(match.group(1))
        except ValueError:
            return None
    match = FIG_ID_RE.search(fig_id)
    if match:
        try:
            return int(match.group(1))
        except ValueError:
            return None
    return None



def extract_mdpi_title(root: ET.Element) -> str:
    for elem in root.iter():
        if _local_name(elem.tag) == "article-title":
            text = _elem_text(elem)
            if text:
                return text
    return "NR"



def extract_mdpi_abstract(root: ET.Element) -> List[str]:
    paragraphs: List[str] = []
    for abstract in root.iter():
        if _local_name(abstract.tag) != "abstract":
            continue
        for child in abstract.iter():
            if _local_name(child.tag) == "p":
                text = _elem_text(child)
                if text:
                    paragraphs.append(text)
        if paragraphs:
            break
    return paragraphs



def _collect_body_blocks_safe(body: ET.Element) -> List[str]:
    """
    Walk the <body> tree in document order, collecting section titles and paragraphs,
    while skipping display objects and back-matter-like structures.
    """
    blocks: List[str] = []

    def walk(elem: ET.Element, parent_name: str = "") -> None:
        lname = _local_name(elem.tag)
        if lname in EXCLUDED_BODY_TAGS:
            return

        for child in list(elem):
            cname = _local_name(child.tag)
            if cname in EXCLUDED_BODY_TAGS:
                continue
            if cname == "title" and lname == "sec":
                text = _elem_text(child)
                if text:
                    blocks.append(text)
                continue
            if cname == "p":
                text = _elem_text(child)
                if text:
                    blocks.append(text)
                continue
            walk(child, lname)

    walk(body)
    return blocks



def extract_mdpi_body_blocks(root: ET.Element) -> List[str]:
    for elem in root.iter():
        if _local_name(elem.tag) == "body":
            return _collect_body_blocks_safe(elem)
    return []



def extract_mdpi_figure_captions(root: ET.Element) -> List[Dict[str, Any]]:
    figures: List[Dict[str, Any]] = []
    seen: Set[int] = set()

    for fig in root.iter():
        if _local_name(fig.tag) != "fig":
            continue

        fig_id = normalize_space(fig.get("id", ""))
        label_text = ""
        caption_text = ""
        graphic_href = ""

        for child in list(fig):
            cname = _local_name(child.tag)
            if cname == "label":
                label_text = _elem_text(child)
            elif cname == "caption":
                caption_parts: List[str] = []
                for cap_child in child.iter():
                    if _local_name(cap_child.tag) == "p":
                        text = _elem_text(cap_child)
                        if text:
                            caption_parts.append(text)
                if not caption_parts:
                    fallback = _elem_text(child)
                    if fallback:
                        caption_parts.append(fallback)
                caption_text = normalize_space(" ".join(caption_parts))
            elif cname == "graphic":
                graphic_href = normalize_space(child.get(XLINK_HREF, "") or child.get("href", ""))

        figure_number = parse_figure_number(label_text, fig_id)
        if figure_number is None or figure_number in seen:
            continue
        seen.add(figure_number)

        label = label_text or f"Figure {figure_number}"
        full_caption = normalize_space(f"{label}. {caption_text}" if caption_text else label)
        figures.append(
            {
                "figure_number": figure_number,
                "label": label,
                "caption": caption_text,
                "full_caption": full_caption,
                "graphic_href": graphic_href,
                "fig_id": fig_id,
            }
        )

    figures.sort(key=lambda row: row["figure_number"])
    return figures


# -----------------------------------------------------------------------------
# Table extraction
# -----------------------------------------------------------------------------


def parse_table_number(raw_label: str, table_id: str = "") -> Optional[int]:
    label_norm = normalize_space(raw_label)
    if label_norm:
        match = TABLE_LABEL_RE.search(label_norm)
        if match:
            try:
                return int(match.group(1))
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



def _extract_mdpi_table_caption(container: ET.Element) -> str:
    for child in list(container):
        if _local_name(child.tag) != "caption":
            continue

        parts: List[str] = []
        for sub in child.iter():
            if _local_name(sub.tag) == "p":
                text = _elem_text(sub)
                if text:
                    parts.append(text)

        if parts:
            return normalize_space(" ".join(parts))

        fallback = _elem_text(child)
        if fallback:
            return fallback

    return ""



def _extract_html_table_rows(section_elem: Optional[ET.Element]) -> List[List[str]]:
    rows: List[List[str]] = []
    if section_elem is None:
        return rows

    for row in section_elem.iter():
        if _local_name(row.tag) != "tr":
            continue
        cells: List[str] = []
        for cell in list(row):
            if _local_name(cell.tag) not in {"th", "td"}:
                continue
            cells.append(_elem_text(cell))
        if cells:
            rows.append(cells)
    return rows



def _extract_mdpi_table_fallback_rows(table_elem: ET.Element) -> Tuple[List[List[str]], List[List[str]]]:
    header_rows: List[List[str]] = []
    body_rows: List[List[str]] = []

    for row in table_elem.iter():
        if _local_name(row.tag) != "tr":
            continue

        cell_texts: List[str] = []
        cell_tags: List[str] = []
        for cell in list(row):
            lname = _local_name(cell.tag)
            if lname not in {"th", "td"}:
                continue
            cell_tags.append(lname)
            cell_texts.append(_elem_text(cell))

        if not cell_texts:
            continue

        if cell_tags and all(tag == "th" for tag in cell_tags) and not body_rows:
            header_rows.append(cell_texts)
        else:
            body_rows.append(cell_texts)

    return header_rows, body_rows



def _extract_mdpi_table_notes(container: ET.Element) -> List[str]:
    notes: List[str] = []

    for foot in container.iter():
        if _local_name(foot.tag) != "table-wrap-foot":
            continue
        for fn in foot.iter():
            if _local_name(fn.tag) != "fn":
                continue

            label = ""
            body_parts: List[str] = []
            for child in list(fn):
                cname = _local_name(child.tag)
                if cname == "label":
                    label = _elem_text(child)
                else:
                    text = _elem_text(child)
                    if text:
                        body_parts.append(text)

            note_text = normalize_space(" ".join(body_parts)) if body_parts else _elem_text(fn)
            if label and note_text.startswith(label):
                note_text = normalize_space(note_text[len(label):])

            if label and note_text:
                notes.append(f"{label}. {note_text}")
            elif note_text:
                notes.append(note_text)

    return notes



def _format_table_text(table: Dict[str, Any]) -> str:
    lines: List[str] = []
    if table.get("full_caption"):
        lines.append(str(table["full_caption"]))

    header_rows = table.get("header_rows", []) or []
    body_rows = table.get("body_rows", []) or []
    notes = table.get("notes", []) or []

    for idx, row in enumerate(header_rows, start=1):
        row_text = " | ".join(cell for cell in row if normalize_space(cell))
        if row_text:
            prefix = "Header" if len(header_rows) == 1 else f"Header {idx}"
            lines.append(f"{prefix}: {row_text}")

    for idx, row in enumerate(body_rows, start=1):
        row_text = " | ".join(cell for cell in row if normalize_space(cell))
        if row_text:
            lines.append(f"Row {idx}: {row_text}")

    for note in notes:
        if note:
            lines.append(f"Note: {note}")

    return "\n".join(lines).strip()



def extract_mdpi_tables(root: ET.Element) -> List[Dict[str, Any]]:
    tables: List[Dict[str, Any]] = []
    seen_numbers: Set[int] = set()
    next_number = 1

    for elem in root.iter():
        if _local_name(elem.tag) != "table-wrap":
            continue

        table_id = normalize_space(elem.get("id", ""))
        raw_label = ""
        for child in list(elem):
            if _local_name(child.tag) == "label":
                raw_label = _elem_text(child)
                break

        table_number = parse_table_number(raw_label, table_id)
        if table_number is None:
            while next_number in seen_numbers:
                next_number += 1
            table_number = next_number
        seen_numbers.add(table_number)
        next_number = max(next_number, table_number + 1)

        label = f"Table {table_number}"
        title = _extract_mdpi_table_caption(elem)
        full_caption = normalize_space(f"{label}. {title}" if title else label)

        table_elem = None
        for child in list(elem):
            if _local_name(child.tag) == "table":
                table_elem = child
                break

        thead = None
        tbody = None
        if table_elem is not None:
            for sub in table_elem.iter():
                lname = _local_name(sub.tag)
                if lname == "thead" and thead is None:
                    thead = sub
                elif lname == "tbody" and tbody is None:
                    tbody = sub

        header_rows = _extract_html_table_rows(thead)
        body_rows = _extract_html_table_rows(tbody)

        if not header_rows and not body_rows and table_elem is not None:
            header_rows, body_rows = _extract_mdpi_table_fallback_rows(table_elem)

        notes = _extract_mdpi_table_notes(elem)

        row_records: List[Dict[str, Any]] = []
        row_index = 1
        for row in header_rows:
            row_records.append({"row_type": "header", "row_index": row_index, "cells": row})
            row_index += 1
        for row in body_rows:
            row_records.append({"row_type": "body", "row_index": row_index, "cells": row})
            row_index += 1

        table = {
            "sort_index": len(tables) + 1,
            "table_number": table_number,
            "label": label,
            "title": title,
            "full_caption": full_caption,
            "table_id": table_id,
            "raw_label": raw_label,
            "header_rows": header_rows,
            "body_rows": body_rows,
            "notes": notes,
            "rows": row_records,
        }
        table["table_text"] = _format_table_text(table)
        tables.append(table)

    tables.sort(key=lambda row: (row["table_number"], row["sort_index"]))
    return tables



def build_table_sentence_lines(tables: List[Dict[str, Any]]) -> List[str]:
    lines: List[str] = []
    for table in tables:
        label = table.get("label", "Table")
        if table.get("full_caption"):
            lines.append(str(table["full_caption"]))

        header_rows = table.get("header_rows", []) or []
        for idx, row in enumerate(header_rows, start=1):
            row_text = " | ".join(cell for cell in row if normalize_space(cell))
            if row_text:
                prefix = f"{label} header" if len(header_rows) == 1 else f"{label} header {idx}"
                lines.append(f"{prefix}: {row_text}")

        for idx, row in enumerate(table.get("body_rows", []) or [], start=1):
            row_text = " | ".join(cell for cell in row if normalize_space(cell))
            if row_text:
                lines.append(f"{label} row {idx}: {row_text}")

        for note in table.get("notes", []) or []:
            if note:
                lines.append(f"{label} note: {note}")

    return [normalize_space(line) for line in lines if normalize_space(line)]


# -----------------------------------------------------------------------------
# Main XML parser
# -----------------------------------------------------------------------------


def parse_mdpi_xml(xml_path: Path) -> Tuple[str, str, List[str], List[Dict[str, Any]], List[Dict[str, Any]]]:
    tree = ET.parse(xml_path)
    root = tree.getroot()

    title = extract_mdpi_title(root)
    abstract_blocks = extract_mdpi_abstract(root)
    body_blocks = extract_mdpi_body_blocks(root)
    figure_captions = extract_mdpi_figure_captions(root)
    xml_tables = extract_mdpi_tables(root)

    text_blocks: List[str] = []
    if abstract_blocks:
        text_blocks.append("Abstract")
        text_blocks.extend(abstract_blocks)
    text_blocks.extend(body_blocks)

    body_text = "\n\n".join(block for block in text_blocks if normalize_space(block))
    table_text_blocks = [table["table_text"] for table in xml_tables if normalize_space(table.get("table_text", ""))]

    full_parts: List[str] = []
    if title and title != "NR":
        full_parts.append(title)
    if body_text:
        full_parts.append(body_text)
    if table_text_blocks:
        full_parts.append("Tables\n\n" + "\n\n".join(table_text_blocks))

    full_text = "\n\n".join(part for part in full_parts if normalize_space(part))

    sentence_source = body_text if body_text else full_text
    sentences = split_sentences(sentence_source)
    sentences.extend(build_table_sentence_lines(xml_tables))

    return title, full_text, sentences, figure_captions, xml_tables


# -----------------------------------------------------------------------------
# Output writers
# -----------------------------------------------------------------------------


def write_sentences(output_path: Path, title: str, sentences: List[str]) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(f"{title}\n")
        for sent in sentences:
            f.write(sent + "\n")



def write_figure_caption_files(output_dir: Path, captions: List[Dict[str, Any]]) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)

    txt_path = output_dir / "figure_captions.txt"
    with open(txt_path, "w", encoding="utf-8") as f:
        for row in captions:
            f.write(row.get("full_caption", "") + "\n")

    csv_path = output_dir / "figure_captions.csv"
    with open(csv_path, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=["figure_number", "label", "caption", "full_caption", "graphic_href", "fig_id"],
        )
        writer.writeheader()
        for row in captions:
            writer.writerow(
                {
                    "figure_number": row.get("figure_number", ""),
                    "label": row.get("label", ""),
                    "caption": row.get("caption", ""),
                    "full_caption": row.get("full_caption", ""),
                    "graphic_href": row.get("graphic_href", ""),
                    "fig_id": row.get("fig_id", ""),
                }
            )



def write_table_files(output_dir: Path, tables: List[Dict[str, Any]]) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)

    captions_txt = output_dir / "table_captions.txt"
    with open(captions_txt, "w", encoding="utf-8") as f:
        for row in tables:
            f.write(row.get("full_caption", "") + "\n")

    captions_csv = output_dir / "table_captions.csv"
    with open(captions_csv, "w", encoding="utf-8", newline="") as f:
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

    tables_json = output_dir / "tables.json"
    with open(tables_json, "w", encoding="utf-8") as f:
        json.dump(tables, f, ensure_ascii=False, indent=2)

    tables_text = output_dir / "tables_text.txt"
    with open(tables_text, "w", encoding="utf-8") as f:
        for idx, row in enumerate(tables, start=1):
            if idx > 1:
                f.write("\n\n")
            f.write(row.get("table_text", ""))

    max_cols = 0
    for table in tables:
        for row in table.get("rows", []) or []:
            max_cols = max(max_cols, len(row.get("cells", []) or []))

    tables_rows_csv = output_dir / "tables_rows.csv"
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

    with open(tables_rows_csv, "w", encoding="utf-8", newline="") as f:
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
                    "cells_joined": " | ".join(cells),
                    "cells_json": json.dumps(cells, ensure_ascii=False),
                }
                for idx, cell in enumerate(cells, start=1):
                    record[f"cell_{idx}"] = cell
                writer.writerow(record)


# -----------------------------------------------------------------------------
# Figure-image extraction
# -----------------------------------------------------------------------------

def find_matching_pdf(paper_name: str, pdf_root: Path) -> Optional[Path]:
    paper_dir = pdf_root / paper_name
    preferred = paper_dir / f"{paper_name}.pdf"
    if preferred.is_file():
        return preferred
    pdfs = sorted(paper_dir.glob("*.pdf")) if paper_dir.is_dir() else []
    return pdfs[0] if pdfs else None


def recreate_images_dir(images_dir: Path) -> None:
    if images_dir.exists():
        shutil.rmtree(images_dir, ignore_errors=True)
    images_dir.mkdir(parents=True, exist_ok=True)


def convert_to_jpeg(source: Path | bytes, destination: Path) -> Tuple[int, int]:
    """Convert image bytes or a local image to a white-background JPEG."""
    source_data: Any = io.BytesIO(source) if isinstance(source, bytes) else source
    with Image.open(source_data) as image:
        image.load()
        transparent = image.mode in ("RGBA", "LA") or (image.mode == "P" and "transparency" in image.info)
        if transparent:
            rgba = image.convert("RGBA")
            result = Image.new("RGB", rgba.size, "white")
            result.paste(rgba, mask=rgba.getchannel("A"))
        else:
            result = image.convert("RGB")
        result.save(destination, "JPEG", quality=95)
        return result.size


def resolve_xml_graphic(xml_path: Path, href: str) -> Optional[Path]:
    """Find a graphic referenced by xlink:href, including extension-less links."""
    if not href:
        return None
    href_path = Path(href.replace("\\", "/"))
    candidates = [xml_path.parent / href_path, xml_path.parent / href_path.name]
    if not href_path.suffix:
        candidates.extend((xml_path.parent / href_path.name).with_suffix(ext) for ext in IMAGE_EXTENSIONS)
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    matches = sorted(p for p in xml_path.parent.rglob(f"{href_path.stem}.*") if p.suffix.lower() in IMAGE_EXTENSIONS)
    return matches[0] if matches else None


def write_image_manifest(images_dir: Path, rows: List[Dict[str, Any]]) -> None:
    with open(images_dir / "image_manifest.json", "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=2)
    fields = ["figure_index", "filename", "source", "original_file", "page", "xref", "width", "height", "caption", "status", "error"]
    with open(images_dir / "image_manifest.csv", "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        writer.writerows({key: row.get(key, "") for key in fields} for row in rows)


def extract_local_xml_graphics(xml_path: Path, images_dir: Path, captions: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    for caption in captions:
        number = caption["figure_number"]
        source = resolve_xml_graphic(xml_path, caption.get("graphic_href", ""))
        if source is None:
            continue
        row: Dict[str, Any] = {"figure_index": number, "filename": f"figure_{number}.jpeg", "source": "xml_graphic", "original_file": source.name, "page": "", "xref": "", "caption": caption["full_caption"], "status": "saved", "error": ""}
        try:
            row["width"], row["height"] = convert_to_jpeg(source, images_dir / row["filename"])
        except Exception as error:
            row.update({"width": "", "height": "", "status": "failed", "error": str(error)})
        rows.append(row)
    return rows


def extract_pdf_images(pdf_path: Path, images_dir: Path, captions: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Extract substantial embedded PDF images and pair them with XML captions in order."""
    rows: List[Dict[str, Any]] = []
    seen_xrefs: Set[int] = set()
    document = fitz.open(pdf_path)
    try:
        for page_number, page in enumerate(document, start=1):
            for image_info in page.get_images(full=True):
                xref = image_info[0]
                if xref in seen_xrefs:
                    continue
                seen_xrefs.add(xref)
                base = document.extract_image(xref)
                width, height = int(base.get("width", 0)), int(base.get("height", 0))
                if width < 120 or height < 120 or width * height < 15000:
                    continue
                number = len(rows) + 1
                row: Dict[str, Any] = {"figure_index": number, "filename": f"figure_{number}.jpeg", "source": "pdf", "original_file": "", "page": page_number, "xref": xref, "caption": "", "status": "saved", "error": ""}
                try:
                    row["width"], row["height"] = convert_to_jpeg(base["image"], images_dir / row["filename"])
                except Exception as error:
                    row.update({"width": "", "height": "", "status": "failed", "error": str(error)})
                rows.append(row)
    finally:
        document.close()

    saved = [row for row in rows if row["status"] == "saved"]
    if len(saved) == len(captions):
        for row, caption in zip(saved, captions):
            old_path = images_dir / row["filename"]
            row["figure_index"] = caption["figure_number"]
            row["filename"] = f"figure_{caption['figure_number']}.jpeg"
            new_path = images_dir / row["filename"]
            if old_path != new_path and old_path.exists():
                old_path.rename(new_path)
            row["caption"] = caption["full_caption"]
    return rows


# -----------------------------------------------------------------------------
# Pipeline
# -----------------------------------------------------------------------------


def process_all_mdpi_xmls(
    xml_root: Path,
    pdf_root: Path,
    output_root: Optional[Path] = None,
    clear_images: bool = False,
) -> None:
    if output_root is None:
        output_root = xml_root

    paper_count = 0
    fail_count = 0

    for paper_name, xml_path in iter_paper_xml_files(xml_root):
        paper_count += 1
        try:
            title, full_text, sentences, xml_captions, xml_tables = parse_mdpi_xml(xml_path)

            paper_dir = xml_path.parent if output_root == xml_root else output_root / paper_name
            paper_dir.mkdir(parents=True, exist_ok=True)

            write_sentences(paper_dir / "sentences.txt", title, sentences)
            with open(paper_dir / "fulltext.txt", "w", encoding="utf-8") as f:
                f.write(full_text)
            write_figure_caption_files(paper_dir, xml_captions)
            write_table_files(paper_dir, xml_tables)

            images_dir = pdf_root / paper_name / "images"
            if clear_images or not images_dir.exists():
                recreate_images_dir(images_dir)
            else:
                images_dir.mkdir(parents=True, exist_ok=True)
            write_figure_caption_files(images_dir, xml_captions)

            image_rows = extract_local_xml_graphics(xml_path, images_dir, xml_captions)
            saved_local = [row for row in image_rows if row["status"] == "saved"]
            if len(saved_local) < len(xml_captions):
                pdf_path = find_matching_pdf(paper_name, pdf_root)
                if pdf_path is not None:
                    recreate_images_dir(images_dir)
                    write_figure_caption_files(images_dir, xml_captions)
                    image_rows = extract_pdf_images(pdf_path, images_dir, xml_captions)
            write_image_manifest(images_dir, image_rows)
            saved_images = sum(row["status"] == "saved" for row in image_rows)

            print(
                f"[OK] {paper_name}: "
                f"{len(sentences)} sentence line(s) -> {paper_dir / 'sentences.txt'} | "
                f"{len(xml_captions)} figure caption(s) -> {paper_dir / 'figure_captions.txt'} | "
                f"{len(xml_tables)} table(s) -> {paper_dir / 'table_captions.txt'} | "
                f"{saved_images} image(s) -> {images_dir}"
            )

        except Exception as e:
            fail_count += 1
            print(f"[FAIL] {paper_name} ({xml_path}): {e}")

    print(f"\nDone. Processed {paper_count} paper(s), failed {fail_count}.")


# -----------------------------------------------------------------------------
# CLI
# -----------------------------------------------------------------------------


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Convert MDPI XML files into text, table, caption, and figure-image outputs."
        )
    )
    parser.add_argument(
        "--pdf-root",
        type=Path,
        default=Path("PDFs"),
        help="Root containing Paper X PDF folders and image outputs.",
    )
    parser.add_argument(
        "--xml-root",
        type=Path,
        default=Path("XMLs"),
        help="Root folder containing Paper 1, Paper 2, ... XML subfolders.",
    )
    parser.add_argument(
        "--clear-images",
        action="store_true",
        help="Delete each existing PDFs/Paper X/images folder before processing.",
    )
    parser.add_argument(
        "--output-root",
        type=Path,
        default=None,
        help=(
            "Optional separate output folder for sentences/fulltext/captions/tables. "
            "If omitted, files are written inside each XML paper folder."
        ),
    )
    args = parser.parse_args()

    if not args.xml_root.exists():
        raise FileNotFoundError(f"XML root not found: {args.xml_root}")

    process_all_mdpi_xmls(
        xml_root=args.xml_root,
        pdf_root=args.pdf_root,
        output_root=args.output_root,
        clear_images=args.clear_images,
    )


if __name__ == "__main__":
    main()
