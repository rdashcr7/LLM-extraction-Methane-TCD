# -*- coding: utf-8 -*-
"""
Convert XML article files into text artifacts and extract figure/scheme images
from matching PDFs or from locally available XML-referenced graphic files.

This version is intended for Wiley XML and also keeps support for MDPI/JATS-like
XML where possible.

Expected folder layout:
    XMLs/Paper 1/Paper 1.xml
    XMLs/Paper 2/Paper 2.xml
    ...

    PDFs/Paper 1/Paper 1.pdf
    PDFs/Paper 2/Paper 2.pdf
    ...

Outputs:
    XMLs/Paper X/sentences.txt
    XMLs/Paper X/fulltext.txt
    XMLs/Paper X/figure_captions.txt
    XMLs/Paper X/figure_captions.csv

    PDFs/Paper X/images/figure_1.jpeg
    PDFs/Paper X/images/scheme_1.jpeg
    PDFs/Paper X/images/... 
    PDFs/Paper X/images/image_manifest.csv
    PDFs/Paper X/images/image_manifest.json
    PDFs/Paper X/images/figure_captions.txt
    PDFs/Paper X/images/figure_captions.csv

Behavior:
1. Parse XML directly with ElementTree.
2. Support Wiley XML (<figure>, <mediaResource>) and MDPI/JATS-like XML
   (<fig>, <graphic>).
3. Extract article title, abstract, section titles, and body paragraphs.
4. Split text into sentences using ChemDataExtractor if available, otherwise a
   regex fallback.
5. Extract figure/scheme captions from XML.
6. If local figure files referenced by the XML are present, convert them to JPEG
   and save them with XML-based filenames.
7. Otherwise, fall back to extracting images from the matching PDF and try to
   match them to XML figure/scheme numbers using nearby PDF text.
"""

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
    raise ImportError(
        "This script requires PyMuPDF. Install it with: pip install pymupdf"
    ) from e

try:
    from PIL import Image
except ImportError as e:
    raise ImportError(
        "This script requires Pillow. Install it with: pip install pillow"
    ) from e

try:
    from chemdataextractor import Document as CDEDocument

    CDE_AVAILABLE = True
except Exception:
    CDE_AVAILABLE = False
    CDEDocument = None  # type: ignore[assignment]


# PDF image filtering thresholds.
MIN_WIDTH = 120
MIN_HEIGHT = 120
MIN_AREA = 15_000
MAX_ASPECT_RATIO = 10.0

# PDF caption proximity thresholds in points.
MAX_BELOW_GAP = 180
MAX_ABOVE_GAP = 120
MAX_LOOSE_GAP = 320

XML_ID = "{http://www.w3.org/XML/1998/namespace}id"
XLINK_HREF = "{http://www.w3.org/1999/xlink}href"

OBJECT_LABEL_RE = re.compile(
    r"\b(?P<type>Fig(?:ure)?|Scheme)\.?\s*(?P<number>\d+)\b", re.IGNORECASE
)
STARTS_WITH_OBJECT_RE = re.compile(
    r"^\s*(?:Fig(?:ure)?|Scheme)\.?\s*\d+\b", re.IGNORECASE
)
TRAILING_NUMBER_RE = re.compile(r"(\d+)(?!.*\d)")
FIG_ID_NUMBER_RE = re.compile(r"(?:^|[-_])(?:fig|f)(?:[-_])?0*(\d+)\b", re.IGNORECASE)

IMAGE_EXTENSIONS = {".tif", ".tiff", ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"}
SECTION_TAGS = {"sec", "section"}
OBJECT_TAGS = {"fig", "figure"}
EXCLUDED_BODY_TAGS = {
    "fig",
    "figure",
    "caption",
    "graphic",
    "mediaresourcegroup",
    "mediaresource",
    "table-wrap",
    "table-wrap-group",
    "table",
    "tabular",
    "tgroup",
    "tbody",
    "thead",
    "ref-list",
    "bibliography",
    "bib",
    "disp-quote",
    "boxed-text",
    "app-group",
    "app",
    "supplementary-material",
    "supportinginformation",
    "supportinginfoitem",
    "supportinginfo",
    "notegroup",
}


def normalize_space(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def _local_name(tag: str) -> str:
    if "}" in tag:
        tag = tag.split("}", 1)[1]
    if ":" in tag:
        tag = tag.split(":", 1)[1]
    return tag.lower()


def _elem_text(elem: ET.Element) -> str:
    return normalize_space("".join(elem.itertext()))


def slugify(text: str) -> str:
    value = re.sub(r"[^a-z0-9]+", "_", normalize_space(text).lower())
    value = value.strip("_")
    return value or "figure"


def display_object_type(object_type: str) -> str:
    return normalize_space(object_type.replace("_", " ")).title() or "Figure"


# -----------------------------------------------------------------------------
# Folder helpers
# -----------------------------------------------------------------------------


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


def find_matching_pdf(paper_name: str, pdf_root: Path) -> Optional[Path]:
    paper_dir = pdf_root / paper_name
    preferred = paper_dir / f"{paper_name}.pdf"
    if preferred.exists():
        return preferred

    pdfs = sorted(paper_dir.glob("*.pdf")) if paper_dir.exists() else []
    return pdfs[0] if pdfs else None


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
        r"\b(Fig|Figure|Scheme|Eq|Ref|Dr|Mr|Ms|Prof|vs)\.",
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
# XML parsing helpers
# -----------------------------------------------------------------------------


def detect_xml_flavor(root: ET.Element) -> str:
    root_name = _local_name(root.tag)
    if root_name == "component":
        return "wiley"
    if root_name in {"article", "articles", "article-set"}:
        return "jats_like"
    return "generic"


def normalize_object_type(text: str) -> str:
    t = normalize_space(text).lower()
    if "scheme" in t:
        return "scheme"
    if "fig" in t:
        return "figure"
    return slugify(t or "figure")


def parse_object_number(label: str, fig_id: str = "", object_type: str = "figure") -> Optional[int]:
    label_norm = normalize_space(label)
    if label_norm:
        match = OBJECT_LABEL_RE.search(label_norm)
        if match:
            try:
                return int(match.group("number"))
            except ValueError:
                return None

        # Wiley labels are often just "1", "2", ...
        match = re.search(r"\b(\d+)\b", label_norm)
        if match:
            try:
                return int(match.group(1))
            except ValueError:
                return None

    fig_id_norm = normalize_space(fig_id)
    if fig_id_norm:
        match = FIG_ID_NUMBER_RE.search(fig_id_norm)
        if match:
            try:
                number = int(match.group(1))
            except ValueError:
                number = None
            if number is not None:
                if object_type != "figure" and number >= 5000:
                    reduced = number % 1000
                    if reduced:
                        return reduced
                return number

        match = TRAILING_NUMBER_RE.search(fig_id_norm)
        if match:
            try:
                number = int(match.group(1))
            except ValueError:
                return None
            if object_type != "figure" and number >= 5000:
                reduced = number % 1000
                if reduced:
                    return reduced
            return number

    return None


def build_display_label(object_type: str, raw_label: str, object_number: Optional[int]) -> str:
    raw = normalize_space(raw_label)
    if OBJECT_LABEL_RE.search(raw):
        return raw
    if raw.isdigit():
        return f"{display_object_type(object_type)} {raw}"
    if raw:
        if re.fullmatch(r"[0-9A-Za-z.\-]+", raw):
            return f"{display_object_type(object_type)} {raw}"
        return raw
    if object_number is not None:
        return f"{display_object_type(object_type)} {object_number}"
    return display_object_type(object_type)


def extract_article_title(root: ET.Element) -> str:
    # Wiley: <contentMeta><titleGroup><title type="main">...
    for elem in root.iter():
        if _local_name(elem.tag) == "contentmeta":
            for child in elem.iter():
                if _local_name(child.tag) == "title" and normalize_space(child.get("type", "")).lower() == "main":
                    text = _elem_text(child)
                    if text:
                        return text
            break

    # JATS/MDPI: <article-title>
    for elem in root.iter():
        if _local_name(elem.tag) == "article-title":
            text = _elem_text(elem)
            if text:
                return text

    return "NR"


def extract_abstract(root: ET.Element) -> Tuple[str, List[str]]:
    for abstract in root.iter():
        if _local_name(abstract.tag) != "abstract":
            continue

        heading = "Abstract"
        paragraphs: List[str] = []

        for child in list(abstract):
            cname = _local_name(child.tag)
            if cname == "title" and not paragraphs:
                title_text = _elem_text(child)
                if title_text:
                    heading = title_text
            elif cname == "p":
                text = _elem_text(child)
                if text:
                    paragraphs.append(text)

        if not paragraphs:
            fallback = _elem_text(abstract)
            if fallback:
                paragraphs = [fallback]

        return heading, paragraphs

    return "Abstract", []


def extract_body_blocks(root: ET.Element) -> List[str]:
    body = None
    for elem in root.iter():
        if _local_name(elem.tag) == "body":
            body = elem
            break
    if body is None:
        return []

    blocks: List[str] = []

    def walk(elem: ET.Element, parent_name: str = "") -> None:
        lname = _local_name(elem.tag)
        if lname in EXCLUDED_BODY_TAGS:
            return

        if lname == "title" and parent_name in SECTION_TAGS:
            text = _elem_text(elem)
            if text:
                blocks.append(text)
            return

        if lname == "p":
            text = _elem_text(elem)
            if text:
                blocks.append(text)
            return

        for child in list(elem):
            walk(child, lname)

    walk(body)
    return blocks


def _extract_caption_text(caption_elem: ET.Element) -> str:
    parts: List[str] = []
    for child in caption_elem.iter():
        if _local_name(child.tag) == "p":
            text = _elem_text(child)
            if text:
                parts.append(text)
    if parts:
        return normalize_space(" ".join(parts))

    fallback = _elem_text(caption_elem)
    return normalize_space(fallback)


def _prefer_graphic_href(hrefs: Sequence[Tuple[str, Dict[str, str]]]) -> str:
    def score(item: Tuple[str, Dict[str, str]]) -> Tuple[int, int, int, int]:
        href, attrs = item
        href_norm = normalize_space(href)
        suffix = Path(href_norm).suffix.lower()
        mime = normalize_space(attrs.get("mimeType", "") or attrs.get("mimetype", "")).lower()
        rendition = normalize_space(attrs.get("rendition", "")).lower()
        is_urn = href_norm.startswith("urn:")
        looks_like_image = int(bool(suffix in IMAGE_EXTENSIONS or mime.startswith("image/")))
        has_hires = int("hires" in rendition or "web" in rendition)
        has_path = int("/" in href_norm or "." in Path(href_norm).name)
        return (0 if is_urn else 1, looks_like_image, has_hires, has_path)

    if not hrefs:
        return ""
    href, _attrs = max(hrefs, key=score)
    return normalize_space(href)


def _extract_graphic_href(object_elem: ET.Element) -> str:
    candidates: List[Tuple[str, Dict[str, str]]] = []

    for child in list(object_elem):
        cname = _local_name(child.tag)
        if cname == "graphic":
            href = normalize_space(child.get(XLINK_HREF, "") or child.get("href", ""))
            if href:
                return href
        elif cname == "mediaresourcegroup":
            for media in list(child):
                if _local_name(media.tag) != "mediaresource":
                    continue
                href = normalize_space(media.get("href", ""))
                if href:
                    candidates.append((href, dict(media.attrib)))

    return _prefer_graphic_href(candidates)


def extract_xml_objects(root: ET.Element) -> List[Dict[str, Any]]:
    objects: List[Dict[str, Any]] = []
    seen_keys: Set[Tuple[str, int]] = set()
    next_by_type: Dict[str, int] = {}

    for elem in root.iter():
        if _local_name(elem.tag) not in OBJECT_TAGS:
            continue

        fig_id = normalize_space(elem.get(XML_ID, "") or elem.get("id", ""))
        raw_label = ""
        object_name = ""
        caption_text = ""

        for child in list(elem):
            cname = _local_name(child.tag)
            if cname == "label":
                raw_label = _elem_text(child)
            elif cname == "caption":
                caption_text = _extract_caption_text(child)
            elif cname == "title" and not object_name:
                object_name = _elem_text(child)
            elif cname == "titlegroup" and not object_name:
                for title in child.iter():
                    if _local_name(title.tag) == "title":
                        text = _elem_text(title)
                        if text:
                            object_name = text
                            break

        object_type = normalize_object_type(object_name or raw_label or "figure")
        object_number = parse_object_number(raw_label, fig_id, object_type)

        if object_number is None:
            object_number = next_by_type.get(object_type, 0) + 1
        next_by_type[object_type] = max(next_by_type.get(object_type, 0), object_number)

        key = (object_type, object_number)
        if key in seen_keys:
            continue
        seen_keys.add(key)

        label = build_display_label(object_type, raw_label, object_number)
        graphic_href = _extract_graphic_href(elem)
        full_caption = normalize_space(f"{label}. {caption_text}" if caption_text else label)
        filename = f"{slugify(object_type)}_{object_number}.jpeg"

        objects.append(
            {
                "sort_index": len(objects) + 1,
                "object_type": object_type,
                "object_number": object_number,
                "figure_number": object_number,
                "filename": filename,
                "raw_label": raw_label,
                "label": label,
                "caption": caption_text,
                "full_caption": full_caption,
                "graphic_href": graphic_href,
                "fig_id": fig_id,
            }
        )

    objects.sort(key=lambda row: (row["sort_index"], row["object_type"], row["object_number"]))
    return objects


def parse_article_xml(xml_path: Path) -> Tuple[str, str, List[str], List[Dict[str, Any]]]:
    tree = ET.parse(xml_path)
    root = tree.getroot()

    title = extract_article_title(root)
    abstract_heading, abstract_blocks = extract_abstract(root)
    body_blocks = extract_body_blocks(root)
    xml_objects = extract_xml_objects(root)

    text_blocks: List[str] = []
    if abstract_blocks:
        text_blocks.append(abstract_heading or "Abstract")
        text_blocks.extend(abstract_blocks)
    text_blocks.extend(body_blocks)

    body_text = "\n\n".join(block for block in text_blocks if normalize_space(block))
    full_text = f"{title}\n\n{body_text}" if title and title != "NR" and body_text else body_text or title
    sentences = split_sentences(body_text or full_text)
    return title, full_text, sentences, xml_objects


# -----------------------------------------------------------------------------
# Output writers
# -----------------------------------------------------------------------------


def write_sentences(output_path: Path, title: str, sentences: List[str]) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(f"{title}\n")
        for sent in sentences:
            f.write(sent + "\n")


def write_caption_files(output_dir: Path, captions: List[Dict[str, Any]]) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)

    txt_path = output_dir / "figure_captions.txt"
    with open(txt_path, "w", encoding="utf-8") as f:
        for row in captions:
            f.write(row.get("full_caption", "") + "\n")

    csv_path = output_dir / "figure_captions.csv"
    with open(csv_path, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=[
                "sort_index",
                "object_type",
                "object_number",
                "figure_number",
                "filename",
                "raw_label",
                "label",
                "caption",
                "full_caption",
                "graphic_href",
                "fig_id",
            ],
        )
        writer.writeheader()
        for row in captions:
            writer.writerow(
                {
                    "sort_index": row.get("sort_index", ""),
                    "object_type": row.get("object_type", ""),
                    "object_number": row.get("object_number", ""),
                    "figure_number": row.get("figure_number", ""),
                    "filename": row.get("filename", ""),
                    "raw_label": row.get("raw_label", ""),
                    "label": row.get("label", ""),
                    "caption": row.get("caption", ""),
                    "full_caption": row.get("full_caption", ""),
                    "graphic_href": row.get("graphic_href", ""),
                    "fig_id": row.get("fig_id", ""),
                }
            )


# -----------------------------------------------------------------------------
# Image helpers
# -----------------------------------------------------------------------------


def _recreate_images_dir(images_dir: Path) -> None:
    if images_dir.exists():
        shutil.rmtree(images_dir)
    images_dir.mkdir(parents=True, exist_ok=True)


def _bytes_to_jpeg(image_bytes: bytes, output_path: Path) -> Tuple[int, int]:
    with Image.open(io.BytesIO(image_bytes)) as img:
        img.load()
        if img.mode in ("RGBA", "LA"):
            background = Image.new("RGB", img.size, (255, 255, 255))
            alpha = img.getchannel("A") if "A" in img.getbands() else None
            background.paste(img.convert("RGBA"), mask=alpha)
            img = background
        else:
            img = img.convert("RGB")

        output_path.parent.mkdir(parents=True, exist_ok=True)
        img.save(output_path, format="JPEG", quality=95)
        return img.size


def _file_to_jpeg(input_path: Path, output_path: Path) -> Tuple[int, int]:
    with Image.open(input_path) as img:
        img.load()
        if img.mode in ("RGBA", "LA"):
            background = Image.new("RGB", img.size, (255, 255, 255))
            alpha = img.getchannel("A") if "A" in img.getbands() else None
            background.paste(img.convert("RGBA"), mask=alpha)
            img = background
        else:
            img = img.convert("RGB")

        output_path.parent.mkdir(parents=True, exist_ok=True)
        img.save(output_path, format="JPEG", quality=95)
        return img.size


def write_figure_manifest(images_dir: Path, manifest: List[Dict[str, Any]]) -> None:
    images_dir.mkdir(parents=True, exist_ok=True)

    json_path = images_dir / "image_manifest.json"
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)

    csv_path = images_dir / "image_manifest.csv"
    fieldnames = [
        "figure_index",
        "object_type",
        "object_number",
        "filename",
        "source",
        "page",
        "xref",
        "width",
        "height",
        "object_type_from_xml",
        "object_number_from_xml",
        "object_type_from_pdf_text",
        "object_number_from_pdf_text",
        "graphic_href",
        "graphic_file",
        "label",
        "caption",
        "caption_match_method",
        "kept_because",
        "filter_reason",
        "status",
        "error",
    ]
    with open(csv_path, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for row in manifest:
            writer.writerow({key: row.get(key, "") for key in fieldnames})


# -----------------------------------------------------------------------------
# Local XML graphic extraction
# -----------------------------------------------------------------------------


def resolve_local_graphic_path(xml_path: Path, graphic_href: str) -> Optional[Path]:
    graphic_href = normalize_space(graphic_href)
    if not graphic_href:
        return None

    graphic_href = re.sub(r"^file:", "", graphic_href, flags=re.IGNORECASE)
    if graphic_href.startswith("urn:"):
        return None

    xml_dir = xml_path.parent

    direct = xml_dir / graphic_href
    if direct.exists() and direct.is_file():
        return direct

    href_name = Path(graphic_href).name
    href_stem = Path(href_name).stem.lower()

    direct_name = xml_dir / href_name
    if direct_name.exists() and direct_name.is_file():
        return direct_name

    for match in xml_dir.rglob(href_name):
        if match.is_file():
            return match

    for ext in IMAGE_EXTENSIONS:
        candidate = xml_dir / f"{Path(href_name).stem}{ext}"
        if candidate.exists() and candidate.is_file():
            return candidate

    for match in xml_dir.rglob(f"{Path(href_name).stem}.*"):
        if match.is_file() and match.suffix.lower() in IMAGE_EXTENSIONS:
            return match

    for match in xml_dir.rglob("*"):
        if match.is_file() and match.suffix.lower() in IMAGE_EXTENSIONS and match.stem.lower() == href_stem:
            return match

    return None


def extract_local_graphics_from_xml(
    xml_path: Path,
    images_dir: Path,
    xml_objects: List[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    manifest: List[Dict[str, Any]] = []

    for row in xml_objects:
        object_type = row["object_type"]
        object_number = row["object_number"]
        filename = row["filename"]
        graphic_href = row.get("graphic_href", "")
        local_path = resolve_local_graphic_path(xml_path, graphic_href)

        if local_path is None:
            manifest.append(
                {
                    "figure_index": object_number,
                    "object_type": object_type,
                    "object_number": object_number,
                    "filename": filename,
                    "source": "xml_graphic",
                    "page": "",
                    "xref": "",
                    "width": "",
                    "height": "",
                    "object_type_from_xml": object_type,
                    "object_number_from_xml": object_number,
                    "object_type_from_pdf_text": "",
                    "object_number_from_pdf_text": "",
                    "graphic_href": graphic_href,
                    "graphic_file": "",
                    "label": row.get("label", ""),
                    "caption": row.get("full_caption", ""),
                    "caption_match_method": "xml_direct",
                    "kept_because": "",
                    "filter_reason": "local_graphic_missing",
                    "status": "missing",
                    "error": "",
                }
            )
            continue

        output_path = images_dir / filename
        try:
            width, height = _file_to_jpeg(local_path, output_path)
            manifest.append(
                {
                    "figure_index": object_number,
                    "object_type": object_type,
                    "object_number": object_number,
                    "filename": filename,
                    "source": "xml_graphic",
                    "page": "",
                    "xref": "",
                    "width": width,
                    "height": height,
                    "object_type_from_xml": object_type,
                    "object_number_from_xml": object_number,
                    "object_type_from_pdf_text": "",
                    "object_number_from_pdf_text": "",
                    "graphic_href": graphic_href,
                    "graphic_file": str(local_path),
                    "label": row.get("label", ""),
                    "caption": row.get("full_caption", ""),
                    "caption_match_method": "xml_direct",
                    "kept_because": "xml_graphic_file",
                    "filter_reason": "",
                    "status": "saved",
                    "error": "",
                }
            )
        except Exception as e:
            manifest.append(
                {
                    "figure_index": object_number,
                    "object_type": object_type,
                    "object_number": object_number,
                    "filename": filename,
                    "source": "xml_graphic",
                    "page": "",
                    "xref": "",
                    "width": "",
                    "height": "",
                    "object_type_from_xml": object_type,
                    "object_number_from_xml": object_number,
                    "object_type_from_pdf_text": "",
                    "object_number_from_pdf_text": "",
                    "graphic_href": graphic_href,
                    "graphic_file": str(local_path),
                    "label": row.get("label", ""),
                    "caption": row.get("full_caption", ""),
                    "caption_match_method": "xml_direct",
                    "kept_because": "xml_graphic_file",
                    "filter_reason": "",
                    "status": "failed",
                    "error": str(e),
                }
            )

    return manifest


# -----------------------------------------------------------------------------
# PDF image extraction fallback
# -----------------------------------------------------------------------------


def is_pdf_figure_by_size(base_image: Dict[str, Any]) -> Tuple[bool, str]:
    width = int(base_image.get("width", 0) or 0)
    height = int(base_image.get("height", 0) or 0)

    if width < MIN_WIDTH:
        return False, f"too narrow ({width} px wide, min {MIN_WIDTH})"
    if height < MIN_HEIGHT:
        return False, f"too short ({height} px tall, min {MIN_HEIGHT})"
    if width * height < MIN_AREA:
        return False, f"area too small ({width * height} px2, min {MIN_AREA})"

    aspect = max(width, height) / max(min(width, height), 1)
    if aspect > MAX_ASPECT_RATIO:
        return False, f"aspect ratio too extreme ({aspect:.1f}:1, max {MAX_ASPECT_RATIO})"

    return True, ""


def _extract_text_blocks(page: fitz.Page) -> List[Dict[str, Any]]:
    blocks: List[Dict[str, Any]] = []
    for raw in page.get_text("blocks"):
        x0, y0, x1, y1, text = raw[:5]
        clean = normalize_space(text)
        if not clean:
            continue
        blocks.append({"rect": fitz.Rect(x0, y0, x1, y1), "text": clean})
    return blocks


def _candidate_objects_from_text(text: str) -> List[Tuple[str, int]]:
    candidates: List[Tuple[str, int]] = []
    for match in OBJECT_LABEL_RE.finditer(text):
        try:
            number = int(match.group("number"))
        except ValueError:
            continue
        candidates.append((normalize_object_type(match.group("type")), number))
    return candidates


def _block_score_for_image(block_rect: fitz.Rect, image_rect: fitz.Rect, text: str) -> Optional[Tuple[int, int, float, float]]:
    objects = _candidate_objects_from_text(text)
    if not objects:
        return None

    x_overlap = max(0.0, min(block_rect.x1, image_rect.x1) - max(block_rect.x0, image_rect.x0))
    image_width = max(image_rect.width, 1.0)
    overlap_ratio = x_overlap / image_width

    below_gap = block_rect.y0 - image_rect.y1
    above_gap = image_rect.y0 - block_rect.y1

    if 0 <= below_gap <= MAX_BELOW_GAP and overlap_ratio > 0.10:
        region_rank = 0
        gap_rank = below_gap
    elif 0 <= above_gap <= MAX_ABOVE_GAP and overlap_ratio > 0.10:
        region_rank = 1
        gap_rank = above_gap
    elif 0 <= below_gap <= MAX_LOOSE_GAP:
        region_rank = 2
        gap_rank = below_gap
    elif 0 <= above_gap <= MAX_LOOSE_GAP:
        region_rank = 3
        gap_rank = above_gap
    else:
        return None

    starts_with_rank = 0 if STARTS_WITH_OBJECT_RE.match(text) else 1
    overlap_rank = -overlap_ratio
    return (region_rank, starts_with_rank, gap_rank, overlap_rank)


def guess_object_from_pdf_page(page: fitz.Page, image_rect: fitz.Rect) -> Optional[Tuple[str, int]]:
    candidates: List[Tuple[int, int, float, float, str, int, str]] = []
    for block in _extract_text_blocks(page):
        score = _block_score_for_image(block["rect"], image_rect, block["text"])
        if score is None:
            continue
        for object_type, number in _candidate_objects_from_text(block["text"]):
            candidates.append((*score, object_type, number, block["text"]))

    if not candidates:
        return None

    candidates.sort(key=lambda item: (item[0], item[1], item[2], item[3], item[4], item[5]))
    best = candidates[0]
    return best[4], best[5]


def _next_unused_number(used_keys: Set[Tuple[str, int]], object_type: str = "figure", start: int = 1) -> int:
    n = start
    while (object_type, n) in used_keys:
        n += 1
    return n


def attach_captions_to_manifest(
    manifest: List[Dict[str, Any]],
    xml_objects: List[Dict[str, Any]],
) -> None:
    xml_lookup = {(row["object_type"], row["object_number"]): row for row in xml_objects}
    used_xml_keys: Set[Tuple[str, int]] = set()

    for row in manifest:
        object_type = normalize_space(str(row.get("object_type_from_pdf_text", ""))).lower()
        object_number_raw = row.get("object_number_from_pdf_text")
        if object_type and isinstance(object_number_raw, int):
            key = (object_type, object_number_raw)
            if key in xml_lookup:
                xml_row = xml_lookup[key]
                row["caption"] = xml_row["full_caption"]
                row["caption_match_method"] = "object_from_pdf_text"
                row["object_type_from_xml"] = xml_row["object_type"]
                row["object_number_from_xml"] = xml_row["object_number"]
                row["label"] = xml_row["label"]
                used_xml_keys.add(key)

    remaining_xml = [row for row in xml_objects if (row["object_type"], row["object_number"]) not in used_xml_keys]
    remaining_images = [row for row in manifest if not row.get("caption")]

    if remaining_images and remaining_xml and len(remaining_images) == len(remaining_xml):
        remaining_images.sort(key=lambda row: (int(row.get("page", 0) or 0), int(row.get("figure_index", 0) or 0)))
        remaining_xml.sort(key=lambda row: row.get("sort_index", 0))
        for image_row, xml_row in zip(remaining_images, remaining_xml):
            image_row["caption"] = xml_row["full_caption"]
            image_row["caption_match_method"] = "sequential_fallback"
            image_row["object_type_from_xml"] = xml_row["object_type"]
            image_row["object_number_from_xml"] = xml_row["object_number"]
            image_row["label"] = xml_row["label"]

    for row in manifest:
        row.setdefault("caption", "")
        row.setdefault("caption_match_method", "")
        row.setdefault("object_type_from_xml", "")
        row.setdefault("object_number_from_xml", "")
        row.setdefault("label", "")


def rename_manifest_files_to_xml_labels(
    images_dir: Path,
    manifest: List[Dict[str, Any]],
    xml_objects: List[Dict[str, Any]],
) -> None:
    xml_lookup = {(row["object_type"], row["object_number"]): row for row in xml_objects}

    for row in manifest:
        if row.get("status") != "saved":
            continue

        key: Optional[Tuple[str, int]] = None
        xml_type = normalize_space(str(row.get("object_type_from_xml", ""))).lower()
        xml_num = row.get("object_number_from_xml")
        if xml_type and isinstance(xml_num, int):
            key = (xml_type, xml_num)
        else:
            own_type = normalize_space(str(row.get("object_type", ""))).lower()
            own_num = row.get("object_number")
            if own_type and isinstance(own_num, int):
                key = (own_type, own_num)

        if key is None or key not in xml_lookup:
            continue

        desired_filename = xml_lookup[key]["filename"]
        current_filename = str(row.get("filename", ""))
        if not desired_filename or desired_filename == current_filename:
            continue

        src = images_dir / current_filename
        dst = images_dir / desired_filename
        if not src.exists() or dst.exists():
            continue

        try:
            src.rename(dst)
            row["filename"] = desired_filename
            row["object_type"] = xml_lookup[key]["object_type"]
            row["object_number"] = xml_lookup[key]["object_number"]
            row["figure_index"] = xml_lookup[key]["object_number"]
        except Exception as e:
            existing_error = normalize_space(str(row.get("error", "")))
            rename_error = f"rename_failed: {e}"
            row["error"] = f"{existing_error} | {rename_error}".strip(" |")


def extract_figures_from_pdf_as_jpeg(
    pdf_path: Path,
    images_dir: Path,
    xml_objects: Optional[List[Dict[str, Any]]] = None,
) -> List[Dict[str, Any]]:
    doc = fitz.open(str(pdf_path))
    seen_xrefs: Set[int] = set()
    used_keys: Set[Tuple[str, int]] = set()
    manifest: List[Dict[str, Any]] = []
    xml_lookup = {(row["object_type"], row["object_number"]): row for row in (xml_objects or [])}

    try:
        for page_num, page in enumerate(doc, start=1):
            for img in page.get_images(full=True):
                xref = img[0]
                if xref in seen_xrefs:
                    continue
                seen_xrefs.add(xref)

                try:
                    base_image = doc.extract_image(xref)
                except Exception:
                    continue

                image_rects: List[fitz.Rect] = []
                try:
                    image_rects = page.get_image_rects(xref)
                except Exception:
                    image_rects = []

                guessed_object: Optional[Tuple[str, int]] = None
                if image_rects:
                    guessed_object = guess_object_from_pdf_page(page, image_rects[0])

                keep_by_size, filter_reason = is_pdf_figure_by_size(base_image)
                if not keep_by_size and guessed_object is None:
                    continue

                if guessed_object is not None and guessed_object not in used_keys:
                    object_type, object_number = guessed_object
                    if guessed_object in xml_lookup:
                        filename = xml_lookup[guessed_object]["filename"]
                        label = xml_lookup[guessed_object]["label"]
                    else:
                        filename = f"{slugify(object_type)}_{object_number}.jpeg"
                        label = build_display_label(object_type, str(object_number), object_number)
                    kept_because = "pdf_text_label"
                else:
                    object_type = "figure"
                    object_number = _next_unused_number(used_keys, object_type="figure")
                    filename = f"figure_{object_number}.jpeg"
                    label = build_display_label(object_type, str(object_number), object_number)
                    kept_because = "size_filter" if keep_by_size else "pdf_text_label_duplicate_number"

                used_keys.add((object_type, object_number))
                output_path = images_dir / filename

                try:
                    width, height = _bytes_to_jpeg(base_image["image"], output_path)
                except Exception as e:
                    manifest.append(
                        {
                            "figure_index": object_number,
                            "object_type": object_type,
                            "object_number": object_number,
                            "filename": filename,
                            "source": "pdf",
                            "page": page_num,
                            "xref": xref,
                            "width": "",
                            "height": "",
                            "object_type_from_xml": "",
                            "object_number_from_xml": "",
                            "object_type_from_pdf_text": guessed_object[0] if guessed_object else "",
                            "object_number_from_pdf_text": guessed_object[1] if guessed_object else "",
                            "graphic_href": "",
                            "graphic_file": "",
                            "label": label,
                            "caption": "",
                            "caption_match_method": "",
                            "kept_because": kept_because,
                            "filter_reason": filter_reason,
                            "status": "failed",
                            "error": str(e),
                        }
                    )
                    continue

                manifest.append(
                    {
                        "figure_index": object_number,
                        "object_type": object_type,
                        "object_number": object_number,
                        "filename": filename,
                        "source": "pdf",
                        "page": page_num,
                        "xref": xref,
                        "width": width,
                        "height": height,
                        "object_type_from_xml": "",
                        "object_number_from_xml": "",
                        "object_type_from_pdf_text": guessed_object[0] if guessed_object else "",
                        "object_number_from_pdf_text": guessed_object[1] if guessed_object else "",
                        "graphic_href": "",
                        "graphic_file": "",
                        "label": label,
                        "caption": "",
                        "caption_match_method": "",
                        "kept_because": kept_because,
                        "filter_reason": filter_reason,
                        "status": "saved",
                        "error": "",
                    }
                )
    finally:
        doc.close()

    saved_rows = [row for row in manifest if row.get("status") == "saved"]
    if xml_objects and saved_rows:
        attach_captions_to_manifest(saved_rows, xml_objects)
        rename_manifest_files_to_xml_labels(images_dir, saved_rows, xml_objects)

    if manifest:
        write_figure_manifest(images_dir, manifest)
    return manifest


# -----------------------------------------------------------------------------
# Pipeline
# -----------------------------------------------------------------------------


def process_all_xmls(
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
            title, full_text, sentences, xml_objects = parse_article_xml(xml_path)

            paper_dir = xml_path.parent if output_root == xml_root else output_root / paper_name
            paper_dir.mkdir(parents=True, exist_ok=True)

            write_sentences(paper_dir / "sentences.txt", title, sentences)
            with open(paper_dir / "fulltext.txt", "w", encoding="utf-8") as f:
                f.write(full_text)
            write_caption_files(paper_dir, xml_objects)

            pdf_path = find_matching_pdf(paper_name, pdf_root)
            images_dir = pdf_root / paper_name / "images"
            if clear_images or not images_dir.exists():
                _recreate_images_dir(images_dir)
            else:
                images_dir.mkdir(parents=True, exist_ok=True)

            write_caption_files(images_dir, xml_objects)

            local_manifest = extract_local_graphics_from_xml(
                xml_path=xml_path,
                images_dir=images_dir,
                xml_objects=xml_objects,
            )
            local_saved_rows = [row for row in local_manifest if row.get("status") == "saved"]

            if xml_objects and local_saved_rows and len(local_saved_rows) == len(xml_objects):
                write_figure_manifest(images_dir, local_manifest)
                print(
                    f"[OK] {paper_name}: {len(sentences)} sentences -> {paper_dir / 'sentences.txt'} | "
                    f"{len(xml_objects)} XML object caption(s) extracted -> {paper_dir / 'figure_captions.txt'} | "
                    f"{len(local_saved_rows)} object image(s) saved from local XML graphics -> {images_dir}"
                )
                continue

            _recreate_images_dir(images_dir)
            write_caption_files(images_dir, xml_objects)

            if pdf_path is not None:
                pdf_manifest = extract_figures_from_pdf_as_jpeg(
                    pdf_path=pdf_path,
                    images_dir=images_dir,
                    xml_objects=xml_objects,
                )
                saved_pdf_rows = [row for row in pdf_manifest if row.get("status") == "saved"]
                matched_caption_rows = [row for row in saved_pdf_rows if row.get("caption")]
                print(
                    f"[OK] {paper_name}: {len(sentences)} sentences -> {paper_dir / 'sentences.txt'} | "
                    f"{len(xml_objects)} XML object caption(s) extracted -> {paper_dir / 'figure_captions.txt'} | "
                    f"{len(saved_pdf_rows)} figure-like image(s) saved from PDF -> {images_dir} | "
                    f"{len(matched_caption_rows)} image(s) matched to XML captions"
                )
            else:
                if local_manifest:
                    write_figure_manifest(images_dir, local_manifest)
                print(
                    f"[OK] {paper_name}: {len(sentences)} sentences -> {paper_dir / 'sentences.txt'} | "
                    f"{len(xml_objects)} XML object caption(s) extracted -> {paper_dir / 'figure_captions.txt'} | "
                    f"no matching PDF found, so no figures were saved"
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
            "Convert XML files into sentence/fulltext/caption files and extract "
            "figure/scheme images from local XML graphics or matching PDFs."
        )
    )
    parser.add_argument(
        "--xml-root",
        type=Path,
        default=Path("XMLs"),
        help="Root folder containing Paper 1, Paper 2, ... XML subfolders.",
    )
    parser.add_argument(
        "--pdf-root",
        type=Path,
        default=Path("PDFs"),
        help="Root folder containing Paper 1, Paper 2, ... PDF subfolders.",
    )
    parser.add_argument(
        "--output-root",
        type=Path,
        default=None,
        help=(
            "Optional separate output folder for sentences/fulltext/captions. "
            "If omitted, files are written inside each XML paper folder."
        ),
    )
    parser.add_argument(
        "--clear-images",
        action="store_true",
        help="Recreate each PDFs/Paper X/images folder before saving fresh figure files.",
    )
    args = parser.parse_args()

    if not args.xml_root.exists():
        raise FileNotFoundError(f"XML root not found: {args.xml_root}")
    if not args.pdf_root.exists():
        raise FileNotFoundError(f"PDF root not found: {args.pdf_root}")

    process_all_xmls(
        xml_root=args.xml_root,
        pdf_root=args.pdf_root,
        output_root=args.output_root,
        clear_images=args.clear_images,
    )


if __name__ == "__main__":
    main()
