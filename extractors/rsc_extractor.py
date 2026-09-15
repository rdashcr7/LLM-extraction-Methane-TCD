# -*- coding: utf-8 -*-
"""
RSC structured methane-decomposition data extractor.

This script is adapted from the attached MDPI data_extractor8.py pipeline. It
uses the same two-pass strategy:

1. Read RSC HTML-derived article text, figure captions, and table text from
   HTMLs/<Paper>/.
2. Read curated/extracted figure images from PDFs/<Paper>/images/.
3. Use the OpenAI Responses API to extract paper-level constants from text/tables
   and datapoints from figure images plus XML caption context.
4. Write a consolidated CSV/XLSX and an LLM usage summary.

Expected default layout:

    RSC_DOIs.xlsx
    HTMLs/
      Paper 1/
        Paper 1.xml
        fulltext.txt
        sentences.txt
        figure_captions.csv
        table_captions.csv
        tables.json
        tables_text.txt
        tables_rows.csv
      Paper 2/
        ...

    PDFs/
      Paper 1/
        images/
          figure_1.png|jpg|jpeg|webp
          figure_2.png|jpg|jpeg|webp
      Paper 2/
        images/
          ...

The loader is intentionally tolerant of earlier RSC artifact names too,
including xml_article_text.txt, xml_figure_reference.csv, and xml_tables_text.txt.

"""

import argparse
import base64
import json
import os
import random
import re
import sys
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from openai import OpenAI
import pandas as pd
from dotenv import load_dotenv
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

REPOSITORY_DIR = Path(__file__).resolve().parents[1]
if str(REPOSITORY_DIR) not in sys.path:
    sys.path.insert(0, str(REPOSITORY_DIR))
from bedrock_runtime import converse as bedrock_converse, selected_provider



# OpenAI-specific structured extractor. Journal-specific loaders and the current
# panel-first extraction logic are preserved; only the LLM transport layer changed.
load_dotenv()

OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "")
OPENAI_BASE_URL = os.environ.get("OPENAI_BASE_URL", "https://api.openai.com/v1")
MODEL_ID = os.environ.get("OPENAI_MODEL_ID", os.environ.get("MODEL_ID", "gpt-5.5"))
OPENAI_REASONING_EFFORT = os.environ.get("OPENAI_REASONING_EFFORT", "low")
OPENAI_TEXT_VERBOSITY = os.environ.get("OPENAI_TEXT_VERBOSITY", "low")
OPENAI_STORE_RESPONSES = os.environ.get("OPENAI_STORE_RESPONSES", "false").strip().lower() in {"1", "true", "yes"}


def _build_openai_client():
    if not OPENAI_API_KEY:
        return None
    kwargs = {"api_key": OPENAI_API_KEY}
    if OPENAI_BASE_URL:
        kwargs["base_url"] = OPENAI_BASE_URL.rstrip("/")
    return OpenAI(**kwargs)


openai_client = _build_openai_client()

# Image limits are conservative local preprocessing limits before sending images
# to the OpenAI Responses API. GPT-5.5 preserves more visual detail by default;
# these limits prevent very large PDFs/figures from producing unnecessary tokens.
BEDROCK_IMAGE_MAX_BYTES = 3_500_000
BEDROCK_IMAGE_MAX_SIDE = 8000
SOFT_IMAGE_MAX_BYTES = int(os.environ.get("OPENAI_IMAGE_SOFT_MAX_BYTES", os.environ.get("BEDROCK_IMAGE_SOFT_MAX_BYTES", "2500000")))
SOFT_IMAGE_MAX_SIDE = int(os.environ.get("OPENAI_IMAGE_SOFT_MAX_SIDE", os.environ.get("BEDROCK_IMAGE_SOFT_MAX_SIDE", "1600")))
SOFT_IMAGE_JPEG_QUALITY = int(os.environ.get("OPENAI_IMAGE_JPEG_QUALITY", os.environ.get("BEDROCK_IMAGE_JPEG_QUALITY", "90")))
FIGURE_MAX_OUTPUT_TOKENS = int(os.environ.get("FIGURE_MAX_OUTPUT_TOKENS", os.environ.get("OPENAI_FIGURE_MAX_OUTPUT_TOKENS", "12000")))

# Application-level request pacing and throttling retry settings.
LLM_CALL_DELAY_SECONDS = float(os.environ.get("LLM_CALL_DELAY_SECONDS", "1.0"))
LLM_THROTTLE_MAX_RETRIES = int(os.environ.get("LLM_THROTTLE_MAX_RETRIES", "6"))
LLM_THROTTLE_BASE_DELAY_SECONDS = float(os.environ.get("LLM_THROTTLE_BASE_DELAY_SECONDS", "15"))
LLM_THROTTLE_MAX_DELAY_SECONDS = float(os.environ.get("LLM_THROTTLE_MAX_DELAY_SECONDS", "180"))
LLM_THROTTLE_JITTER_SECONDS = float(os.environ.get("LLM_THROTTLE_JITTER_SECONDS", "3"))

CURRENT_PAPER_FOR_USAGE: Optional[str] = None
LLM_USAGE_BY_PAPER: Dict[str, Dict[str, int]] = {}
CATALYST_SYNTHESIS_PROFILE_CACHE: Dict[Tuple[str, str, str, str, str], Dict[str, Any]] = {}
# True pass-1 constants captured before figure extraction. Used for post-run audit reports.
PASS1_CONSTANTS_BY_PAPER: Dict[str, Dict[str, Any]] = {}
# Raw panel/series discovery decisions captured for every figure that reaches the LLM.
# This lets audit report 07 show how many relevant panels the LLM identified,
# even when later extraction/post-processing returns zero rows.
PANEL_DISCOVERY_AUDIT_BY_PAPER: Dict[str, List[Dict[str, Any]]] = {}
# Image paths actually considered for each paper; used by hard-skip and LLM-rejection audits.
AUDIT_IMAGE_PATHS_BY_PAPER: Dict[str, List[str]] = {}
# Bump when output columns/profile schema changes so stale progress files are ignored.
PROGRESS_VERSION = 130

USAGE_COUNTER_KEYS = [
    "llm_calls",
    "input_tokens",
    "output_tokens",
]

CATALYST_SYNTHESIS_PROFILE_FIELDS = [
    "Precursor_Salts",
    "Drying_Temperature_C_or_K",
    "Drying_Time_hr_min_s",
    "Calcination_Temperature_C_or_K",
    "Calcination_Temperature_Ramp_Rate_C_or_K_per_min",
    "Calcination_Time_hr_min_s",
    "Calcination_Atmosphere",
    "Reduction_Temperature_C_or_K",
    "Reduction_Temperature_Ramp_Rate_C_or_K_per_min",
    "Reduction_Time_hr_min_s",
    "Reduction_Atmosphere",
    "Cooling_Rate_C_or_K_per_min",
]

CATALYST_SUPPORT_PROFILE_FIELDS = [
    "Support_Type",
    "Support_Surface_Area",
]

CATALYST_PROFILE_FIELDS = CATALYST_SUPPORT_PROFILE_FIELDS + CATALYST_SYNTHESIS_PROFILE_FIELDS


def _empty_usage_bucket() -> Dict[str, int]:
    return {key: 0 for key in USAGE_COUNTER_KEYS}


def _usage_bucket(paper_name: Optional[str]) -> Optional[Dict[str, int]]:
    if not paper_name:
        return None
    bucket = LLM_USAGE_BY_PAPER.setdefault(paper_name, _empty_usage_bucket())
    for key in USAGE_COUNTER_KEYS:
        bucket.setdefault(key, 0)
    return bucket


def _progress_path(output_prefix: str) -> str:
    return f"{output_prefix}_progress.json"


def _save_progress(progress_path: str, completed_papers: List[str], all_rows: List[Dict[str, Any]]) -> None:
    payload = {
        "version": PROGRESS_VERSION,
        "completed_papers": completed_papers,
        "all_rows": all_rows,
        "llm_usage_by_paper": LLM_USAGE_BY_PAPER,
        "pass1_constants_by_paper": PASS1_CONSTANTS_BY_PAPER,
        "panel_discovery_audit_by_paper": PANEL_DISCOVERY_AUDIT_BY_PAPER,
        "audit_image_paths_by_paper": AUDIT_IMAGE_PATHS_BY_PAPER,
    }
    with open(progress_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False)


def _load_progress(progress_path: str) -> Dict[str, Any]:
    with open(progress_path, "r", encoding="utf-8") as f:
        payload = json.load(f)
    if payload.get("version") != PROGRESS_VERSION:
        return {"completed_papers": [], "all_rows": [], "llm_usage_by_paper": {}}
    return payload


def _set_current_paper_for_usage(paper_name: Optional[str]) -> None:
    global CURRENT_PAPER_FOR_USAGE
    CURRENT_PAPER_FOR_USAGE = paper_name
    _usage_bucket(paper_name)


def _extract_usage_from_converse_response(response: Dict[str, Any]) -> Dict[str, int]:
    usage = response.get("usage") or response.get("tokenUsage") or response.get("metrics", {}).get("usage") or response.get("metrics", {}).get("tokenUsage") or {}

    def get_int(*names: str) -> int:
        for name in names:
            if name in usage:
                try:
                    return int(usage.get(name, 0) or 0)
                except Exception:
                    return 0
        return 0

    return {
        "input_tokens": get_int("inputTokens", "input_tokens"),
        "output_tokens": get_int("outputTokens", "output_tokens"),
    }


def _count_input_tokens(messages: list, system_prompt: str) -> int:
    # The OpenAI Responses API returns usage after each call; no separate count-token call is used here.
    return 0


def _clone_messages_without_images(messages: list) -> list:
    stripped = []
    for msg in messages:
        content = []
        for item in msg.get("content", []):
            if isinstance(item, dict) and "text" in item:
                content.append({"text": item["text"]})
        stripped.append({"role": msg.get("role", "user"), "content": content})
    return stripped


def _classify_input_type(messages: list) -> str:
    has_text = False
    has_image = False
    for msg in messages:
        for item in msg.get("content", []):
            if isinstance(item, dict) and "text" in item and str(item.get("text", "")).strip():
                has_text = True
            if isinstance(item, dict) and "image" in item:
                has_image = True
    if has_image and has_text:
        return "Figure + Text"
    if has_image:
        return "Figure only"
    return "Text only"


def _record_llm_usage(response: Dict[str, Any], messages: list, system_prompt: str) -> None:
    bucket = _usage_bucket(CURRENT_PAPER_FOR_USAGE)
    if bucket is None:
        return
    usage = _extract_usage_from_converse_response(response)
    bucket["llm_calls"] += 1
    for key in USAGE_COUNTER_KEYS:
        if key == "llm_calls":
            continue
        bucket[key] = int(bucket.get(key, 0) or 0) + int(usage.get(key, 0) or 0)


COLUMNS = [
    "Paper",
    "Catalyst_Label",
    "Series_Label",
    "Marker_Color",
    "Marker_Shape",
    "DataPoint_Index",
    "Heating_Mode",
    "CH4_Conversion_%",
    "Temperature_C_or_K",
    "Temperature_source",
    "Temperature_snippet",
    "Pressure_bar_or_atm_or_Pa",
    "X_Axis_Label",
    "Reaction_Time_hr_min_s",
    "Feed_Flow_Rate_mL_or_L_per_min",
    "CH4_Feed_Flow_Rate_mL_or_L_per_min",
    "Feed_Gas_Mixture",
    "Space_Velocity",
    "Catalyst_Mass_g_or_mg",
    "Active_Metal",
    "Metal_Weight_%",
    "Catalyst_Class",
    "Promoter_Type",
    "Support_Type",
    "Support_Surface_Area",
    "Support_Description",
    "Precursor_Salts",
    "Drying_Temperature_C_or_K",
    "Drying_Time_hr_min_s",
    "Calcination_Temperature_C_or_K",
    "Calcination_Temperature_Ramp_Rate_C_or_K_per_min",
    "Calcination_Time_hr_min_s",
    "Calcination_Atmosphere",
    "Reduction_Temperature_C_or_K",
    "Reduction_Temperature_Ramp_Rate_C_or_K_per_min",
    "Reduction_Time_hr_min_s",
    "Reduction_Atmosphere",
    "Cooling_Rate_C_or_K_per_min",
    "Y_Axis_Label",
    "Figure_Sources",
    "Matched_XML_Figure_Label",
    "Matched_XML_Figure_Order",
    "Caption_Match_Method",
]


TEXT_CONSTANT_FIELDS = [
    "Title",
    "DOI",
    "Catalyst_Label",
    "Heating_Mode",
    "Pressure_bar_or_atm_or_Pa",
    "Reaction_Time_hr_min_s",
    "Feed_Flow_Rate_mL_or_L_per_min",
    "CH4_Feed_Flow_Rate_mL_or_L_per_min",
    "Feed_Gas_Mixture",
    "Space_Velocity",
    "Catalyst_Mass_g_or_mg",
    "Temperature_C_or_K",
    "Precursor_Salts",
    "Drying_Temperature_C_or_K",
    "Drying_Time_hr_min_s",
    "Calcination_Temperature_C_or_K",
    "Calcination_Temperature_Ramp_Rate_C_or_K_per_min",
    "Calcination_Time_hr_min_s",
    "Calcination_Atmosphere",
    "Reduction_Temperature_C_or_K",
    "Reduction_Temperature_Ramp_Rate_C_or_K_per_min",
    "Reduction_Time_hr_min_s",
    "Reduction_Atmosphere",
    "Cooling_Rate_C_or_K_per_min",
]


PROMPT_TEXT_CONSTANT_FIELDS = [
    "Catalyst_Label",
    "Heating_Mode",
    "Pressure_bar_or_atm_or_Pa",
    "Reaction_Time_hr_min_s",
    "Feed_Flow_Rate_mL_or_L_per_min",
    "CH4_Feed_Flow_Rate_mL_or_L_per_min",
    "Feed_Gas_Mixture",
    "Space_Velocity",
    "Catalyst_Mass_g_or_mg",
    "Temperature_C_or_K",
    "Precursor_Salts",
    "Drying_Temperature_C_or_K",
    "Drying_Time_hr_min_s",
    "Calcination_Temperature_C_or_K",
    "Calcination_Temperature_Ramp_Rate_C_or_K_per_min",
    "Calcination_Time_hr_min_s",
    "Calcination_Atmosphere",
    "Reduction_Temperature_C_or_K",
    "Reduction_Temperature_Ramp_Rate_C_or_K_per_min",
    "Reduction_Time_hr_min_s",
    "Reduction_Atmosphere",
    "Cooling_Rate_C_or_K_per_min",
]


FIGURE_BASE_KEYS = [
    "Catalyst_Label",
    "Series_Label",
    "Marker_Color",
    "Marker_Shape",
    "X_Axis_Label",
    "CH4_Conversion_%",
    "Reaction_Time_hr_min_s",
    "Temperature_C_or_K",
    "Pressure_bar_or_atm_or_Pa",
    "Feed_Flow_Rate_mL_or_L_per_min",
    "CH4_Feed_Flow_Rate_mL_or_L_per_min",
    "Feed_Gas_Mixture",
    "Space_Velocity",
    "Catalyst_Mass_g_or_mg",
    "Y_Axis_Label",
]

FIGURE_EXTRACTION_KEYS = FIGURE_BASE_KEYS + [
    "Heating_Mode",
    "Precursor_Salts",
    "Drying_Temperature_C_or_K",
    "Drying_Time_hr_min_s",
    "Calcination_Temperature_C_or_K",
    "Calcination_Temperature_Ramp_Rate_C_or_K_per_min",
    "Calcination_Time_hr_min_s",
    "Calcination_Atmosphere",
    "Reduction_Temperature_C_or_K",
    "Reduction_Temperature_Ramp_Rate_C_or_K_per_min",
    "Reduction_Time_hr_min_s",
    "Reduction_Atmosphere",
    "Cooling_Rate_C_or_K_per_min",
]


# Compact schema used only for LLM figure-output rows. Catalyst/support/synthesis
# profile fields are filled later from paper constants and catalyst-profile
# enrichment, so repeating them in every datapoint row wastes output tokens and
# caused dense curves to truncate into one parsed row.
FIGURE_LLM_POINT_FIELDS = [
    "Catalyst_Label",
    "Series_Label",
    "Marker_Color",
    "Marker_Shape",
    "X_Axis_Label",
    "Y_Axis_Label",
    "CH4_Conversion_%",
    "Reaction_Time_hr_min_s",
    "Temperature_C_or_K",
    "Pressure_bar_or_atm_or_Pa",
    "Feed_Flow_Rate_mL_or_L_per_min",
    "CH4_Feed_Flow_Rate_mL_or_L_per_min",
    "Feed_Gas_Mixture",
    "Space_Velocity",
    "Catalyst_Mass_g_or_mg",
    "Metal_Weight_%",
    "Support_Type",
    "Heating_Mode",
]
FIGURE_SCHEMA_MODE = os.environ.get("FIGURE_SCHEMA_MODE", "compact").strip().lower()


INTERNAL_FIELD_SOURCE_PREFIX = "_Field_Source_"


TEMP_PATTERN = re.compile(
    r"(?P<val>[-+]?\d{1,4}(?:\.\d+)?)\s*"
    r"(?:°|º|◦)?\s*"
    r"(?P<unit>"
    r"C|K|F|R|"
    r"°C|°F|°K|"
    r"degC|degF|degK|"          
    r"celsius|kelvin|fahrenheit"
    r")\b",
    flags=re.IGNORECASE,
)

FLOW_PATTERN = re.compile(
    r"(?P<val>[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?)\s*"
    r"(?P<unit>"    
    r"mL/min|ml/min|L/min|l/min|mL/h|ml/h|L/h|l/h|"
    r"mL\s*min-1|ml\s*min-1|L\s*min-1|l\s*min-1|"
    r"mL\s*h-1|ml\s*h-1|L\s*h-1|l\s*h-1|"
    r"mL\s*hr-1|ml\s*hr-1|L\s*hr-1|l\s*hr-1|"
    r"cc/min|cm3/min|cc\s*min-1|cm3\s*min-1|"
    r"cc/h|cm3/h|cc\s*h-1|cm3\s*h-1|cc\s*hr-1|cm3\s*hr-1|"    
    r"sccm|sccs|slm|SLM|slpm|SLPM|lpm|LPM|SCCM|"
    r"m3/s|m3/min|cm3/s|"
    r"g/s|g/min|kg/s|kg/h|mg/s|"
    r"mol/s|mmol/s|mol/min|kmol/h"
    r")\b",
    flags=re.IGNORECASE,
)


GAS_COMPONENT_PERCENT_PATTERN = re.compile(
    r"(?P<pct>[-+]?\d*\.?\d+)\s*(?:vol\s*%|%)\s*(?P<gas>CH4|methane|N2|nitrogen|Ar|argon|He|helium|H2|hydrogen|CO2|carbon\s+dioxide|CO|O2|oxygen|steam|water\s+vapor|H2O|air)\b",
    flags=re.IGNORECASE,
)

GAS_PERCENT_COMPONENT_PATTERN = re.compile(
    r"(?P<gas>CH4|methane|N2|nitrogen|Ar|argon|He|helium|H2|hydrogen|CO2|carbon\s+dioxide|CO|O2|oxygen|steam|water\s+vapor|H2O|air)\s*(?:=|:)?\s*(?P<pct>[-+]?\d*\.?\d+)\s*(?:vol\s*%|%)\b",
    flags=re.IGNORECASE,
)

MARKER_COLOR_CANONICAL = {
    "red": "red",
    "blue": "blue",
    "green": "green",
    "orange": "orange",
    "yellow": "yellow",
    "purple": "purple",
    "violet": "purple",
    "magenta": "magenta",
    "pink": "pink",
    "gray": "gray",
    "grey": "gray",
    "black": "black",
    "brown": "brown",
    "cyan": "cyan",
    "teal": "cyan",
    "white": "white",
    "gold": "yellow",
    "golden": "yellow",
}


MARKER_SHAPE_CANONICAL = {
    "circle": "circle",
    "round": "circle",
    "square": "square",
    "triangle": "triangle",
    "triangular": "triangle",
    "diamond": "diamond",
    "star": "star",
    "asterisk": "star",
    "plus": "plus",
    "cross": "cross",
    "x": "x",
    "hexagon": "hexagon",
    "pentagon": "pentagon",
    "dot": "dot",
}

MARKER_FILL_CANONICAL = {
    "filled": "filled",
    "solid": "filled",
    "closed": "filled",
    "open": "open",
    "hollow": "open",
    "unfilled": "open",
    "empty": "open",
}


OXIDE_LEFT_PATTERN = re.compile(
    r"\b(?:"
    r"NiO|FeO|Fe2O3|Fe3O4|CoO|Co3O4|CuO|PdO|PtO|Rh2O3|RuO2|IrO2|"
    r"MnO2|Cr2O3|ZnO|MoO3|WO3|CeO2|La2O3|ZrO2|MgO|CaO|TiO2|V2O5"
    r")\b",
    flags=re.IGNORECASE,
)

PRECURSOR_NAME_PATTERN = re.compile(
    r"\b(?:nickel|iron(?:\s*\([^\)]*\))?|cobalt|copper|molybdenum|tungsten|chromium|"
    r"manganese|zinc|palladium|platinum|rhodium|ruthenium|iridium|silver|lanthanum|"
    r"cerium|yttrium|zirconium|magnesium|calcium|potassium|sodium|cesium|lithium)"
    r"\s+(?:nitrate|chloride|acetate|sulfate|carbonate|bicarbonate|citrate|oxalate|"
    r"hydroxide|formate|propionate)"
    r"(?:\s+(?:mono|di|tri|tetra|penta|hexa|hepta|octa|nona)hydrate)?\b",
    flags=re.IGNORECASE,
)

PRECURSOR_FORMULA_PATTERN = re.compile(
    r"\b(?:"
    r"[A-Z][a-z]?(?:\([A-Za-z0-9]+\)\d*)+(?:[·⋅\.][xX]?\d*H2O)?|"
    r"[A-Z][a-z]?(?:Cl|NO3|CH3COO|OAc|SO4|CO3|C2O4|OH)[A-Za-z0-9()·⋅\.xX-]*"
    r")\b"
)


PROMOTER_ELEMENTS = {
    "K", "Na", "Cs", "Li",
    "Mg", "Ca", "Sr", "Ba",
    "Ce", "La", "Y", "Zr",
    "Ti", "V", "P", "B", "Pd",
}

CATALYST_TOKEN_PATTERN = re.compile(
    r"\b(?:"
    r"(?:\d+\s*[A-Z][a-z]?)(?:-\d+\s*[A-Z][a-z]?)*[@/][A-Za-z0-9][A-Za-z0-9\-\./()]+"
    r"|"
    r"[A-Z][a-z]?(?:[-+][A-Z][a-z]?)*[/@][A-Za-z0-9][A-Za-z0-9\-\./()]+"
    r")\b"
)

CONDITION_ONLY_PATTERNS = [
    r"^\~?\d+(?:\.\d+)?\s*(?:°\s*C|°C|C|K|°F|F|R|degC|degF|degK)$",
    r"^\~?\d+(?:\.\d+)?\s*(?:s|sec|secs|second|seconds|"
    r"m|min|mins|minute|minutes|"
    r"h|hr|hrs|hour|hours|"
    r"d|day|days)$",
    r"^\~?\d+(?:\.\d+)?\s*(?:"
    r"bar|atm|pa|kpa|mpa|"
    r"torr|mmhg|psi|"
    r"mbar|millibar"
    r")$",
    r"^\~?\d+(?:\.\d+)?\s*(?:"
    r"ml/min|l/min|cc/min|cm3/min|"
    r"sccm|slm|slpm|lpm|"
    r"g|mg|kg|"
    r"mol|mmol|kmol|"
    r"mol/s|mol/min|"
    r"g/s|kg/h"
    r")$",
    r"^\~?\d+(?:\.\d+)?\s*(?:"
    r"wt%|vol%|mol%|at%|"
    r"M|mM|uM|nM"
    r")$",
]


# Prompt text is externalized under ./prompts by default. The code keeps only
# prompt file references so prompt wording can be edited without changing Python.
PROMPTS_DIR = Path(os.environ.get("PROMPTS_DIR", "./prompts"))
PROMPT_FILES = {
    "text_system": "text_system.txt",
    "figure_system": "figure_system.txt",
    "constants_extraction": "constants_extraction.txt",
    "constants_field_definitions": "constants_field_definitions.txt",
    "figure_extraction": "figure_extraction.txt",
    "panel_series_discovery": "panel_series_discovery.txt",
    "single_series_extraction": "single_series_extraction.txt",
    "catalyst_profile_enrichment": "catalyst_profile_enrichment.txt",
    "strict_retry_block": "strict_retry_block.txt",
    "previous_bad_output_block": "previous_bad_output_block.txt",
    "expected_series_block": "expected_series_block.txt",
    "expected_series_counts_block": "expected_series_counts_block.txt",
    "expected_panel_specs_block": "expected_panel_specs_block.txt",
}
_PROMPT_CACHE: Dict[str, str] = {}


def _set_prompts_dir(path: str) -> None:
    global PROMPTS_DIR
    PROMPTS_DIR = Path(path)
    _PROMPT_CACHE.clear()


def _prompt_file_path(prompt_name: str) -> Path:
    if prompt_name not in PROMPT_FILES:
        raise KeyError(f"Unknown prompt template: {prompt_name}")
    return PROMPTS_DIR / PROMPT_FILES[prompt_name]


def _load_prompt(prompt_name: str) -> str:
    if prompt_name in _PROMPT_CACHE:
        return _PROMPT_CACHE[prompt_name]
    path = _prompt_file_path(prompt_name)
    if not path.exists():
        raise FileNotFoundError(
            f"Prompt file not found: {path}. Use --prompts-dir to point to a directory containing prompt templates."
        )
    prompt_text = path.read_text(encoding="utf-8").strip()
    _PROMPT_CACHE[prompt_name] = prompt_text
    return prompt_text


def _render_prompt_template(prompt_name: str, **values: Any) -> str:
    prompt_text = _load_prompt(prompt_name)
    for key, value in values.items():
        prompt_text = prompt_text.replace("{{" + key + "}}", _safe_str(value))
    # Drop lines containing optional placeholders left blank.
    prompt_text = re.sub(r"^.*\{\{[A-Z0-9_]+\}\}.*$", "", prompt_text, flags=re.MULTILINE)
    prompt_text = re.sub(r"\n{3,}", "\n\n", prompt_text).strip()
    return prompt_text


def _prompt_ref(prompt_name: str) -> str:
    return f"PROMPT_FILE:{prompt_name}"


def _resolve_prompt_reference(prompt_text_or_ref: str) -> str:
    if isinstance(prompt_text_or_ref, str) and prompt_text_or_ref.startswith("PROMPT_FILE:"):
        return _load_prompt(prompt_text_or_ref.split(":", 1)[1])
    return prompt_text_or_ref


SEND_ALL_FIGURES_TO_LLM = False


def _set_send_all_figures_to_llm(value: bool) -> None:
    global SEND_ALL_FIGURES_TO_LLM
    SEND_ALL_FIGURES_TO_LLM = bool(value)


HEATING_MODE_CONTROLLED_VOCAB = (
    'Heating_Mode should identify the actual reaction heat source when reported; '
    'examples include microwave, electric furnace, tubular furnace, muffle furnace, '
    'induction heating, plasma-assisted, solar heating, oil bath, hot plate, and unresolved furnace, '
    'but the value is not restricted to these examples.'
)

TEXT_SYSTEM_PROMPT = _prompt_ref("text_system")
FIGURE_SYSTEM_PROMPT = _prompt_ref("figure_system")
SYSTEM_PROMPT = TEXT_SYSTEM_PROMPT


INTERNAL_NOTES_KEY = "_Internal_Notes"
FIGURE_CONTEXT_MAX_CHARS = int(os.environ.get("FIGURE_CONTEXT_MAX_CHARS", "5000"))
CH4_CONVERSION_FINALIZED_KEY = "_CH4_Conversion_Finalized"

# Normalize unicode chemistry/axis text such as CH₄, CH⁴, CH₄(%) so
# deterministic text rules recognize methane labels consistently.
UNICODE_SUBSCRIPT_SUPERSCRIPT_TRANSLATION = str.maketrans({
    "₀": "0", "₁": "1", "₂": "2", "₃": "3", "₄": "4",
    "₅": "5", "₆": "6", "₇": "7", "₈": "8", "₉": "9",
    "⁰": "0", "¹": "1", "²": "2", "³": "3", "⁴": "4",
    "⁵": "5", "⁶": "6", "⁷": "7", "⁸": "8", "⁹": "9",
    "％": "%",
})


NONCHARACTER_RE = re.compile(r"[\uFFFE\uFFFF]")
CONTROL_ARTIFACT_RE = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F]")

TEXT_ARTIFACT_TRANSLATION = str.maketrans({
    "−": "-",
    "–": "-",
    "—": "-",
    "‐": "-",
    "‑": "-",
    "⁻": "-",
    "％": "%",
    "∼": "~",
    "≈": "~",
})


def _clean_text_artifacts(value: Any) -> str:
    """Remove invalid/control extraction artifacts while preserving scientific text."""
    if value is None:
        return ""
    s = str(value)
    s = s.translate(UNICODE_SUBSCRIPT_SUPERSCRIPT_TRANSLATION)
    s = s.translate(TEXT_ARTIFACT_TRANSLATION)
    s = s.replace("\ufeff", "").replace("\ufffd", "")
    s = NONCHARACTER_RE.sub("", s)
    # Common PDF artifact: a control character where a minus/superscript minus
    # should be, e.g. min\x011, g\x011, or \x01196 °C.
    s = re.sub(r"(?<=[A-Za-z0-9\)\]])[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F]+(?=\d)", "-", s)
    s = re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F]+(?=\d)", "-", s)
    s = CONTROL_ARTIFACT_RE.sub("", s)
    return s


def _safe_str(v: Any) -> str:
    return _clean_text_artifacts(v)


def _norm_text(v: Any) -> str:
    s = _safe_str(v).translate(UNICODE_SUBSCRIPT_SUPERSCRIPT_TRANSLATION)
    s = s.replace("–", "-").replace("—", "-").replace("−", "-")
    s = s.replace("◦", "°").replace("º", "°").replace("℃", "°C")
    s = re.sub(r"\s+", " ", s).strip()
    return s.lower()


def _display_norm(v: Any) -> str:
    s = _safe_str(v).strip()
    return s if s else "NR"


def _compact_text_for_prompt(text: Any, max_chars: Optional[int] = None) -> str:
    s = re.sub(r"\s+", " ", _safe_str(text)).strip()
    if max_chars is not None and max_chars > 0 and len(s) > max_chars:
        s = s[:max_chars].rstrip() + "..."
    return s


def normalize_blank_to_NR(v: Any) -> Any:
    if v is None:
        return "NR"
    if isinstance(v, float) and pd.isna(v):
        return "NR"
    if isinstance(v, str) and v.strip() == "":
        return "NR"
    return v


def _field_source_key(field_name: str) -> str:
    return f"{INTERNAL_FIELD_SOURCE_PREFIX}{field_name}"


def _set_row_field_sources(row: Dict[str, Any], record: Dict[str, Any], constants: Dict[str, Any], fields: List[str]) -> None:
    for fld in fields:
        rec_val = _display_norm(record.get(fld, "NR"))
        const_val = _display_norm(constants.get(fld, "NR"))
        if rec_val != "NR":
            row[_field_source_key(fld)] = "figure"
        elif const_val != "NR":
            row[_field_source_key(fld)] = "paper"
        else:
            row[_field_source_key(fld)] = "NR"


def _get_row_field_source(row: Dict[str, Any], field_name: str) -> str:
    return _display_norm(row.get(_field_source_key(field_name), "NR"))


def _append_note(existing: str, extra: str) -> str:
    existing = existing if existing not in (None, "", "NR") else ""
    extra = _safe_str(extra).strip()
    if not extra:
        return existing or "NR"
    if existing:
        if extra in existing:
            return existing
        return f"{existing} | {extra}"
    return extra


def _merge_notes_uniquely(*values: Any) -> str:
    seen: List[str] = []
    for v in values:
        s = _display_norm(v)
        if s == "NR":
            continue
        if s not in seen:
            seen.append(s)
    return " | ".join(seen) if seen else "NR"


def _parse_first_float(val: Any) -> Optional[float]:
    if val is None:
        return None
    s = _safe_str(val).strip()
    if s in ("NR", "ERROR_unphysical"):
        return None
    s = s.replace("≈", "").replace("~", "")
    m = re.search(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?", s)
    if not m:
        return None
    try:
        return float(m.group(0))
    except ValueError:
        return None


def _parse_time_to_minutes(val: Any) -> Optional[float]:
    if val is None:
        return None
    s = _safe_str(val).strip()
    if s in ("NR", "ERROR_unphysical"):
        return None
    s = s.replace("≈", "").replace("~", "").strip()
    m = re.search(r"([-+]?\d*\.?\d+(?:[eE][-+]?\d+)?)\s*([a-zA-Zµμ°/\-0-9]*)", s)
    if not m:
        return None
    try:
        x = float(m.group(1))
    except ValueError:
        return None
    unit = (m.group(2) or "").lower()
    if unit in ("h", "hr", "hrs", "hour", "hours"):
        return x * 60.0
    if unit in ("min", "mins", "m"):
        return x
    if unit in ("s", "sec", "secs", "second", "seconds"):
        return x / 60.0
    return x


def _canonical_time_unit(unit: Any) -> Optional[str]:
    u = _norm_text(unit)
    if u in {"h", "hr", "hrs", "hour", "hours"}:
        return "h"
    if u in {"min", "mins", "minute", "minutes", "m"}:
        return "min"
    if u in {"s", "sec", "secs", "second", "seconds"}:
        return "s"
    return None


def _infer_time_unit_from_axis_label(label: Any) -> Optional[str]:
    text = _safe_str(label)
    if not text:
        return None
    norm = _norm_text(text)
    # Prefer explicit bracketed/slashed axis units, e.g. Time (min), TOS / h, time [s].
    unit_patterns = [
        ("min", r"(?:\(|\[|/|\b)(?:min|mins|minute|minutes)(?:\)|\]|\b)"),
        ("h", r"(?:\(|\[|/|\b)(?:h|hr|hrs|hour|hours)(?:\)|\]|\b)"),
        ("s", r"(?:\(|\[|/|\b)(?:s|sec|secs|second|seconds)(?:\)|\]|\b)"),
    ]
    for unit, pat in unit_patterns:
        if re.search(pat, norm):
            return unit
    return None


def _normalize_reaction_time_field(value: Any, x_axis_label: Any = "", context_text: Any = "") -> str:
    s = _safe_str(value).strip()
    if not s or s == "NR":
        return "NR"
    s = s.replace("≈", "~").replace("−", "-")
    s = re.sub(r"\s+", " ", s).strip()
    approx = "~" if "~" in s else ""
    m = TIME_HINT_PATTERN.search(s)
    if m:
        val = _format_number_clean(float(m.group("val")))
        unit = _canonical_time_unit(m.group("unit"))
        return f"{approx}{val} {unit}" if unit else "NR"

    # Unitless numeric reaction times are allowed only when the selected time axis supplies the unit.
    num = _parse_first_float(s)
    if num is None:
        return "NR"
    unit = _infer_time_unit_from_axis_label(x_axis_label)
    if unit is None:
        unit = _infer_time_unit_from_axis_label(context_text)
    if unit is None:
        return "NR"
    return f"{approx}{_format_number_clean(num)} {unit}"


def _finalize_reaction_time_field(row: Dict[str, Any]) -> None:
    old_val = _display_norm(row.get("Reaction_Time_hr_min_s", "NR"))
    new_val = _normalize_reaction_time_field(
        old_val,
        row.get("X_Axis_Label", ""),
        row.get("_Caption_Context", ""),
    )
    if old_val != "NR" and new_val == "NR":
        row[INTERNAL_NOTES_KEY] = _append_note(
            row.get(INTERNAL_NOTES_KEY, "NR"),
            f"Reaction_Time_hr_min_s cleared because no supported time unit was available: {old_val}.",
        )
    elif old_val != "NR" and old_val != new_val:
        row[INTERNAL_NOTES_KEY] = _append_note(
            row.get(INTERNAL_NOTES_KEY, "NR"),
            f"Reaction_Time_hr_min_s normalized with explicit units: {old_val} -> {new_val}.",
        )
    row["Reaction_Time_hr_min_s"] = new_val


def _format_number_clean(x: float) -> str:
    if abs(x - round(x)) < 1e-9:
        return str(int(round(x)))
    return f"{x:.6f}".rstrip("0").rstrip(".")


def _to_celsius(value: float, unit_raw: str) -> float:
    unit = _safe_str(unit_raw).strip().lower().replace("°", "")
    if unit in {"c", "degc", "celsius"}:
        return value
    if unit in {"k", "degk", "kelvin"}:
        return value - 273.15
    if unit in {"f", "degf", "fahrenheit"}:
        return (value - 32.0) * 5.0 / 9.0
    if unit in {"r", "rankine"}:
        return (value - 491.67) * 5.0 / 9.0
    return value


def _normalize_multistep_separator(value: Any) -> str:
    s = _safe_str(value).strip()
    if not s or s == "NR":
        return "NR"
    s = s.replace(" followed by ", " + ").replace(" then ", " + ")
    s = s.replace(";", " + ")
    s = re.sub(r"\s*->\s*", " + ", s)
    s = re.sub(r"\s+plus\s+", " + ", s, flags=re.IGNORECASE)
    s = re.sub(r"\s*\+\s*", " + ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _normalize_multistep_temperature_field(value: Any) -> str:
    s = _normalize_multistep_separator(value)
    if s == "NR":
        return "NR"

    def repl(match: re.Match) -> str:
        raw = match.group(0)
        if _temperature_match_is_rate_context(s, match):
            return raw
        try:
            val = float(match.group("val"))
        except Exception:
            return raw
        c_val = _to_celsius(val, match.group("unit"))
        return f"{_format_number_clean(c_val)} °C"

    s = TEMP_PATTERN.sub(repl, s)
    s = re.sub(r"\s+", " ", s).strip()

    parts = [p.strip() for p in re.split(r"\s+\+\s+", s) if p.strip()]
    normalized_parts: List[str] = []
    for part in parts or [s]:
        if TEMP_PATTERN.search(part):
            normalized_parts.append(part)
            continue
        m_num = re.fullmatch(r"(?P<approx>[~]?)\s*(?P<val>[-+]?\d{1,4}(?:\.\d+)?)", part)
        if m_num:
            normalized_parts.append(f"{m_num.group('approx')}{_format_number_clean(float(m_num.group('val')))} °C")
            continue
        if _norm_text(part) in {"not reported", "not mentioned", "none", "unknown", "na", "n/a", "-"}:
            normalized_parts.append("NR")
        else:
            # Temperature fields must carry explicit units; unsupported text is safer as NR.
            normalized_parts.append("NR")
    normalized_parts = [p for p in normalized_parts if p and p != "NR"]
    return " + ".join(normalized_parts) if normalized_parts else "NR"


def _normalize_multistep_time_field(value: Any) -> str:
    s = _normalize_multistep_separator(value)
    if not s or s == "NR":
        return "NR"
    parts = [p.strip() for p in re.split(r"\s+\+\s+", s) if p.strip()]
    units = []
    for part in parts:
        m = TIME_HINT_PATTERN.search(part)
        if m:
            unit = _canonical_time_unit(m.group("unit"))
            if unit:
                units.append(unit)
    inferred_unit = units[0] if units and len(set(units)) == 1 else None
    out: List[str] = []
    for part in parts or [s]:
        approx = "~" if "~" in part else ""
        m = TIME_HINT_PATTERN.search(part)
        if m:
            unit = _canonical_time_unit(m.group("unit"))
            if unit:
                out.append(f"{approx}{_format_number_clean(float(m.group('val')))} {unit}")
            continue
        m_num = re.fullmatch(r"[~]?\s*([-+]?\d*\.?\d+(?:[eE][-+]?\d+)?)", part)
        if m_num:
            val = float(m_num.group(1))
            if abs(val) < 1e-12:
                out.append("0 h")
            elif inferred_unit:
                out.append(f"{approx}{_format_number_clean(val)} {inferred_unit}")
            else:
                out.append("NR")
            continue
        if _norm_text(part) not in {"not reported", "not mentioned", "none", "unknown", "na", "n/a", "-"}:
            out.append("NR")
    out = [p for p in out if p and p != "NR"]
    return " + ".join(out) if out else "NR"


def _temperature_match_is_rate_context(s: str, match: re.Match) -> bool:
    tail = s[match.end(): match.end() + 20].lower()
    if re.match(r"\s*(?:/|per\s+)?\s*(?:min|minute|minutes|h|hr|hrs|hour|hours)\b", tail):
        return True
    if re.match(r"\s*(?:min|minute|minutes|h|hr|hrs|hour|hours)\s*-?1\b", tail):
        return True
    lead = s[max(0, match.start() - 12): match.start()].lower()
    if "δh" in lead or "Δh" in lead or re.search(r"(?:delta\s*h|\bdh)\s*$", lead):
        return True
    return False


def _temperature_text_has_range(text: Any) -> bool:
    s = _safe_str(text)
    if not s:
        return False
    s = s.replace("◦", "°").replace("º", "°").replace("℃", "°C")
    if re.search(r"\d{2,4}(?:\.\d+)?\s*[-–]\s*\d{2,4}(?:\.\d+)?\s*(?:°\s*)?(?:C|K|F|R)\b", s, flags=re.IGNORECASE):
        return True
    if re.search(r"between\s+\d{2,4}(?:\.\d+)?\s*(?:°\s*)?(?:C|K|F|R)?\s+and\s+\d{2,4}(?:\.\d+)?\s*(?:°\s*)?(?:C|K|F|R)\b", s, flags=re.IGNORECASE):
        return True
    return False


def _context_mentions_temperature_ramp(text: Any) -> bool:
    t = _norm_text(text)
    if not t:
        return False
    return (
        "room temperature" in t
        and ("rate of" in t or "ramps" in t or "ramp" in t or "increases gradually" in t)
        and ("designated reaction temperature" in t or "reaction temperature" in t)
    )


def _extract_temperature_tokens(text: Any) -> List[str]:
    s = _safe_str(text)
    if not s:
        return []
    s = s.replace("◦", "°").replace("º", "°").replace("℃", "°C")
    temps: List[str] = []
    seen = set()
    for m in TEMP_PATTERN.finditer(s):
        if _temperature_match_is_rate_context(s, m):
            continue
        try:
            val = float(m.group("val"))
        except ValueError:
            continue
        c_val = _to_celsius(val, m.group("unit"))
        tok = f"{_format_number_clean(c_val)} °C"
        if tok not in seen:
            seen.add(tok)
            temps.append(tok)
    return temps


def _extract_single_temperature_with_snippet(text: str) -> Tuple[Optional[str], Optional[str]]:
    if not text:
        return None, None
    if _temperature_text_has_range(text):
        return None, None
    temps = _extract_temperature_tokens(text)
    unique = list(dict.fromkeys(temps))
    if len(unique) != 1:
        return None, None
    temp = unique[0]
    sentences = _sentence_tokenize(text)
    for sent in sentences:
        if _temperature_text_has_range(sent):
            continue
        sent_temps = _extract_temperature_tokens(sent)
        if sent_temps and temp in sent_temps:
            snippet = re.sub(r"\s+", " ", sent).strip()
            if len(snippet) > 140:
                snippet = snippet[:137].rstrip() + "..."
            return temp, snippet
    snippet = re.sub(r"\s+", " ", text).strip()
    if len(snippet) > 140:
        snippet = snippet[:137].rstrip() + "..."
    return temp, snippet


def _normalize_flow_unit(unit: str) -> Optional[str]:
    u = re.sub(r"\s+", "", _safe_str(unit)).lower()
    if u in {"sccm"}:
        return "sccm"
    if u in {"slpm", "slm", "lpm"}:
        return "SLPM"
    if u in {"ml/min", "mlmin-1", "cc/min", "ccmin-1", "cm3/min", "cm3min-1"}:
        return "mL/min"
    if u in {"ml/h", "mlh-1", "mlhr-1", "cc/h", "cch-1", "cchr-1", "cm3/h", "cm3h-1", "cm3hr-1"}:
        return "mL/h"
    if u in {"l/min", "lmin-1"}:
        return "L/min"
    if u in {"l/h", "lh-1", "lhr-1"}:
        return "L/h"
    return None


def _parse_flow_rate_string(val: Any) -> Optional[Tuple[float, str, bool]]:
    s = _safe_str(val).strip()
    if not s or s == "NR":
        return None
    m = FLOW_PATTERN.search(s)
    if not m:
        return None
    try:
        x = float(m.group("val"))
    except ValueError:
        return None
    unit = _normalize_flow_unit(m.group("unit"))
    if not unit:
        return None
    approx = ("~" in s) or ("≈" in s)
    return x, unit, approx


def _format_flow_rate(value: float, unit: str, approx: bool = False) -> str:
    prefix = "~" if approx else ""
    return f"{prefix}{_format_number_clean(value)} {unit}"


def _standardize_space_velocity(value: Any) -> str:
    s = _safe_str(value).strip()
    if not s or s == "NR":
        return "NR"
    s = s.replace("–", "-").replace("—", "-").replace("−", "-")
    s = re.sub(r"\s+", " ", s).strip()
    return s


MASS_PATTERN = re.compile(
    r"(?P<val>[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?)\s*"
    r"(?P<unit>kg|g|mg|ug|μg|µg)\b",
    flags=re.IGNORECASE,
)


def _normalize_mass_unit(unit: str) -> Optional[str]:
    u = re.sub(r"\s+", "", _safe_str(unit)).lower()
    if u == "kg":
        return "kg"
    if u == "g":
        return "g"
    if u == "mg":
        return "mg"
    if u in {"ug", "μg", "µg"}:
        return "ug"
    return None


def _parse_mass_string(val: Any) -> Optional[Tuple[float, str, bool]]:
    s = _safe_str(val).strip()
    if not s or s == "NR":
        return None
    matches = list(MASS_PATTERN.finditer(s))
    if len(matches) != 1:
        return None
    m = matches[0]
    try:
        x = float(m.group("val"))
    except ValueError:
        return None
    unit = _normalize_mass_unit(m.group("unit"))
    if not unit:
        return None
    approx = ("~" in s) or ("≈" in s)
    return x, unit, approx


def _space_velocity_unit_from_ratio(flow_unit: str, mass_unit: str) -> str:
    if flow_unit in {"mL/min", "L/min"}:
        volume_unit, time_unit = flow_unit.split("/", 1)
        return f"{volume_unit} {mass_unit}^-1 {time_unit}^-1"
    return f"{flow_unit} {mass_unit}^-1"


def _format_space_velocity(value: float, unit: str, approx: bool = False) -> str:
    prefix = "~" if approx else ""
    return f"{prefix}{_format_number_clean(value)} {unit}"


def _derive_space_velocity_from_flow_and_mass(flow_str: Any, mass_str: Any) -> Optional[str]:
    parsed_flow = _parse_flow_rate_string(flow_str)
    parsed_mass = _parse_mass_string(mass_str)
    if not parsed_flow or not parsed_mass:
        return None
    flow_value, flow_unit, flow_approx = parsed_flow
    mass_value, mass_unit, mass_approx = parsed_mass
    if mass_value == 0:
        return None
    sv_value = flow_value / mass_value
    sv_unit = _space_velocity_unit_from_ratio(flow_unit, mass_unit)
    return _format_space_velocity(sv_value, sv_unit, approx=(flow_approx or mass_approx))


def _space_velocity_string_is_ambiguous(value: Any) -> bool:
    s = _safe_str(value).strip()
    if not s or s == "NR":
        return False
    nums = re.findall(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?", s)
    lowered = _norm_text(s)
    if len(nums) > 1 and any(sep in lowered for sep in [",", ";", " and "]):
        return True
    return False


def _extract_ch4_fraction_candidates(text: Any) -> List[float]:
    s = _safe_str(text)
    if not s:
        return []
    norm = s.replace("–", "-").replace("—", "-").replace("−", "-")
    fracs: List[float] = []

    patterns = [
        r"(\d+(?:\.\d+)?)\s*(?:vol\s*%|%)\s*(?:CH4|methane)\b",
        r"(?:CH4|methane)\s*(?:=|:)?\s*(\d+(?:\.\d+)?)\s*(?:vol\s*%|%)",
        r"(?:CH4|methane)\s+concentration\s*(?:of|=|:)?\s*(\d+(?:\.\d+)?)\s*(?:vol\s*%|%)",
        r"(\d+(?:\.\d+)?)\s*(?:vol\s*%|%)\s*(?:as\s+the\s+reactant|reactant)\b",
    ]
    for pat in patterns:
        for m in re.finditer(pat, norm, flags=re.IGNORECASE):
            try:
                pct = float(m.group(1))
            except ValueError:
                continue
            if 0 <= pct <= 100:
                fracs.append(pct / 100.0)

    ratio_patterns = [
        r"CH4\s*[/:-]\s*[A-Za-z0-9]+\s*[=:]\s*(\d+(?:\.\d+)?)\s*[/:-]\s*(\d+(?:\.\d+)?)",
        r"methane\s*[/:-]\s*[A-Za-z0-9]+\s*[=:]\s*(\d+(?:\.\d+)?)\s*[/:-]\s*(\d+(?:\.\d+)?)",
    ]
    for pat in ratio_patterns:
        for m in re.finditer(pat, norm, flags=re.IGNORECASE):
            try:
                a = float(m.group(1))
                b = float(m.group(2))
            except ValueError:
                continue
            if a >= 0 and b >= 0 and (a + b) > 0:
                fracs.append(a / (a + b))

    out: List[float] = []
    seen = set()
    for frac in fracs:
        key = round(frac, 6)
        if key not in seen:
            seen.add(key)
            out.append(frac)
    return out


def _extract_single_ch4_fraction(text: Any) -> Optional[float]:
    vals = _extract_ch4_fraction_candidates(text)
    if len(vals) == 1:
        return vals[0]
    return None


def _flow_string_has_multiple_values(text: Any) -> bool:
    s = _safe_str(text).strip()
    if not s or s == "NR":
        return False
    num_tokens = re.findall(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?", s)
    if len(num_tokens) <= 1:
        return False
    lowered = _norm_text(s)
    if len(FLOW_PATTERN.findall(s)) > 1:
        return True
    return any(sep in lowered for sep in [",", ";", " and ", "/", "\\"])


def _extract_single_direct_flow_from_text(text: Any) -> Optional[str]:
    s = _safe_str(text).strip()
    if not s or s == "NR" or _flow_string_has_multiple_values(s):
        return None
    parsed = _parse_flow_rate_string(s)
    if not parsed:
        return None
    value, unit, approx = parsed
    return _format_flow_rate(value, unit, approx=approx)


def _derive_series_specific_ch4_flow(series_text: Any, total_flow_str: Any, context_text: Any = "") -> Optional[str]:
    series_norm = _norm_text(series_text)
    ctx_norm = _norm_text(context_text)

    direct = _extract_single_direct_flow_from_text(series_text)
    if direct and (
        "ch4" in series_norm
        or "methane" in series_norm
        or "methane flow" in ctx_norm
        or "ch4 flow" in ctx_norm
        or "methane feed flow" in ctx_norm
    ):
        return direct

    frac = _extract_single_ch4_fraction(series_text)
    if frac is None:
        return None
    parsed_total = _parse_flow_rate_string(total_flow_str)
    if not parsed_total:
        return None
    value, unit, approx = parsed_total
    return _format_flow_rate(value * frac, unit, approx=approx)


def _derive_ch4_flow_from_total_flow(total_flow_str: Any, context_text: Any) -> Optional[str]:
    parsed_flow = _parse_flow_rate_string(total_flow_str)
    if not parsed_flow:
        return None
    frac = _extract_single_ch4_fraction(context_text)
    if frac is None:
        return None
    value, unit, approx = parsed_flow
    ch4_flow = value * frac
    return _format_flow_rate(ch4_flow, unit, approx=approx)


def _canonical_gas_name(name: str) -> str:
    s = _norm_text(name)
    mapping = {
        "ch4": "CH4",
        "methane": "CH4",
        "n2": "N2",
        "nitrogen": "N2",
        "ar": "Ar",
        "argon": "Ar",
        "he": "He",
        "helium": "He",
        "h2": "H2",
        "hydrogen": "H2",
        "co2": "CO2",
        "carbon dioxide": "CO2",
        "co": "CO",
        "o2": "O2",
        "oxygen": "O2",
        "steam": "H2O",
        "water vapor": "H2O",
        "h2o": "H2O",
        "air": "Air",
    }
    return mapping.get(s, name.strip())


def _format_feed_gas_mixture(components: List[Tuple[str, float]]) -> str:
    parts = []
    for gas, pct in components:
        parts.append(f"{_format_number_clean(pct)}% {gas}")
    return " + ".join(parts) if parts else "NR"


def _extract_feed_gas_mixture_candidates(text: Any) -> List[str]:
    s = _safe_str(text)
    if not s:
        return []
    candidates: List[str] = []
    seen = set()

    def add_candidate(components: List[Tuple[str, float]]):
        if len(components) < 2:
            return
        total = sum(p for _g, p in components)
        if not (90 <= total <= 110):
            return
        norm_components = []
        seen_local = set()
        for gas, pct in components:
            gasc = _canonical_gas_name(gas)
            key = (gasc, round(float(pct), 6))
            if key not in seen_local:
                seen_local.add(key)
                norm_components.append((gasc, float(pct)))
        mix = _format_feed_gas_mixture(norm_components)
        key = _norm_text(mix)
        if key and key not in seen:
            seen.add(key)
            candidates.append(mix)

    for sent in _sentence_tokenize(s):
        comps: List[Tuple[str, float]] = []
        for m in GAS_COMPONENT_PERCENT_PATTERN.finditer(sent):
            try:
                pct = float(m.group("pct"))
            except ValueError:
                continue
            comps.append((_canonical_gas_name(m.group("gas")), pct))
        if len(comps) >= 2:
            add_candidate(comps)
            continue
        comps = []
        for m in GAS_PERCENT_COMPONENT_PATTERN.finditer(sent):
            try:
                pct = float(m.group("pct"))
            except ValueError:
                continue
            comps.append((_canonical_gas_name(m.group("gas")), pct))
        if len(comps) >= 2:
            add_candidate(comps)
            continue

        ratio_match = re.search(
            r"(?P<gases>(?:CH4|methane|N2|nitrogen|Ar|argon|He|helium|H2|hydrogen|CO2|carbon\s+dioxide|CO|O2|oxygen|air)(?:\s*[/:-]\s*(?:CH4|methane|N2|nitrogen|Ar|argon|He|helium|H2|hydrogen|CO2|carbon\s+dioxide|CO|O2|oxygen|air))+)",
            sent,
            flags=re.IGNORECASE,
        )
        if ratio_match:
            tail = sent[ratio_match.end():]
            nums_match = re.search(r"(?:=|:)\s*(?P<nums>\d+(?:\.\d+)?(?:\s*[/:-]\s*\d+(?:\.\d+)?)+)", tail)
            if nums_match:
                gases = [g.strip() for g in re.split(r"\s*[/:-]\s*", ratio_match.group("gases"))]
                nums = [float(x) for x in re.split(r"\s*[/:-]\s*", nums_match.group("nums"))]
                if len(gases) == len(nums) and len(gases) >= 2 and sum(nums) > 0:
                    total = sum(nums)
                    comps = [(_canonical_gas_name(g), n / total * 100.0) for g, n in zip(gases, nums)]
                    add_candidate(comps)

    return candidates


def _extract_single_feed_gas_mixture(text: Any) -> Optional[str]:
    vals = _extract_feed_gas_mixture_candidates(text)
    if len(vals) == 1:
        return vals[0]
    return None


def _derive_binary_feed_mixture_from_ch4_fraction(series_text: Any, context_text: Any) -> Optional[str]:
    frac = _extract_single_ch4_fraction(series_text)
    if frac is None:
        frac = _extract_single_ch4_fraction(context_text)
    if frac is None:
        return None
    context = " ".join([_safe_str(series_text), _safe_str(context_text)])
    gas_names = set()
    for tok in ["N2", "nitrogen", "Ar", "argon", "He", "helium", "CO2", "carbon dioxide", "CO", "O2", "oxygen", "air"]:
        if re.search(rf"\b{re.escape(tok)}\b", context, flags=re.IGNORECASE):
            gas_names.add(_canonical_gas_name(tok))
    if len(gas_names) != 1:
        return None
    balance_gas = list(gas_names)[0]
    ch4_pct = frac * 100.0
    return _format_feed_gas_mixture([("CH4", ch4_pct), (balance_gas, 100.0 - ch4_pct)])


def _standardize_feed_gas_mixture(value: Any) -> str:
    s = _safe_str(value).strip()
    if not s or s == "NR":
        return "NR"
    single = _extract_single_feed_gas_mixture(s)
    if single:
        return single
    return re.sub(r"\s+", " ", s).strip()


def _standardize_marker_color(value: Any) -> str:
    s = _norm_text(value)
    if not s or s == "nr":
        return "NR"
    for key, val in MARKER_COLOR_CANONICAL.items():
        if key in s:
            return val
    if s in {"filled", "open", "hollow"}:
        return "NR"
    return "other"


def _standardize_marker_shape(value: Any) -> str:
    s = _norm_text(value)
    if not s or s == "nr":
        return "NR"

    shape = None
    for key, val in sorted(MARKER_SHAPE_CANONICAL.items(), key=lambda kv: len(kv[0]), reverse=True):
        if re.search(rf"\b{re.escape(key)}\b", s):
            shape = val
            break

    fill = None
    for key, val in MARKER_FILL_CANONICAL.items():
        if re.search(rf"\b{re.escape(key)}\b", s):
            fill = val
            break

    if shape is None:
        if fill == "filled":
            return "filled marker"
        if fill == "open":
            return "open marker"
        return "other"

    if fill and shape not in {"plus", "cross", "x"}:
        return f"{fill} {shape}"
    return shape


def _standardize_heating_mode(value: Any) -> str:
    s = _norm_text(value)
    if not s or s == "nr":
        return "NR"
    if "microwave" in s:
        return "microwave"
    if "induction" in s:
        return "induction heating"
    if "plasma" in s:
        return "plasma-assisted"
    if "solar" in s:
        return "solar heating"
    if any(tok in s for tok in ["muffle furnace", "muffle-heated", "muffle"]):
        return "muffle furnace"
    if any(tok in s for tok in ["tubular furnace", "tube furnace", "quartz tube furnace", "tubular reactor"]):
        return "tubular furnace"
    if any(tok in s for tok in ["electric furnace", "electric heating", "electric heated", "electrically heated"]):
        return "electric furnace"
    if any(tok in s for tok in ["conventional heating", "conventional"]):
        return "~furnace"
    if any(tok in s for tok in ["oil bath", "sand bath", "heating mantle", "hot plate"]):
        return "others"
    if "furnace" in s:
        return "~furnace"
    return "others"


def _infer_heating_mode(*texts: Any) -> str:
    joined = " ".join([_safe_str(t) for t in texts if _safe_str(t).strip()])
    return _standardize_heating_mode(joined)


def _normalize_precursor_salts(value: Any) -> str:
    s = _safe_str(value).strip()
    if not s or s == "NR":
        return "NR"
    s = s.replace(" and ", " + ").replace(" AND ", " + ")
    s = s.replace(";", " + ").replace(",", " + ")
    parts = [re.sub(r"\s+", " ", p).strip(" +") for p in re.split(r"\s*\+\s*", s) if re.sub(r"\s+", " ", p).strip(" +")]
    out: List[str] = []
    seen = set()
    for p in parts:
        key = _norm_text(p)
        if key and key not in seen:
            seen.add(key)
            out.append(p)
    return " + ".join(out) if out else "NR"


def _extract_precursor_salts_from_text(text: Any) -> Optional[str]:
    s = _safe_str(text)
    if not s:
        return None

    formulas: List[str] = []
    for m in PRECURSOR_FORMULA_PATTERN.finditer(s):
        tok = re.sub(r"\s+", "", m.group(0))
        if any(key in tok.upper() for key in ["NO3", "CL", "CH3COO", "OAC", "SO4", "CO3", "C2O4", "OH"]):
            formulas.append(tok)

    names: List[str] = []
    for m in PRECURSOR_NAME_PATTERN.finditer(s):
        tok = re.sub(r"\s+", " ", m.group(0)).strip()
        names.append(tok)

    candidates = formulas if formulas else names
    cleaned: List[str] = []
    seen = set()
    for c in candidates:
        key = _norm_text(c)
        if key and key not in seen:
            seen.add(key)
            cleaned.append(c)
    if not cleaned:
        return None
    return _normalize_precursor_salts(" + ".join(cleaned))


IGNORED_PRECURSOR_MATCH_ELEMENTS = {"H", "C", "N", "O", "P", "S", "Cl", "Br", "I", "F"}

ELEMENT_NAME_HINTS = {
    "nickel": "Ni",
    "iron": "Fe",
    "ferric": "Fe",
    "ferrous": "Fe",
    "cobalt": "Co",
    "copper": "Cu",
    "molybdenum": "Mo",
    "tungsten": "W",
    "chromium": "Cr",
    "manganese": "Mn",
    "zinc": "Zn",
    "palladium": "Pd",
    "platinum": "Pt",
    "rhodium": "Rh",
    "ruthenium": "Ru",
    "iridium": "Ir",
    "silver": "Ag",
    "gold": "Au",
    "lanthanum": "La",
    "cerium": "Ce",
    "yttrium": "Y",
    "zirconium": "Zr",
    "magnesium": "Mg",
    "calcium": "Ca",
    "potassium": "K",
    "sodium": "Na",
    "cesium": "Cs",
    "lithium": "Li",
    "aluminum": "Al",
    "aluminium": "Al",
    "alumina": "Al",
    "silicon": "Si",
    "silica": "Si",
    "orthosilicate": "Si",
    "titania": "Ti",
    "zirconia": "Zr",
    "ceria": "Ce",
    "lanthana": "La",
    "magnesia": "Mg",
    "calcia": "Ca",
}

PROCESS_SENTENCE_KEYWORDS = {
    "Drying": ["dry", "dried", "drying"],
    "Calcination": ["calcined", "calcination", "calcine", "calcining"],
    "Reduction": ["reduced", "reduction", "reduce", "pre-reduced", "prereduced", "h2-reduced", "h2 reduced"],
}

PROCESS_GAS_HINT_GROUPS = {
    "H2": ["h2", "hydrogen"],
    "N2": ["n2", "nitrogen"],
    "Ar": ["ar", "argon"],
    "Air": ["air"],
    "O2": ["o2", "oxygen"],
    "He": ["he", "helium"],
    "CO": ["co", "carbon monoxide"],
    "CO2": ["co2", "carbon dioxide"],
    "CH4": ["ch4", "methane"],
    "NH3": ["nh3", "ammonia"],
    "H2O": ["h2o", "steam", "water vapor", "water vapour"],
    "Vacuum": ["vacuum"],
}

GAS_ALIAS_TO_CANONICAL = {
    alias: canonical
    for canonical, aliases in PROCESS_GAS_HINT_GROUPS.items()
    for alias in aliases + [canonical]
}

SURFACE_AREA_PATTERN = re.compile(
    r"(?P<val>[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?)\s*"
    r"(?P<unit>m(?:\^?2|2|²)\s*(?:/|\s*)\s*g(?:\s*[-^−]?1)?)",
    flags=re.IGNORECASE,
)

SI_AL_RATIO_PATTERN = re.compile(
    r"(?:(?P<name>Si\s*/\s*Al|SiO2\s*/\s*Al2O3|SAR)\s*(?:ratio)?\s*(?:=|:|of)?\s*|"
    r"(?P<name2>Si\s*/\s*Al|SiO2\s*/\s*Al2O3|SAR)\s*(?:ratio)?)"
    r"(?P<val>[-+]?\d*\.?\d+(?:\s*[-–]\s*[-+]?\d*\.?\d+)?)",
    flags=re.IGNORECASE,
)

ZEOLITE_HINT_RE = re.compile(
    r"\b(?:zeolite|zsm[- ]?5|hzsm[- ]?5|h[- ]?zsm[- ]?5|zsm|mfi|usy|y[- ]?zeolite|hy|mordenite|beta|bea|sapo|ssz|fau)\b",
    flags=re.IGNORECASE,
)

COOLING_RATE_PATTERN = re.compile(
    r"(?P<val>[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?)\s*"
    r"(?P<unit>(?:°|º|◦)?\s*(?:C|K)|deg\s*C|deg\s*K)\s*"
    r"(?:/|per\s+|\s*)\s*(?P<time>min|minute|h|hr|hour|s|sec|second)(?:\s*[-^−]?1)?",
    flags=re.IGNORECASE,
)

TIME_HINT_PATTERN = re.compile(
    r"(?P<val>[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?)\s*"
    r"(?P<unit>h|hr|hrs|hour|hours|min|mins|minute|minutes|s|sec|secs|second|seconds)\b",
    flags=re.IGNORECASE,
)


SEMANTIC_SUPPORT_ALIASES = {
    # Carbon supports and carbon-derived supports
    "activated carbon": "activated carbon",
    "active carbon": "activated carbon",
    "ac": "activated carbon",
    "carbon black": "carbon black",
    "cb": "carbon black",
    "carbon nanotubes": "carbon nanotubes",
    "carbon nanotube": "carbon nanotubes",
    "cnts": "carbon nanotubes",
    "cnt": "carbon nanotubes",
    "mwcnt": "multi-walled carbon nanotubes",
    "mwcnts": "multi-walled carbon nanotubes",
    "swcnt": "single-walled carbon nanotubes",
    "swcnts": "single-walled carbon nanotubes",
    "carbon nanofiber": "carbon nanofibers",
    "carbon nanofibers": "carbon nanofibers",
    "cnf": "carbon nanofibers",
    "carbon fiber": "carbon fiber",
    "carbon fibre": "carbon fiber",
    "cf": "carbon fiber",
    "graphene oxide": "graphene oxide",
    "reduced graphene oxide": "reduced graphene oxide",
    "graphene": "graphene",
    "graphite": "graphite",
    "biochar": "biochar",
    "activated coke": "activated coke",
    "mesoporous carbon": "mesoporous carbon",

    # Common metal oxide and metalloid oxide supports
    "al2o3": "alumina (Al2O3)",
    "γ-al2o3": "gamma-alumina (γ-Al2O3)",
    "gamma-alumina": "gamma-alumina (γ-Al2O3)",
    "gamma alumina": "gamma-alumina (γ-Al2O3)",
    "g alumina": "gamma-alumina (γ-Al2O3)",
    "alumina": "alumina (Al2O3)",
    "aluminium oxide": "alumina (Al2O3)",
    "aluminum oxide": "alumina (Al2O3)",
    "sio2": "silica (SiO2)",
    "silica": "silica (SiO2)",
    "silicon dioxide": "silica (SiO2)",
    "tio2": "titania (TiO2)",
    "titania": "titania (TiO2)",
    "titanium dioxide": "titania (TiO2)",
    "zro2": "zirconia (ZrO2)",
    "zirconia": "zirconia (ZrO2)",
    "zirconium dioxide": "zirconia (ZrO2)",
    "ceo2": "ceria (CeO2)",
    "ceria": "ceria (CeO2)",
    "cerium dioxide": "ceria (CeO2)",
    "mgo": "magnesia (MgO)",
    "magnesia": "magnesia (MgO)",
    "magnesium oxide": "magnesia (MgO)",
    "cao": "calcia (CaO)",
    "calcia": "calcia (CaO)",
    "calcium oxide": "calcia (CaO)",
    "la2o3": "lanthana (La2O3)",
    "lanthana": "lanthana (La2O3)",
    "zno": "zinc oxide (ZnO)",

    # Mixed oxides and spinel-like supports
    "silica-alumina": "silica-alumina (SiO2-Al2O3)",
    "silica alumina": "silica-alumina (SiO2-Al2O3)",
    "sio2-al2o3": "silica-alumina (SiO2-Al2O3)",
    "aluminosilicate": "aluminosilicate",
    "magnesium aluminate spinel": "magnesium aluminate spinel (MgAl2O4)",
    "magnesium aluminate": "magnesium aluminate (MgAl2O4)",
    "mgal2o4": "magnesium aluminate (MgAl2O4)",
    "ceria-zirconia": "ceria-zirconia (CeO2-ZrO2)",
    "ceo2-zro2": "ceria-zirconia (CeO2-ZrO2)",
    "cerium zirconium oxide": "ceria-zirconia (CeO2-ZrO2)",

    # Zeolites and molecular sieves
    "h-zsm-5": "HZSM-5 zeolite",
    "hzsm-5": "HZSM-5 zeolite",
    "zsm-5": "ZSM-5 zeolite",
    "zsm5": "ZSM-5 zeolite",
    "usy": "USY zeolite",
    "hy": "Y zeolite",
    "y zeolite": "Y zeolite",
    "beta zeolite": "Beta zeolite",
    "bea zeolite": "Beta zeolite",
    "mordenite": "mordenite zeolite",
    "sapo": "SAPO zeolite",
    "ssz": "SSZ zeolite",
    "fau": "FAU zeolite",
    "mfi": "MFI zeolite",
    "zeolite": "zeolite",
    "molecular sieve": "molecular sieve",
    "molecular sieves": "molecular sieve",

    # Ordered mesoporous silicas
    "sba-15": "SBA-15 mesoporous silica",
    "mcm-41": "MCM-41 mesoporous silica",
    "mcm-48": "MCM-48 mesoporous silica",
    "kit-6": "KIT-6 mesoporous silica",
    "mesoporous silica": "mesoporous silica",

    # Carbides, nitrides, borides, and related ceramic supports
    "sic": "silicon carbide (SiC)",
    "silicon carbide": "silicon carbide (SiC)",
    "wc": "tungsten carbide (WC)",
    "tungsten carbide": "tungsten carbide (WC)",
    "mo2c": "molybdenum carbide (Mo2C)",
    "molybdenum carbide": "molybdenum carbide (Mo2C)",
    "tic": "titanium carbide (TiC)",
    "titanium carbide": "titanium carbide (TiC)",
    "zrc": "zirconium carbide (ZrC)",
    "zirconium carbide": "zirconium carbide (ZrC)",
    "b4c": "boron carbide (B4C)",
    "boron carbide": "boron carbide (B4C)",
    "bn": "boron nitride (BN)",
    "boron nitride": "boron nitride (BN)",
    "si3n4": "silicon nitride (Si3N4)",
    "silicon nitride": "silicon nitride (Si3N4)",
    "aln": "aluminum nitride (AlN)",
    "aluminum nitride": "aluminum nitride (AlN)",
    "aluminium nitride": "aluminum nitride (AlN)",
    "tin nitride": "titanium nitride (TiN)",
    "titanium nitride": "titanium nitride (TiN)",
    "mo2n": "molybdenum nitride (Mo2N)",
    "molybdenum nitride": "molybdenum nitride (Mo2N)",
    "tib2": "titanium diboride (TiB2)",
    "titanium diboride": "titanium diboride (TiB2)",
    "zrb2": "zirconium diboride (ZrB2)",
    "zirconium diboride": "zirconium diboride (ZrB2)",

    # Clay, mineral, ceramic, monolith, foam, and metal supports
    "cordierite monolith": "cordierite monolith",
    "cordierite": "cordierite",
    "ceramic monolith": "ceramic monolith",
    "monolith": "monolith",
    "ceramic foam": "ceramic foam",
    "metal foam": "metal foam",
    "nickel foam": "nickel foam",
    "fecral": "FeCrAl alloy support",
    "stainless steel": "stainless steel support",
    "kaolin": "kaolin clay",
    "bentonite": "bentonite clay",
    "montmorillonite": "montmorillonite clay",
    "sepiolite": "sepiolite clay",
    "attapulgite": "attapulgite clay",
    "palygorskite": "palygorskite clay",
    "halloysite": "halloysite clay",
    "diatomite": "diatomite",
    "natural clay": "natural clay",

    # Metal-organic frameworks
    "metal organic framework": "metal-organic framework (MOF)",
    "metal-organic framework": "metal-organic framework (MOF)",
    "mof": "metal-organic framework (MOF)",
    "zif-8": "ZIF-8 metal-organic framework",
    "uio-66": "UiO-66 metal-organic framework",
    "mil-101": "MIL-101 metal-organic framework",
    "hkust-1": "HKUST-1 metal-organic framework",
}


METAL_OXIDE_SUPPORT_TOKENS = {
    "al2o3", "tio2", "zro2", "mgo", "cao", "ceo2", "la2o3", "moo3", "wo3",
    "zno", "feo", "fe2o3", "fe3o4", "coo", "co3o4", "cuo", "mno2", "cr2o3", "sno2", "pbo", "pbo2",
}

METALLOID_OXIDE_SUPPORT_TOKENS = {
    "sio2", "b2o3", "geo2", "silica", "silicate", "mesoporous silica", "sba-15", "mcm-41", "mcm-48",
}

CARBIDE_SUPPORT_TOKENS = {
    "sic", "wc", "mo2c", "tic", "zrc", "b4c", "silicon carbide", "tungsten carbide",
    "molybdenum carbide", "titanium carbide", "zirconium carbide", "boron carbide",
}

BORIDE_SUPPORT_TOKENS = {
    "tib2", "zrb2", "lab6", "boride", "diboride", "hexaboride",
}

NITRIDE_SUPPORT_TOKENS = {
    "si3n4", "bn", "aln", "tin nitride", "mo2n", "nitride", "boron nitride", "silicon nitride",
    "aluminum nitride", "aluminium nitride", "titanium nitride", "molybdenum nitride",
}

MOF_SUPPORT_TOKENS = {
    "mof", "metal organic framework", "metal-organic framework", "zif", "zif-8", "uio", "uio-66", "mil-101", "hkust",
}

MIXED_METAL_OXIDE_SUPPORT_TOKENS = {
    "mixed metal oxide", "mixed oxide", "spinel", "perovskite", "hydrotalcite", "aluminate",
    "mgal2o4", "ceo2-zro2", "ce-zr", "laalo3", "sralo3", "lamno3", "lafeo3", "lacoo3",
}


def _support_text_has_token(text: Any, tokens: set[str]) -> bool:
    low = _norm_text(text)
    compact = re.sub(r"[^a-z0-9]+", "", low)
    for tok in tokens:
        tok_low = _norm_text(tok)
        tok_compact = re.sub(r"[^a-z0-9]+", "", tok_low)
        if not tok_compact:
            continue
        if len(tok_compact) <= 3:
            # Short formula tokens such as BN, WC, and TiN must be bounded so
            # words like "carbon" or "tin oxide" do not create false hits.
            if re.search(rf"(?<![a-z0-9]){re.escape(tok_compact)}(?![a-z0-9])", low):
                return True
            continue
        if tok_low in low or tok_compact in compact:
            return True
    return False


def _support_has_mixed_oxide_formula(text: Any) -> bool:
    low = _norm_text(text)
    if re.search(r"\b(?:[a-z]{1,2}\d*o\d*\s*[-/]\s*){1,}[a-z]{1,2}\d*o\d*\b", low):
        oxide_parts = re.findall(r"[a-z]{1,2}\d*o\d*", low)
        metal_like = [p for p in oxide_parts if p not in {"sio2", "b2o3", "geo2"}]
        return len(set(metal_like)) >= 2
    if re.search(r"\b[a-z]{1,2}[a-z]{1,2}\d*o\d+\b", low):
        compact = re.sub(r"[^a-z0-9]+", "", low)
        return any(tok in compact for tok in MIXED_METAL_OXIDE_SUPPORT_TOKENS if re.search(r"\d", tok))
    return False

SEMANTIC_METAL_NAME_TO_SYMBOL = {
    "nickel": "Ni",
    "iron": "Fe",
    "ferrous": "Fe",
    "ferric": "Fe",
    "cobalt": "Co",
    "cobaltous": "Co",
    "cobaltic": "Co",
    "copper": "Cu",
    "cuprous": "Cu",
    "cupric": "Cu",
    "manganese": "Mn",
    "manganous": "Mn",
    "manganic": "Mn",
    "chromium": "Cr",
    "chromous": "Cr",
    "chromic": "Cr",
    "zinc": "Zn",
    "magnesium": "Mg",
    "calcium": "Ca",
    "cadmium": "Cd",
    "titanium": "Ti",
    "zirconium": "Zr",
    "cerium": "Ce",
    "lanthanum": "La",
    "molybdenum": "Mo",
    "tungsten": "W",
    "tin": "Sn",
    "stannous": "Sn",
    "stannic": "Sn",
    "lead": "Pb",
    "plumbous": "Pb",
    "plumbic": "Pb",
    "silver": "Ag",
    "gold": "Au",
    "palladium": "Pd",
    "platinum": "Pt",
    "rhodium": "Rh",
    "ruthenium": "Ru",
    "iridium": "Ir",
    "aluminum": "Al",
    "aluminium": "Al",
}

SEMANTIC_EXACT_CATALYST_FORMULAS = {
    "ferrous oxide": "FeO",
    "ferric oxide": "Fe2O3",
    "magnetite": "Fe3O4",
    "hematite": "Fe2O3",
    "wustite": "FeO",
    "nickel oxide": "NiO",
    "nickelous oxide": "NiO",
    "cobaltous oxide": "CoO",
    "cobaltic oxide": "Co2O3",
    "cuprous oxide": "Cu2O",
    "cupric oxide": "CuO",
    "manganous oxide": "MnO",
    "manganic oxide": "Mn2O3",
    "chromous oxide": "CrO",
    "chromic oxide": "Cr2O3",
    "zinc oxide": "ZnO",
    "magnesium oxide": "MgO",
    "calcium oxide": "CaO",
    "cerium oxide": "CeO2",
    "lanthanum oxide": "La2O3",
    "titanium dioxide": "TiO2",
    "zirconium oxide": "ZrO2",
    "molybdenum oxide": "MoO3",
    "tungsten oxide": "WO3",
    "stannous oxide": "SnO",
    "stannic oxide": "SnO2",
    "plumbous oxide": "PbO",
    "plumbic oxide": "PbO2",
}

SEMANTIC_AMBIGUOUS_OXIDE_DEFAULTS = {
    "iron oxide": "Fe2O3",
    "copper oxide": "CuO",
    "cobalt oxide": "Co3O4",
    "manganese oxide": "MnO2",
    "chromium oxide": "Cr2O3",
    "tin oxide": "SnO2",
    "lead oxide": "PbO",
}

SEMANTIC_AMBIGUOUS_OXIDE_OPTIONS = {
    "iron oxide": ["FeO", "Fe2O3", "Fe3O4"],
    "copper oxide": ["Cu2O", "CuO"],
    "cobalt oxide": ["CoO", "Co2O3", "Co3O4"],
    "manganese oxide": ["MnO", "Mn2O3", "MnO2", "Mn3O4"],
    "chromium oxide": ["CrO", "Cr2O3"],
    "tin oxide": ["SnO", "SnO2"],
    "lead oxide": ["PbO", "PbO2"],
}

FERRITE_FORMULAS_BY_METAL_NAME = {
    "nickel": "NiFe2O4",
    "cobalt": "CoFe2O4",
    "manganese": "MnFe2O4",
    "copper": "CuFe2O4",
    "zinc": "ZnFe2O4",
    "magnesium": "MgFe2O4",
    "calcium": "CaFe2O4",
    "cadmium": "CdFe2O4",
}

_GENERIC_CATALYST_DESCRIPTOR_PATTERNS = [
    r"\b(?:unsupported|supported|mesoporous|macroporous|microporous|nanoporous|nanostructured|nano-sized|nanosized|porous|fresh|spent|used|prepared|as-prepared|as prepared|calcined|reduced|phase-pure|phase pure|monometallic|bimetallic|trimetallic|multimetallic|commercial|synthesized|as-synthesized)\b",
    r"\b(?:catalyst|catalysts|sample|samples|material|materials|powder|powders|nanoparticle|nanoparticles|particle|particles)\b",
]


def _support_alias_present(text: Any, alias: str) -> bool:
    norm = _norm_text(text)
    alias_norm = _norm_text(alias)
    if not norm or not alias_norm:
        return False
    compact_norm = re.sub(r"[^a-z0-9]+", "", norm)
    compact_alias = re.sub(r"[^a-z0-9]+", "", alias_norm)
    if not compact_alias:
        return False
    # Short support abbreviations such as AC, CB, CF, BN, and WC must be
    # bounded so they do not match inside unrelated words.
    if len(compact_alias) <= 3:
        pattern = rf"(?<![a-z0-9]){re.escape(alias_norm)}(?![a-z0-9])"
        return bool(re.search(pattern, norm))
    return alias_norm in norm or compact_alias in compact_norm


def _canonicalize_support_name(support: Any) -> str:
    raw = _safe_str(support).strip()
    if not raw or raw == "NR":
        return "NR"

    # Common composite carbon labels used in catalyst names. These are kept
    # human-readable so a CSV reader does not need to know paper-specific codes.
    has_ac = _support_alias_present(raw, "AC") or _support_alias_present(raw, "activated carbon")
    has_cb = _support_alias_present(raw, "CB") or _support_alias_present(raw, "carbon black")
    if has_ac and has_cb:
        return "carbon black/activated carbon composite"

    # If a support appears as a monolith, keep that geometry in the support name.
    low = _norm_text(raw)
    if "monolith" in low:
        if _support_alias_present(raw, "cordierite"):
            return "cordierite monolith"
        if _support_alias_present(raw, "SiC") or _support_alias_present(raw, "silicon carbide"):
            return "silicon carbide (SiC) monolith"
        if "ceramic" in low:
            return "ceramic monolith"
        if "metal" in low or "fecral" in low or "stainless" in low:
            return "metallic monolith"
        return raw

    for name, canonical in sorted(SEMANTIC_SUPPORT_ALIASES.items(), key=lambda kv: len(re.sub(r"[^a-z0-9]+", "", _norm_text(kv[0]))), reverse=True):
        if _support_alias_present(raw, name):
            return canonical
    return raw


def _extract_si_al_ratio_descriptor(text: Any) -> Optional[str]:
    s = _safe_str(text)
    if not s:
        return None
    for m in SI_AL_RATIO_PATTERN.finditer(s):
        ratio_name = (m.group("name") or m.group("name2") or "Si/Al").strip()
        ratio_name = re.sub(r"\s+", "", ratio_name)
        if ratio_name.lower() == "sar":
            ratio_name = "Si/Al"
        elif ratio_name.lower() == "sio2/al2o3":
            ratio_name = "SiO2/Al2O3"
        else:
            ratio_name = "Si/Al"
        val = re.sub(r"\s*[-–]\s*", "-", _safe_str(m.group("val")).strip())
        if val:
            return f"{ratio_name} ratio = {val}"
    return None


def _is_zeolite_support_text(value: Any) -> bool:
    s = _safe_str(value)
    if not s or s == "NR":
        return False
    return bool(ZEOLITE_HINT_RE.search(s)) or _norm_text(_canonicalize_support_name(s)) == "zeolite"


def _normalize_support_type_field(value: Any, context_text: Any = "") -> str:
    raw = _safe_str(value).strip()
    if not raw or raw == "NR":
        return "NR"
    combined = " ".join([raw, _safe_str(context_text)])
    base = _canonicalize_support_name(raw)
    # For defined shorthand such as AC or CB, raw alone is usually enough. If
    # the raw value is a generic word/code, use the surrounding context as a
    # fallback only when it provides a recognized support identity.
    if _norm_text(base) == _norm_text(raw) and len(re.sub(r"[^A-Za-z0-9]+", "", raw)) <= 4:
        context_base = _canonicalize_support_name(combined)
        if _norm_text(context_base) != _norm_text(combined):
            base = context_base
    if _is_zeolite_support_text(combined):
        if "zeolite" not in _norm_text(base):
            base = _canonicalize_support_name(combined)
            if "zeolite" not in _norm_text(base):
                base = "zeolite"
        ratio = _extract_si_al_ratio_descriptor(combined)
        return f"{base} ({ratio})" if ratio and ratio not in base else base
    return base


def _support_value_more_informative(candidate: Any, current: Any) -> bool:
    cand = _display_norm(candidate)
    cur = _display_norm(current)
    if cand == "NR":
        return False
    if cur == "NR":
        return True
    if _is_zeolite_support_text(cand) and _is_zeolite_support_text(cur):
        return bool(_extract_si_al_ratio_descriptor(cand)) and not bool(_extract_si_al_ratio_descriptor(cur))
    if _norm_text(cand) == _norm_text(cur):
        return False
    return len(cand) > len(cur) and _norm_text(cur) in _norm_text(cand)


def _normalize_surface_area_field(value: Any) -> str:
    s = _safe_str(value).strip()
    if not s or s == "NR":
        return "NR"
    s = s.replace("m²", "m2").replace("M²", "m2").replace("m^2", "m2")
    s = s.replace("g−1", "g-1").replace("g⁻¹", "g-1")
    s = re.sub(r"\s+", " ", s).strip()
    m = SURFACE_AREA_PATTERN.search(s)
    if m:
        val = _format_number_clean(float(m.group("val")))
        return f"{val} m2/g"
    return s


def _normalize_temperature_rate_field(value: Any) -> str:
    s = _safe_str(value).strip()
    if not s or s == "NR":
        return "NR"
    s = s.replace("℃", "°C").replace("◦", "°").replace("º", "°").replace("−", "-")
    s = re.sub(r"\s+", " ", s).strip()
    m = COOLING_RATE_PATTERN.search(s)
    if m:
        val = _format_number_clean(float(m.group("val")))
        unit = m.group("unit").replace(" ", "")
        unit = unit.replace("degC", "°C").replace("degK", "K")
        if unit.upper() == "C":
            unit = "°C"
        time_unit = m.group("time").lower()
        time_unit = {"minute": "min", "hr": "h", "hour": "h", "sec": "s", "second": "s"}.get(time_unit, time_unit)
        return f"{val} {unit}/{time_unit}"
    m_num = re.fullmatch(r"(?P<approx>[~]?)\s*(?P<val>[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?)", s)
    if m_num:
        return f"{m_num.group('approx')}{_format_number_clean(float(m_num.group('val')))} °C/min"
    if _norm_text(s) in {"not reported", "not mentioned", "none", "unknown", "na", "n/a", "-"}:
        return "NR"
    return "NR"


def _normalize_cooling_rate_field(value: Any) -> str:
    return _normalize_temperature_rate_field(value)


def _normalize_temperature_ramp_rate_field(value: Any) -> str:
    return _normalize_temperature_rate_field(value)


def _extract_temperature_rate_from_text(value: Any) -> Optional[str]:
    normalized = _normalize_temperature_rate_field(value)
    return None if normalized == "NR" or normalized == _safe_str(value).strip() and not COOLING_RATE_PATTERN.search(_safe_str(value)) else normalized


def _canonical_process_gas_name(name: Any) -> str:
    raw = _safe_str(name).strip()
    if not raw:
        return ""
    compact = re.sub(r"[^A-Za-z0-9]", "", raw).lower()
    for canonical, aliases in PROCESS_GAS_HINT_GROUPS.items():
        for alias in aliases + [canonical]:
            alias_compact = re.sub(r"[^A-Za-z0-9]", "", alias).lower()
            if compact == alias_compact:
                return canonical
    return _canonical_gas_name(raw)


def _gas_alias_regex_group() -> str:
    aliases: List[str] = []
    for canonical, names in PROCESS_GAS_HINT_GROUPS.items():
        aliases.extend(names + [canonical])
    aliases = sorted(set(aliases), key=len, reverse=True)
    return "(?:" + "|".join(re.escape(a) for a in aliases) + ")"


def _find_process_gases(text: Any) -> List[str]:
    s = _safe_str(text)
    if not s:
        return []
    gases: List[str] = []
    group = _gas_alias_regex_group()
    for m in re.finditer(rf"(?<![A-Za-z0-9])(?P<gas>{group})(?![A-Za-z0-9])", s, flags=re.IGNORECASE):
        raw_gas = m.group("gas")
        gas = _canonical_process_gas_name(raw_gas)
        # Avoid treating cobalt abbreviations such as "Co" as carbon monoxide.
        if gas == "CO" and raw_gas != "CO" and "carbon" not in raw_gas.lower():
            continue
        if gas and gas not in gases:
            gases.append(gas)
    return gases


def _format_atmosphere_components(components: List[Tuple[str, Optional[float]]]) -> str:
    clean: List[Tuple[str, Optional[float]]] = []
    seen = set()
    for gas, pct in components:
        gas = _canonical_process_gas_name(gas)
        if not gas:
            continue
        key = (gas, None if pct is None else round(float(pct), 6))
        if key in seen:
            continue
        seen.add(key)
        clean.append((gas, None if pct is None else float(pct)))
    if not clean:
        return "NR"
    parts = []
    for gas, pct in clean:
        if pct is None:
            parts.append(gas)
        else:
            parts.append(f"{_format_number_clean(pct)}% {gas}")
    return " + ".join(parts)


def _normalize_single_atmosphere_part(part: Any) -> str:
    p = _safe_str(part).strip()
    if not p or p == "NR":
        return "NR"
    p = p.replace("vol.%", "%").replace("vol%", "%").replace("mole%", "%").replace("mol%", "%")
    p = re.sub(r"\b(?:under|with|using|flowing|static|a|an|the|atmosphere|gas|mixture|stream|flow|balance|balanced|diluted)\b", " ", p, flags=re.IGNORECASE)
    p = p.replace("/", " / ").replace("+", " + ")
    p = re.sub(r"\s+", " ", p).strip(" .;,()")

    gases = _find_process_gases(p)
    if not gases:
        return "NR"

    if "Vacuum" in gases:
        return "Vacuum"
    if gases == ["Air"]:
        return "Air"

    group = _gas_alias_regex_group()
    pct_components: List[Tuple[str, Optional[float]]] = []
    pct_gases = set()
    for m in re.finditer(rf"(?P<pct>[-+]?\d*\.?\d+)\s*%\s*(?P<gas>{group})", p, flags=re.IGNORECASE):
        gas = _canonical_process_gas_name(m.group("gas"))
        pct = float(m.group("pct"))
        pct_components.append((gas, pct))
        pct_gases.add(gas)
    for m in re.finditer(rf"(?P<gas>{group})\s*(?:=|:)?\s*(?P<pct>[-+]?\d*\.?\d+)\s*%", p, flags=re.IGNORECASE):
        gas = _canonical_process_gas_name(m.group("gas"))
        pct = float(m.group("pct"))
        if gas not in pct_gases:
            pct_components.append((gas, pct))
            pct_gases.add(gas)

    if pct_components:
        total_pct = sum(pct for _gas, pct in pct_components if pct is not None)
        missing_gases = [g for g in gases if g not in pct_gases and g not in {"Air", "Vacuum"}]
        if len(missing_gases) == 1 and 0 <= total_pct <= 100:
            remainder = 100.0 - total_pct
            if remainder > 0:
                pct_components.append((missing_gases[0], remainder))
        return _format_atmosphere_components(pct_components)

    if len(gases) == 1 and gases[0] not in {"Air", "Vacuum"}:
        return f"100% {gases[0]}"
    return _format_atmosphere_components([(g, None) for g in gases])


def _normalize_atmosphere_field(value: Any) -> str:
    s = _safe_str(value).strip()
    if not s or s == "NR":
        return "NR"
    sequential_hint = bool(re.search(r";|,\s*then|followed\s+by|then", s, flags=re.IGNORECASE))
    s = s.replace(";", " + ").replace(", then", " + ").replace(" followed by ", " + ")
    s = re.sub(r"\s*\+\s*", " + ", s)
    # A plus sign commonly denotes a gas mixture, e.g. 5% H2 + 95% Ar.
    # Split only for clear sequential/multistep wording.
    parts = _split_plus_parts(s) if sequential_hint and " + " in s else [s]
    norm_parts: List[str] = []
    for part in parts:
        normalized = _normalize_single_atmosphere_part(part)
        if normalized != "NR":
            norm_parts.append(normalized)
    return " + ".join(dict.fromkeys(norm_parts)) if norm_parts else "NR"


def _extract_atmosphere_from_text(value: Any) -> Optional[str]:
    s = _safe_str(value)
    if not s or s == "NR":
        return None
    normalized = _normalize_atmosphere_field(s)
    return None if normalized == "NR" else normalized


def _strip_atmosphere_from_process_value(value: Any) -> str:
    s = _safe_str(value).strip()
    if not s or s == "NR":
        return "NR"
    gas_pat = r"(?:H2|hydrogen|N2|nitrogen|Ar|argon|Air|air|O2|oxygen|He|helium|Vacuum|vacuum)"
    s = re.sub(rf"\s*\([^)]*\b{gas_pat}\b[^)]*\)", "", s, flags=re.IGNORECASE)
    s = re.sub(rf"\s+(?:under|in|with|using|flowing|static)\s+{gas_pat}\b", "", s, flags=re.IGNORECASE)
    return re.sub(r"\s+", " ", s).strip() or "NR"


def _strip_semantic_catalyst_descriptors(text: Any) -> str:
    cleaned = _norm_text(text)
    if not cleaned or cleaned == "nr":
        return ""
    for pat in _GENERIC_CATALYST_DESCRIPTOR_PATTERNS:
        cleaned = re.sub(pat, " ", cleaned, flags=re.IGNORECASE)
    cleaned = cleaned.replace("/", " /").replace("@", " @")
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" -_,;:.()[]{}")
    return cleaned


def _formula_to_loadings(formula: str) -> Dict[str, float]:
    loadings: Dict[str, float] = {}
    if not formula or formula == "NR":
        return loadings
    for el in re.findall(r"[A-Z][a-z]?", formula):
        if el == "O":
            continue
        loadings.setdefault(el, 0.0)
    return loadings


def _ordered_semantic_metal_symbols(text: Any) -> List[str]:
    low = _norm_text(text)
    hits: List[Tuple[int, str]] = []
    for name, symbol in SEMANTIC_METAL_NAME_TO_SYMBOL.items():
        for m in re.finditer(rf"\b{re.escape(name)}\b", low):
            hits.append((m.start(), symbol))
    hits.sort(key=lambda item: item[0])
    ordered: List[str] = []
    seen = set()
    for _pos, symbol in hits:
        if symbol not in seen:
            seen.add(symbol)
            ordered.append(symbol)
    return ordered


def _resolve_ambiguous_oxide_formula(label_phrase: str, context_text: Any = "") -> str:
    phrase_low = _norm_text(label_phrase)
    default_formula = SEMANTIC_AMBIGUOUS_OXIDE_DEFAULTS.get(phrase_low, "NR")
    ctx = _norm_text(context_text)
    options = SEMANTIC_AMBIGUOUS_OXIDE_OPTIONS.get(phrase_low, [])
    explicit_hits = [formula for formula in options if re.search(rf"\b{re.escape(formula.lower())}\b", ctx)]
    if len(explicit_hits) == 1:
        return explicit_hits[0]
    adjective_overrides = {
        "iron oxide": [("ferrous oxide", "FeO"), ("ferric oxide", "Fe2O3"), ("magnetite", "Fe3O4"), ("hematite", "Fe2O3"), ("wustite", "FeO")],
        "copper oxide": [("cuprous oxide", "Cu2O"), ("cupric oxide", "CuO")],
        "cobalt oxide": [("cobaltous oxide", "CoO"), ("cobaltic oxide", "Co2O3")],
        "manganese oxide": [("manganous oxide", "MnO"), ("manganic oxide", "Mn2O3")],
        "chromium oxide": [("chromous oxide", "CrO"), ("chromic oxide", "Cr2O3")],
        "tin oxide": [("stannous oxide", "SnO"), ("stannic oxide", "SnO2")],
        "lead oxide": [("plumbous oxide", "PbO"), ("plumbic oxide", "PbO2")],
    }
    for trigger, formula in adjective_overrides.get(phrase_low, []):
        if trigger in ctx:
            return formula
    return default_formula


def _semantic_catalyst_components(label: str, context_text: Any = "") -> Tuple[str, Dict[str, float], str]:
    raw = _safe_str(label).strip()
    if not raw or raw == "NR":
        return "NR", {}, "NR"

    norm = raw.replace("–", "-").replace("—", "-").replace("−", "-")
    support = "NR"
    left = norm
    if "@" in norm:
        left, support = norm.split("@", 1)
    elif "/" in norm:
        left, support = norm.split("/", 1)
    support = _canonicalize_support_name(support)

    semantic_left = _strip_semantic_catalyst_descriptors(left)

    symbol_oxide_match = re.fullmatch(r"([A-Z][a-z]?)(?:[-+ ]([A-Z][a-z]?))+\s*oxide", left.strip())
    if symbol_oxide_match:
        symbols = re.findall(r"[A-Z][a-z]?", left.split("oxide", 1)[0])
        symbols = [sym for sym in symbols if sym != "O"]
        if len(symbols) >= 2:
            ordered = []
            seen = set()
            for sym in symbols:
                if sym not in seen:
                    seen.add(sym)
                    ordered.append(sym)
            canonical = "-".join(ordered) + " oxide"
            return canonical, {sym: 0.0 for sym in ordered}, support

    symbol_match = re.fullmatch(r"([A-Z][a-z]?)(?:[-+ ]([A-Z][a-z]?))+", left.strip())
    if symbol_match:
        symbols = re.findall(r"[A-Z][a-z]?", left)
        ordered = []
        seen = set()
        for sym in symbols:
            if sym != "O" and sym not in seen:
                seen.add(sym)
                ordered.append(sym)
        if len(ordered) >= 2:
            canonical = "-".join(ordered)
            return canonical, {sym: 0.0 for sym in ordered}, support

    for phrase, formula in sorted(SEMANTIC_EXACT_CATALYST_FORMULAS.items(), key=lambda kv: len(kv[0]), reverse=True):
        if re.search(rf"\b{re.escape(phrase)}\b", semantic_left):
            return formula, _formula_to_loadings(formula), support

    for metal_name, formula in sorted(FERRITE_FORMULAS_BY_METAL_NAME.items(), key=lambda kv: len(kv[0]), reverse=True):
        if re.search(rf"\b{re.escape(metal_name)}\s+ferrit(?:e|es)\b", semantic_left):
            return formula, _formula_to_loadings(formula), support

    if "oxide" in semantic_left:
        ordered_symbols = _ordered_semantic_metal_symbols(semantic_left)
        if len(ordered_symbols) >= 2:
            canonical = "-".join(ordered_symbols) + " oxide"
            return canonical, {sym: 0.0 for sym in ordered_symbols}, support

    for phrase in sorted(SEMANTIC_AMBIGUOUS_OXIDE_DEFAULTS.keys(), key=len, reverse=True):
        if re.search(rf"\b{re.escape(phrase)}\b", semantic_left):
            formula = _resolve_ambiguous_oxide_formula(phrase, " ".join([left, _safe_str(context_text)]))
            if formula != "NR":
                return formula, _formula_to_loadings(formula), support

    if "oxide" in semantic_left:
        ordered_symbols = _ordered_semantic_metal_symbols(semantic_left)
        if len(ordered_symbols) == 1:
            sym = ordered_symbols[0]
            default_formula = {
                "Fe": "Fe2O3", "Cu": "CuO", "Co": "Co3O4", "Mn": "MnO2", "Cr": "Cr2O3",
                "Ni": "NiO", "Zn": "ZnO", "Mg": "MgO", "Ca": "CaO", "Ce": "CeO2", "La": "La2O3",
                "Ti": "TiO2", "Zr": "ZrO2", "Mo": "MoO3", "W": "WO3", "Sn": "SnO2", "Pb": "PbO",
            }.get(sym)
            if default_formula:
                return default_formula, _formula_to_loadings(default_formula), support

    ordered_symbols = _ordered_semantic_metal_symbols(semantic_left)
    if len(ordered_symbols) >= 2:
        canonical = "-".join(ordered_symbols)
        return canonical, {sym: 0.0 for sym in ordered_symbols}, support
    if len(ordered_symbols) == 1 and not _looks_like_condition_label(raw):
        sym = ordered_symbols[0]
        return sym, {sym: 0.0}, support

    return raw, {}, support


def _canonicalize_catalyst_label_with_context(label: str, context_text: Any = "") -> str:
    canonical, loadings, support = _semantic_catalyst_components(label, context_text)
    raw = _safe_str(label).strip()
    if loadings:
        if support not in ("", "NR") and "@" not in canonical and "/" not in canonical:
            return f"{canonical}/{support}"
        return canonical
    return raw if raw else "NR"


def _infer_catalyst_label_from_text(title_text: Any, full_text: Any, precursor_salts: Any = "NR") -> str:
    title = _safe_str(title_text).strip()
    if title:
        title_low = _norm_text(title)
        if not ("catalysts" in title_low and any(sep in title_low for sep in [",", " and ", ";"])):
            cand = _canonicalize_catalyst_label_with_context(title, title)
            if cand not in ("NR", "") and parse_catalyst_label(cand)[0]:
                return cand

    sentences = _sentence_tokenize(_safe_str(full_text))
    priority_sents: List[str] = []
    for sent in sentences[:80]:
        low = _norm_text(sent)
        if any(tok in low for tok in [" ferrite", " oxide", "catalyst", "catalysts", "used to catalyze", "synthesized", "prepared"]) or re.search(r"\b(?:ni|fe|co|cu|mn|zn|mg|ca)[a-z0-9]*2?o[0-9]*\b", low):
            priority_sents.append(sent)
    for sent in priority_sents:
        low = _norm_text(sent)
        if "catalysts" in low and any(sep in low for sep in [",", " and ", ";"]):
            continue
        cand = _canonicalize_catalyst_label_with_context(sent, sent)
        if cand not in ("NR", "") and parse_catalyst_label(cand)[0]:
            return cand

    corpus_low = _norm_text(" ".join([_safe_str(title_text), _safe_str(full_text)]))
    if ("catalysts" in corpus_low and any(sep in corpus_low for sep in [",", " and ", ";"])) or re.search(r"\b(?:different|various|several|multiple)\s+catalysts\b", corpus_low):
        return "NR"

    prec_norm = _normalize_precursor_salts(precursor_salts)
    elems = [el for el in _extract_relevant_elements(prec_norm) if el not in {"Al", "Si", "C"}]
    elems = list(dict.fromkeys(elems))
    text_low = _norm_text(" ".join([_safe_str(title_text), _safe_str(full_text)]))
    if len(elems) == 2 and "Fe" in elems and ("ferrite" in text_low or "spinel" in text_low):
        other = [el for el in elems if el != "Fe"][0]
        ferrite_formula_map_by_symbol = {SEMANTIC_METAL_NAME_TO_SYMBOL[name]: formula for name, formula in FERRITE_FORMULAS_BY_METAL_NAME.items() if name in SEMANTIC_METAL_NAME_TO_SYMBOL}
        formula = ferrite_formula_map_by_symbol.get(other)
        if formula:
            return formula
    if len(elems) >= 2:
        return "-".join(elems)
    if len(elems) == 1:
        return elems[0]
    return "NR"


def _split_catalyst_left_and_support(label: Any) -> Tuple[str, str]:
    raw = _safe_str(label).strip()
    if not raw or raw == "NR":
        return "", "NR"
    norm = raw.replace("–", "-").replace("—", "-").replace("−", "-")
    support = "NR"
    left = norm
    if "@" in norm:
        left, support = norm.split("@", 1)
    elif "/" in norm:
        left, support = norm.split("/", 1)
    return re.sub(r"\s+", "", left), _canonicalize_support_name(support)


def _explicit_numeric_composition_loadings(label: Any) -> Dict[str, float]:
    left, _support = _split_catalyst_left_and_support(label)
    if not left:
        return {}

    inline_matches = list(re.finditer(r"([A-Z][a-z]?)(\d+(?:\.\d+)?)", left))
    inline_reconstructed = "".join(m.group(1) + m.group(2) for m in inline_matches)
    inline_non_o = [m.group(1) for m in inline_matches if m.group(1) != "O"]
    if inline_matches and inline_reconstructed == left and len(set(inline_non_o)) >= 2 and not any(m.group(1) == "O" for m in inline_matches):
        loadings: Dict[str, float] = {}
        for m in inline_matches:
            loadings[m.group(1)] = float(m.group(2))
        return loadings

    segments = [seg for seg in re.split(r"[-+]", left) if seg]
    if len(segments) < 2:
        return {}
    loadings: Dict[str, float] = {}
    numeric_count = 0
    plain_count = 0
    for seg in segments:
        seg_clean = re.sub(r"(?i)wt%", "", seg)
        m = re.fullmatch(r"(\d+(?:\.\d+)?)([A-Z][a-z]?)", seg_clean)
        if m:
            wt, el = m.group(1), m.group(2)
            if el == "O":
                return {}
            loadings[el] = float(wt)
            numeric_count += 1
            continue
        m = re.fullmatch(r"([A-Z][a-z]?)(\d+(?:\.\d+)?)", seg_clean)
        if m:
            el, wt = m.group(1), m.group(2)
            if el == "O":
                return {}
            loadings[el] = float(wt)
            numeric_count += 1
            continue
        m = re.fullmatch(r"([A-Z][a-z]?)", seg_clean)
        if m:
            if m.group(1) == "O":
                return {}
            loadings.setdefault(m.group(1), 0.0)
            plain_count += 1
            continue
        return {}
    if loadings and numeric_count == len(loadings) and plain_count == 0 and len(loadings) >= 2:
        return loadings
    return {}


def _drop_explicit_zero_weight_metals(loadings: Dict[str, float], label: Any) -> Dict[str, float]:
    explicit = _explicit_numeric_composition_loadings(label)
    if not explicit:
        return dict(loadings)
    filtered = {el: wt for el, wt in explicit.items() if float(wt) > 0}
    return filtered if filtered else dict(loadings)


def _resolve_row_catalyst_identity(row: Dict[str, Any], context_text: Any = "") -> Tuple[str, Dict[str, float], str]:
    raw_label = _safe_str(row.get("Catalyst_Label", "NR")).strip()
    series_label = _safe_str(row.get("Series_Label", "NR")).strip()

    raw_canonical = _canonicalize_catalyst_label_with_context(raw_label, context_text) if raw_label not in ("", "NR") else "NR"
    raw_loadings, raw_support = parse_catalyst_label(raw_canonical)
    raw_loadings = _drop_explicit_zero_weight_metals(raw_loadings, raw_canonical)
    raw_explicit = bool(_explicit_numeric_composition_loadings(raw_canonical))

    series_canonical = _canonicalize_catalyst_label_with_context(series_label, context_text) if series_label not in ("", "NR") else "NR"
    series_loadings, series_support = parse_catalyst_label(series_canonical)
    series_loadings = _drop_explicit_zero_weight_metals(series_loadings, series_canonical)
    series_explicit = bool(_explicit_numeric_composition_loadings(series_canonical))

    raw_generic = raw_label in ("", "NR") or _looks_like_generic_phase_or_context_label(raw_label) or not raw_loadings
    series_specific = _looks_like_specific_series_identity_label(series_label) and bool(series_loadings)

    use_series = False
    if series_specific:
        if raw_generic:
            use_series = True
        elif series_explicit and not raw_explicit:
            use_series = True
        elif series_loadings and raw_loadings and set(series_loadings.keys()) != set(raw_loadings.keys()):
            use_series = True
        elif series_explicit and raw_explicit and _norm_text(series_canonical) != _norm_text(raw_canonical):
            use_series = True

    chosen_label = series_canonical if use_series else raw_canonical
    chosen_loadings = series_loadings if use_series else raw_loadings
    chosen_support = series_support if use_series else raw_support

    fallback_support = _canonicalize_support_name(row.get("Support_Type", "NR"))
    if chosen_support in ("", "NR"):
        if raw_support not in ("", "NR"):
            chosen_support = raw_support
        elif fallback_support not in ("", "NR"):
            chosen_support = fallback_support

    if chosen_label in ("", "NR") and series_specific:
        chosen_label = series_canonical
        chosen_loadings = series_loadings
    if chosen_label in ("", "NR"):
        chosen_label = raw_label if raw_label else "NR"

    return chosen_label, chosen_loadings, chosen_support


def _split_plus_parts(value: Any) -> List[str]:
    s = re.sub(r"\s+", " ", _safe_str(value)).strip()
    if not s or s == "NR":
        return []
    return [part.strip() for part in re.split(r"\s+\+\s+", s) if part.strip()]


def _extract_formula_like_elements(text: Any) -> set[str]:
    raw = _safe_str(text)
    elems: set[str] = set()
    for token in re.findall(r"\b[A-Z][A-Za-z0-9()·⋅\.\-]*\b", raw):
        if not (re.fullmatch(r"[A-Z][a-z]?", token) or re.search(r"[\d()·⋅]", token) or token.isupper()):
            continue
        for sym in re.findall(r"[A-Z][a-z]?", token):
            if sym not in IGNORED_PRECURSOR_MATCH_ELEMENTS:
                elems.add(sym)
    return elems


def _extract_named_elements(text: Any) -> set[str]:
    low = _norm_text(text)
    elems: set[str] = set()
    for needle, symbol in ELEMENT_NAME_HINTS.items():
        if needle in low:
            elems.add(symbol)
    return elems


def _extract_relevant_elements(text: Any) -> set[str]:
    elems = _extract_formula_like_elements(text) | _extract_named_elements(text)
    elems.discard("C")
    return elems


def _row_relevant_elements(row: Dict[str, Any]) -> set[str]:
    elems: set[str] = set()
    context_text = _safe_str(row.get("_Caption_Context", ""))
    resolved_label, loadings, support = _resolve_row_catalyst_identity(row, context_text)
    for el in loadings.keys():
        if el not in IGNORED_PRECURSOR_MATCH_ELEMENTS and el != "C":
            elems.add(el)
    for fld in ["Active_Metal", "Promoter_Type", "Support_Type"]:
        elems.update(_extract_relevant_elements(row.get(fld, "")))
    if support not in ("", "NR"):
        elems.update(_extract_relevant_elements(support))
    return elems


def _filter_precursor_salts_for_row(value: Any, row: Dict[str, Any]) -> str:
    normalized = _normalize_precursor_salts(value)
    parts = _split_plus_parts(normalized)
    if normalized == "NR" or len(parts) <= 1:
        return normalized
    row_elements = _row_relevant_elements(row)
    if not row_elements:
        return normalized
    matched: List[str] = []
    seen = set()
    for part in parts:
        elems = _extract_relevant_elements(part)
        if elems and elems.intersection(row_elements):
            key = _norm_text(part)
            if key not in seen:
                seen.add(key)
                matched.append(part)
    return " + ".join(matched) if matched else normalized


def _process_evidence_text(row: Dict[str, Any], context_text: str) -> str:
    return ". ".join([
        _safe_str(row.get("Catalyst_Label", "")),
        _safe_str(row.get("Series_Label", "")),
        _safe_str(context_text),
    ])


def _process_relevant_sentences(text: Any, process_name: str) -> List[str]:
    keywords = PROCESS_SENTENCE_KEYWORDS.get(process_name, [])
    if not keywords:
        return []
    hits: List[str] = []
    for sent in _sentence_tokenize(_safe_str(text)):
        low = _norm_text(sent)
        if any(keyword in low for keyword in keywords):
            hits.append(sent)
    return hits


def _extract_process_temperature_hints(text: Any, process_name: str) -> set[str]:
    hints: set[str] = set()
    for sent in _process_relevant_sentences(text, process_name):
        for temp in _extract_temperature_tokens(sent):
            hints.add(temp)
    return hints


def _extract_process_time_hints(text: Any, process_name: str) -> List[float]:
    hints: List[float] = []
    seen = set()
    for sent in _process_relevant_sentences(text, process_name):
        for match in TIME_HINT_PATTERN.finditer(sent):
            minutes = _parse_time_to_minutes(match.group(0))
            if minutes is None:
                continue
            key = round(minutes, 6)
            if key not in seen:
                seen.add(key)
                hints.append(minutes)
    return hints


def _extract_process_gas_hints(text: Any, process_name: str) -> set[str]:
    sentences = _process_relevant_sentences(text, process_name)
    if not sentences:
        return set()
    evidence = " ".join(sentences)
    low = _norm_text(evidence)
    gases: set[str] = set()
    for canonical, aliases in PROCESS_GAS_HINT_GROUPS.items():
        if any(alias in low for alias in aliases):
            gases.add(canonical)
    return gases


def _segment_matches_gases(segment: Any, gas_hints: set[str]) -> bool:
    low = _norm_text(segment)
    for canonical in gas_hints:
        aliases = PROCESS_GAS_HINT_GROUPS.get(canonical, [canonical.lower()])
        if any(alias in low for alias in aliases):
            return True
    return False


def _time_matches_hints(time_value: Any, time_hints: List[float]) -> bool:
    if not time_hints:
        return False
    value = _parse_time_to_minutes(time_value)
    if value is None:
        return False
    return any(abs(value - hint) <= max(0.25, 0.05 * max(abs(hint), 1.0)) for hint in time_hints)


def _select_relevant_step_indices(
    row: Dict[str, Any],
    context_text: str,
    process_name: str,
    temperature_value: Any,
    time_value: Any,
    atmosphere_value: Any = "NR",
) -> List[int]:
    temp_norm = _normalize_multistep_temperature_field(temperature_value)
    time_norm = _normalize_multistep_time_field(time_value)
    atm_norm = _normalize_atmosphere_field(atmosphere_value)
    temp_parts = _split_plus_parts(temp_norm)
    time_parts = _split_plus_parts(time_norm)
    atm_parts = _split_plus_parts(atm_norm)
    n_parts = max(len(temp_parts), len(time_parts), len(atm_parts))
    if n_parts <= 1:
        return []

    evidence = _process_evidence_text(row, context_text)
    label_evidence = ". ".join([
        _safe_str(row.get("Series_Label", "")),
        _safe_str(row.get("Catalyst_Label", "")),
    ])
    temp_hints = _extract_process_temperature_hints(label_evidence, process_name) or _extract_process_temperature_hints(evidence, process_name)
    time_hints = _extract_process_time_hints(label_evidence, process_name) or _extract_process_time_hints(evidence, process_name)
    gas_hints = _extract_process_gas_hints(label_evidence, process_name) or _extract_process_gas_hints(evidence, process_name)

    scores: List[int] = []
    for idx in range(n_parts):
        score = 0
        temp_part = temp_parts[idx] if idx < len(temp_parts) else ""
        time_part = time_parts[idx] if idx < len(time_parts) else ""
        atm_part = atm_parts[idx] if idx < len(atm_parts) else ""
        if temp_hints and set(_extract_temperature_tokens(temp_part)).intersection(temp_hints):
            score += 2
        if _time_matches_hints(time_part, time_hints):
            score += 2
        if gas_hints and (
            _segment_matches_gases(temp_part, gas_hints)
            or _segment_matches_gases(time_part, gas_hints)
            or _segment_matches_gases(atm_part, gas_hints)
        ):
            score += 1
        scores.append(score)

    max_score = max(scores) if scores else 0
    if max_score <= 0:
        return []
    return [idx for idx, score in enumerate(scores) if score == max_score and score > 0]


def _select_parts_by_index(parts: List[str], keep_indices: List[int]) -> str:
    if not parts:
        return "NR"
    if len(parts) <= 1 or not keep_indices:
        return " + ".join(parts)
    selected: List[str] = []
    seen = set()
    for idx in keep_indices:
        if 0 <= idx < len(parts):
            part = parts[idx]
            key = _norm_text(part)
            if key not in seen:
                seen.add(key)
                selected.append(part)
    return " + ".join(selected) if selected else " + ".join(parts)


def _filter_process_values_for_row(
    row: Dict[str, Any],
    context_text: str,
    process_name: str,
    temperature_value: Any,
    time_value: Any,
    atmosphere_value: Any = "NR",
) -> Tuple[str, str, str]:
    temp_norm = _normalize_multistep_temperature_field(_strip_atmosphere_from_process_value(temperature_value))
    time_norm = _normalize_multistep_time_field(_strip_atmosphere_from_process_value(time_value))
    atm_norm = _normalize_atmosphere_field(atmosphere_value)
    if atm_norm == "NR":
        inferred = _extract_atmosphere_from_text(" ".join([_safe_str(temperature_value), _safe_str(time_value)]))
        if inferred:
            atm_norm = _normalize_atmosphere_field(inferred)
    keep_indices = _select_relevant_step_indices(row, context_text, process_name, temp_norm, time_norm, atm_norm)
    if not keep_indices:
        return temp_norm, time_norm, atm_norm
    temp_parts = _split_plus_parts(temp_norm)
    time_parts = _split_plus_parts(time_norm)
    atm_parts = _split_plus_parts(atm_norm)
    temp_out = _select_parts_by_index(temp_parts, keep_indices) if len(temp_parts) > 1 else temp_norm
    time_out = _select_parts_by_index(time_parts, keep_indices) if len(time_parts) > 1 else time_norm
    atm_out = _select_parts_by_index(atm_parts, keep_indices) if len(atm_parts) > 1 else atm_norm
    return temp_out, time_out, atm_out


def _filter_process_pair_values_for_row(
    row: Dict[str, Any],
    context_text: str,
    process_name: str,
    temperature_value: Any,
    time_value: Any,
) -> Tuple[str, str]:
    temp_out, time_out, _atm_out = _filter_process_values_for_row(
        row, context_text, process_name, temperature_value, time_value, "NR"
    )
    return temp_out, time_out


def _filter_catalyst_profile_for_row(profile: Dict[str, Any], row: Dict[str, Any], context_text: str) -> Dict[str, Any]:
    filtered = _normalize_catalyst_synthesis_profile(profile or {})
    filtered["Precursor_Salts"] = _filter_precursor_salts_for_row(filtered.get("Precursor_Salts", "NR"), row)
    for process_name, temp_field, time_field, atm_field in [
        ("Drying", "Drying_Temperature_C_or_K", "Drying_Time_hr_min_s", None),
        ("Calcination", "Calcination_Temperature_C_or_K", "Calcination_Time_hr_min_s", "Calcination_Atmosphere"),
        ("Reduction", "Reduction_Temperature_C_or_K", "Reduction_Time_hr_min_s", "Reduction_Atmosphere"),
    ]:
        temp_val, time_val, atm_val = _filter_process_values_for_row(
            row,
            context_text,
            process_name,
            filtered.get(temp_field, "NR"),
            filtered.get(time_field, "NR"),
            filtered.get(atm_field, "NR") if atm_field else "NR",
        )
        filtered[temp_field] = temp_val
        filtered[time_field] = time_val
        if atm_field:
            filtered[atm_field] = atm_val
    return _normalize_catalyst_synthesis_profile(filtered)


def _left_catalyst_token(label: Any) -> str:
    norm = _safe_str(label).replace("–", "-").replace("—", "-").replace("−", "-")
    norm = re.sub(r"\s+", "", norm)
    left = norm.split("@", 1)[0]
    left = left.split("/", 1)[0]
    return left


def _is_oxide_catalyst_label(label: Any) -> bool:
    canonical = _canonicalize_catalyst_label_with_context(_safe_str(label), _safe_str(label))
    left = _left_catalyst_token(canonical)
    if not left:
        return False
    if bool(OXIDE_LEFT_PATTERN.search(left)):
        return True
    if re.search(r"O\d*$", left):
        return True
    low = _norm_text(canonical)
    return any(tok in low for tok in [" oxide", "ferrite", " spinel"])

def _count_metal_elements_for_class(row: Dict[str, Any], label: Any) -> int:
    metals: set[str] = set()
    active = _safe_str(row.get("Active_Metal", ""))
    for tok in re.findall(r"\b([A-Z][a-z]?)\b", active):
        if tok != "C":
            metals.add(tok)
    promoter = _safe_str(row.get("Promoter_Type", ""))
    for tok in re.findall(r"\b([A-Z][a-z]?)\b", promoter):
        if tok != "None":
            metals.add(tok)
    if not metals:
        canonical = _canonicalize_catalyst_label_with_context(_safe_str(label), _safe_str(label))
        loadings, _support = parse_catalyst_label(canonical)
        loadings = _drop_explicit_zero_weight_metals(loadings, canonical)
        for tok in loadings.keys():
            if tok != "C":
                metals.add(tok)
    return len(metals)


def infer_catalyst_class(row: Dict[str, Any], context_text: str = "") -> str:
    label = _safe_str(row.get("Catalyst_Label", "NR")).strip()
    if _is_simple_carbon_catalyst_label(label):
        return "carbon material"
    if _is_foam_variant_label(label):
        return "metallic"

    oxide = _is_oxide_catalyst_label(label)
    metal_count = _count_metal_elements_for_class(row, label)

    if oxide:
        if metal_count <= 1:
            return "metal oxide"
        if metal_count == 2:
            return "bimetallic oxide"
        if metal_count == 3:
            return "ternary oxide"
        return "multimetal oxide"

    if metal_count <= 0:
        txt = _norm_text(" ".join([label, context_text, _safe_str(row.get("Support_Type", "NR"))]))
        if "foam" in txt:
            return "metallic"
        if "biochar" in txt or txt in {"ac", "mbc"}:
            return "carbon material"
        return "NR"
    if metal_count == 1:
        return "metallic"
    if metal_count == 2:
        return "bimetallic"
    if metal_count == 3:
        return "ternary alloy"
    return "multimetallic alloy"


def _is_direct_use_catalyst(row: Dict[str, Any], context_text: str = "") -> bool:
    label = _safe_str(row.get("Catalyst_Label", "NR")).strip()
    cclass = _norm_text(row.get("Catalyst_Class", "NR"))
    txt = _norm_text(" ".join([label, _safe_str(row.get("Series_Label", "")), context_text, _safe_str(row.get("Support_Type", ""))]))
    if _is_simple_carbon_catalyst_label(label):
        return True
    if _is_foam_variant_label(label):
        return True
    if cclass in {"carbon material", "metallic"} and ("nickel foam" in txt or "biochar" in txt or _is_simple_carbon_catalyst_label(label) or _is_foam_variant_label(label)):
        return True
    return False

def _looks_like_journal_line(s: str) -> bool:
    t = _norm_text(s)
    return any(tok in t for tok in [
        "journal", "volume", "issue", "issn", "available online",
        "received", "accepted", "published by", "copyright",
    ])


def _choose_better_title(old_val: str, new_val: str) -> str:
    candidates = [x for x in [_display_norm(old_val), _display_norm(new_val)] if x != "NR"]
    if not candidates:
        return "NR"
    if len(candidates) == 1:
        return candidates[0]

    def score(x: str) -> Tuple[int, int]:
        t = _norm_text(x)
        n_words = len(t.split())
        s = 0
        if _looks_like_journal_line(x):
            s -= 5
        if 4 <= n_words <= 30:
            s += 2
        if any(tok in t for tok in ["methane", "catalyst", "decomposition", "hydrogen"]):
            s += 2
        if len(x) > 20:
            s += 1
        return s, len(x)

    return sorted(candidates, key=score, reverse=True)[0]


def safe_paper_slug(paper: str) -> str:
    s = _safe_str(paper).strip()
    s = re.sub(r"[^\w\-.]+", "_", s)
    return s or "paper"


def _sentence_tokenize(text: str) -> List[str]:
    s = re.sub(r"\s+", " ", _safe_str(text)).strip()
    if not s:
        return []

    protected = s
    abbreviation_patterns = [
        r"\bfigs?\.",
        r"\beqs?\.",
        r"\brefs?\.",
        r"\bet al\.",
        r"\be\.g\.",
        r"\bi\.e\.",
        r"\bvs\.",
    ]
    for pat in abbreviation_patterns:
        protected = re.sub(pat, lambda m: m.group(0).replace(".", "<DOT>"), protected, flags=re.IGNORECASE)

    parts = re.split(r"(?<=[\.\?\!])\s+", protected)
    out: List[str] = []
    for p in parts:
        p = p.replace("<DOT>", ".").strip()
        if p:
            out.append(p)
    return out


def _strip_code_fences(s: str) -> str:
    s = re.sub(r"```(?:json)?", "", s, flags=re.IGNORECASE).strip()
    s = s.rstrip("`").strip()
    return s


def _salvage_json_array_items(text: str) -> List[Dict[str, Any]]:
    """Recover complete dict objects from a malformed/truncated JSON array.

    Older extraction code tried to parse the whole array first. If the model hit
    max output tokens before the closing bracket, json.loads failed and the
    fallback returned only the first complete {object}. This function walks the
    array and returns every complete leading object instead.
    """
    cleaned = _strip_code_fences(text)
    start = cleaned.find("[")
    if start < 0:
        return []
    dec = json.JSONDecoder()
    i = start + 1
    out: List[Dict[str, Any]] = []
    n = len(cleaned)
    while i < n:
        while i < n and cleaned[i] in " \t\r\n,":
            i += 1
        if i >= n or cleaned[i] == "]":
            break
        if cleaned[i] != "{":
            i += 1
            continue
        try:
            obj, end = dec.raw_decode(cleaned[i:])
        except json.JSONDecodeError:
            break
        if isinstance(obj, dict):
            out.append(obj)
        i += end
    return out


def _extract_first_json_value(text: str) -> Any:
    cleaned = _strip_code_fences(text)
    dec = json.JSONDecoder()

    # Prefer a complete JSON value when available.
    for i, ch in enumerate(cleaned):
        if ch in "[{":
            try:
                obj, _ = dec.raw_decode(cleaned[i:])
                return obj
            except json.JSONDecodeError:
                # If the response starts an array but is truncated, salvage all
                # complete row objects instead of falling through to the first row.
                if ch == "[":
                    salvaged = _salvage_json_array_items(cleaned[i:])
                    if salvaged:
                        print(f"      [salvage] recovered {len(salvaged)} complete JSON row object(s) from a malformed/truncated array")
                        return salvaged
                continue
    # Last-resort salvage if prose preceded a truncated array.
    salvaged = _salvage_json_array_items(cleaned)
    if salvaged:
        print(f"      [salvage] recovered {len(salvaged)} complete JSON row object(s) from a malformed/truncated array")
        return salvaged
    raise ValueError("No JSON value found in model response.")

def _to_list_of_dicts(obj: Any) -> List[Dict[str, Any]]:
    if obj is None:
        return []
    if isinstance(obj, list):
        return [x for x in obj if isinstance(x, dict)]
    if isinstance(obj, dict):
        if "records" in obj and isinstance(obj["records"], list):
            return [x for x in obj["records"] if isinstance(x, dict)]
        return [obj]
    return []



def _propagate_common_fields(records: List[Dict[str, Any]], keys: List[str]) -> List[Dict[str, Any]]:
    """Fill NR values from values common to the same panel/axis group only.

    Earlier versions propagated the first non-NR value across all records from a
    figure. That is unsafe when one image contains multiple CH4-conversion panels
    with different x-axes, legends, temperatures, supports, or metal loadings.
    """
    if not records:
        return records

    axis_pairs = {
        (_display_norm(r.get("X_Axis_Label", "NR")), _display_norm(r.get("Y_Axis_Label", "NR")))
        for r in records
    }
    if len(axis_pairs) > 1:
        grouped: Dict[Tuple[str, str], List[Dict[str, Any]]] = {}
        order: List[Tuple[str, str]] = []
        for r in records:
            key = (_display_norm(r.get("X_Axis_Label", "NR")), _display_norm(r.get("Y_Axis_Label", "NR")))
            if key not in grouped:
                grouped[key] = []
                order.append(key)
            grouped[key].append(r)
        out: List[Dict[str, Any]] = []
        for key in order:
            out.extend(_propagate_common_fields(grouped[key], keys))
        return out

    common: Dict[str, Any] = {}
    for k in keys:
        vals: List[Any] = []
        for r in records:
            v = r.get(k, "NR")
            if v not in (None, "", "NR"):
                vals.append(v)
        unique_norms = {_norm_text(v) for v in vals}
        if len(vals) == 1 or len(unique_norms) == 1:
            common[k] = vals[0]
    if not common:
        return records
    for r in records:
        for k, v in common.items():
            if r.get(k, "NR") in (None, "", "NR"):
                r[k] = v
    return records

def natural_key(s: str) -> List[Any]:
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", s)]


def get_image_paths(images_dir: str) -> List[str]:
    if not os.path.isdir(images_dir):
        return []
    supported = {".png", ".jpg", ".jpeg", ".gif", ".webp"}
    paths = [
        os.path.join(images_dir, fn)
        for fn in os.listdir(images_dir)
        if Path(fn).suffix.lower() in supported
    ]
    return sorted(paths, key=lambda p: natural_key(os.path.basename(p)))


def read_paper_mapping(mapping_xlsx: str) -> List[Dict[str, str]]:
    if not os.path.exists(mapping_xlsx):
        raise FileNotFoundError(f"Mapping workbook not found: {mapping_xlsx}")
    df = pd.read_excel(mapping_xlsx)
    expected = {"Paper", "DOIs"}
    missing = expected.difference(df.columns)
    if missing:
        raise ValueError(f"Mapping workbook must contain columns {sorted(expected)}. Missing: {sorted(missing)}")
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


def _read_text_file(path: Path) -> str:
    if not path.exists():
        return ""
    return path.read_text(encoding="utf-8", errors="ignore")



def _first_existing_file(base_dir: Path, names: List[str]) -> Optional[Path]:
    for name in names:
        path = base_dir / name
        if path.exists():
            return path
    return None


def _read_first_text_file(base_dir: Path, names: List[str]) -> str:
    path = _first_existing_file(base_dir, names)
    return _read_text_file(path).strip() if path else ""


def _parse_int_like(value: Any, fallback: int) -> int:
    raw = _safe_str(value).strip()
    if not raw:
        return fallback
    match = re.search(r"\d+", raw)
    if not match:
        return fallback
    try:
        return int(match.group(0))
    except Exception:
        return fallback


def _html_unescape(value: Any) -> str:
    try:
        from html import unescape
        return unescape(_safe_str(value))
    except Exception:
        return _safe_str(value)


def _first_html_file(paper_dir: Path) -> Optional[Path]:
    preferred = ["article.html", "article.htm", "fulltext.html", "fulltext.htm", "paper.html", "paper.htm", "index.html", "index.htm"]
    explicit = _first_existing_file(paper_dir, preferred)
    if explicit:
        return explicit
    htmls = sorted(list(paper_dir.glob("*.html")) + list(paper_dir.glob("*.htm")), key=lambda p: natural_key(p.name))
    return htmls[0] if htmls else None


def _rsc_figure_number_from_image_name(path_or_name: Any) -> Optional[int]:
    name = os.path.basename(_safe_str(path_or_name))
    stem = Path(name).stem
    # RSC saved HTML assets are commonly named like <article-id>-f1.gif.
    for pat in [r"(?:^|[-_])f(\d+)(?:$|[-_])", r"(?:^|[-_])fig(?:ure)?[-_ ]?(\d+)(?:$|[-_])", r"(?:^|[-_])ga(\d+)(?:$|[-_])"]:
        m = re.search(pat, stem, flags=re.IGNORECASE)
        if m:
            try:
                return int(m.group(1))
            except Exception:
                return None
    return None


def get_rsc_html_image_paths(paper_html_dir: str) -> List[str]:
    """Find RSC figure images saved by browsers under HTMLs/<Paper>/<Paper>_files/.

    RSC article exports commonly place images in a sibling assets folder such as
    `Paper 1_files/`, with filenames ending in `-f1.gif`, `-f2.gif`, etc. This
    helper returns those images in figure-number order. It intentionally filters
    to names that look like figure assets so logos/icons are not sent to the LLM.
    """
    root = Path(paper_html_dir)
    if not root.exists():
        return []
    supported = {".png", ".jpg", ".jpeg", ".gif", ".webp"}
    search_dirs: List[Path] = []
    expected_files_dir = root / f"{root.name}_files"
    if expected_files_dir.is_dir():
        search_dirs.append(expected_files_dir)
    search_dirs.extend([p for p in sorted(root.glob("*_files"), key=lambda p: natural_key(p.name)) if p.is_dir() and p not in search_dirs])
    # Some extractors store the gifs directly in HTMLs/<Paper>/.
    search_dirs.append(root)

    paths: List[Path] = []
    seen = set()
    for d in search_dirs:
        for p in d.iterdir() if d.exists() else []:
            if not p.is_file() or p.suffix.lower() not in supported:
                continue
            fig_num = _rsc_figure_number_from_image_name(p.name)
            if fig_num is None:
                continue
            key = str(p.resolve())
            if key not in seen:
                seen.add(key)
                paths.append(p)

    def sort_key(p: Path) -> List[Any]:
        n = _rsc_figure_number_from_image_name(p.name)
        return [n if n is not None else 10**9] + natural_key(p.name)

    return [str(p) for p in sorted(paths, key=sort_key)]


def _normalize_rsc_figure_label(raw_label: Any, fallback_num: int) -> Tuple[str, int]:
    raw = re.sub(r"\s+", " ", _safe_str(raw_label)).strip()
    num = _parse_int_like(raw, fallback_num)
    if raw:
        m = re.search(r"\bfig(?:ure)?\.?\s*(\d+)\b", raw, flags=re.IGNORECASE)
        if m:
            num = int(m.group(1))
            return f"Fig. {num}", num
        if re.fullmatch(r"\d+(?:\.0+)?", raw):
            return f"Fig. {num}", num
        return raw, num
    return f"Fig. {num}", num


def _load_rsc_caption_rows_from_csv(paper_dir: Path) -> List[Dict[str, Any]]:
    csv_path = _first_existing_file(paper_dir, ["figure_captions.csv", "html_figure_captions.csv", "figures.csv", "rsc_figure_captions.csv"])
    if not csv_path:
        return []
    try:
        df = pd.read_csv(csv_path).fillna("")
    except Exception as e:
        if isinstance(e, RateLimitReachedError):
            raise
        print(f"  [warning] Could not read RSC figure captions for {paper_dir.name}: {e}")
        return []
    rows: List[Dict[str, Any]] = []
    for _, rec in df.iterrows():
        raw_num = (_safe_str(rec.get("figure_number", "")).strip() or _safe_str(rec.get("object_number", "")).strip() or _safe_str(rec.get("order", "")).strip() or _safe_str(rec.get("sort_index", "")).strip())
        fig_num = _parse_int_like(raw_num, len(rows) + 1)
        label_raw = (_safe_str(rec.get("label", "")).strip() or _safe_str(rec.get("raw_label", "")).strip() or raw_num)
        label, fig_num = _normalize_rsc_figure_label(label_raw, fig_num)
        caption = (_safe_str(rec.get("caption", "")).strip() or _safe_str(rec.get("title", "")).strip() or _safe_str(rec.get("raw_text", "")).strip())
        full_caption = _safe_str(rec.get("full_caption", "")).strip()
        if not full_caption:
            full_caption = f"{label}. {caption}".strip() if caption else label
        filename = _safe_str(rec.get("filename", rec.get("file", rec.get("image", "")))).strip()
        if not filename:
            # Let filename matching work for RSC asset names when possible.
            filename = f"f{fig_num}.gif"
        rows.append({
            "label": label,
            "caption": caption or full_caption,
            "order": fig_num,
            "object_number": fig_num,
            "object_type": "figure",
            "filename": filename,
            "raw_label": label_raw or label,
            "graphic_href": _safe_str(rec.get("graphic_href", rec.get("src", ""))).strip(),
            "fig_id": _safe_str(rec.get("fig_id", rec.get("id", ""))).strip(),
            "full_caption": full_caption,
        })
    rows.sort(key=lambda row: (int(row.get("object_number", row.get("order", 0)) or 0), _safe_str(row.get("label", ""))))
    return rows


def _strip_html_with_regex(html_text: str) -> Tuple[str, List[str], str, List[Dict[str, Any]], str]:
    clean_html = re.sub(r"(?is)<(script|style|noscript).*?>.*?</\1>", " ", html_text)
    title = ""
    m_title = re.search(r"(?is)<title[^>]*>(.*?)</title>", clean_html)
    if m_title:
        title = re.sub(r"\s+", " ", _html_unescape(re.sub(r"<[^>]+>", " ", m_title.group(1)))).strip()
    paragraph_bits = re.findall(r"(?is)<(?:p|h1|h2|h3|h4|li|td|th)[^>]*>(.*?)</(?:p|h1|h2|h3|h4|li|td|th)>", clean_html)
    paragraphs: List[str] = []
    for bit in paragraph_bits:
        txt = re.sub(r"\s+", " ", _html_unescape(re.sub(r"<[^>]+>", " ", bit))).strip()
        if txt:
            paragraphs.append(txt)
    if not paragraphs:
        txt = re.sub(r"\s+", " ", _html_unescape(re.sub(r"<[^>]+>", " ", clean_html))).strip()
        paragraphs = _sentence_tokenize(txt) if txt else []
    full_text = "\n".join(paragraphs)
    doi = "NR"
    m_doi = re.search(r"\b10\.\d{4,9}/[-._;()/:A-Za-z0-9]+", full_text)
    if m_doi:
        doi = m_doi.group(0).rstrip(".,;) ")
    figures: List[Dict[str, Any]] = []
    # Captions from text. Conservative: start at Fig. N and stop at next Fig. M.
    for m in re.finditer(r"\bFig(?:ure)?\.?\s*(\d+)\b(.{0,1400}?)(?=\bFig(?:ure)?\.?\s*\d+\b|$)", full_text, flags=re.I | re.S):
        fig_num = int(m.group(1))
        caption = re.sub(r"\s+", " ", m.group(0)).strip()
        if len(caption) < 8:
            continue
        figures.append({"label": f"Fig. {fig_num}", "caption": caption, "order": fig_num, "object_number": fig_num, "object_type": "figure", "filename": f"f{fig_num}.gif", "raw_label": f"Fig. {fig_num}", "graphic_href": "", "fig_id": "", "full_caption": caption})
    return title, paragraphs, full_text, figures, doi


def _parse_rsc_html_text_and_figures(html_text: str) -> Tuple[str, List[str], str, List[Dict[str, Any]], str]:
    try:
        from bs4 import BeautifulSoup  # type: ignore
        soup = BeautifulSoup(html_text, "html.parser")
        for tag in soup(["script", "style", "noscript"]):
            tag.decompose()
        title = ""
        if soup.title and soup.title.get_text(" ", strip=True):
            title = re.sub(r"\s+", " ", soup.title.get_text(" ", strip=True)).strip()
        doi = "NR"
        for attrs in [{"name": re.compile(r"^(citation_doi|dc\.identifier|doi)$", re.I)}, {"property": re.compile(r"doi", re.I)}]:
            tag = soup.find("meta", attrs=attrs)
            if tag and tag.get("content"):
                doi = re.sub(r"^doi:\s*", "", _safe_str(tag.get("content")).strip(), flags=re.I)
                break
        article_node = soup.find("article") or soup.find(id=re.compile(r"article|content|main", re.I)) or soup.body or soup
        paragraphs: List[str] = []
        for node in article_node.find_all(["h1", "h2", "h3", "h4", "p", "li", "td", "th"]):
            txt = re.sub(r"\s+", " ", node.get_text(" ", strip=True)).strip()
            if txt:
                paragraphs.append(txt)
        if title and (not paragraphs or _norm_text(paragraphs[0]) != _norm_text(title)):
            paragraphs.insert(0, title)
        full_text = "\n".join(paragraphs)
        if doi == "NR":
            m_doi = re.search(r"\b10\.\d{4,9}/[-._;()/:A-Za-z0-9]+", full_text)
            if m_doi:
                doi = m_doi.group(0).rstrip(".,;) ")
        figures: List[Dict[str, Any]] = []
        seen_nums = set()
        # Prefer local figure/caption containers.
        candidate_nodes = soup.find_all(["figure", "div", "section", "li", "p"])
        for node in candidate_nodes:
            node_text = re.sub(r"\s+", " ", node.get_text(" ", strip=True)).strip()
            node_id = _safe_str(node.get("id", ""))
            cls = " ".join([_safe_str(c) for c in (node.get("class") or [])])
            m_label = re.search(r"\bFig(?:ure)?\.?\s*(\d+)\b", node_text, flags=re.I)
            m_id = re.search(r"fig(?:ure)?[-_ ]?(\d+)", node_id, flags=re.I)
            if not (m_label or m_id or re.search(r"caption|figure", cls, flags=re.I)):
                continue
            fig_num = int((m_label or m_id).group(1)) if (m_label or m_id) else None
            if fig_num is None:
                src = " ".join([_safe_str(img.get("src", "")) for img in node.find_all("img")])
                fig_num = _rsc_figure_number_from_image_name(src) if src else None
            if fig_num is None or fig_num in seen_nums:
                continue
            if len(node_text) > 1800:
                mm = re.search(rf"\bFig(?:ure)?\.?\s*{fig_num}\b(.{{0,1500}})", node_text, flags=re.I | re.S)
                node_text = re.sub(r"\s+", " ", (mm.group(0) if mm else node_text[:1500])).strip()
            if len(node_text) < 8:
                continue
            seen_nums.add(fig_num)
            img = node.find("img")
            src = _safe_str(img.get("src", "")).strip() if img else ""
            figures.append({"label": f"Fig. {fig_num}", "caption": node_text, "order": fig_num, "object_number": fig_num, "object_type": "figure", "filename": os.path.basename(src) or f"f{fig_num}.gif", "raw_label": f"Fig. {fig_num}", "graphic_href": src, "fig_id": node_id, "full_caption": node_text})
        figures.sort(key=lambda row: (int(row.get("object_number", row.get("order", 0)) or 0), _safe_str(row.get("label", ""))))
        return title, paragraphs, full_text, figures, doi
    except Exception:
        return _strip_html_with_regex(html_text)


def _load_rsc_caption_rows(paper_dir: Path) -> List[Dict[str, Any]]:
    rows = _load_rsc_caption_rows_from_csv(paper_dir)
    if rows:
        return rows
    html_path = _first_html_file(paper_dir)
    if not html_path:
        return []
    _title, _paras, _full, figures, _doi = _parse_rsc_html_text_and_figures(_read_text_file(html_path))
    return figures


_load_elsevier_caption_rows = _load_rsc_caption_rows
_load_wiley_caption_rows = _load_rsc_caption_rows


def load_parsed_xml(parsed_dir: str, paper_name: str) -> Dict[str, Any]:
    """Load RSC HTML-derived artifacts into the normalized internal article dict."""
    paper_dir = Path(parsed_dir) / paper_name
    if not paper_dir.exists():
        raise FileNotFoundError(f"RSC HTML artifact folder not found for {paper_name}: {paper_dir}")

    full_text = _read_first_text_file(paper_dir, ["fulltext.txt", "full_text.txt", "html_article_text.txt", "article_text.txt", "rsc_article_text.txt"])
    title = paper_name
    sentences: List[str] = []
    doi = "NR"

    sentences_path = _first_existing_file(paper_dir, ["sentences.txt", "article_sentences.txt", "html_sentences.txt"])
    if sentences_path:
        raw_lines = [re.sub(r"\s+", " ", line).strip() for line in _read_text_file(sentences_path).splitlines()]
        raw_lines = [line for line in raw_lines if line]
        if raw_lines:
            title = raw_lines[0]
            sentences = raw_lines[1:]

    tables_text = _read_first_text_file(paper_dir, ["tables_text.txt", "table_text.txt", "html_tables_text.txt", "tables.txt"])
    if tables_text and tables_text not in full_text:
        full_text = (full_text + "\n\n" + tables_text).strip()

    figures = _load_rsc_caption_rows_from_csv(paper_dir)
    html_path = _first_html_file(paper_dir)
    if html_path:
        html_title, html_paragraphs, html_full_text, html_figures, html_doi = _parse_rsc_html_text_and_figures(_read_text_file(html_path))
        if _display_norm(html_doi) != "NR":
            doi = html_doi
        if not full_text and html_full_text:
            full_text = html_full_text
        if not sentences and html_paragraphs:
            sentences = html_paragraphs
        if title == paper_name and html_title:
            title = html_title
        if not figures and html_figures:
            figures = html_figures

    if not full_text and sentences:
        full_text = " ".join(sentences)
    if not sentences and full_text:
        sentences = _sentence_tokenize(full_text)
    if tables_text:
        seen_sentences = {_norm_text(s) for s in sentences}
        for line in tables_text.splitlines():
            clean = re.sub(r"\s+", " ", _safe_str(line)).strip()
            key = _norm_text(clean)
            if clean and key and key not in seen_sentences:
                sentences.append(clean)
                seen_sentences.add(key)
    if _display_norm(doi) == "NR":
        m_doi = re.search(r"\b10\.\d{4,9}/[-._;()/:A-Za-z0-9]+", full_text)
        if m_doi:
            doi = m_doi.group(0).rstrip(".,;) ")
    return {"paper": paper_name, "title": title or paper_name, "doi": doi, "paragraphs": sentences, "full_text": full_text, "figures": figures}



def _figure_label_patterns(label: str) -> List[re.Pattern]:
    patterns: List[re.Pattern] = []
    if not label or label == "NR":
        return patterns
    label_norm = label.replace("–", "-").replace("—", "-").replace("−", "-")
    nums = re.findall(r"\d+", label_norm)
    if nums:
        for n in nums:
            patterns.append(re.compile(rf"\bfig(?:ure)?\.?\s*{re.escape(n)}\b", flags=re.IGNORECASE))
    else:
        patterns.append(re.compile(re.escape(label_norm), flags=re.IGNORECASE))
    return patterns


def _figure_numbers_in_text(text: str) -> List[str]:
    return re.findall(r"\bfig(?:ure)?\.?\s*(\d+)\b", _safe_str(text), flags=re.IGNORECASE)


def _contains_non_target_figure_reference(text: str, target_nums: List[str]) -> bool:
    nums = _figure_numbers_in_text(text)
    if not nums:
        return False
    targets = {str(x) for x in target_nums}
    return any(str(n) not in targets for n in nums)


CONTEXT_BRIDGE_KEYWORDS = [
    "same test procedures",
    "same test procedure",
    "same procedures",
    "same procedure",
    "same conditions",
    "same operating conditions",
    "all the cmd experiments",
    "all cmd experiments",
]

CONDITION_CONTEXT_KEYWORDS = [
    "reaction temperature",
    "temperature",
    "methane concentration",
    "ch4 concentration",
    "feed flow",
    "total flow",
    "sccm",
    "ml/min",
    "l/min",
    "catalyst",
    "wt%",
    "pressure",
    "time-on-stream",
    "time on stream",
]


def _is_bridge_paragraph(text: str) -> bool:
    t = _norm_text(text)
    return any(key in t for key in CONTEXT_BRIDGE_KEYWORDS)


def _has_condition_signal(text: str) -> bool:
    t = _norm_text(text)
    return any(key in t for key in CONDITION_CONTEXT_KEYWORDS)


def find_figure_mentions(parsed_xml: Dict[str, Any], figure_label: str, max_hits: int = 4) -> List[str]:
    paragraphs = parsed_xml.get("paragraphs", []) or []
    if not paragraphs:
        return []
    pats = _figure_label_patterns(figure_label)
    if not pats:
        return []
    hits: List[str] = []
    for p in paragraphs:
        text = re.sub(r"\s+", " ", _safe_str(p)).strip()
        if not text:
            continue
        if any(pat.search(text) for pat in pats):
            hits.append(text)
            if len(hits) >= max_hits:
                break
    return hits


def _find_figure_mention_indices(parsed_xml: Dict[str, Any], figure_label: str, max_hits: int = 6) -> List[int]:
    paragraphs = parsed_xml.get("paragraphs", []) or []
    if not paragraphs:
        return []
    pats = _figure_label_patterns(figure_label)
    if not pats:
        return []
    idxs: List[int] = []
    for i, p in enumerate(paragraphs):
        text = re.sub(r"\s+", " ", _safe_str(p)).strip()
        if text and any(pat.search(text) for pat in pats):
            idxs.append(i)
            if len(idxs) >= max_hits:
                break
    return idxs


def _extract_figure_caption_text(figure_info: Optional[Dict[str, Any]]) -> str:
    if not figure_info:
        return ""
    return re.sub(r"\s+", " ", _safe_str(figure_info.get("caption", ""))).strip()


def _sentence_records_from_parsed_xml(parsed_xml: Dict[str, Any]) -> List[Dict[str, Any]]:
    paragraphs = [re.sub(r"\s+", " ", _safe_str(p)).strip() for p in (parsed_xml.get("paragraphs", []) or [])]
    records: List[Dict[str, Any]] = []
    for para_idx, paragraph in enumerate(paragraphs):
        if not paragraph:
            continue
        for sent_idx, sent in enumerate(_sentence_tokenize(paragraph)):
            clean = re.sub(r"\s+", " ", _safe_str(sent)).strip()
            if not clean:
                continue
            records.append({
                "text": clean,
                "paragraph_idx": para_idx,
                "sentence_idx": sent_idx,
            })
    return records


def _find_figure_mention_sentence_indices(parsed_xml: Dict[str, Any], figure_label: str, max_hits: int = 10) -> List[int]:
    sentence_records = _sentence_records_from_parsed_xml(parsed_xml)
    if not sentence_records:
        return []
    pats = _figure_label_patterns(figure_label)
    if not pats:
        return []
    idxs: List[int] = []
    for i, rec in enumerate(sentence_records):
        txt = rec.get("text", "")
        if txt and any(pat.search(txt) for pat in pats):
            idxs.append(i)
            if len(idxs) >= max_hits:
                break
    return idxs


def _extract_condition_sentences_from_paragraph(text: str, bridge_only: bool = False) -> str:
    sentences = _sentence_tokenize(text)
    keep: List[str] = []
    for sent in sentences:
        st = _norm_text(sent)
        if bridge_only:
            if _is_bridge_paragraph(sent):
                keep.append(re.sub(r"\s+", " ", sent).strip())
            continue
        if _is_bridge_paragraph(sent):
            keep.append(re.sub(r"\s+", " ", sent).strip())
            continue
        if any(key in st for key in [
            "typical experimental test",
            "methane concentration",
            "ch4 concentration",
            "reaction temperature",
            "designated reaction temperature",
            "total flow",
            "feed flow",
            "sccm",
            "ml/min",
            "l/min",
            "pressure",
            "as the catalyst",
        ]):
            keep.append(re.sub(r"\s+", " ", sent).strip())
    out: List[str] = []
    seen = set()
    for sent in keep:
        key = _norm_text(sent)
        if key and key not in seen:
            seen.add(key)
            out.append(sent)
    return " ".join(out).strip()


def _find_shared_procedure_context(parsed_xml: Dict[str, Any]) -> str:
    paragraphs = [re.sub(r"\s+", " ", _safe_str(p)).strip() for p in (parsed_xml.get("paragraphs", []) or [])]
    if not paragraphs:
        return ""
    bridge_idxs = [i for i, p in enumerate(paragraphs) if _is_bridge_paragraph(p)]
    if not bridge_idxs:
        return ""

    parts: List[str] = []
    seen = set()

    def add_part(txt: str) -> None:
        clean = re.sub(r"\s+", " ", _safe_str(txt)).strip()
        key = _norm_text(clean)
        if clean and key and key not in seen:
            seen.add(key)
            parts.append(clean)

    for i in bridge_idxs[:2]:
        bridge_text = _extract_condition_sentences_from_paragraph(paragraphs[i], bridge_only=True)
        if bridge_text:
            add_part(bridge_text)
        for j in range(max(0, i - 5), i):
            txt = paragraphs[j]
            if not txt:
                continue
            if "typical experimental test" in _norm_text(txt):
                cond = _extract_condition_sentences_from_paragraph(txt)
                if cond:
                    add_part(cond)
                    break

    return " ".join(parts).strip()


def build_xml_figure_context(parsed_xml: Dict[str, Any], figure_info: Optional[Dict[str, Any]]) -> str:
    if not figure_info:
        return ""
    label = _display_norm(figure_info.get("label", "NR"))
    caption = _display_norm(figure_info.get("caption", "NR"))
    sentence_records = _sentence_records_from_parsed_xml(parsed_xml)
    parts: List[str] = []
    seen = set()

    def add_part(txt: str) -> None:
        clean = re.sub(r"\s+", " ", _safe_str(txt)).strip()
        key = _norm_text(clean)
        if clean and key and key not in seen:
            seen.add(key)
            parts.append(clean)

    if caption != "NR":
        add_part(f"{label}. {caption}" if label != "NR" else caption)
    if not sentence_records or label == "NR":
        return " ".join(parts).strip()

    target_nums = re.findall(r"\d+", label)
    mention_idxs = _find_figure_mention_sentence_indices(parsed_xml, label, max_hits=10)
    for idx in mention_idxs:
        start = max(0, idx - 10)
        end = min(len(sentence_records), idx + 11)
        for j in range(start, end):
            txt = sentence_records[j].get("text", "")
            if not txt:
                continue
            if _contains_non_target_figure_reference(txt, target_nums):
                continue
            add_part(txt)

    combined = " ".join(parts).strip()
    shared_ctx = _find_shared_procedure_context(parsed_xml)
    if shared_ctx:
        need_temp = _extract_single_temperature_with_snippet(combined)[0] is None
        need_frac = _extract_single_ch4_fraction(combined) is None
        if need_temp or need_frac:
            for sent in _sentence_tokenize(shared_ctx):
                add_part(sent)

    combined = _compact_text_for_prompt(" ".join(parts), max_chars=FIGURE_CONTEXT_MAX_CHARS)
    return combined


def _extract_candidate_catalyst_labels(text: str) -> List[str]:
    if not text:
        return []
    hits = []
    norm = text.replace("–", "-").replace("—", "-").replace("−", "-")
    for m in CATALYST_TOKEN_PATTERN.finditer(norm):
        val = re.sub(r"\s+", "", m.group(0))
        hits.append(val)
    out = []
    seen = set()
    for h in hits:
        if h not in seen:
            out.append(h)
            seen.add(h)
    return out


def _is_simple_carbon_catalyst_label(label: str) -> bool:
    s_norm = _norm_text(label)
    return s_norm in {"ac", "mbc"} or bool(re.fullmatch(r"bc[- ]?\d{3}", s_norm))


def _is_foam_variant_label(label: str) -> bool:
    s_norm = _norm_text(label)
    return bool(re.search(r"\b\d+\s*ppi\b", s_norm) or "nickel foam" in s_norm or re.fullmatch(r"nf[- ]?\d+", s_norm))


def _canonicalize_foam_label(label: str) -> str:
    s = _safe_str(label).strip()
    m = re.search(r"(\d+)\s*ppi", s, flags=re.IGNORECASE)
    if m:
        return f"Nickel foam {m.group(1)} PPI"
    m = re.search(r"nf[- ]?(\d+)", s, flags=re.IGNORECASE)
    if m:
        return f"Nickel foam {m.group(1)} PPI"
    return s or "NR"


def _extract_context_catalyst_candidates(text: str) -> List[str]:
    if not text:
        return []
    norm = _safe_str(text).replace("–", "-").replace("—", "-").replace("−", "-")
    hits: List[str] = []
    patterns = [
        r"\b(?:\d+\s*wt\s*%?\s*)?(?:NiO|Fe2O3|Fe3O4|CoO|CuO)(?:\s*[-/]\s*(?:Al2O3|AC|MC|SiC))?\b",
        r"\b(?:Ni\d+(?:-Fe\d+)+|Fe\d+(?:-Ni\d+)+|Ni\d+|Fe\d+)\b",
        r"\b(?:\d+Ni(?:-\d+Fe)+|\d+Fe(?:-\d+Ni)+|\d+Ni|\d+Fe)(?:[@/]\w+)?\b",
        r"\b(?:AC|MBC|BC-\d{3})\b",
        r"\b(?:Nickel\s+foam(?:\s+with)?(?:\s+a)?\s*PPI(?:\s+of)?\s*\d+|\d+\s*PPI|NF-?\d+)\b",
    ]
    for pat in patterns:
        for m in re.finditer(pat, norm, flags=re.IGNORECASE):
            val = re.sub(r"\s+", " ", m.group(0)).strip()
            if _is_foam_variant_label(val):
                val = _canonicalize_foam_label(val)
            hits.append(val)
    out: List[str] = []
    seen = set()
    for h in hits:
        key = _norm_text(h)
        if key and key not in seen:
            seen.add(key)
            out.append(h)
    return out


def _infer_support_from_context(text: str) -> str:
    s = _norm_text(text)
    if not s or s == "nr":
        return "NR"
    if "al2o3" in s:
        return "Al2O3"
    if "activated carbon" in s or re.search(r"\bac\b", s):
        return "AC"
    if "mesoporous carbon" in s or re.search(r"\bmc\b", s):
        return "MC"
    if "silicon carbide" in s or re.search(r"\bsic\b", s):
        return "SiC"
    if "nickel foam" in s or "ppi" in s:
        return "Nickel foam"
    if ZEOLITE_HINT_RE.search(_safe_str(text)):
        ratio = _extract_si_al_ratio_descriptor(text)
        return f"Zeolite ({ratio})" if ratio else "Zeolite"
    return "NR"


def _looks_like_condition_label(label: str) -> bool:
    s = _safe_str(label).strip()
    if not s or s == "NR":
        return False
    if _is_simple_carbon_catalyst_label(s) or _is_foam_variant_label(s):
        return False
    s_norm = _norm_text(s)
    if any(re.fullmatch(pat, s, flags=re.IGNORECASE) for pat in CONDITION_ONLY_PATTERNS):
        return True
    if s_norm in {"microwave", "electric furnace", "tubular furnace", "muffle furnace", "induction heating", "plasma-assisted", "solar heating", "~furnace", "others", "conventional", "mw", "cw"}:
        return True
    if any(tok in s_norm for tok in ["ml/min", "sccm", "slpm", "lpm", "vol%", "feed", "pressure", "temperature"]) and "@" not in s and "/" not in s:
        return True
    if re.search(r"^\d{2,4}\s*°?\s*c$", s_norm):
        return True
    return False


def _is_valid_catalyst_label(label: str) -> bool:
    s = _safe_str(label).strip()
    if not s or s == "NR":
        return False
    if _is_simple_carbon_catalyst_label(s) or _is_foam_variant_label(s):
        return True
    norm = s.replace("–", "-").replace("—", "-").replace("−", "-")
    if re.search(r"\b(?:Ni\d+(?:-Fe\d+)+|Fe\d+(?:-Ni\d+)+|Ni\d+|Fe\d+)\b", norm):
        return True
    if re.search(r"\b(?:\d+Ni(?:-\d+Fe)+|\d+Fe(?:-\d+Ni)+|\d+Ni|\d+Fe)(?:[@/]\w+)?\b", norm):
        return True
    if re.search(r"\b(?:\d+\s*wt\s*%?\s*)?(?:NiO|Fe2O3|Fe3O4|CoO|CuO)(?:\s*[-/]\s*(?:Al2O3|AC|MC|SiC))?\b", norm, flags=re.IGNORECASE):
        return True
    if _looks_like_condition_label(s):
        return False
    if CATALYST_TOKEN_PATTERN.search(s):
        return True
    loadings, support = parse_catalyst_label(s)
    if loadings:
        return True
    if re.search(r"[A-Z][a-z]?\s*[@/]\s*[A-Za-z0-9]", s):
        return True
    return False


def _looks_like_generic_phase_or_context_label(label: str) -> bool:
    s = _safe_str(label).strip()
    if not s or s == "NR":
        return False
    if _is_simple_carbon_catalyst_label(s) or _is_foam_variant_label(s):
        return False
    low = _norm_text(s)
    compact = re.sub(r"[\s_\-]+", "", low)
    generic_tokens = {
        "catalyst", "catalysts", "nanocatalyst", "nanocatalysts", "nanoparticles", "nanopowder",
        "nc", "ncs", "mixedoxide", "oxide", "metaloxide", "support"
    }
    if compact in generic_tokens:
        return True
    support_like = {"al2o3", "gal2o3", "gammaal2o3", "sio2", "sic", "ceo2", "zro2", "mgo", "cao", "tio2", "la2o3", "nickelfoam", "ac", "mc", "cnt"}
    if compact in support_like:
        return True
    if re.fullmatch(r"(?:g|gamma)?\s*[- ]?al2o3", low):
        return True
    if compact in {"alumina", "silica", "titania", "zirconia", "ceria", "magnesia", "calcia", "lanthana"}:
        return True
    return False

def _looks_like_specific_series_identity_label(label: str) -> bool:
    s = _safe_str(label).strip()
    if not s or s == "NR":
        return False
    if _looks_like_condition_label(s):
        return False
    if re.fullmatch(r"(?:curve|series)\s*\d+", _norm_text(s)):
        return False
    if _is_simple_carbon_catalyst_label(s) or _is_foam_variant_label(s):
        return True
    low = _norm_text(s)
    if "%" in s and re.search(r"\b(?:ch4|methane|h2|hydrogen|n2|nitrogen|co2|oxygen|steam|argon)\b", low):
        return False
    if any(tok in low for tok in ["pressure", "temperature", "time", "tos", "flow", "feed", "sccm", "ml/min", "l/min", "ghsv", "whsv", "atm", "bar", "kpa", "mpa"]):
        return False
    canonical = _canonicalize_catalyst_label_with_context(s, s)
    loadings, support = parse_catalyst_label(canonical)
    distinct_metals = [el for el in loadings.keys() if el != "O"]
    if support != "NR":
        return True
    if canonical != s and bool(loadings):
        return True
    if len(set(distinct_metals)) >= 2:
        return True
    if len(set(distinct_metals)) == 1 and not _looks_like_generic_phase_or_context_label(s):
        return True
    return False

def _preferred_series_label_as_catalyst(series_label: str, raw_label: str) -> Optional[str]:
    series = _safe_str(series_label).strip()
    raw = _safe_str(raw_label).strip()
    if not _looks_like_specific_series_identity_label(series):
        return None
    if raw in ("", "NR") or _looks_like_generic_phase_or_context_label(raw):
        return _canonicalize_foam_label(series) if _is_foam_variant_label(series) else series
    return None


def _enforce_single_panel_consistency(records: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    if len(records) <= 1:
        return records
    groups: Dict[Tuple[str, str], List[Dict[str, Any]]] = {}
    for rec in records:
        key = (_display_norm(rec.get("X_Axis_Label", "NR")), _display_norm(rec.get("Y_Axis_Label", "NR")))
        groups.setdefault(key, []).append(rec)
    if len(groups) <= 1:
        return records

    def _pair_score(item: Tuple[Tuple[str, str], List[Dict[str, Any]]]) -> Tuple[int, int, int, int, int]:
        (x_label, y_label), recs = item
        y_norm = _norm_text(y_label)
        return (
            len(recs),
            int("conversion" in y_norm or "xch4" in y_norm or _is_unreacted_ch4_axis(y_norm)),
            int(_x_axis_kind(x_label) != "other"),
            int(_norm_text(x_label) != "nr"),
            int(_norm_text(y_label) != "nr"),
        )

    best_pair, best_records = max(groups.items(), key=_pair_score)
    dropped_count = sum(len(v) for k, v in groups.items() if k != best_pair)
    if dropped_count > 0:
        note = (
            f"Rows from other panel axes were dropped to keep one selected target panel: "
            f"kept X_Axis_Label='{best_pair[0]}' and Y_Axis_Label='{best_pair[1]}'."
        )
        for rec in best_records:
            rec[INTERNAL_NOTES_KEY] = _append_note(rec.get(INTERNAL_NOTES_KEY, "NR"), note)
    return best_records

def _x_axis_kind(x_axis_label: Any) -> str:
    x = _norm_text(x_axis_label)
    if not x or x == "nr":
        return "other"

    if any(tok in x for tok in ["space velocity", "ghsv", "whsv", "gshv", "weight hourly space velocity", "gas hourly space velocity"]):
        return "space_velocity"
    if ("ch4" in x or "methane" in x) and any(tok in x for tok in ["flow", "feed", "rate", "ml/min", "l/min", "sccm", "slpm", "lpm"]):
        return "ch4_flow"
    if any(tok in x for tok in ["catalyst mass", "mass of catalyst", "catalyst weight", "weight of catalyst"]):
        return "catalyst_mass"
    if any(tok in x for tok in ["pressure", "bar", "atm", "kpa", "mpa", "pa"]):
        return "pressure"
    if "temperature" in x or re.search(r"(?:°c|\bc\b|\bk\b|\bf\b)", x):
        if "1000/t" not in x and "1/t" not in x:
            return "temperature"
    if "time" in x or "time on stream" in x or re.search(r"\btos\b", x):
        return "time"
    if any(tok in x for tok in ["flow rate", "feed flow", "total flow", "feed rate", "gas flow", "ml/min", "l/min", "sccm", "slpm", "lpm"]):
        return "total_flow"
    if "mass" in x and "catalyst" in x:
        return "catalyst_mass"
    return "other"


def _x_axis_numeric_value(record: Dict[str, Any]) -> Optional[float]:
    kind = _x_axis_kind(record.get("X_Axis_Label", "NR"))
    if kind == "time":
        return _parse_time_to_minutes(record.get("Reaction_Time_hr_min_s", "NR"))
    if kind == "temperature":
        return _parse_first_float(record.get("Temperature_C_or_K", "NR"))
    if kind == "ch4_flow":
        return _parse_first_float(record.get("CH4_Feed_Flow_Rate_mL_or_L_per_min", "NR"))
    if kind == "total_flow":
        return _parse_first_float(record.get("Feed_Flow_Rate_mL_or_L_per_min", "NR"))
    if kind == "catalyst_mass":
        return _parse_first_float(record.get("Catalyst_Mass_g_or_mg", "NR"))
    if kind == "pressure":
        return _parse_first_float(record.get("Pressure_bar_or_atm_or_Pa", "NR"))
    if kind == "space_velocity":
        return _parse_first_float(record.get("Space_Velocity", "NR"))

    for key in [
        "Reaction_Time_hr_min_s",
        "Temperature_C_or_K",
        "CH4_Feed_Flow_Rate_mL_or_L_per_min",
        "Feed_Flow_Rate_mL_or_L_per_min",
        "Catalyst_Mass_g_or_mg",
        "Pressure_bar_or_atm_or_Pa",
        "Space_Velocity",
    ]:
        if key == "Reaction_Time_hr_min_s":
            val = _parse_time_to_minutes(record.get(key, "NR"))
        else:
            val = _parse_first_float(record.get(key, "NR"))
        if val is not None:
            return val
    return None


def _compact_prompt_constants(constants_json: Dict[str, Any], keys: Optional[List[str]] = None) -> str:
    default_keys = [
        "Catalyst_Label",
        "Heating_Mode",
        "Pressure_bar_or_atm_or_Pa",
        "Reaction_Time_hr_min_s",
        "Feed_Flow_Rate_mL_or_L_per_min",
        "CH4_Feed_Flow_Rate_mL_or_L_per_min",
        "Feed_Gas_Mixture",
        "Space_Velocity",
        "Catalyst_Mass_g_or_mg",
        "Temperature_C_or_K",
        "Precursor_Salts",
    ]
    chosen_keys = keys or default_keys
    compact: Dict[str, Any] = {}
    for key in chosen_keys:
        val = _display_norm(constants_json.get(key, "NR"))
        if val != "NR":
            compact[key] = val
    return json.dumps(compact, ensure_ascii=False, separators=(",", ":")) if compact else "{}"


def _json_key_list(keys: List[str]) -> str:
    return json.dumps(list(keys), ensure_ascii=False, separators=(',', ':'))


def _prompt_join(*parts: str) -> str:
    cleaned: List[str] = []
    for part in parts:
        s = _safe_str(part).strip()
        if s:
            cleaned.append(s)
    return "\n".join(cleaned)


def _compact_prompt_string(value: Any, max_chars: int = 1200) -> str:
    s = re.sub(r'\s+', ' ', _safe_str(value)).strip()
    if not s:
        return 'NR'
    return s if len(s) <= max_chars else s[: max_chars - 3].rstrip() + '...'


def _compact_prompt_list(values: List[str], max_items: int = 20) -> str:
    clean: List[str] = []
    seen = set()
    for item in values:
        s = re.sub(r'\s+', ' ', _safe_str(item)).strip()
        if not s:
            continue
        key = _norm_text(s)
        if key in seen:
            continue
        seen.add(key)
        clean.append(s)
        if len(clean) >= max_items:
            break
    return json.dumps(clean, ensure_ascii=False, separators=(',', ':'))


# Figure-extraction instructions are stored in prompt template files under PROMPTS_DIR.
# See figure_extraction.txt, panel_series_discovery.txt, and single_series_extraction.txt.



def _constants_extraction_user_prompt(chunk_idx: int, n_chunks: int) -> str:
    fields = _json_key_list(PROMPT_TEXT_CONSTANT_FIELDS)
    definitions = _load_prompt("constants_field_definitions")
    return _render_prompt_template(
        "constants_extraction",
        CHUNK_IDX=chunk_idx,
        N_CHUNKS=n_chunks,
        OUTPUT_KEYS=fields,
        FIELD_DEFINITIONS=definitions,
    )


def _figure_extraction_user_prompt(
    fig_filename: str,
    matched_xml_label: str,
    matched_xml_order: Any,
    caption_context: str,
    constants_json: Dict[str, Any],
    *,
    strict: bool = False,
    previous_extraction_summary: Optional[str] = None,
    expected_series_labels: Optional[List[str]] = None,
    expected_series_counts: Optional[Dict[str, Any]] = None,
    expected_panel_specs: Optional[List[Dict[str, Any]]] = None,
    figure_keys: Optional[List[str]] = None,
) -> str:
    keys = _json_key_list(figure_keys or FIGURE_EXTRACTION_KEYS)
    const_snip = _compact_prompt_constants(constants_json, keys=PROMPT_TEXT_CONSTANT_FIELDS)
    ctx = _compact_text_for_prompt(caption_context if caption_context else 'NR', max_chars=FIGURE_CONTEXT_MAX_CHARS)

    strict_block = _render_prompt_template("strict_retry_block") if strict else ""
    prev_block = (
        _render_prompt_template(
            "previous_bad_output_block",
            PREVIOUS_BAD_OUTPUT=_compact_text_for_prompt(previous_extraction_summary, max_chars=1200),
        )
        if previous_extraction_summary else ""
    )
    expected = [s for s in (expected_series_labels or []) if _safe_str(s).strip()]
    expected_series_block = (
        _render_prompt_template(
            "expected_series_block",
            EXPECTED_SERIES=json.dumps(expected, ensure_ascii=False, separators=(",", ":")),
        )
        if expected else ""
    )
    expected_counts_clean = {str(k): v for k, v in (expected_series_counts or {}).items() if _safe_str(k).strip()}
    expected_counts_block = (
        _render_prompt_template(
            "expected_series_counts_block",
            EXPECTED_SERIES_COUNTS=json.dumps(expected_counts_clean, ensure_ascii=False, separators=(",", ":")),
        )
        if expected_counts_clean else ""
    )
    panel_specs_block = (
        _render_prompt_template(
            "expected_panel_specs_block",
            EXPECTED_PANEL_SPECS=json.dumps(expected_panel_specs, ensure_ascii=False, separators=(",", ":")),
        )
        if expected_panel_specs else ""
    )
    return _render_prompt_template(
        "figure_extraction",
        FIGURE_FILENAME=fig_filename,
        MATCHED_XML_LABEL=matched_xml_label,
        MATCHED_XML_ORDER=matched_xml_order,
        FIGURE_CONTEXT=ctx,
        PAPER_CONSTANTS=const_snip,
        STRICT_RETRY_BLOCK=strict_block,
        PREVIOUS_BAD_OUTPUT_BLOCK=prev_block,
        EXPECTED_SERIES_BLOCK=expected_series_block,
        EXPECTED_SERIES_COUNTS_BLOCK=expected_counts_block,
        EXPECTED_PANEL_SPECS_BLOCK=panel_specs_block,
        OUTPUT_KEYS=keys,
    )


def _panel_series_discovery_user_prompt(
    fig_filename: str,
    matched_xml_label: str,
    matched_xml_order: Any,
    caption_context: str,
) -> str:
    ctx = _compact_text_for_prompt(caption_context if caption_context else 'NR', max_chars=FIGURE_CONTEXT_MAX_CHARS)
    discovery_keys = _json_key_list(["has_multiple_panels", "all_panel_count", "panels", "is_target_conversion_figure", "notes"])
    return _render_prompt_template(
        "panel_series_discovery",
        FIGURE_FILENAME=fig_filename,
        MATCHED_XML_LABEL=matched_xml_label,
        MATCHED_XML_ORDER=matched_xml_order,
        FIGURE_CONTEXT=ctx,
        OUTPUT_KEYS=discovery_keys,
    )


def _single_series_extraction_user_prompt(
    fig_filename: str,
    matched_xml_label: str,
    matched_xml_order: Any,
    caption_context: str,
    constants_json: Dict[str, Any],
    target_series_label: str,
    figure_keys: Optional[List[str]] = None,
    *,
    panel_spec: Optional[Dict[str, Any]] = None,
    series_spec: Optional[Dict[str, Any]] = None,
    expected_count: Optional[int] = None,
    strict: bool = False,
    previous_extraction_summary: Optional[str] = None,
) -> str:
    keys = _json_key_list(figure_keys or FIGURE_EXTRACTION_KEYS)
    const_snip = _compact_prompt_constants(constants_json, keys=PROMPT_TEXT_CONSTANT_FIELDS)
    ctx = _compact_text_for_prompt(caption_context if caption_context else 'NR', max_chars=FIGURE_CONTEXT_MAX_CHARS)
    locked = json.dumps(_safe_str(target_series_label), ensure_ascii=False)
    panel_json = json.dumps(panel_spec or {}, ensure_ascii=False, separators=(",", ":"))
    series_json = json.dumps(series_spec or {}, ensure_ascii=False, separators=(",", ":"))
    expected_count_text = str(expected_count) if expected_count is not None else "NR"
    strict_block = _render_prompt_template("strict_retry_block") if strict else ""
    prev_block = (
        _render_prompt_template(
            "previous_bad_output_block",
            PREVIOUS_BAD_OUTPUT=_compact_text_for_prompt(previous_extraction_summary, max_chars=1200),
        )
        if previous_extraction_summary else ""
    )
    return _render_prompt_template(
        "single_series_extraction",
        FIGURE_FILENAME=fig_filename,
        MATCHED_XML_LABEL=matched_xml_label,
        MATCHED_XML_ORDER=matched_xml_order,
        TARGET_SERIES_LABEL=locked,
        TARGET_PANEL_JSON=panel_json,
        TARGET_SERIES_JSON=series_json,
        EXPECTED_MARKER_COUNT=expected_count_text,
        FIGURE_CONTEXT=ctx,
        PAPER_CONSTANTS=const_snip,
        STRICT_RETRY_BLOCK=strict_block,
        PREVIOUS_BAD_OUTPUT_BLOCK=prev_block,
        OUTPUT_KEYS=keys,
    )

def _looks_like_equilibrium_or_max_conversion(text: Any) -> bool:
    t = _norm_text(text)
    if not t or t == "nr":
        return False
    phrases = [
        "equilibrium conversion",
        "thermodynamic equilibrium",
        "equilibrium methane conversion",
        "maximum conversion",
        "maximum methane conversion",
        "maximum xch4",
        "max conversion",
        "max methane conversion",
        "equilibrium xch4",
    ]
    return any(p in t for p in phrases)


def _text_has_methane_or_ch4(text: Any) -> bool:
    t = _norm_text(text)
    if not t or t == "nr":
        return False
    # _norm_text translates CH₄/CH⁴ -> CH4.
    return bool(re.search(r"\b(?:ch\s*4|methane)\b", t, flags=re.IGNORECASE))


def _is_unreacted_ch4_axis(text: Any) -> bool:
    """
    True when the selected y-axis/metric is the amount of methane that remains
    rather than methane converted. This intentionally includes common paper
    labels such as:
      - Amount of unreacted CH4 (%)
      - Unreacted methane vs time
      - Remaining/unconverted/residual methane
      - Methane slip
      - Outlet/effluent CH4 concentration/content/fraction
    """
    t = _norm_text(text)
    if not t or t == "nr" or not _text_has_methane_or_ch4(t):
        return False

    direct_remaining_terms = [
        "unreacted",
        "unconverted",
        "unreacted methane",
        "unreacted ch4",
        "methane unreacted",
        "ch4 unreacted",
        "remaining methane",
        "remaining ch4",
        "methane remaining",
        "ch4 remaining",
        "residual methane",
        "residual ch4",
        "methane slip",
        "ch4 slip",
    ]
    if any(term in t for term in direct_remaining_terms):
        return True

    # Patterns cover word-order variants such as "amount of methane unreacted".
    remaining_patterns = [
        r"\b(?:amount|content|concentration|fraction|percentage|percent)\s+of\s+(?:remaining|unreacted|unconverted|residual)\s+(?:methane|ch\s*4)\b",
        r"\b(?:amount|content|concentration|fraction|percentage|percent)\s+of\s+(?:methane|ch\s*4)\s+(?:remaining|unreacted|unconverted|residual)\b",
        r"\b(?:methane|ch\s*4)\s+(?:remaining|unreacted|unconverted|residual)\s+(?:amount|content|concentration|fraction|percentage|percent)?\b",
        r"\b(?:remaining|unreacted|unconverted|residual)\s+(?:amount|content|concentration|fraction|percentage|percent)?\s*(?:of\s+)?(?:methane|ch\s*4)\b",
    ]
    if any(re.search(pat, t, flags=re.IGNORECASE) for pat in remaining_patterns):
        return True

    # For methane decomposition, outlet/effluent/product-gas methane is the
    # unreacted methane stream, but require a concentration-like term so a
    # generic sentence mentioning "outlet" is not enough.
    outlet_terms = ["outlet", "effluent", "off-gas", "off gas", "product gas", "exit gas", "exhaust"]
    concentration_terms = ["concentration", "content", "fraction", "percentage", "percent", "%", "amount"]
    if any(term in t for term in outlet_terms) and any(term in t for term in concentration_terms):
        return True

    return False


def _is_fraction_like_conversion_axis(text: Any) -> bool:
    t = _norm_text(text)
    if not t or "nr" == t:
        return False
    raw = _safe_str(text).translate(UNICODE_SUBSCRIPT_SUPERSCRIPT_TRANSLATION)
    has_percent_unit = "%" in raw or "percent" in t
    if _is_unreacted_ch4_axis(t):
        return "fraction" in t or not has_percent_unit
    return (("conversion" in t or "xch4" in t) and _text_has_methane_or_ch4(t) and ("fraction" in t or not has_percent_unit))


def _row_has_unreacted_ch4_metric(row: Dict[str, Any], context_text: Any = "") -> bool:
    y_label = row.get("Y_Axis_Label", "NR")
    if _is_unreacted_ch4_axis(y_label):
        return True

    # Conservative fallback: if the selected y-axis is a generic methane amount
    # label and the figure/caption/notes explicitly say this panel reports
    # unreacted/remaining methane, treat it as an unreacted-methane metric.
    y = _norm_text(y_label)
    generic_methane_amount_axis = (
        _text_has_methane_or_ch4(y)
        and any(tok in y for tok in ["amount", "content", "concentration", "fraction", "percentage", "percent", "%"])
        and not any(tok in y for tok in ["conversion", "xch4", "yield", "selectivity", "rate"])
    )
    if not generic_methane_amount_axis:
        return False

    context = " ".join([
        _safe_str(context_text),
        _safe_str(row.get("_Caption_Context", "")),
        _safe_str(row.get("Matched_XML_Figure_Label", "")),
        _safe_str(row.get("Caption_Match_Method", "")),
    ])
    return _is_unreacted_ch4_axis(context)


def _normalize_ch4_percent_value_from_axis(raw: Any, axis_text: Any) -> Optional[float]:
    val = _parse_first_float(raw)
    if val is None:
        return None
    axis = _safe_str(axis_text).translate(UNICODE_SUBSCRIPT_SUPERSCRIPT_TRANSLATION)
    # If the visible axis is explicitly a fraction or lacks a percent unit, a
    # value <= 1.2 is interpreted as fraction. This also handles occasional LLM
    # outputs like 0.75 for a fraction-scale axis.
    if _is_fraction_like_conversion_axis(axis_text) and val <= 1.2:
        return val * 100.0
    return val



def _dynamic_figure_extraction_keys(constants: Dict[str, Any]) -> List[str]:
    if FIGURE_SCHEMA_MODE == "compact":
        return list(FIGURE_LLM_POINT_FIELDS)
    keys = list(FIGURE_BASE_KEYS)
    excluded = {"Title", "DOI"}
    for fld in TEXT_CONSTANT_FIELDS:
        if fld in excluded or fld in keys:
            continue
        if _display_norm(constants.get(fld, "NR")) == "NR":
            keys.append(fld)
    return keys

def _infer_reduction_state(row: Dict[str, Any], context_text: str = "") -> str:
    text = " ".join([
        _safe_str(row.get("Catalyst_Label", "")),
        _safe_str(row.get("Series_Label", "")),
        _safe_str(context_text),
    ])
    t = _norm_text(text)
    if any(tok in t for tok in ["unreduced", "non-reduced", "not reduced", "as-prepared", "fresh catalyst", "calcined catalyst"]):
        return "unreduced"
    if any(tok in t for tok in ["pre-reduced", "h2-reduced", "reduced form", "reduced catalyst", "after reduction"]):
        return "reduced"
    return "unknown"


def _finalize_ch4_conversion_fields(row: Dict[str, Any], context_text: Any = "") -> None:
    """
    Normalize CH4_Conversion_% after figure extraction.

    For normal conversion axes, a fraction such as 0.42 becomes 42%.
    For unreacted/remaining/outlet methane axes, the stored value is converted
    deterministically to conversion using:
        CH4 conversion (%) = 100 - unreacted CH4 (%)

    The internal flag makes the operation safe to call more than once in the
    same run and lets the final output pass repair rows loaded from old progress
    files without double-converting freshly finalized rows.
    """
    if row.get(CH4_CONVERSION_FINALIZED_KEY) == "yes":
        return

    raw = row.get("CH4_Conversion_%", "NR")
    val = _parse_first_float(raw)
    if val is None:
        return

    y_label = row.get("Y_Axis_Label", "NR")
    approx = _safe_str(raw).strip().startswith("~") or "≈" in _safe_str(raw)
    converted = None

    if _row_has_unreacted_ch4_metric(row, context_text):
        unreacted_pct = _normalize_ch4_percent_value_from_axis(raw, y_label)
        if unreacted_pct is None:
            return
        converted = 100.0 - unreacted_pct
        note = f"CH4_Conversion_% converted from unreacted methane value {_safe_str(raw)} using 100 - value."
        row[INTERNAL_NOTES_KEY] = _append_note(row.get(INTERNAL_NOTES_KEY, "NR"), note)
    elif _is_fraction_like_conversion_axis(y_label) and val <= 1.2:
        converted = val * 100.0

    if converted is None:
        return

    converted = max(0.0, min(100.0, converted))
    row["CH4_Conversion_%"] = f"{'~' if approx else ''}{_format_number_clean(converted)}"
    row[CH4_CONVERSION_FINALIZED_KEY] = "yes"


TARGET_CONVERSION_CUES = [
    "methane conversion",
    "methane conversions",
    "ch4 conversion",
    "ch4 conversions",
    "conversion of methane",
    "conversions of methane",
    "conversion of ch4",
    "conversions of ch4",
    "xch4",
    "methane remaining",
    "ch4 remaining",
    "methane unreacted",
    "ch4 unreacted",
    "methane unconverted",
    "ch4 unconverted",
    "residual methane",
    "residual ch4",
    "methane slip",
    "ch4 slip",
    "outlet methane concentration",
    "outlet ch4 concentration",
    "effluent methane concentration",
    "effluent ch4 concentration",
]

WEAK_PROCESS_TARGET_CUES = [
    "methane decomposition",
    "decomposition of methane",
    "catalytic methane decomposition",
    "catalytic decomposition of methane",
    "ch4 decomposition",
    "decomposition of ch4",
    "methane cracking",
    "cracking of methane",
    "ch4 cracking",
    "methane pyrolysis",
    "pyrolysis of methane",
    "ch4 pyrolysis",
    "cmd",
]

WEAK_ACTIVITY_TARGET_CUES = [
    "tos",
    "tos study",
    "time on stream",
    "time-on-stream",
    "stability",
    "stability test",
    "stability study",
    "deactivation",
    "durability",
    "durability test",
    "durability study",
    "lifetime",
]

STRONG_ACTIVITY_TARGET_CUES = [
    "catalyst activity",
    "catalyst activities",
    "catalytic activity",
    "catalytic activities",
    "activity of catalyst",
    "activities of catalyst",
    "activity of catalysts",
    "activities of catalysts",
    "activity of all catalyst",
    "activities of all catalyst",
    "activity of all catalysts",
    "activities of all catalysts",
]

SOFT_NON_TARGET_ALLOWABLE_GROUPS = {"hydrogen metric"}

NON_TARGET_FIGURE_CUE_GROUPS = {
    "adsorption/desorption": [
        "quantity adsorbed",
        "adsorption",
        "desorption",
        "adsorption desorption",
        "p/p°",
        "p/p0",
        "p/po",
        "isotherm",
    ],
    "carbon deposition/yield": [
        "carbon deposited",
        "carbon yield",
        "amount of carbon",
        "carbon deposition",
        "carbon formed",
        "growth of carbon",
        "carbon filament",
        "carbon filaments",
        "growth of carbon filament",
        "growth of carbon filaments",
        "filamentous carbon",
        "carbon nanofiber",
        "carbon nanofibers",
        "carbon nanotube",
        "carbon nanotubes",
        "carbon morphology",
        "coke morphology",
    ],
    "hydrogen metric": [
        "h2 production rate",
        "h2 yield",
        "h2 yields",
        "h2 productivity",
        "hydrogen production",
        "hydrogen productions",
        "hydrogen yield",
        "hydrogen yields",
        "hydrogen productivity",
        "hydrogen selectivity",
        "hydrogen selectivities",
        "h2 evolution",
        "hydrogen evolution",
    ],
    "arrhenius/kinetic": [
        "lnr",
        "ln(rate)",
        "ln rate",
        "1000/t",
        "arrhenius",
        "activation energy",
        "apparent activation energy",
    ],
    "equilibrium/max conversion": [
        "equilibrium conversion",
        "thermodynamic equilibrium",
        "equilibrium methane conversion",
        "maximum conversion",
        "maximum methane conversion",
        "maximum xch4",
        "max conversion",
        "max methane conversion",
        "equilibrium xch4",
    ],
    "schematic/setup": [
        "schematic",
        "schematics",
        "schematic diagram",
        "schematic representation",
        "schematic illustration",
        "conceptual diagram",
        "conceptual illustration",
        "diagrammatic illustration",
        "graphical illustration",
        "experimental setup",
        "experimental set up",
        "setup schematic",
        "setup diagram",
        "set up schematic",
        "set up diagram",
        "reactor schematic",
        "reactor setup",
        "reactor diagram",
        "apparatus schematic",
        "apparatus diagram",
        "process flow diagram",
        "flow diagram",
        "flowsheet",
        "block diagram",
    ],
    "characterization": [
        "xrd",
        "x-ray diffraction",
        "raman",
        "ftir",
        "ir",
        "infra red",
        "infra-red",
        "infra_red",
        "infrared",
        "infra red spectroscopy",
        "infra-red spectroscopy",
        "infrared spectroscopy",
        "tga",
        "Thermogravimetric Analysis",
        "tpo",
        "tpr",
        "tpd",
        "temperature programmed desorption",
        "temperature-programmed desorption",
        "temperature programmed disorption",
        "temperature-programmed disorption",
        "nh3 tpd",
        "co2 tpd",
        "tem",
        "stem",
        "sem",
        "xps",
        "eds",
        "edx",
        "sem_edx",
        "sem-edx",
        "sem edx",
        "tem_edx",
        "tem-edx",
        "tem edx",
        "stem_edx",
        "stem-edx",
        "stem edx",
        "nmr",
        "nuclear magnetic resonance",
        "mapping",
        "elemental map",
        "microscopy",
        "micrograph",
        "micrographs",
        "morphology",
        "surface morphology",
        "bet",
        "bet surface area",
        "brunauer emmett teller",
        "bjh",
        "xanes",
        "exafs",
        "uv vis",
    ],
}


EXPLICIT_TARGET_METRIC_PAIR_PATTERN = re.compile(
    r"\b(?:methane|ch4)\b.{0,40}\b(?:conversion(?:s)?|remaining|unreacted|unconverted|residual|slip)\b"
    r"|\b(?:conversion(?:s)?|remaining|unreacted|unconverted|residual|slip)\b.{0,40}\b(?:methane|ch4)\b",
    flags=re.IGNORECASE,
)

EXPLICIT_TARGET_STREAM_PATTERN = re.compile(
    r"\b(?:outlet|effluent|off[- ]gas|product gas)\b.{0,40}\b(?:methane|ch4)\b.{0,20}\b(?:concentration|content|fraction|percentage|percent)\b"
    r"|\b(?:methane|ch4)\b.{0,20}\b(?:concentration|content|fraction|percentage|percent)\b.{0,40}\b(?:outlet|effluent|off[- ]gas|product gas)\b",
    flags=re.IGNORECASE,
)

EXPLICIT_TARGET_REMAINING_AMOUNT_PATTERN = re.compile(
    r"\b(?:amount|content|concentration|fraction|percentage|percent)\s+of\s+(?:remaining|unreacted|unconverted|residual)\s+(?:methane|ch4)\b"
    r"|\b(?:remaining|unreacted|unconverted|residual)\s+(?:methane|ch4)\s+(?:amount|content|concentration|fraction|percentage|percent)?\b",
    flags=re.IGNORECASE,
)

EXPLICIT_TARGET_PERCENT_AXIS_PATTERN = re.compile(
    r"\b(?:methane|ch4)\s*[\(\[]\s*%\s*[\)\]]"
    r"|[\(\[]\s*%\s*[\)\]]\s*(?:methane|ch4)\b",
    flags=re.IGNORECASE,
)


def _cue_regex(cue: str) -> re.Pattern:
    parts: List[str] = []
    for ch in _safe_str(cue).strip():
        if ch.isspace() or ch in "-_/":
            parts.append(r"[\s_\-/]*")
        else:
            parts.append(re.escape(ch))
    pattern = "".join(parts)
    return re.compile(rf"(?<!\w){pattern}(?!\w)", flags=re.IGNORECASE)


def _cue_present(text: str, cue: str) -> bool:
    return bool(_cue_regex(cue).search(text))


def _remove_cues(text: str, cues: List[str]) -> str:
    out = text
    for cue in cues:
        out = _cue_regex(cue).sub(" ", out)
    out = re.sub(r"\s+", " ", out).strip()
    return out


def _has_target_conversion_cue(text: Any) -> bool:
    t = _norm_text(text)
    if not t or t == "nr":
        return False
    all_non_target_cues = [cue for cues in NON_TARGET_FIGURE_CUE_GROUPS.values() for cue in cues]
    stripped = _remove_cues(t, all_non_target_cues)
    if any(_cue_present(stripped, cue) for cue in TARGET_CONVERSION_CUES):
        return True
    if EXPLICIT_TARGET_STREAM_PATTERN.search(stripped) or EXPLICIT_TARGET_REMAINING_AMOUNT_PATTERN.search(stripped):
        return True
    if _is_unreacted_ch4_axis(stripped):
        return True
    return (("conversion" in stripped or "xch4" in stripped) and ("methane" in stripped or "ch4" in stripped))


def _caption_has_explicit_target_metric(text: Any) -> bool:
    t = _norm_text(text)
    if not t or t == "nr":
        return False
    compact = t.replace(" ", "")
    if "xch4" in compact:
        return True
    if any(_cue_present(t, cue) for cue in TARGET_CONVERSION_CUES):
        return True
    if EXPLICIT_TARGET_METRIC_PAIR_PATTERN.search(t):
        return True
    if EXPLICIT_TARGET_STREAM_PATTERN.search(t):
        return True
    if EXPLICIT_TARGET_REMAINING_AMOUNT_PATTERN.search(t):
        return True
    if EXPLICIT_TARGET_PERCENT_AXIS_PATTERN.search(t):
        return True
    return False


def _has_weak_process_target_cue(text: Any) -> bool:
    t = _norm_text(text)
    if not t or t == "nr":
        return False
    return any(_cue_present(t, cue) for cue in WEAK_PROCESS_TARGET_CUES)


def _has_weak_activity_target_cue(text: Any) -> bool:
    t = _norm_text(text)
    if not t or t == "nr":
        return False
    return any(_cue_present(t, cue) for cue in WEAK_ACTIVITY_TARGET_CUES)


def _has_strong_activity_target_cue(text: Any) -> bool:
    t = _norm_text(text)
    if not t or t == "nr":
        return False
    return any(_cue_present(t, cue) for cue in STRONG_ACTIVITY_TARGET_CUES)


def _caption_has_methane_or_ch4_context(text: Any) -> bool:
    t = _norm_text(text)
    if not t or t == "nr":
        return False
    if any(_cue_present(t, cue) for cue in WEAK_PROCESS_TARGET_CUES):
        return True
    return bool(re.search(r"\b(?:methane|ch4)\b", t))


def _caption_has_multi_panel_signal(text: Any) -> bool:
    raw = _safe_str(text)
    if not raw:
        return False
    paren_labels = {lab.lower() for lab in re.findall(r"\(([a-f])\)", raw, flags=re.IGNORECASE)}
    if len(paren_labels) >= 2:
        return True
    panel_labels = {lab.lower() for lab in re.findall(r"\bpanel\s*([a-f])\b", raw, flags=re.IGNORECASE)}
    if len(panel_labels) >= 2:
        return True
    bare_labels = {lab.lower() for lab in re.findall(r"(?:^|[\s,;])([a-f])\)", raw, flags=re.IGNORECASE)}
    return len(bare_labels) >= 2


def _non_target_figure_signal_hits(text: Any) -> List[str]:
    t = _norm_text(text)
    if not t or t == "nr":
        return []
    hits: List[str] = []
    for label, cues in NON_TARGET_FIGURE_CUE_GROUPS.items():
        if any(_cue_present(t, cue) for cue in cues):
            hits.append(label)
    return hits


def _hard_non_target_signal_hits(text: Any) -> List[str]:
    return [
        hit for hit in _non_target_figure_signal_hits(text)
        if hit not in SOFT_NON_TARGET_ALLOWABLE_GROUPS
    ]


def _caption_has_likely_activity_target_signal(text: Any) -> bool:
    if _hard_non_target_signal_hits(text):
        return False
    if _has_strong_activity_target_cue(text):
        return True
    return _caption_has_methane_or_ch4_context(text) and _has_weak_activity_target_cue(text)


def _should_hard_skip_non_target_figure(text: Any) -> Tuple[bool, List[str]]:
    raw = _safe_str(text).strip()
    if not raw or _norm_text(raw) == "nr":
        return False, []
    hits = _non_target_figure_signal_hits(raw)
    hard_hits = _hard_non_target_signal_hits(raw)
    explicit_target = _caption_has_explicit_target_metric(raw)
    weak_process_target = _has_weak_process_target_cue(raw)
    weak_activity_target = _caption_has_likely_activity_target_signal(raw)
    multi_panel = _caption_has_multi_panel_signal(raw)

    if explicit_target:
        return False, hits
    if hits:
        if not hard_hits and weak_activity_target and multi_panel:
            keep_reasons = list(dict.fromkeys(hits + ["caption has activity/TOS cue in a multi-panel catalyst-performance figure"]))
            return False, keep_reasons
        return True, list(dict.fromkeys(hits))
    if weak_activity_target:
        return False, ["caption has activity/TOS cue consistent with a catalyst-performance figure"]
    if weak_process_target:
        return False, ["caption has methane-decomposition/process cue"]
    return True, ["caption lacks explicit or likely CH4-target output cue"]


def _caption_looks_target_conversion(text: str) -> bool:
    return _caption_has_explicit_target_metric(text) or _has_target_conversion_cue(text)


def _caption_looks_likely_ch4_target_process(text: Any) -> bool:
    raw = _safe_str(text)
    if _caption_looks_target_conversion(raw):
        return True
    if _caption_has_likely_activity_target_signal(raw):
        hits = _non_target_figure_signal_hits(raw)
        if not hits or _caption_has_multi_panel_signal(raw):
            return True
    return _has_weak_process_target_cue(raw) and not _non_target_figure_signal_hits(raw)


LATER_CYCLE_OR_REUSE_PHRASES = [
    "recycled catalyst",
    "recycled",
    "recycling",
    "recyclability",
    "re-use",
    "reuse",
    "reused",
    "reusability",
    "used catalyst",
    "spent catalyst",
    "regenerated catalyst",
    "regenerated",
    "after regeneration",
    "after reuse",
    "after recycling",
]


LATER_CYCLE_OR_REUSE_PATTERNS = [
    re.compile(r"\b(?:2nd|second|3rd|third|4th|fourth|5th|fifth|6th|sixth|7th|seventh|8th|eighth|9th|ninth|\d+(?:st|nd|rd|th))\s*(?:cycle|cycles|run|runs|pass|passes|use|uses)\b", flags=re.IGNORECASE),
    re.compile(r"\b(?:cycle|cycles|run|runs|pass|passes|use|uses)\s*(?:2|3|4|5|6|7|8|9|[1-9]\d+)\b", flags=re.IGNORECASE),
]


FIRST_CYCLE_OR_FRESH_PATTERNS = [
    re.compile(r"\b(?:1st|first)\s*(?:cycle|run|pass|use)\b", flags=re.IGNORECASE),
    re.compile(r"\b(?:cycle|run|pass|use)\s*1\b", flags=re.IGNORECASE),
    re.compile(r"\bfresh catalyst\b", flags=re.IGNORECASE),
    re.compile(r"\bpristine catalyst\b", flags=re.IGNORECASE),
    re.compile(r"\bnew catalyst\b", flags=re.IGNORECASE),
]


def _looks_like_recycled_or_later_cycle(text: Any) -> bool:
    t = _norm_text(text)
    if not t or t == "nr":
        return False
    if any(phrase in t for phrase in LATER_CYCLE_OR_REUSE_PHRASES):
        return True
    if any(pat.search(t) for pat in LATER_CYCLE_OR_REUSE_PATTERNS):
        return True
    if any(pat.search(t) for pat in FIRST_CYCLE_OR_FRESH_PATTERNS):
        return False
    return False


def _should_drop_recycled_or_later_cycle_record(row: Dict[str, Any]) -> bool:
    texts = [
        _safe_str(row.get("Series_Label", "")),
        _safe_str(row.get("Catalyst_Label", "")),
    ]
    combined = " | ".join([t for t in texts if t and t != "NR"])
    return _looks_like_recycled_or_later_cycle(combined)


def _is_target_conversion_axis(y_axis_label: Any, caption_context: str = "", caption_text: str = "") -> bool:
    if _looks_like_equilibrium_or_max_conversion(y_axis_label):
        return False
    hard_skip, _hits = _should_hard_skip_non_target_figure(caption_text)
    if hard_skip:
        return False
    y = _norm_text(y_axis_label)
    y_compact = y.replace(" ", "")
    if any(tok in y for tok in ["hydrogen", "h2 yield", "h2 production", "hydrogen yield", "hydrogen production"]):
        return False
    if re.search(r"\bh2\b", y):
        return False
    if _is_unreacted_ch4_axis(y):
        return True
    if "xch4" in y_compact:
        return True
    if "methane conversion" in y or "ch4 conversion" in y:
        return True
    caption_signal_text = caption_context if caption_context else caption_text
    if "conversion" in y and _caption_looks_target_conversion(caption_signal_text):
        return True
    if (not y or y == "nr") and _caption_looks_target_conversion(caption_signal_text):
        return True
    return False



def _parse_expected_count(value: Any) -> Optional[int]:
    if value in (None, "", "NR"):
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        try:
            n = int(round(float(value)))
            return n if n > 0 else None
        except Exception:
            return None
    m = re.search(r"\d+", _safe_str(value))
    if not m:
        return None
    try:
        n = int(m.group(0))
        return n if n > 0 else None
    except Exception:
        return None


def _normalize_panel_series_specs(discovered: Any) -> List[Dict[str, Any]]:
    """Normalize panel/legend/series discovery output into a stable list.

    Expected modern shape:
      {"panels":[{"panel_id":"a", "x_axis_label":"...", "y_axis_label":"...",
                   "legend_field":"temperature", "series":[{"label":"650", "field_value":"650 °C", "point_count":6}]}]}
    Older shape with top-level series_labels is also accepted.
    """
    if not isinstance(discovered, dict):
        return []
    panels_raw = discovered.get("panels")
    if not isinstance(panels_raw, list):
        if discovered.get("series_labels"):
            panels_raw = [{
                "panel_id": discovered.get("panel_id", "panel 1"),
                "x_axis_label": discovered.get("x_axis_label", "NR"),
                "y_axis_label": discovered.get("y_axis_label", "NR"),
                "legend_field": discovered.get("legend_field", "NR"),
                "series": [
                    {"label": s, "field_value": s, "point_count": (discovered.get("series_point_counts", {}) or {}).get(s, "NR")}
                    for s in (discovered.get("series_labels") or [])
                ],
                "panel_context_fields": discovered.get("panel_context_fields", {}),
            }]
        else:
            return []
    panels: List[Dict[str, Any]] = []
    for i, p in enumerate(panels_raw, start=1):
        if not isinstance(p, dict):
            continue
        panel_id = _display_norm(p.get("panel_id", p.get("panel_label", f"panel {i}")))
        x_label = _display_norm(p.get("x_axis_label", "NR"))
        y_label = _display_norm(p.get("y_axis_label", "NR"))
        legend_field = _display_norm(p.get("legend_field", "NR"))
        if not _is_target_conversion_axis(y_label, _safe_str(p), _safe_str(p)):
            # Keep panels whose y label is unclear only if the global discovery says target and labels mention conversion.
            if "conversion" not in _norm_text(y_label) and "xch4" not in _norm_text(y_label):
                continue
        series_raw = p.get("series", p.get("series_labels", []))
        series_out: List[Dict[str, Any]] = []
        if isinstance(series_raw, dict):
            series_raw = [{"label": k, **(v if isinstance(v, dict) else {"field_value": v})} for k, v in series_raw.items()]
        if isinstance(series_raw, list):
            for j, s in enumerate(series_raw, start=1):
                if isinstance(s, dict):
                    label = _display_norm(s.get("label", s.get("series_label", s.get("name", f"curve {j}"))))
                    field_value = _display_norm(s.get("field_value", s.get("value", label)))
                    count = _parse_expected_count(s.get("point_count", s.get("marker_count", s.get("datapoint_count", s.get("n_points", "NR")))))
                    marker_color = _display_norm(s.get("marker_color", s.get("color", "NR")))
                    marker_shape = _display_norm(s.get("marker_shape", s.get("shape", "NR")))
                    line_style = _display_norm(s.get("line_style", "NR"))
                    extraction_mode = _display_norm(s.get("extraction_mode", "visible markers"))
                else:
                    label = _display_norm(s)
                    field_value = label
                    count = None
                    marker_color = marker_shape = line_style = "NR"
                    extraction_mode = "visible markers"
                if label in ("", "NR"):
                    label = f"curve {j}"
                if _looks_like_recycled_or_later_cycle(label):
                    continue
                series_out.append({
                    "label": label,
                    "field_value": field_value,
                    "point_count": count if count is not None else "NR",
                    "marker_color": marker_color,
                    "marker_shape": marker_shape,
                    "line_style": line_style,
                    "extraction_mode": extraction_mode,
                })
        if not series_out:
            # A single visible curve with no legend should still be extractable.
            series_out = [{
                "label": "curve 1",
                "field_value": "NR",
                "point_count": _parse_expected_count(p.get("point_count", p.get("marker_count", "NR"))) or "NR",
                "marker_color": "NR",
                "marker_shape": "NR",
                "line_style": "NR",
                "extraction_mode": "visible markers or 10 evenly spaced line points",
            }]
        panels.append({
            "panel_id": panel_id,
            "x_axis_label": x_label,
            "y_axis_label": y_label,
            "legend_field": legend_field,
            "series": series_out,
            "panel_context_fields": p.get("panel_context_fields", p.get("context_fields", {})) if isinstance(p.get("panel_context_fields", p.get("context_fields", {})), dict) else {},
            "notes": _display_norm(p.get("notes", "NR")),
        })
    return panels


def _flatten_expected_series_from_panel_specs(panel_specs: List[Dict[str, Any]]) -> List[str]:
    labels: List[str] = []
    for p in panel_specs or []:
        for s in p.get("series", []) or []:
            label = _display_norm(s.get("label", "NR")) if isinstance(s, dict) else _display_norm(s)
            if label != "NR" and label not in labels:
                labels.append(label)
    return labels


def _expected_counts_from_panel_specs(panel_specs: List[Dict[str, Any]]) -> Dict[str, Any]:
    counts: Dict[str, Any] = {}
    for p in panel_specs or []:
        panel_id = _display_norm(p.get("panel_id", "NR"))
        for s in p.get("series", []) or []:
            if not isinstance(s, dict):
                continue
            label = _display_norm(s.get("label", "NR"))
            count = _parse_expected_count(s.get("point_count", "NR"))
            if label != "NR" and count is not None:
                counts[f"{panel_id}::{label}"] = count
                counts[label] = count
    return counts


def _record_count_meets_expected(records: List[Dict[str, Any]], expected_count: Optional[int]) -> bool:
    if expected_count is None:
        return True
    return len(records or []) >= expected_count




def _safe_json_compact(obj: Any) -> str:
    try:
        return json.dumps(obj, ensure_ascii=False, separators=(",", ":"))
    except Exception:
        return _safe_str(obj)


def _series_labels_from_panel_specs_for_audit(panel_specs: List[Dict[str, Any]]) -> List[str]:
    labels: List[str] = []
    for panel in panel_specs or []:
        if not isinstance(panel, dict):
            continue
        for series in panel.get("series", []) or []:
            if isinstance(series, dict):
                label = _display_norm(series.get("label", "NR"))
            else:
                label = _display_norm(series)
            if label != "NR" and label not in labels:
                labels.append(label)
    return labels


def _panel_ids_from_panel_specs_for_audit(panel_specs: List[Dict[str, Any]]) -> List[str]:
    ids: List[str] = []
    for i, panel in enumerate(panel_specs or [], start=1):
        if not isinstance(panel, dict):
            continue
        pid = _display_norm(panel.get("panel_id", f"panel {i}"))
        if pid != "NR" and pid not in ids:
            ids.append(pid)
    return ids


def _record_panel_discovery_audit(
    *,
    paper: Any,
    fig_filename: str,
    matched_label: Any,
    matched_order: Any,
    match_method: Any,
    caption_text: Any,
    caption_context: Any,
    discovered: Any,
    panel_specs: List[Dict[str, Any]],
    status: str,
    error: Any = "NR",
) -> None:
    """Store raw panel/series discovery information for post-run audit reports."""
    paper_name = _display_norm(paper)
    if not paper_name or paper_name == "NR":
        paper_name = _display_norm(CURRENT_PAPER_FOR_USAGE)
    if not paper_name or paper_name == "NR":
        paper_name = "UNKNOWN_PAPER"
    if not isinstance(panel_specs, list):
        panel_specs = []
    if not isinstance(discovered, dict):
        discovered = {}

    series_labels = _series_labels_from_panel_specs_for_audit(panel_specs)
    panel_ids = _panel_ids_from_panel_specs_for_audit(panel_specs)
    panel_series_counts: List[str] = []
    for i, panel in enumerate(panel_specs, start=1):
        if not isinstance(panel, dict):
            continue
        pid = _display_norm(panel.get("panel_id", f"panel {i}"))
        series_count = len([x for x in (panel.get("series", []) or []) if isinstance(x, dict)])
        panel_series_counts.append(f"{pid}: {series_count}")

    row = {
        "Paper": paper_name,
        "Figure": _display_norm(fig_filename),
        "Matched_XML_Figure_Label": _display_norm(matched_label),
        "Matched_XML_Figure_Order": _display_norm(matched_order),
        "Caption_Match_Method": _display_norm(match_method),
        "LLMProcessed": "yes",
        "Discovery_Status": _display_norm(status),
        "Has_Multiple_Panels": _display_norm(discovered.get("has_multiple_panels", "NR")),
        "All_Panel_Count": _display_norm(discovered.get("all_panel_count", "NR")),
        "Relevant_Panel_Count": len(panel_specs),
        "Relevant_Panel_IDs": " | ".join(panel_ids) if panel_ids else "NR",
        "Discovered_Series_Count": len(series_labels),
        "Discovered_Series_Labels": " | ".join(series_labels) if series_labels else "NR",
        "Panel_Series_Counts": " | ".join(panel_series_counts) if panel_series_counts else "NR",
        "Is_Target_Conversion_Figure": _display_norm(discovered.get("is_target_conversion_figure", "NR")),
        "Discovery_Notes": _display_norm(discovered.get("notes", "NR")),
        "Discovery_Error": _display_norm(error),
        "Caption": _display_norm(caption_text),
        "Caption_Context_Chars": len(_safe_str(caption_context)),
        "Panel_Specs_JSON": _safe_json_compact(panel_specs),
        "Discovery_JSON": _safe_json_compact(discovered),
    }
    existing = PANEL_DISCOVERY_AUDIT_BY_PAPER.setdefault(paper_name, [])
    fig_norm = _norm_text(fig_filename)
    existing[:] = [r for r in existing if _norm_text(r.get("Figure", "")) != fig_norm]
    existing.append(row)


def _discover_relevant_panels_and_series(
    image_path: str,
    fig_filename: str,
    matched_xml_label: str,
    matched_xml_order: Any,
    caption_context: str,
) -> Dict[str, Any]:
    content = [
        encode_image_for_converse(image_path),
        {"text": _panel_series_discovery_user_prompt(
            fig_filename=fig_filename,
            matched_xml_label=matched_xml_label,
            matched_xml_order=matched_xml_order,
            caption_context=caption_context,
        )},
    ]
    msgs = [{"role": "user", "content": content}]
    raw = _converse(msgs, system_prompt=FIGURE_SYSTEM_PROMPT, max_tokens=2200)
    obj = _extract_first_json_value(raw)
    if not isinstance(obj, dict):
        return {"panels": [], "is_target_conversion_figure": False, "notes": "Expected JSON object."}
    obj["panels"] = _normalize_panel_series_specs(obj)
    if obj.get("panels"):
        obj["is_target_conversion_figure"] = True
    return obj

def _needs_series_recovery(records: List[Dict[str, Any]], caption_context: str) -> bool:
    caption_target = _caption_looks_target_conversion(caption_context)
    likely_target_process = _caption_looks_likely_ch4_target_process(caption_context)
    if not records:
        return likely_target_process
    x_kind = _x_axis_kind(records[0].get("X_Axis_Label", ""))
    y_text = _norm_text(records[0].get("Y_Axis_Label", ""))
    target_y = ("conversion" in y_text) or ("methane" in y_text) or ("ch4" in y_text) or ("xch4" in y_text)
    target_x = x_kind in {"time", "temperature", "ch4_flow", "total_flow", "space_velocity", "catalyst_mass", "pressure"}
    unique_series = {
        _series_group_key(r) for r in records
        if _series_group_key(r) not in ("", "NR")
    }
    if (likely_target_process or (target_y and target_x)) and len(records) <= 2:
        return True
    if (caption_target or (target_y and target_x)) and len(unique_series) <= 1:
        return True
    return False


def _dedupe_records_simple(records: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    seen = set()
    out: List[Dict[str, Any]] = []
    for r in records:
        key = (
            _safe_str(r.get("Catalyst_Label", "NR")).strip(),
            _safe_str(r.get("Series_Label", "NR")).strip(),
            round(_x_axis_numeric_value(r), 6) if _x_axis_numeric_value(r) is not None else None,
            round(_parse_first_float(r.get("CH4_Conversion_%", "NR")), 6) if _parse_first_float(r.get("CH4_Conversion_%", "NR")) is not None else None,
            _safe_str(r.get("X_Axis_Label", "NR")).strip(),
            _safe_str(r.get("Y_Axis_Label", "NR")).strip(),
        )
        if key in seen:
            continue
        seen.add(key)
        out.append(r)
    return out


def _series_recovery_still_needed(records: List[Dict[str, Any]], expected_series_labels: List[str]) -> bool:
    if not expected_series_labels:
        return False
    found = {
        _norm_text(_series_group_key(r)) for r in records
        if _series_group_key(r) not in ("", "NR")
    }
    expected = {_norm_text(x) for x in expected_series_labels if _safe_str(x).strip()}
    if not records:
        return True
    if len(records) <= 2 and len(expected) >= 2:
        return True
    if expected and not expected.issubset(found):
        return True
    return False


def encode_image_for_converse(
    img_path: str,
    soft_max_side: int = SOFT_IMAGE_MAX_SIDE,
    soft_max_bytes: int = SOFT_IMAGE_MAX_BYTES,
    jpeg_quality: int = SOFT_IMAGE_JPEG_QUALITY,
) -> dict:
    from PIL import Image as PILImage
    import io

    img_path_obj = Path(img_path)
    orig_bytes = img_path_obj.read_bytes()
    orig_kb = len(orig_bytes) // 1024
    ext = img_path_obj.suffix.lower().lstrip(".")
    if ext == "jpg":
        ext = "jpeg"

    with PILImage.open(img_path) as im:
        width, height = im.size
        if (
            ext in {"jpeg", "png", "gif", "webp"}
            and len(orig_bytes) <= soft_max_bytes
            and max(width, height) <= soft_max_side
        ):
            print(f"    [img] {img_path_obj.name}: {orig_kb} KB -> {orig_kb} KB ({width}x{height} px, kept original)")
            return {"image": {"format": ext, "source": {"bytes": orig_bytes}}}

        if im.mode in ("RGBA", "LA", "P"):
            bg = PILImage.new("RGB", im.size, (255, 255, 255))
            alpha = im.convert("RGBA").split()[3]
            bg.paste(im.convert("RGBA"), mask=alpha)
            im = bg
        elif im.mode != "RGB":
            im = im.convert("RGB")

        w, h = im.size
        if max(w, h) > soft_max_side:
            scale = soft_max_side / max(w, h)
            im = im.resize((max(1, int(w * scale)), max(1, int(h * scale))), PILImage.LANCZOS)

        quality = jpeg_quality
        data = b""
        while True:
            buf = io.BytesIO()
            im.save(buf, format="JPEG", quality=quality, optimize=False)
            data = buf.getvalue()
            if len(data) <= BEDROCK_IMAGE_MAX_BYTES and max(im.size) <= BEDROCK_IMAGE_MAX_SIDE:
                break
            if quality > 78:
                quality -= 5
                continue
            scale = 0.9
            new_w = max(1, int(im.size[0] * scale))
            new_h = max(1, int(im.size[1] * scale))
            if new_w == im.size[0] and new_h == im.size[1]:
                break
            im = im.resize((new_w, new_h), PILImage.LANCZOS)

    compr_kb = len(data) // 1024
    print(f"    [img] {img_path_obj.name}: {orig_kb} KB -> {compr_kb} KB ({im.size[0]}x{im.size[1]} px)")
    return {"image": {"format": "jpeg", "source": {"bytes": data}}}




def _set_openai_settings(
    model_id: Optional[str] = None,
    base_url: Optional[str] = None,
    reasoning_effort: Optional[str] = None,
    text_verbosity: Optional[str] = None,
) -> None:
    """Configure OpenAI model/runtime settings and rebuild the client if needed."""
    global MODEL_ID, OPENAI_BASE_URL, OPENAI_REASONING_EFFORT, OPENAI_TEXT_VERBOSITY, openai_client
    if model_id:
        MODEL_ID = str(model_id).strip()
    if base_url:
        OPENAI_BASE_URL = str(base_url).strip().rstrip("/")
    if reasoning_effort:
        OPENAI_REASONING_EFFORT = str(reasoning_effort).strip()
    if text_verbosity:
        OPENAI_TEXT_VERBOSITY = str(text_verbosity).strip()
    openai_client = _build_openai_client()




def _set_figure_extraction_settings(
    figure_max_output_tokens: Optional[int] = None,
    figure_schema_mode: Optional[str] = None,
) -> None:
    """Configure figure extraction output budget and schema mode."""
    global FIGURE_MAX_OUTPUT_TOKENS, FIGURE_SCHEMA_MODE
    if figure_max_output_tokens is not None:
        FIGURE_MAX_OUTPUT_TOKENS = max(1, int(figure_max_output_tokens))
    if figure_schema_mode:
        mode = str(figure_schema_mode).strip().lower()
        if mode not in {"compact", "full"}:
            raise ValueError("figure_schema_mode must be 'compact' or 'full'")
        FIGURE_SCHEMA_MODE = mode

def _set_llm_rate_limit_settings(
    call_delay_seconds: Optional[float] = None,
    throttle_max_retries: Optional[int] = None,
    throttle_base_delay_seconds: Optional[float] = None,
    throttle_max_delay_seconds: Optional[float] = None,
    throttle_jitter_seconds: Optional[float] = None,
) -> None:
    """Configure application-level pacing and throttling retry behavior."""
    global LLM_CALL_DELAY_SECONDS
    global LLM_THROTTLE_MAX_RETRIES
    global LLM_THROTTLE_BASE_DELAY_SECONDS
    global LLM_THROTTLE_MAX_DELAY_SECONDS
    global LLM_THROTTLE_JITTER_SECONDS

    if call_delay_seconds is not None:
        LLM_CALL_DELAY_SECONDS = max(0.0, float(call_delay_seconds))
    if throttle_max_retries is not None:
        LLM_THROTTLE_MAX_RETRIES = max(0, int(throttle_max_retries))
    if throttle_base_delay_seconds is not None:
        LLM_THROTTLE_BASE_DELAY_SECONDS = max(0.0, float(throttle_base_delay_seconds))
    if throttle_max_delay_seconds is not None:
        LLM_THROTTLE_MAX_DELAY_SECONDS = max(0.0, float(throttle_max_delay_seconds))
    if throttle_jitter_seconds is not None:
        LLM_THROTTLE_JITTER_SECONDS = max(0.0, float(throttle_jitter_seconds))


class RateLimitReachedError(RuntimeError):
    """Raised when OpenAI rate/quota errors persist after application-level retries."""


CURRENT_PROGRESS_PATH_FOR_RATE_LIMIT: Optional[str] = None
CURRENT_COMPLETED_PAPERS_FOR_RATE_LIMIT: Optional[List[str]] = None
CURRENT_ALL_ROWS_FOR_RATE_LIMIT: Optional[List[Dict[str, Any]]] = None


def _register_progress_state_for_rate_limit(progress_path: str, completed_papers: List[str], all_rows: List[Dict[str, Any]]) -> None:
    global CURRENT_PROGRESS_PATH_FOR_RATE_LIMIT
    global CURRENT_COMPLETED_PAPERS_FOR_RATE_LIMIT
    global CURRENT_ALL_ROWS_FOR_RATE_LIMIT
    CURRENT_PROGRESS_PATH_FOR_RATE_LIMIT = progress_path
    CURRENT_COMPLETED_PAPERS_FOR_RATE_LIMIT = completed_papers
    CURRENT_ALL_ROWS_FOR_RATE_LIMIT = all_rows


def _save_progress_on_rate_limit(reason: str) -> None:
    if not CURRENT_PROGRESS_PATH_FOR_RATE_LIMIT or CURRENT_COMPLETED_PAPERS_FOR_RATE_LIMIT is None or CURRENT_ALL_ROWS_FOR_RATE_LIMIT is None:
        return
    try:
        _finalize_output_rows(CURRENT_ALL_ROWS_FOR_RATE_LIMIT)
    except Exception:
        pass
    _save_progress(CURRENT_PROGRESS_PATH_FOR_RATE_LIMIT, CURRENT_COMPLETED_PAPERS_FOR_RATE_LIMIT, CURRENT_ALL_ROWS_FOR_RATE_LIMIT)
    print(f"      [rate-limit] Progress saved to {CURRENT_PROGRESS_PATH_FOR_RATE_LIMIT}. Current in-progress paper is not marked complete. Reason: {reason}")


def _is_throttling_exception(exc: BaseException) -> bool:
    """Return True when an exception is an OpenAI rate-limit/quota error."""
    if isinstance(exc, RateLimitReachedError):
        return True
    status_code = getattr(exc, "status_code", None)
    if status_code == 429:
        return True
    code = _safe_str(getattr(exc, "code", "")).lower()
    if code in {"rate_limit_exceeded", "insufficient_quota", "tokens_per_day_exceeded", "requests_per_day_exceeded"}:
        return True
    response = getattr(exc, "response", None)
    if isinstance(response, dict):
        err = response.get("error", {}) or response.get("Error", {}) or {}
        err_code = _safe_str(err.get("code", err.get("Code", ""))).lower()
        if err_code in {"rate_limit_exceeded", "insufficient_quota", "too_many_requests"}:
            return True
    text = _safe_str(exc).lower()
    return any(
        token in text
        for token in [
            "rate limit", "rate_limit", "too many requests", "429", "insufficient quota",
            "quota", "tokens per day", "requests per day", "per day", "exceeded your current quota",
        ]
    )


def _throttle_sleep_seconds(retry_index: int) -> float:
    """Exponential backoff with jitter for an application-level retry."""
    base = LLM_THROTTLE_BASE_DELAY_SECONDS * (2 ** max(0, retry_index - 1))
    wait = min(base, LLM_THROTTLE_MAX_DELAY_SECONDS)
    if LLM_THROTTLE_JITTER_SECONDS > 0:
        wait += random.uniform(0.0, LLM_THROTTLE_JITTER_SECONDS)
    return max(0.0, wait)


def _openai_usage_payload(response: Any) -> Dict[str, Any]:
    usage = getattr(response, "usage", None)
    if usage is None:
        return {"usage": {"inputTokens": 0, "outputTokens": 0}}
    input_tokens = int(getattr(usage, "input_tokens", 0) or 0)
    output_tokens = int(getattr(usage, "output_tokens", 0) or 0)
    return {"usage": {"inputTokens": input_tokens, "outputTokens": output_tokens}}


def _bedrock_messages_to_openai_input(messages: list) -> list:
    converted = []
    for msg in messages:
        content_out = []
        for item in msg.get("content", []):
            if isinstance(item, dict) and "text" in item:
                content_out.append({"type": "input_text", "text": _safe_str(item.get("text", ""))})
            elif isinstance(item, dict) and "image" in item:
                image = item.get("image", {}) or {}
                fmt = _safe_str(image.get("format", "jpeg")).lower() or "jpeg"
                if fmt == "jpg":
                    fmt = "jpeg"
                data = image.get("source", {}).get("bytes", b"")
                if not isinstance(data, (bytes, bytearray)):
                    data = bytes(data)
                mime = f"image/{fmt}"
                b64 = base64.b64encode(data).decode("ascii")
                content_out.append({"type": "input_image", "image_url": f"data:{mime};base64,{b64}"})
        converted.append({"role": msg.get("role", "user"), "content": content_out})
    return converted


def _response_output_text(response: Any) -> str:
    direct = getattr(response, "output_text", None)
    if isinstance(direct, str) and direct.strip():
        return direct.strip()
    parts: List[str] = []
    for item in getattr(response, "output", []) or []:
        for content in getattr(item, "content", []) or []:
            text = getattr(content, "text", None)
            if isinstance(text, str):
                parts.append(text)
    return "\n".join(parts).strip()


def _openai_converse(messages: list, system_prompt: str = TEXT_SYSTEM_PROMPT, max_tokens: int = 3200, temperature: float = 0.0) -> str:
    if openai_client is None:
        raise RuntimeError(
            "OpenAI client is not configured. Set OPENAI_API_KEY in your .env or environment. "
            "Optionally set OPENAI_BASE_URL=https://api.openai.com/v1 and OPENAI_MODEL_ID=gpt-5.5."
        )
    resolved_system_prompt = _resolve_prompt_reference(system_prompt)
    total_attempts = max(1, LLM_THROTTLE_MAX_RETRIES + 1)
    last_exc: Optional[BaseException] = None
    openai_input = _bedrock_messages_to_openai_input(messages)

    for attempt in range(1, total_attempts + 1):
        if LLM_CALL_DELAY_SECONDS > 0:
            time.sleep(LLM_CALL_DELAY_SECONDS)
        try:
            kwargs = {
                "model": MODEL_ID,
                "instructions": resolved_system_prompt,
                "input": openai_input,
                "max_output_tokens": max_tokens,
                "store": OPENAI_STORE_RESPONSES,
            }
            if OPENAI_REASONING_EFFORT and OPENAI_REASONING_EFFORT.lower() not in {"", "none", "off", "false"}:
                kwargs["reasoning"] = {"effort": OPENAI_REASONING_EFFORT}
            if OPENAI_TEXT_VERBOSITY and OPENAI_TEXT_VERBOSITY.lower() not in {"", "default", "none"}:
                kwargs["text"] = {"verbosity": OPENAI_TEXT_VERBOSITY}
            # Temperature is retained for compatibility with non-reasoning models, but
            # omitted for GPT-5 series by default because reasoning controls are preferred.
            if os.environ.get("OPENAI_USE_TEMPERATURE", "0").strip().lower() in {"1", "true", "yes"}:
                kwargs["temperature"] = temperature
            response = openai_client.responses.create(**kwargs)
            _record_llm_usage(_openai_usage_payload(response), messages, resolved_system_prompt)
            status = _safe_str(getattr(response, "status", "")).lower()
            incomplete_details = getattr(response, "incomplete_details", None)
            if status == "incomplete" or incomplete_details is not None:
                print(f"      [warning] OpenAI response incomplete/status={status or 'unknown'}; details={incomplete_details}; parser salvage will be used if possible")
            text = _response_output_text(response)
            if not text:
                raise RuntimeError("OpenAI response contained no output text.")
            return text
        except Exception as exc:
            if not _is_throttling_exception(exc):
                raise
            last_exc = exc
            if attempt >= total_attempts:
                break
            retry_index = attempt
            sleep_s = _throttle_sleep_seconds(retry_index)
            print(
                f"      [throttle] OpenAI rate/quota limit; sleeping {sleep_s:.1f}s "
                f"before application retry {retry_index}/{LLM_THROTTLE_MAX_RETRIES}"
            )
            time.sleep(sleep_s)

    reason = _safe_str(last_exc) if last_exc is not None else "unknown OpenAI rate/quota error"
    _save_progress_on_rate_limit(reason)
    raise RateLimitReachedError(f"OpenAI rate/quota limit reached after {total_attempts} attempt(s): {reason}") from last_exc

def _raise_bedrock_rate_limit(reason: str) -> None:
    _save_progress_on_rate_limit(reason)
    raise RateLimitReachedError(f"Bedrock throttling/rate/quota limit reached: {reason}")


def _converse(messages: list, system_prompt: str = TEXT_SYSTEM_PROMPT, max_tokens: int = 3200, temperature: float = 0.0) -> str:
    if selected_provider() == "openai":
        return _openai_converse(messages, system_prompt=system_prompt, max_tokens=max_tokens, temperature=temperature)
    resolved_system_prompt = _resolve_prompt_reference(system_prompt)
    return bedrock_converse(
        messages=messages, system_prompt=resolved_system_prompt, max_tokens=max_tokens, temperature=temperature,
        call_delay_seconds=LLM_CALL_DELAY_SECONDS, max_retries=LLM_THROTTLE_MAX_RETRIES,
        base_delay_seconds=LLM_THROTTLE_BASE_DELAY_SECONDS, max_delay_seconds=LLM_THROTTLE_MAX_DELAY_SECONDS,
        jitter_seconds=LLM_THROTTLE_JITTER_SECONDS, is_throttling=_is_throttling_exception,
        record_usage=lambda response: _record_llm_usage(response, messages, resolved_system_prompt),
        raise_rate_limit=_raise_bedrock_rate_limit,
    )


def _chunk_text(text: str, max_chars: int = 40_000) -> List[str]:
    return [text[i:i + max_chars] for i in range(0, len(text), max_chars)]


def _merge_constant_dicts(dicts: List[Dict[str, Any]]) -> Dict[str, Any]:
    merged: Dict[str, Any] = {k: "NR" for k in TEXT_CONSTANT_FIELDS}
    notes: List[str] = []
    temp_values_seen: List[str] = []
    heating_values_seen: List[str] = []
    catalyst_values_seen: List[str] = []

    for d in dicts:
        for k in TEXT_CONSTANT_FIELDS:
            new_val = normalize_blank_to_NR(d.get(k, "NR"))
            old_val = normalize_blank_to_NR(merged.get(k, "NR"))

            if k == INTERNAL_NOTES_KEY:
                merged[k] = _merge_notes_uniquely(old_val, new_val)
                continue

            if k == "Title":
                merged[k] = _choose_better_title(_safe_str(old_val), _safe_str(new_val))
                continue

            if k == "Catalyst_Label":
                cand = _canonicalize_catalyst_label_with_context(_safe_str(new_val), _safe_str(new_val)) if new_val != "NR" else "NR"
                if old_val == "NR" and cand != "NR":
                    merged[k] = cand
                elif old_val != "NR" and cand != "NR" and _norm_text(old_val) != _norm_text(cand):
                    merged[k] = old_val if len(_safe_str(old_val)) >= len(_safe_str(cand)) else cand
                    notes.append(f"Conflict on {k}: kept '{merged[k]}'.")
                continue

            if k == "Heating_Mode":
                if old_val == "NR" and new_val != "NR":
                    merged[k] = _standardize_heating_mode(new_val)
                elif new_val != "NR":
                    cand = _standardize_heating_mode(new_val)
                    if cand != "NR":
                        merged[k] = cand if old_val == "NR" else merged[k]
                continue

            if k == "Precursor_Salts":
                if old_val == "NR" and new_val != "NR":
                    merged[k] = _normalize_precursor_salts(new_val)
                elif old_val != "NR" and new_val != "NR":
                    merged[k] = _normalize_precursor_salts(" + ".join([_safe_str(old_val), _safe_str(new_val)]))
                continue

            if k == "Feed_Gas_Mixture":
                if old_val == "NR" and new_val != "NR":
                    merged[k] = _standardize_feed_gas_mixture(new_val)
                elif old_val != "NR" and new_val != "NR" and _norm_text(old_val) != _norm_text(new_val):
                    chosen = old_val if len(_extract_feed_gas_mixture_candidates(old_val)) >= len(_extract_feed_gas_mixture_candidates(new_val)) else new_val
                    merged[k] = _standardize_feed_gas_mixture(chosen)
                continue

            if old_val == "NR" and new_val != "NR":
                merged[k] = new_val
            elif old_val != "NR" and new_val != "NR" and _safe_str(new_val).strip() != _safe_str(old_val).strip():
                chosen = new_val if len(_safe_str(new_val)) > len(_safe_str(old_val)) else old_val
                merged[k] = chosen
                if k != "DOI":
                    notes.append(f"Conflict on {k}: kept '{chosen}'.")

        temps = _extract_temperature_tokens(d.get("Temperature_C_or_K", "NR"))
        temp_values_seen.extend(temps)
        hm = _standardize_heating_mode(d.get("Heating_Mode", "NR"))
        if hm != "NR":
            heating_values_seen.append(hm)
        catalyst_val = _canonicalize_catalyst_label_with_context(_safe_str(d.get("Catalyst_Label", "NR")), _safe_str(d.get("Catalyst_Label", "NR")))
        if catalyst_val != "NR" and parse_catalyst_label(catalyst_val)[0]:
            catalyst_values_seen.append(catalyst_val)

    unique_temps = sorted(set(temp_values_seen))
    if len(unique_temps) > 1:
        notes.append(
            f"Multiple distinct temperatures found in HTML/article text ({', '.join(unique_temps)}). "
            f"Set Temperature_C_or_K = NR to force figure-level resolution."
        )
        merged["Temperature_C_or_K"] = "NR"

    unique_heating = sorted(set(heating_values_seen))
    unique_catalysts = sorted(set(catalyst_values_seen))
    if len(unique_catalysts) > 1:
        notes.append(
            f"Multiple distinct catalyst labels found in HTML/article text ({', '.join(unique_catalysts)}). "
            f"Set Catalyst_Label = NR to force figure-level resolution."
        )
        merged["Catalyst_Label"] = "NR"
    elif len(unique_catalysts) == 1 and merged.get("Catalyst_Label", "NR") == "NR":
        merged["Catalyst_Label"] = unique_catalysts[0]

    if len(unique_heating) > 1:
        notes.append(
            f"Multiple heating modes found in HTML/article text ({', '.join(unique_heating)}). "
            f"Set Heating_Mode = NR to force figure-level resolution."
        )
        merged["Heating_Mode"] = "NR"
    elif len(unique_heating) == 1 and merged.get("Heating_Mode", "NR") == "NR":
        merged["Heating_Mode"] = unique_heating[0]

    if notes:
        merged[INTERNAL_NOTES_KEY] = _merge_notes_uniquely(merged.get(INTERNAL_NOTES_KEY, "NR"), " ".join(notes))

    return merged


def extract_paper_constants_from_xml(parsed_xml: Dict[str, Any]) -> Dict[str, Any]:
    full_text = parsed_xml.get("full_text", "") or ""
    prompt_text = _compact_text_for_prompt(full_text)
    chunks = _chunk_text(prompt_text, max_chars=40_000)
    partial_dicts: List[Dict[str, Any]] = []

    for i, chunk in enumerate(chunks, start=1):
        print(f"    [pass 1] constants from article-text chunk {i}/{len(chunks)} ({len(chunk):,} chars)")
        prompt = _constants_extraction_user_prompt(i, len(chunks))
        msgs = [{
            "role": "user",
            "content": [{"text": prompt + "\n\nTEXT_CHUNK:\n" + chunk}],
        }]

        raw = _converse(msgs, system_prompt=TEXT_SYSTEM_PROMPT, max_tokens=1400)
        try:
            obj = _extract_first_json_value(raw)
            if not isinstance(obj, dict):
                raise ValueError("Expected JSON object for constants.")
            cleaned = {k: obj.get(k, "NR") for k in TEXT_CONSTANT_FIELDS}
            partial_dicts.append(cleaned)
        except Exception as e:
            if isinstance(e, RateLimitReachedError):
                raise
            print(f"      [warning] constants parse failed for article-text chunk {i}: {e}")
            partial_dicts.append({k: "NR" for k in TEXT_CONSTANT_FIELDS})

    merged = _merge_constant_dicts(partial_dicts)
    merged.setdefault(INTERNAL_NOTES_KEY, "NR")
    for k in TEXT_CONSTANT_FIELDS:
        merged.setdefault(k, "NR")

    xml_title = _display_norm(parsed_xml.get("title", "NR"))
    xml_doi = _display_norm(parsed_xml.get("doi", "NR"))
    merged["Title"] = xml_title if xml_title != "NR" else merged.get("Title", "NR")
    merged["DOI"] = xml_doi if xml_doi != "NR" else merged.get("DOI", "NR")

    if merged.get("Catalyst_Label", "NR") != "NR":
        merged["Catalyst_Label"] = _canonicalize_catalyst_label_with_context(
            merged.get("Catalyst_Label", "NR"),
            " ".join([_safe_str(parsed_xml.get("title", "")), full_text]),
        )
    if merged.get("Catalyst_Label", "NR") == "NR":
        inferred_catalyst = _infer_catalyst_label_from_text(
            parsed_xml.get("title", ""),
            full_text,
            merged.get("Precursor_Salts", "NR"),
        )
        if inferred_catalyst != "NR":
            merged["Catalyst_Label"] = inferred_catalyst
            merged[INTERNAL_NOTES_KEY] = _append_note(
                merged.get(INTERNAL_NOTES_KEY, "NR"),
                f"Catalyst_Label inferred from article text: {inferred_catalyst}.",
            )

    if merged.get("Temperature_C_or_K", "NR") == "NR":
        temp, snippet = _extract_single_temperature_with_snippet(full_text)
        if temp:
            merged["Temperature_C_or_K"] = temp
            merged[INTERNAL_NOTES_KEY] = _append_note(
                merged.get(INTERNAL_NOTES_KEY, "NR"),
                f"Temperature_C_or_K inferred from article text: {temp}. Snippet: {snippet}",
            )

    if merged.get("Heating_Mode", "NR") == "NR":
        inferred_mode = _infer_heating_mode(full_text)
        if inferred_mode != "NR":
            merged["Heating_Mode"] = inferred_mode
            merged[INTERNAL_NOTES_KEY] = _append_note(
                merged.get(INTERNAL_NOTES_KEY, "NR"),
                f"Heating mode inferred from article text: {inferred_mode}.",
            )

    if merged.get("Precursor_Salts", "NR") == "NR":
        precursor_guess = _extract_precursor_salts_from_text(full_text)
        if precursor_guess:
            merged["Precursor_Salts"] = precursor_guess
            merged[INTERNAL_NOTES_KEY] = _append_note(
                merged.get(INTERNAL_NOTES_KEY, "NR"),
                f"Precursor_Salts inferred from HTML/article text: {precursor_guess}.",
            )
    else:
        merged["Precursor_Salts"] = _normalize_precursor_salts(merged.get("Precursor_Salts", "NR"))

    if merged.get("Catalyst_Label", "NR") not in ("NR", None, ""):
        merged["Catalyst_Label"] = _canonicalize_catalyst_label_with_context(
            merged.get("Catalyst_Label", "NR"),
            " ".join([_safe_str(parsed_xml.get("title", "")), full_text]),
        )

    if merged.get("Reaction_Time_hr_min_s", "NR") not in ("NR", None, ""):
        merged["Reaction_Time_hr_min_s"] = _normalize_reaction_time_field(merged.get("Reaction_Time_hr_min_s", "NR"))

    for fld in ["Temperature_C_or_K", "Drying_Temperature_C_or_K", "Calcination_Temperature_C_or_K", "Reduction_Temperature_C_or_K"]:
        if merged.get(fld, "NR") not in ("NR", None, ""):
            merged[fld] = _normalize_multistep_temperature_field(_strip_atmosphere_from_process_value(merged.get(fld, "NR")))
    for fld in ["Drying_Time_hr_min_s", "Calcination_Time_hr_min_s", "Reduction_Time_hr_min_s"]:
        if merged.get(fld, "NR") not in ("NR", None, ""):
            merged[fld] = _normalize_multistep_time_field(_strip_atmosphere_from_process_value(merged.get(fld, "NR")))
    for fld in ["Calcination_Atmosphere", "Reduction_Atmosphere"]:
        if merged.get(fld, "NR") not in ("NR", None, ""):
            merged[fld] = _normalize_atmosphere_field(merged.get(fld, "NR"))
    for fld in ["Calcination_Temperature_Ramp_Rate_C_or_K_per_min", "Reduction_Temperature_Ramp_Rate_C_or_K_per_min"]:
        if merged.get(fld, "NR") not in ("NR", None, ""):
            merged[fld] = _normalize_temperature_ramp_rate_field(merged.get(fld, "NR"))
    if merged.get("Cooling_Rate_C_or_K_per_min", "NR") not in ("NR", None, ""):
        merged["Cooling_Rate_C_or_K_per_min"] = _normalize_cooling_rate_field(merged.get("Cooling_Rate_C_or_K_per_min", "NR"))

    if merged.get("CH4_Feed_Flow_Rate_mL_or_L_per_min", "NR") == "NR":
        derived_ch4_flow = _derive_ch4_flow_from_total_flow(
            merged.get("Feed_Flow_Rate_mL_or_L_per_min", "NR"),
            full_text,
        )
        if derived_ch4_flow:
            merged["CH4_Feed_Flow_Rate_mL_or_L_per_min"] = derived_ch4_flow
            merged[INTERNAL_NOTES_KEY] = _append_note(
                merged.get(INTERNAL_NOTES_KEY, "NR"),
                f"CH4 feed flow derived from HTML/article text using total feed flow and methane fraction: {derived_ch4_flow}.",
            )

    if merged.get("Space_Velocity", "NR") == "NR":
        derived_sv = _derive_space_velocity_from_flow_and_mass(
            merged.get("Feed_Flow_Rate_mL_or_L_per_min", "NR"),
            merged.get("Catalyst_Mass_g_or_mg", "NR"),
        )
        if derived_sv:
            merged["Space_Velocity"] = derived_sv
            merged[INTERNAL_NOTES_KEY] = _append_note(
                merged.get(INTERNAL_NOTES_KEY, "NR"),
                f"Space velocity derived from HTML/article text using feed flow and catalyst mass: {derived_sv}.",
            )

    if merged.get("Feed_Gas_Mixture", "NR") == "NR":
        mix = _extract_single_feed_gas_mixture(full_text)
        if mix:
            merged["Feed_Gas_Mixture"] = mix
            merged[INTERNAL_NOTES_KEY] = _append_note(
                merged.get(INTERNAL_NOTES_KEY, "NR"),
                f"Feed gas mixture inferred from HTML/article text: {mix}.",
            )

    if merged.get("Pressure_bar_or_atm_or_Pa", "NR") == "NR":
        merged["Pressure_bar_or_atm_or_Pa"] = "~1 atm"
        merged[INTERNAL_NOTES_KEY] = _append_note(
            merged.get(INTERNAL_NOTES_KEY, "NR"),
            "Pressure not reported at paper level; defaulted to ~1 atm.",
        )

    return merged



def _object_hint_from_filename(image_path: str) -> Tuple[Optional[str], Optional[int]]:
    base = os.path.basename(image_path)
    # RSC HTML asset names often look like c1xx00000a-f1.gif. Match this
    # before the generic numeric fallback so DOI/manuscript numbers are ignored.
    rsc_num = _rsc_figure_number_from_image_name(base) if '_rsc_figure_number_from_image_name' in globals() else None
    if rsc_num is not None:
        return "figure", rsc_num
    m = re.search(r"scheme[_\s-]*(\d+)", base, flags=re.IGNORECASE)
    if m:
        return "scheme", int(m.group(1))
    m = re.search(r"fig(?:ure)?[_\s-]*(\d+)", base, flags=re.IGNORECASE)
    if m:
        return "figure", int(m.group(1))
    m = re.search(r"(\d+)", base)
    if m:
        return None, int(m.group(1))
    return None, None


def _figure_number_from_filename(image_path: str) -> Optional[int]:
    _obj, num = _object_hint_from_filename(image_path)
    return num


def _object_type_matches(fig: Dict[str, Any], object_hint: Optional[str]) -> bool:
    if not object_hint:
        return True
    fig_type = _norm_text(fig.get("object_type", ""))
    label = _norm_text(fig.get("label", ""))
    if object_hint == "scheme":
        return "scheme" in fig_type or "scheme" in label
    if object_hint == "figure":
        return "scheme" not in fig_type and "scheme" not in label
    return True


def match_image_to_xml_figure_exact(image_path: str, parsed_xml: Dict[str, Any]) -> Tuple[Optional[Dict[str, Any]], str]:
    xml_figures = parsed_xml.get("figures", []) or []
    if not xml_figures:
        return None, "no_xml_figures"

    image_base = os.path.basename(image_path)
    image_stem = Path(image_base).stem.lower()

    # If HTML_and_image_extractor.py used a filename from the XML caption row,
    # prefer that exact filename/stem before falling back to numeric matching.
    exact_file_matches: List[Dict[str, Any]] = []
    for fig in xml_figures:
        xml_filename = _safe_str(fig.get("filename", "")).strip()
        if not xml_filename:
            continue
        xml_stem = Path(xml_filename).stem.lower()
        if xml_filename.lower() == image_base.lower() or xml_stem == image_stem:
            exact_file_matches.append(fig)
    if len(exact_file_matches) == 1:
        return exact_file_matches[0], "filename_xml_exact"
    if len(exact_file_matches) > 1:
        return exact_file_matches[0], "filename_xml_first"

    object_hint, fig_num = _object_hint_from_filename(image_path)
    if fig_num is None:
        return None, "filename_no_number"

    candidates = [fig for fig in xml_figures if _object_type_matches(fig, object_hint)]
    if not candidates:
        candidates = xml_figures

    label_match = []
    for fig in candidates:
        label = _safe_str(fig.get("label", ""))
        nums = re.findall(r"\d+", label)
        if any(int(n) == fig_num for n in nums):
            label_match.append(fig)
    if len(label_match) == 1:
        return label_match[0], "filename_label_exact"
    if len(label_match) > 1:
        return label_match[0], "filename_label_first"

    order_match = []
    for fig in candidates:
        try:
            order_num = int(float(fig.get("order", 0)))
        except Exception:
            continue
        if order_num == fig_num:
            order_match.append(fig)
    if len(order_match) == 1:
        return order_match[0], "filename_order_exact"
    if len(order_match) > 1:
        return order_match[0], "filename_order_first"

    return None, "filename_no_match"

def _series_group_key(record: Dict[str, Any]) -> str:
    series = _safe_str(record.get("Series_Label", "NR")).strip()
    if series not in ("", "NR"):
        return series
    label = _safe_str(record.get("Catalyst_Label", "NR")).strip()
    if label not in ("", "NR"):
        return label
    return "NR"


def _summarize_records_for_retry(records: List[Dict[str, Any]]) -> str:
    if not records:
        return "n=0"
    series_keys = sorted({
        _series_group_key(r)
        for r in records
        if _series_group_key(r) not in ("", "NR")
    })
    x_label = _safe_str(records[0].get("X_Axis_Label", "NR")).strip()
    y_label = _safe_str(records[0].get("Y_Axis_Label", "NR")).strip()
    counts: Dict[str, int] = {}
    x_vals = set()
    for r in records:
        s = _series_group_key(r)
        if s not in ("", "NR"):
            counts[s] = counts.get(s, 0) + 1
        xv = _x_axis_numeric_value(r)
        if xv is not None:
            x_vals.add(round(xv, 6))
    return (
        f"n={len(records)};"
        f"series={json.dumps(series_keys, ensure_ascii=False, separators=(',', ':'))};"
        f"x={json.dumps(x_label, ensure_ascii=False)};"
        f"y={json.dumps(y_label, ensure_ascii=False)};"
        f"counts={json.dumps(counts, ensure_ascii=False, separators=(',',':')) if counts else '{}'};"
        f"xvals={json.dumps(sorted(x_vals), ensure_ascii=False, separators=(',',':')) if x_vals else '[]'}"
    )


def _should_retry_strict(records: List[Dict[str, Any]]) -> bool:
    if not records:
        return False
    series_keys = [_series_group_key(r) for r in records]
    series_keys = [s for s in series_keys if s not in ("", "NR")]
    unique_series = sorted(set(series_keys))

    x_kind = _x_axis_kind(records[0].get("X_Axis_Label", ""))
    y_label = _safe_str(records[0].get("Y_Axis_Label", "")).lower()
    if x_kind not in {"time", "temperature", "ch4_flow", "total_flow", "space_velocity", "catalyst_mass", "pressure"}:
        return False
    if not ("conversion" in y_label or "methane" in y_label or "ch4" in y_label or "xch4" in y_label):
        return False

    x_vals = [_x_axis_numeric_value(r) for r in records]
    x_vals = [x for x in x_vals if x is not None]
    unique_x = set(round(x, 6) for x in x_vals)

    counts: Dict[str, int] = {}
    for r in records:
        s = _series_group_key(r)
        if s in ("", "NR"):
            continue
        counts[s] = counts.get(s, 0) + 1
    max_points = max(counts.values()) if counts else 0
    min_points = min(counts.values()) if counts else 0

    if len(records) <= 2:
        return True
    if len(unique_series) >= 1 and len(records) <= len(unique_series) + 1:
        return True
    if len(unique_x) <= 1:
        return True
    if max_points >= 3 and min_points < max_points:
        return True
    return False



def extract_records_from_figure(image_path: str, parsed_xml: Dict[str, Any], constants: Dict[str, Any], max_tokens: Optional[int] = None) -> List[Dict[str, Any]]:
    if max_tokens is None:
        max_tokens = FIGURE_MAX_OUTPUT_TOKENS
    fig_filename = os.path.basename(image_path)
    figure_keys = _dynamic_figure_extraction_keys(constants)
    matched_fig, match_method = match_image_to_xml_figure_exact(image_path, parsed_xml)
    caption_text = _extract_figure_caption_text(matched_fig)
    hard_skip, skip_hits = _should_hard_skip_non_target_figure(caption_text)
    if hard_skip and not SEND_ALL_FIGURES_TO_LLM:
        reason = ", ".join(skip_hits) if skip_hits else "non-target cues"
        print(f"      [skip] {fig_filename}: hard-skip non-target figure ({reason})")
        return []
    if hard_skip and SEND_ALL_FIGURES_TO_LLM:
        reason = ", ".join(skip_hits) if skip_hits else "non-target cues"
        print(f"      [no-skip] {fig_filename}: sending figure despite skip cues ({reason})")
    caption_context = build_xml_figure_context(parsed_xml, matched_fig)
    caption_candidates = _extract_context_catalyst_candidates(caption_context)
    matched_label = _display_norm(matched_fig.get("label", "NR")) if matched_fig else "NR"
    matched_order = matched_fig.get("order", "NR") if matched_fig else "NR"

    def _recover_catalyst_label(raw_label: str, series_label: str) -> Tuple[str, List[str]]:
        notes: List[str] = []
        label = _safe_str(raw_label).strip()
        series = _safe_str(series_label).strip()

        preferred_series = _preferred_series_label_as_catalyst(series, label)
        if preferred_series:
            notes.append(f"Catalyst_Label locked to Series_Label based on the visible legend/category label: {preferred_series}.")
            return preferred_series, notes

        if _is_foam_variant_label(label):
            return _canonicalize_foam_label(label), notes

        if _is_foam_variant_label(series):
            fixed = _canonicalize_foam_label(series)
            notes.append(f"Catalyst_Label recovered from series label: {fixed}.")
            return fixed, notes

        if _is_valid_catalyst_label(series) and (label in ("", "NR") or not _is_valid_catalyst_label(label)):
            notes.append(f"Catalyst_Label recovered from Series_Label: {series}.")
            return series, notes

        if _is_valid_catalyst_label(label):
            return label, notes

        valid_caption = [c for c in caption_candidates if _is_valid_catalyst_label(c)]
        if len(valid_caption) == 1:
            notes.append(f"Catalyst_Label recovered from HTML/caption context: {valid_caption[0]}.")
            return valid_caption[0], notes

        if len(valid_caption) > 1:
            for cand in valid_caption:
                if not _looks_like_condition_label(cand) and not _is_foam_variant_label(cand):
                    notes.append(f"Catalyst_Label recovered from HTML/caption context: {cand}.")
                    return cand, notes

        if label not in ("", "NR"):
            return label, notes
        return "NR", notes

    def _postprocess_records(records: List[Dict[str, Any]], forced_series_label: Optional[str] = None) -> List[Dict[str, Any]]:
        cleaned_records: List[Dict[str, Any]] = []
        for r in records:
            cleaned: Dict[str, Any] = {}
            for k in figure_keys:
                cleaned[k] = r.get(k, "NR")

            series_label = _safe_str(cleaned.get("Series_Label", "NR")).strip()
            raw_label = _safe_str(cleaned.get("Catalyst_Label", "NR")).strip()
            if forced_series_label and forced_series_label.strip() and forced_series_label != "NR":
                cleaned["Series_Label"] = forced_series_label
                series_label = forced_series_label.strip()
            elif series_label in ("", "NR"):
                cleaned["Series_Label"] = "NR"
                series_label = "NR"

            fixed_label, extra_notes = _recover_catalyst_label(raw_label, series_label)
            cleaned["Catalyst_Label"] = fixed_label
            for note in extra_notes:
                cleaned[INTERNAL_NOTES_KEY] = _append_note(cleaned.get(INTERNAL_NOTES_KEY, "NR"), note)

            cleaned["Series_Label"] = cleaned.get("Series_Label", forced_series_label or "NR")
            cleaned["_Matched_XML_Figure_Label"] = matched_label
            cleaned["_Matched_XML_Figure_Order"] = matched_order
            cleaned["_Caption_Match_Method"] = match_method
            cleaned["_Caption_Context"] = caption_context
            if _should_drop_recycled_or_later_cycle_record(cleaned):
                cleaned[INTERNAL_NOTES_KEY] = _append_note(
                    cleaned.get(INTERNAL_NOTES_KEY, "NR"),
                    "Record omitted because Series_Label/Catalyst_Label indicates a recycled, reused, regenerated, spent, or later-cycle catalyst.",
                )
                continue
            if not _is_target_conversion_axis(cleaned.get("Y_Axis_Label", "NR"), caption_context, caption_text):
                continue
            cleaned_records.append(cleaned)

        cleaned_records = _propagate_common_fields(
            cleaned_records,
            keys=[
                "X_Axis_Label",
                "Y_Axis_Label",
                "Heating_Mode",
                "Temperature_C_or_K",
                "Pressure_bar_or_atm_or_Pa",
                "Reaction_Time_hr_min_s",
                "Feed_Flow_Rate_mL_or_L_per_min",
                "CH4_Feed_Flow_Rate_mL_or_L_per_min",
                "Feed_Gas_Mixture",
                "Catalyst_Mass_g_or_mg",
                "Marker_Color",
                "Marker_Shape",
            ] + [k for k in figure_keys if k not in {"Catalyst_Label", "Series_Label", "Marker_Color", "Marker_Shape", "Heating_Mode", "Temperature_C_or_K", "Pressure_bar_or_atm_or_Pa", "X_Axis_Label", "CH4_Conversion_%", "Reaction_Time_hr_min_s", "Feed_Flow_Rate_mL_or_L_per_min", "CH4_Feed_Flow_Rate_mL_or_L_per_min", "Feed_Gas_Mixture", "Space_Velocity", "Catalyst_Mass_g_or_mg", "Y_Axis_Label"}],
        )
        return cleaned_records

    def _run_once(*, strict: bool, previous_summary: Optional[str], expected_series: Optional[List[str]] = None, expected_counts: Optional[Dict[str, Any]] = None, panel_specs: Optional[List[Dict[str, Any]]] = None) -> List[Dict[str, Any]]:
        content = [
            encode_image_for_converse(image_path),
            {"text": _figure_extraction_user_prompt(
                fig_filename=fig_filename,
                matched_xml_label=matched_label,
                matched_xml_order=matched_order,
                caption_context=caption_context,
                constants_json=constants,
                strict=strict,
                previous_extraction_summary=previous_summary,
                expected_series_labels=expected_series,
                expected_series_counts=expected_counts,
                expected_panel_specs=panel_specs,
                figure_keys=figure_keys,
            )},
        ]
        msgs = [{"role": "user", "content": content}]
        raw = _converse(msgs, system_prompt=FIGURE_SYSTEM_PROMPT, max_tokens=max_tokens)
        obj = _extract_first_json_value(raw)
        records = _to_list_of_dicts(obj)
        return _postprocess_records(records)

    def _run_single_series(panel_spec: Dict[str, Any], series_spec: Dict[str, Any]) -> List[Dict[str, Any]]:
        target_series = _display_norm(series_spec.get("label", "NR"))
        expected_count = _parse_expected_count(series_spec.get("point_count", "NR"))

        def _call(strict: bool = False, previous_summary: Optional[str] = None) -> List[Dict[str, Any]]:
            content = [
                encode_image_for_converse(image_path),
                {"text": _single_series_extraction_user_prompt(
                    fig_filename=fig_filename,
                    matched_xml_label=matched_label,
                    matched_xml_order=matched_order,
                    caption_context=caption_context,
                    constants_json=constants,
                    target_series_label=target_series,
                    figure_keys=figure_keys,
                    panel_spec=panel_spec,
                    series_spec=series_spec,
                    expected_count=expected_count,
                    strict=strict,
                    previous_extraction_summary=previous_summary,
                )},
            ]
            msgs = [{"role": "user", "content": content}]
            raw = _converse(msgs, system_prompt=FIGURE_SYSTEM_PROMPT, max_tokens=max_tokens)
            obj = _extract_first_json_value(raw)
            records = _to_list_of_dicts(obj)
            forced = None if target_series in ("", "NR") else target_series
            return _postprocess_records(records, forced_series_label=forced)

        records = _call(strict=False)
        if expected_count is not None and not _record_count_meets_expected(records, expected_count):
            print(f"      [retry] {fig_filename}: series '{target_series}' expected {expected_count} marker(s), got {len(records)} -> strict marker retry")
            try:
                retry_records = _call(strict=True, previous_summary=_summarize_records_for_retry(records))
                if len(retry_records) >= len(records):
                    records = retry_records
            except Exception as e:
                if isinstance(e, RateLimitReachedError):
                    raise
                print(f"      [warning] strict single-series retry failed for {fig_filename} [{target_series}]: {e}")
        return records

    panel_specs: List[Dict[str, Any]] = []
    discovered: Dict[str, Any] = {}
    discovery_failed = False
    try:
        discovered = _discover_relevant_panels_and_series(
            image_path=image_path,
            fig_filename=fig_filename,
            matched_xml_label=matched_label,
            matched_xml_order=matched_order,
            caption_context=caption_context,
        )
        panel_specs = discovered.get("panels", []) if isinstance(discovered, dict) else []
        has_multiple = discovered.get("has_multiple_panels", "NR") if isinstance(discovered, dict) else "NR"
        all_count = discovered.get("all_panel_count", "NR") if isinstance(discovered, dict) else "NR"
        if panel_specs:
            n_series = sum(len(p.get("series", []) or []) for p in panel_specs)
            print(f"      [panel-series] {fig_filename}: has_multiple_panels={has_multiple}, all_panel_count={all_count}, relevant CH4 panels={len(panel_specs)}, series/curves={n_series}")
        else:
            print(f"      [panel-series] {fig_filename}: has_multiple_panels={has_multiple}, all_panel_count={all_count}, no relevant CH4-conversion panel specs returned")
        _record_panel_discovery_audit(
            paper=parsed_xml.get("paper", CURRENT_PAPER_FOR_USAGE),
            fig_filename=fig_filename,
            matched_label=matched_label,
            matched_order=matched_order,
            match_method=match_method,
            caption_text=caption_text,
            caption_context=caption_context,
            discovered=discovered,
            panel_specs=panel_specs,
            status="ok",
        )
    except Exception as e:
        if isinstance(e, RateLimitReachedError):
            raise
        discovery_failed = True
        print(f"      [warning] panel/series discovery failed for {fig_filename}: {e}")
        panel_specs = []
        _record_panel_discovery_audit(
            paper=parsed_xml.get("paper", CURRENT_PAPER_FOR_USAGE),
            fig_filename=fig_filename,
            matched_label=matched_label,
            matched_order=matched_order,
            match_method=match_method,
            caption_text=caption_text,
            caption_context=caption_context,
            discovered=discovered,
            panel_specs=[],
            status="error",
            error=e,
        )

    def _run_panel_guided_fallback(panel_spec: Dict[str, Any], previous_summary: str) -> List[Dict[str, Any]]:
        one_panel_specs = [panel_spec]
        expected_series = _flatten_expected_series_from_panel_specs(one_panel_specs)
        expected_counts = _expected_counts_from_panel_specs(one_panel_specs)
        return _run_once(
            strict=True,
            previous_summary=previous_summary,
            expected_series=expected_series,
            expected_counts=expected_counts,
            panel_specs=one_panel_specs,
        )

    best_records: List[Dict[str, Any]] = []
    discovery_indicates_non_target = False
    try:
        discovery_indicates_non_target = isinstance(discovered, dict) and discovered.get("is_target_conversion_figure") is False
    except Exception:
        discovery_indicates_non_target = False

    if panel_specs:
        all_panel_records: List[Dict[str, Any]] = []
        for panel_idx, panel_spec in enumerate(panel_specs, start=1):
            panel_id = _display_norm(panel_spec.get("panel_id", f"panel {panel_idx}"))
            series_list = [s for s in (panel_spec.get("series", []) or []) if isinstance(s, dict)]
            print(f"      [panel] {fig_filename}: extracting {panel_id} with {len(series_list)} series/curve(s)")
            panel_records: List[Dict[str, Any]] = []
            for series_spec in series_list:
                label = _display_norm(series_spec.get("label", "NR"))
                if _looks_like_recycled_or_later_cycle(label):
                    continue
                try:
                    recs = _run_single_series(panel_spec, series_spec)
                    if recs:
                        panel_records.extend(recs)
                except Exception as e:
                    if isinstance(e, RateLimitReachedError):
                        raise
                    print(f"      [warning] single-series extraction failed for {fig_filename} [{panel_id} / {label}]: {e}")
            panel_records = _dedupe_records_simple(panel_records)
            if not panel_records:
                try:
                    print(f"      [fallback] {fig_filename}: panel-guided extraction for {panel_id} because per-series extraction returned no valid rows")
                    panel_records = _run_panel_guided_fallback(
                        panel_spec,
                        previous_summary=f"Per-series extraction returned no valid rows for {panel_id}.",
                    )
                except Exception as e:
                    if isinstance(e, RateLimitReachedError):
                        raise
                    print(f"      [warning] panel-guided fallback extraction failed for {fig_filename} [{panel_id}]: {e}")
                    panel_records = []
            if panel_records:
                all_panel_records.extend(panel_records)
        best_records = _dedupe_records_simple(all_panel_records)

    if not best_records and panel_specs:
        try:
            expected_series = _flatten_expected_series_from_panel_specs(panel_specs)
            expected_counts = _expected_counts_from_panel_specs(panel_specs)
            print(f"      [fallback] {fig_filename}: figure-level panel-guided extraction because all panel-specific attempts returned no valid rows")
            best_records = _run_once(
                strict=True,
                previous_summary="Panel/series discovery succeeded but all panel-specific extraction attempts returned no valid rows.",
                expected_series=expected_series,
                expected_counts=expected_counts,
                panel_specs=panel_specs,
            )
        except Exception as e:
            if isinstance(e, RateLimitReachedError):
                raise
            print(f"      [warning] figure-level panel-guided fallback extraction failed for {fig_filename}: {e}")

    if not best_records and discovery_indicates_non_target:
        print(f"      [skip] {fig_filename}: panel discovery found no CH4/methane-conversion panel")
        return []

    if not best_records and not panel_specs and (discovery_failed or _caption_looks_likely_ch4_target_process(caption_context)):
        # Emergency fallback: panel discovery failed or was inconclusive on a caption/context that still looks target, so preserve the old broad extraction path.
        # This call is intentionally after the panel-discovery attempt, never before it.
        try:
            records_1 = _run_once(strict=False, previous_summary="Panel discovery failed or returned no usable panel specs; running emergency broad extraction.")
        except Exception as e:
            if isinstance(e, RateLimitReachedError):
                raise
            print(f"      [warning] figure extraction parse failed for {fig_filename} (emergency fallback attempt 1): {e}")
            records_1 = []
        best_records = records_1
        if _should_retry_strict(records_1):
            summary = _summarize_records_for_retry(records_1)
            print(f"      [retry] {fig_filename}: incomplete emergency fallback extraction detected -> strict retry")
            try:
                records_2 = _run_once(strict=True, previous_summary=summary)
                if len(records_2) >= len(best_records) and records_2:
                    best_records = records_2
            except Exception as e:
                if isinstance(e, RateLimitReachedError):
                    raise
                print(f"      [warning] figure extraction parse failed for {fig_filename} (emergency fallback attempt 2): {e}")

    if not best_records and not panel_specs:
        print(f"      [skip] {fig_filename}: no relevant CH4/methane-conversion panel found after panel discovery")
        return []

    return best_records

def parse_catalyst_label(label: str) -> Tuple[Dict[str, float], str]:
    if not label or label == "NR":
        return {}, "NR"
    norm = label.replace("–", "-").replace("—", "-").replace("−", "-")
    norm = re.sub(r"\s+", "", norm)

    support = "NR"
    if "@" in norm:
        left, support = norm.split("@", 1)
    elif "/" in norm:
        left, support = norm.split("/", 1)
    else:
        left = norm

    support = _canonicalize_support_name(support)

    oxide_map = {
        "nio": "Ni",
        "feo": "Fe",
        "fe2o3": "Fe",
        "fe3o4": "Fe",
        "coo": "Co",
        "co2o3": "Co",
        "co3o4": "Co",
        "cuo": "Cu",
        "cu2o": "Cu",
        "pdo": "Pd",
        "zno": "Zn",
        "mno": "Mn",
        "mn2o3": "Mn",
        "mno2": "Mn",
        "cro": "Cr",
        "cr2o3": "Cr",
        "sno": "Sn",
        "sno2": "Sn",
        "pbo": "Pb",
        "pbo2": "Pb",
        "ceo2": "Ce",
        "tio2": "Ti",
        "zro2": "Zr",
        "mgo": "Mg",
        "cao": "Ca",
        "moo3": "Mo",
        "wo3": "W",
        "la2o3": "La",
        "al2o3": "Al",
    }
    loadings: Dict[str, float] = {}
    left_low = left.lower()
    if left_low in oxide_map:
        loadings.setdefault(oxide_map[left_low], 0.0)
        return loadings, support or "NR"

    inline_matches = list(re.finditer(r"([A-Z][a-z]?)(\d*(?:\.\d+)?)", left))
    reconstructed = "".join(m.group(1) + m.group(2) for m in inline_matches)
    non_o = [m.group(1) for m in inline_matches if m.group(1) != "O"]
    if inline_matches and reconstructed == left and len(set(non_o)) >= 2:
        has_oxygen_term = any(m.group(1) == "O" for m in inline_matches)
        if has_oxygen_term:
            for m in inline_matches:
                el = m.group(1)
                if el == "O":
                    continue
                loadings.setdefault(el, 0.0)
            return loadings, support or "NR"
        for m in inline_matches:
            el = m.group(1)
            if el == "O":
                continue
            wt = float(m.group(2)) if m.group(2) else 0.0
            loadings[el] = wt
        return loadings, support or "NR"

    segments = [seg for seg in re.split(r"[-+]", left) if seg]
    for seg in segments:
        seg_clean = re.sub(r"(?i)wt%", "", seg)
        low = seg_clean.lower()
        if low in oxide_map:
            loadings.setdefault(oxide_map[low], 0.0)
            continue
        m = re.fullmatch(r"(\d+(?:\.\d+)?)([A-Z][a-z]?)(?:O\d*)?", seg_clean)
        if m:
            loadings[m.group(2)] = float(m.group(1))
            continue
        m = re.fullmatch(r"([A-Z][a-z]?)(\d+(?:\.\d+)?)(?:O\d*)?", seg_clean)
        if m:
            loadings[m.group(1)] = float(m.group(2))
            continue
        m = re.fullmatch(r"([A-Z][a-z]?)(?:O\d*)?", seg_clean)
        if m:
            loadings.setdefault(m.group(1), 0.0)
            continue

    _canonical, semantic_loadings, semantic_support = _semantic_catalyst_components(label, label)
    if semantic_support not in ("NR", "") and support == "NR":
        support = semantic_support
    if semantic_loadings and (not loadings or len(semantic_loadings) > len(loadings)):
        return semantic_loadings, support or "NR"

    if not loadings and _is_foam_variant_label(label):
        loadings["Ni"] = 0.0

    if not loadings and _is_simple_carbon_catalyst_label(label):
        loadings["C"] = 0.0

    return loadings, support or "NR"

def classify_support_description(support: str) -> str:
    """Return a compact support class for the normalized Support_Type field."""
    if support == "NR":
        return "NR"
    s = _safe_str(support).strip()
    if not s or s == "NR":
        return "NR"
    low = _norm_text(s)
    if "unsupported" in low or "no support" in low:
        return "unsupported"
    if _is_zeolite_support_text(s) or _support_text_has_token(s, {"zsm", "mordenite", "beta", "bea", "sapo", "fau", "usy"}):
        return "zeolite"
    if _support_text_has_token(s, MOF_SUPPORT_TOKENS):
        return "MOF"
    if "monolith" in low:
        if any(x in low for x in ["metal", "fecral", "stainless", "nickel"]):
            return "metallic monolith"
        if any(x in low for x in ["cordierite", "ceramic", "sic", "silicon carbide", "alumina", "zirconia", "silica"]):
            return "ceramic monolith"
        return "monolith"
    if any(x in low for x in ["foam", "mesh", "foil", "metal foam", "nickel foam"]):
        if "carbon" in low:
            return "carbon-based"
        return "metallic"
    if _support_text_has_token(s, CARBIDE_SUPPORT_TOKENS):
        return "carbide"
    if _support_text_has_token(s, BORIDE_SUPPORT_TOKENS):
        return "boride"
    if re.search(r"(?<![A-Za-z0-9])TiN(?![A-Za-z0-9])", _safe_str(support)):
        return "nitride"
    if _support_text_has_token(s, NITRIDE_SUPPORT_TOKENS):
        return "nitride"
    if _support_text_has_token(s, {"ac", "mc", "cnt", "carbon", "graphite", "graphene", "activated carbon", "biochar", "mbc", "bc-"}):
        return "carbon-based"
    if _support_text_has_token(s, MIXED_METAL_OXIDE_SUPPORT_TOKENS) or _support_has_mixed_oxide_formula(s):
        return "mixed metal oxide"
    if _support_text_has_token(s, METALLOID_OXIDE_SUPPORT_TOKENS):
        return "metalloid oxide"
    if _support_text_has_token(s, METAL_OXIDE_SUPPORT_TOKENS) or "oxide" in low:
        return "metal oxide"
    return "other"


def enrich_row_from_label(row: Dict[str, Any], context_text: str = "") -> Dict[str, Any]:
    label = _safe_str(row.get("Catalyst_Label", "NR")).strip()

    if _is_foam_variant_label(label):
        row["Catalyst_Label"] = _canonicalize_foam_label(label)
        if row.get("Active_Metal", "NR") in ("NR", "", None):
            row["Active_Metal"] = "Ni"
        if row.get("Promoter_Type", "NR") in ("NR", "", None):
            row["Promoter_Type"] = "None"
        if row.get("Support_Type", "NR") in ("NR", "", None):
            row["Support_Type"] = "Nickel foam"
        if row.get("Support_Description", "NR") in ("NR", "", None):
            row["Support_Description"] = "metallic"
        if row.get("Catalyst_Class", "NR") in ("NR", "", None):
            row["Catalyst_Class"] = "metallic"
        return row

    if _is_simple_carbon_catalyst_label(label):
        if row.get("Active_Metal", "NR") in ("NR", "", None):
            row["Active_Metal"] = "C"
        if row.get("Promoter_Type", "NR") in ("NR", "", None):
            row["Promoter_Type"] = "None"
        if row.get("Support_Type", "NR") in ("NR", "", None):
            row["Support_Type"] = label
        if row.get("Support_Description", "NR") in ("NR", "", None):
            row["Support_Description"] = "carbon-based"
        if row.get("Catalyst_Class", "NR") in ("NR", "", None):
            row["Catalyst_Class"] = "carbon material"
        return row

    resolved_label, loadings, support = _resolve_row_catalyst_identity(row, context_text)
    if resolved_label not in ("NR", ""):
        row["Catalyst_Label"] = resolved_label
        label = resolved_label

    series_label = _safe_str(row.get("Series_Label", "NR")).strip()
    force_row_specific = bool(_explicit_numeric_composition_loadings(label))
    if not force_row_specific and series_label not in ("", "NR"):
        series_canonical = _canonicalize_catalyst_label_with_context(series_label, context_text)
        series_loadings, _series_support = parse_catalyst_label(series_canonical)
        series_loadings = _drop_explicit_zero_weight_metals(series_loadings, series_canonical)
        if _looks_like_specific_series_identity_label(series_label) and series_loadings and set(series_loadings.keys()) != set(loadings.keys()):
            force_row_specific = True

    if row.get("Support_Type", "NR") in ("NR", "", None) and support != "NR":
        row["Support_Type"] = support
    if row.get("Support_Type", "NR") not in ("NR", "", None):
        row["Support_Type"] = _normalize_support_type_field(
            row.get("Support_Type", "NR"),
            " ".join([context_text, label, _safe_str(row.get("Series_Label", ""))]),
        )
    if row.get("Support_Description", "NR") in ("NR", "", None):
        row["Support_Description"] = classify_support_description(_safe_str(row.get("Support_Type", "NR")))

    if not loadings:
        if row.get("Promoter_Type", "NR") in ("NR", "", None):
            row["Promoter_Type"] = "None"
        if row.get("Catalyst_Class", "NR") in ("NR", "", None):
            row["Catalyst_Class"] = infer_catalyst_class(row, context_text)
        return row

    promoters = []
    active = []
    for el, wt in loadings.items():
        if el in PROMOTER_ELEMENTS and len(loadings) > 1:
            promoters.append((el, wt))
        else:
            active.append((el, wt))

    if not active:
        active = list(loadings.items())

    if force_row_specific or row.get("Active_Metal", "NR") in ("NR", "", None):
        row["Active_Metal"] = "+".join([el for el, _ in active]) if active else "NR"

    if force_row_specific or row.get("Metal_Weight_%", "NR") in ("NR", "", None):
        act = [f"{_format_number_clean(float(wt))}% {el}" if float(wt) != 0 else el for el, wt in active]
        row["Metal_Weight_%"] = " + ".join(act) if act else "NR"

    if force_row_specific or row.get("Promoter_Type", "NR") in ("NR", "", None):
        if promoters:
            row["Promoter_Type"] = " + ".join(
                [f"{el} ({_format_number_clean(float(wt))} wt%)" if float(wt) != 0 else el for el, wt in promoters]
            )
        else:
            row["Promoter_Type"] = "None"

    if force_row_specific or row.get("Catalyst_Class", "NR") in ("NR", "", None):
        row["Catalyst_Class"] = infer_catalyst_class(row, context_text)

    return row

CATALYST_SYNTHESIS_KEYWORDS = [
    "prepare", "prepared", "preparation", "impregnation", "incipient", "ultrasonic",
    "stirred", "slurry", "dried", "drying", "calcined", "calcination",
    "reduced", "reduction", "precursor", "nitrate", "chloride", "acetate",
    "pyrolysis", "purchased", "commercially available", "table",
    "surface area", "specific surface", "bet", "sbet", "s bet", "si/al", "sio2/al2o3",
    "zeolite", "zsm", "mordenite", "support"
]


def _build_catalyst_synthesis_context(parsed_xml: Dict[str, Any], catalyst_label: str, support_type: str = "NR", series_label: str = "NR") -> str:
    paragraphs = [re.sub(r"\s+", " ", _safe_str(p)).strip() for p in (parsed_xml.get("paragraphs", []) or [])]
    if not paragraphs:
        return ""

    label = _safe_str(catalyst_label).strip()
    support = _safe_str(support_type).strip()
    series = _safe_str(series_label).strip()
    norm_label = _norm_text(label)
    left = _norm_text(_left_catalyst_token(label))
    exact_tokens = {t for t in [norm_label, left, _norm_text(series), _norm_text(support)] if t and t != "nr"}
    support_tokens = set()
    for support_part in [support, _normalize_support_type_field(support), series]:
        if support_part and support_part != "NR":
            support_tokens.add(_norm_text(support_part))
            support_tokens.add(_norm_text(re.sub(r"\([^)]*\)", "", support_part)))
            if _is_zeolite_support_text(support_part):
                support_tokens.update({"zeolite", "zsm", "zsm-5", "hzsm-5", "h-zsm-5", "mordenite", "beta"})
    for guessed in _extract_context_catalyst_candidates(" ".join([label, support, series])):
        support_tokens.add(_norm_text(guessed))

    selected: List[str] = []
    seen = set()
    for p in paragraphs:
        low = _norm_text(p)
        if not low:
            continue
        synth_hit = any(k in low for k in CATALYST_SYNTHESIS_KEYWORDS)
        exact_hit = any(tok and tok in low for tok in exact_tokens)
        support_hit = any(tok and tok in low for tok in support_tokens)
        if exact_hit or (synth_hit and support_hit):
            if low not in seen:
                seen.add(low)
                selected.append(p)

    if not selected:
        for p in paragraphs:
            low = _norm_text(p)
            if any(k in low for k in CATALYST_SYNTHESIS_KEYWORDS):
                if low not in seen:
                    seen.add(low)
                    selected.append(p)
                if len(selected) >= 10:
                    break

    context = " ".join(selected[:15]).strip()
    if len(context) > 9000:
        context = context[:9000].rstrip() + "..."
    return context



def _catalyst_synthesis_profile_user_prompt(
    catalyst_label: str,
    series_label: str,
    support_type: str,
    context_excerpt: str,
    reduction_state: str = "unknown",
) -> str:
    fields = _json_key_list(CATALYST_PROFILE_FIELDS)
    ctx = _compact_text_for_prompt(context_excerpt if context_excerpt else 'NR', max_chars=4200)
    return _render_prompt_template(
        "catalyst_profile_enrichment",
        TARGET_CATALYST=json.dumps(_safe_str(catalyst_label), ensure_ascii=False),
        TARGET_SERIES=json.dumps(_safe_str(series_label) if series_label else "NR", ensure_ascii=False),
        TARGET_SUPPORT=json.dumps(_safe_str(support_type) if support_type else "NR", ensure_ascii=False),
        TARGET_REDUCTION_STATE=json.dumps(_safe_str(reduction_state), ensure_ascii=False),
        CONTEXT=ctx,
        OUTPUT_KEYS=fields,
    )

def _empty_catalyst_profile() -> Dict[str, Any]:
    return {k: "NR" for k in CATALYST_PROFILE_FIELDS}


def _normalize_catalyst_synthesis_profile(profile: Dict[str, Any]) -> Dict[str, Any]:
    out = {k: normalize_blank_to_NR(profile.get(k, "NR")) for k in CATALYST_PROFILE_FIELDS}
    out["Support_Type"] = _normalize_support_type_field(out.get("Support_Type", "NR"), json.dumps(profile, ensure_ascii=False))
    out["Support_Surface_Area"] = _normalize_surface_area_field(out.get("Support_Surface_Area", "NR"))
    out["Precursor_Salts"] = _normalize_precursor_salts(out.get("Precursor_Salts", "NR"))
    for fld in ["Drying_Temperature_C_or_K", "Calcination_Temperature_C_or_K", "Reduction_Temperature_C_or_K"]:
        out[fld] = _normalize_multistep_temperature_field(_strip_atmosphere_from_process_value(out.get(fld, "NR")))
    for fld in ["Drying_Time_hr_min_s", "Calcination_Time_hr_min_s", "Reduction_Time_hr_min_s"]:
        out[fld] = _normalize_multistep_time_field(_strip_atmosphere_from_process_value(out.get(fld, "NR")))
    for atm_field, temp_field, time_field in [
        ("Calcination_Atmosphere", "Calcination_Temperature_C_or_K", "Calcination_Time_hr_min_s"),
        ("Reduction_Atmosphere", "Reduction_Temperature_C_or_K", "Reduction_Time_hr_min_s"),
    ]:
        atm = _normalize_atmosphere_field(out.get(atm_field, "NR"))
        if atm == "NR":
            inferred = _extract_atmosphere_from_text(" ".join([_safe_str(profile.get(temp_field, "")), _safe_str(profile.get(time_field, ""))]))
            if inferred:
                atm = _normalize_atmosphere_field(inferred)
        out[atm_field] = atm
    out["Calcination_Temperature_Ramp_Rate_C_or_K_per_min"] = _normalize_temperature_ramp_rate_field(out.get("Calcination_Temperature_Ramp_Rate_C_or_K_per_min", "NR"))
    out["Reduction_Temperature_Ramp_Rate_C_or_K_per_min"] = _normalize_temperature_ramp_rate_field(out.get("Reduction_Temperature_Ramp_Rate_C_or_K_per_min", "NR"))
    out["Cooling_Rate_C_or_K_per_min"] = _normalize_cooling_rate_field(out.get("Cooling_Rate_C_or_K_per_min", "NR"))
    return out


def _get_catalyst_specific_synthesis_profile(paper_name: str, parsed_xml: Dict[str, Any], row: Dict[str, Any]) -> Dict[str, Any]:
    label = _safe_str(row.get("Catalyst_Label", "NR")).strip()
    if label in ("", "NR") or not _is_valid_catalyst_label(label):
        return _empty_catalyst_profile()
    reduction_state = _infer_reduction_state(row, _safe_str(row.get("_Caption_Context", "")))
    support = _safe_str(row.get("Support_Type", "NR")).strip()
    series = _safe_str(row.get("Series_Label", "NR")).strip()
    cache_key = (paper_name, _norm_text(label), _norm_text(series), _norm_text(support), reduction_state)
    if cache_key in CATALYST_SYNTHESIS_PROFILE_CACHE:
        return CATALYST_SYNTHESIS_PROFILE_CACHE[cache_key]

    if _is_direct_use_catalyst(row, ""):
        profile = _empty_catalyst_profile()
        CATALYST_SYNTHESIS_PROFILE_CACHE[cache_key] = profile
        return profile

    context = _build_catalyst_synthesis_context(parsed_xml, label, support, series)
    if not context:
        profile = _empty_catalyst_profile()
        CATALYST_SYNTHESIS_PROFILE_CACHE[cache_key] = profile
        return profile

    prompt = _catalyst_synthesis_profile_user_prompt(label, series, support, context, reduction_state)
    msgs = [{"role": "user", "content": [{"text": prompt}]}]
    try:
        raw = _converse(msgs, system_prompt=TEXT_SYSTEM_PROMPT, max_tokens=1100)
        obj = _extract_first_json_value(raw)
        profile = obj if isinstance(obj, dict) else _empty_catalyst_profile()
    except Exception:
        profile = _empty_catalyst_profile()
    profile = _normalize_catalyst_synthesis_profile(profile)
    CATALYST_SYNTHESIS_PROFILE_CACHE[cache_key] = profile
    return profile


def _apply_catalyst_specific_synthesis_profile(row: Dict[str, Any], profile: Dict[str, Any]) -> None:
    if not profile:
        return
    for fld in CATALYST_PROFILE_FIELDS:
        val = normalize_blank_to_NR(profile.get(fld, "NR"))
        if val in ("NR", "", None):
            continue
        if fld == "Support_Type":
            val = _normalize_support_type_field(val, row.get("_Caption_Context", ""))
            current_val = _display_norm(row.get(fld, "NR"))
            if _support_value_more_informative(val, current_val):
                row[fld] = val
                row[_field_source_key(fld)] = "profile"
            continue
        if fld == "Support_Surface_Area":
            val = _normalize_surface_area_field(val)
            current_val = _display_norm(row.get(fld, "NR"))
            if current_val == "NR" or current_val != val:
                row[fld] = val
                row[_field_source_key(fld)] = "profile"
            continue
        current_source = _get_row_field_source(row, fld)
        current_val = _display_norm(row.get(fld, "NR"))
        if current_source == "figure" and current_val != "NR":
            continue
        row[fld] = val
        row[_field_source_key(fld)] = "profile"


def _dedupe_and_index_records(records: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    if not records:
        return []
    seen = set()
    deduped: List[Dict[str, Any]] = []
    for r in records:
        label = _safe_str(r.get("Catalyst_Label", "NR")).strip()
        series = _series_group_key(r)
        x_val = _x_axis_numeric_value(r)
        conv = _parse_first_float(r.get("CH4_Conversion_%", "NR"))
        key = (
            label,
            series,
            _safe_str(r.get("X_Axis_Label", "NR")).strip(),
            None if x_val is None else round(x_val, 6),
            None if conv is None else round(conv, 6),
            _safe_str(r.get("Y_Axis_Label", "NR")).strip(),
        )
        if key in seen:
            continue
        seen.add(key)
        deduped.append(r)

    groups: Dict[str, List[Dict[str, Any]]] = {}
    for r in deduped:
        series = _series_group_key(r)
        groups.setdefault(series, []).append(r)

    out: List[Dict[str, Any]] = []
    for series in sorted(groups.keys()):
        lst = groups[series]
        lst_sorted = sorted(
            lst,
            key=lambda rr: _x_axis_numeric_value(rr) if _x_axis_numeric_value(rr) is not None else 1e18,
        )
        for i, r in enumerate(lst_sorted, start=1):
            r["_DataPoint_Index"] = i
            out.append(r)
    return out
def build_row_template(paper_name: str) -> Dict[str, Any]:
    row = {col: "NR" for col in COLUMNS}
    row["Paper"] = paper_name
    row[INTERNAL_NOTES_KEY] = "NR"
    return row


def merge_constants_into_row(row: Dict[str, Any], constants: Dict[str, Any]) -> Dict[str, Any]:
    for k in TEXT_CONSTANT_FIELDS:
        if k in row and row.get(k, "NR") in ("NR", "", None):
            row[k] = constants.get(k, "NR") or "NR"
    return row


def _temperature_string_is_ambiguous(value: Any) -> bool:
    s = _safe_str(value).strip()
    if not s or s == "NR":
        return False
    if _temperature_text_has_range(s):
        return True
    if re.match(r"^\s*[-–]", s):
        return True
    if len(_extract_temperature_tokens(s)) > 1:
        return True
    num_tokens = re.findall(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?", s)
    lowered = _norm_text(s)
    if len(num_tokens) > 1 and any(sep in lowered for sep in [",", ";", " and ", "/"]):
        return True
    if re.search(r"\d\s*[-–]\s*\d", s):
        return True
    return False


def _finalize_temperature_fields(row: Dict[str, Any], constants: Dict[str, Any], context_text: str) -> None:
    x_kind = _x_axis_kind(row.get("X_Axis_Label", "NR"))
    context_temp, context_snippet = _extract_single_temperature_with_snippet(context_text)
    raw_temp_value = row.get("Temperature_C_or_K", "NR")
    const_raw_temp = constants.get("Temperature_C_or_K", "NR")
    row_temps = _extract_temperature_tokens(raw_temp_value)
    const_temps = [] if _temperature_string_is_ambiguous(const_raw_temp) else _extract_temperature_tokens(const_raw_temp)

    if _temperature_string_is_ambiguous(raw_temp_value):
        row[INTERNAL_NOTES_KEY] = _append_note(
            row.get(INTERNAL_NOTES_KEY, "NR"),
            f"Ambiguous temperature string for one datapoint ({raw_temp_value}). Cleared Temperature_C_or_K to NR."
        )
        row["Temperature_C_or_K"] = "NR"
        row["Temperature_source"] = "NR"
        row["Temperature_snippet"] = "NR"
        row_temps = []
    elif len(row_temps) == 1 and _safe_str(raw_temp_value).strip() not in (row_temps[0], f"~{row_temps[0]}"):
        row["Temperature_C_or_K"] = row_temps[0]

    if len(row_temps) > 1:
        row[INTERNAL_NOTES_KEY] = _append_note(
            row.get(INTERNAL_NOTES_KEY, "NR"),
            f"Multiple temperatures returned for one datapoint ({', '.join(row_temps)}). Cleared Temperature_C_or_K to NR."
        )
        row["Temperature_C_or_K"] = "NR"
        row["Temperature_source"] = "NR"
        row["Temperature_snippet"] = "NR"
        row_temps = []

    if x_kind != "temperature" and context_temp:
        if row_temps != [context_temp]:
            if row_temps:
                row[INTERNAL_NOTES_KEY] = _append_note(
                    row.get(INTERNAL_NOTES_KEY, "NR"),
                    f"Temperature overwritten by HTML/caption context: {', '.join(row_temps)} -> {context_temp}."
                )
            row["Temperature_C_or_K"] = context_temp
        row["Temperature_source"] = "xml_caption_context"
        if context_snippet:
            row["Temperature_snippet"] = context_snippet
        if _context_mentions_temperature_ramp(context_text):
            row[INTERNAL_NOTES_KEY] = _append_note(
                row.get(INTERNAL_NOTES_KEY, "NR"),
                "Figure context indicates a temperature ramp from room temperature to the designated reaction temperature."
            )
        return

    if not row_temps and x_kind != "temperature" and len(const_temps) == 1:
        row["Temperature_C_or_K"] = const_temps[0]
        row["Temperature_source"] = "paper_xml_text"
        row["Temperature_snippet"] = const_temps[0]
        row[INTERNAL_NOTES_KEY] = _append_note(
            row.get(INTERNAL_NOTES_KEY, "NR"),
            "Temperature taken from paper-level HTML/article constants (single value)."
        )
        return

    if len(const_temps) > 1 and row.get("Temperature_C_or_K", "NR") == "NR" and x_kind != "temperature":
        row[INTERNAL_NOTES_KEY] = _append_note(
            row.get(INTERNAL_NOTES_KEY, "NR"),
            "Paper-level HTML/article text contains multiple temperatures; figure-specific temperature left as NR."
        )

    if row.get("Temperature_C_or_K", "NR") not in ("NR", None, "") and row.get("Temperature_source", "NR") in ("NR", None, ""):
        # Direct figure-axis/legend temperatures do not need provenance columns.
        row["Temperature_source"] = "NR"
        row["Temperature_snippet"] = "NR"
    _finalize_temperature_provenance_columns(row)


TEMPERATURE_PROVENANCE_SOURCES = {"xml_caption_context", "paper_xml_text"}


def _finalize_temperature_provenance_columns(row: Dict[str, Any]) -> None:
    """Keep temperature source/snippet only when they explain a text-derived reaction temperature."""
    source = _display_norm(row.get("Temperature_source", "NR"))
    if source not in TEMPERATURE_PROVENANCE_SOURCES:
        row["Temperature_source"] = "NR"
        row["Temperature_snippet"] = "NR"
        return
    if _display_norm(row.get("Temperature_C_or_K", "NR")) == "NR":
        row["Temperature_source"] = "NR"
        row["Temperature_snippet"] = "NR"
    elif _display_norm(row.get("Temperature_snippet", "NR")) == "NR":
        row["Temperature_snippet"] = row.get("Temperature_C_or_K", "NR")


def _finalize_ch4_flow_fields(row: Dict[str, Any], constants: Dict[str, Any], context_text: str) -> None:
    x_kind = _x_axis_kind(row.get("X_Axis_Label", "NR"))
    raw_row_ch4 = row.get("CH4_Feed_Flow_Rate_mL_or_L_per_min", "NR")
    row_total = row.get("Feed_Flow_Rate_mL_or_L_per_min", "NR")
    series_text = _safe_str(row.get("Series_Label", "NR"))

    if _flow_string_has_multiple_values(raw_row_ch4):
        row[INTERNAL_NOTES_KEY] = _append_note(
            row.get(INTERNAL_NOTES_KEY, "NR"),
            f"Ambiguous CH4 feed flow string for one datapoint ({raw_row_ch4}). Cleared CH4_Feed_Flow_Rate_mL_or_L_per_min to NR."
        )
        row["CH4_Feed_Flow_Rate_mL_or_L_per_min"] = "NR"
        row_ch4 = None
    else:
        row_ch4 = _parse_flow_rate_string(raw_row_ch4)

    derivation_text = " ".join([
        _safe_str(series_text),
        _safe_str(context_text),
        _safe_str(constants.get(INTERNAL_NOTES_KEY, "NR")),
    ])

    series_specific = _derive_series_specific_ch4_flow(series_text, row_total, context_text)
    if series_specific is None:
        series_specific = _derive_series_specific_ch4_flow(
            series_text,
            constants.get("Feed_Flow_Rate_mL_or_L_per_min", row_total),
            context_text,
        )

    derived_from_context = _derive_ch4_flow_from_total_flow(row_total, derivation_text)
    if derived_from_context is None:
        derived_from_context = _derive_ch4_flow_from_total_flow(
            constants.get("Feed_Flow_Rate_mL_or_L_per_min", row_total),
            derivation_text,
        )

    preferred = series_specific or derived_from_context

    if x_kind == "ch4_flow":
        if preferred and row_ch4 is None:
            row["CH4_Feed_Flow_Rate_mL_or_L_per_min"] = preferred
            row[INTERNAL_NOTES_KEY] = _append_note(
                row.get(INTERNAL_NOTES_KEY, "NR"),
                f"CH4 feed flow filled from deterministic series/context derivation: {preferred}."
            )
        return

    if preferred:
        if row_ch4 is None:
            row["CH4_Feed_Flow_Rate_mL_or_L_per_min"] = preferred
            row[INTERNAL_NOTES_KEY] = _append_note(
                row.get(INTERNAL_NOTES_KEY, "NR"),
                f"CH4 feed flow derived from series/context: {preferred}."
            )
            return
        existing_numeric = _parse_first_float(raw_row_ch4)
        preferred_numeric = _parse_first_float(preferred)
        if existing_numeric is not None and preferred_numeric is not None and abs(existing_numeric - preferred_numeric) > max(1e-6, 0.05 * max(abs(preferred_numeric), 1.0)):
            row[INTERNAL_NOTES_KEY] = _append_note(
                row.get(INTERNAL_NOTES_KEY, "NR"),
                f"CH4 feed flow overwritten by deterministic derivation: {raw_row_ch4} -> {preferred}."
            )
            row["CH4_Feed_Flow_Rate_mL_or_L_per_min"] = preferred
            return

    if row.get("CH4_Feed_Flow_Rate_mL_or_L_per_min", "NR") in ("NR", None, ""):
        const_ch4 = constants.get("CH4_Feed_Flow_Rate_mL_or_L_per_min", "NR")
        if const_ch4 not in ("NR", None, "") and not _flow_string_has_multiple_values(const_ch4):
            row["CH4_Feed_Flow_Rate_mL_or_L_per_min"] = const_ch4
            row[INTERNAL_NOTES_KEY] = _append_note(
                row.get(INTERNAL_NOTES_KEY, "NR"),
                "CH4 feed flow taken from paper-level HTML/article constants."
            )
            return


def _finalize_feed_gas_mixture(row: Dict[str, Any], constants: Dict[str, Any], context_text: str) -> None:
    raw_mix = row.get("Feed_Gas_Mixture", "NR")
    if raw_mix not in ("NR", None, ""):
        row["Feed_Gas_Mixture"] = _standardize_feed_gas_mixture(raw_mix)
        return

    series_text = _safe_str(row.get("Series_Label", "NR"))
    mix = _extract_single_feed_gas_mixture(context_text)
    if mix is None:
        mix = _derive_binary_feed_mixture_from_ch4_fraction(series_text, context_text)
    if mix is None:
        mix = _extract_single_feed_gas_mixture(constants.get("Feed_Gas_Mixture", "NR"))
    if mix is None:
        mix = _derive_binary_feed_mixture_from_ch4_fraction(series_text, constants.get(INTERNAL_NOTES_KEY, "NR"))
    if mix:
        row["Feed_Gas_Mixture"] = mix
        row[INTERNAL_NOTES_KEY] = _append_note(
            row.get(INTERNAL_NOTES_KEY, "NR"),
            f"Feed gas mixture resolved from series/context: {mix}."
        )


def _normalize_pressure_field(value: Any) -> str:
    s = _safe_str(value).strip()
    if not s or s == "NR":
        return "NR"
    low = _norm_text(s)
    if re.search(r"\b(?:atmospheric|atmosphere|ambient pressure|room pressure|normal pressure)\b", low):
        return "1 atm"
    if low in {"atm", "1 atmosphere", "one atmosphere", "atmospheric pressure"}:
        return "1 atm"
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _pressure_string_is_ambiguous(value: Any) -> bool:
    s = _safe_str(value).strip()
    if not s or s == "NR":
        return False
    num_tokens = re.findall(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?", s)
    lowered = _norm_text(s)
    if len(num_tokens) > 1 and any(sep in lowered for sep in [",", ";", " and ", "/"]):
        return True
    if re.search(r"\d\s*[-–]\s*\d", s):
        return True
    return False


def _finalize_pressure_fields(row: Dict[str, Any], constants: Dict[str, Any], context_text: str) -> None:
    x_kind = _x_axis_kind(row.get("X_Axis_Label", "NR"))
    raw_pressure = _normalize_pressure_field(row.get("Pressure_bar_or_atm_or_Pa", "NR"))
    row["Pressure_bar_or_atm_or_Pa"] = raw_pressure
    if _pressure_string_is_ambiguous(raw_pressure):
        row[INTERNAL_NOTES_KEY] = _append_note(
            row.get(INTERNAL_NOTES_KEY, "NR"),
            f"Ambiguous pressure string for one datapoint ({raw_pressure}). Cleared Pressure_bar_or_atm_or_Pa to NR."
        )
        row["Pressure_bar_or_atm_or_Pa"] = "NR"

    if x_kind == "pressure":
        return

    if row.get("Pressure_bar_or_atm_or_Pa", "NR") not in ("NR", None, ""):
        row["Pressure_bar_or_atm_or_Pa"] = _normalize_pressure_field(row.get("Pressure_bar_or_atm_or_Pa", "NR"))
        return

    context_pressure = _safe_str(context_text)
    const_pressure = _safe_str(constants.get("Pressure_bar_or_atm_or_Pa", "NR"))

    if re.search(r"atmospheric|atmosphere|ambient pressure", context_pressure, flags=re.IGNORECASE):
        row["Pressure_bar_or_atm_or_Pa"] = "1 atm"
        row[INTERNAL_NOTES_KEY] = _append_note(row.get(INTERNAL_NOTES_KEY, "NR"), "Pressure set to 1 atm from atmospheric-pressure context.")
        return

    const_pressure = _normalize_pressure_field(const_pressure)
    if const_pressure not in ("NR", None, "") and not _pressure_string_is_ambiguous(const_pressure):
        row["Pressure_bar_or_atm_or_Pa"] = const_pressure
        return

    row["Pressure_bar_or_atm_or_Pa"] = "~1 atm"
    row[INTERNAL_NOTES_KEY] = _append_note(row.get(INTERNAL_NOTES_KEY, "NR"), "Pressure not reported for non-pressure figure; defaulted to ~1 atm.")


def _finalize_marker_color(row: Dict[str, Any]) -> None:
    row["Marker_Color"] = _standardize_marker_color(row.get("Marker_Color", "NR"))


def _finalize_heating_mode(row: Dict[str, Any], constants: Dict[str, Any], context_text: str) -> None:
    row_mode = _standardize_heating_mode(row.get("Heating_Mode", "NR"))
    if row_mode != "NR":
        row["Heating_Mode"] = row_mode
        return

    context_mode = _infer_heating_mode(context_text, row.get(INTERNAL_NOTES_KEY, "NR"))
    if context_mode != "NR":
        row["Heating_Mode"] = context_mode
        row[INTERNAL_NOTES_KEY] = _append_note(
            row.get(INTERNAL_NOTES_KEY, "NR"),
            f"Heating mode inferred from HTML/caption context: {context_mode}.",
        )
        return

    const_mode = _standardize_heating_mode(constants.get("Heating_Mode", "NR"))
    if const_mode != "NR":
        row["Heating_Mode"] = const_mode
        row[INTERNAL_NOTES_KEY] = _append_note(
            row.get(INTERNAL_NOTES_KEY, "NR"),
            "Heating mode taken from paper-level HTML/article constants.",
        )


def _finalize_space_velocity_fields(row: Dict[str, Any], constants: Dict[str, Any], context_text: str) -> None:
    x_kind = _x_axis_kind(row.get("X_Axis_Label", "NR"))
    raw_sv = row.get("Space_Velocity", "NR")
    if _space_velocity_string_is_ambiguous(raw_sv):
        row[INTERNAL_NOTES_KEY] = _append_note(
            row.get(INTERNAL_NOTES_KEY, "NR"),
            f"Ambiguous space velocity string for one datapoint ({raw_sv}). Cleared Space_Velocity to NR."
        )
        row["Space_Velocity"] = "NR"

    if row.get("Space_Velocity", "NR") not in ("NR", None, ""):
        row["Space_Velocity"] = _standardize_space_velocity(row.get("Space_Velocity", "NR"))
        return

    if x_kind == "space_velocity":
        return

    flow_candidates = [
        row.get("Feed_Flow_Rate_mL_or_L_per_min", "NR"),
        constants.get("Feed_Flow_Rate_mL_or_L_per_min", "NR"),
    ]
    mass_candidates = [
        row.get("Catalyst_Mass_g_or_mg", "NR"),
        constants.get("Catalyst_Mass_g_or_mg", "NR"),
    ]
    derived_sv = None
    for flow_candidate in flow_candidates:
        for mass_candidate in mass_candidates:
            derived_sv = _derive_space_velocity_from_flow_and_mass(flow_candidate, mass_candidate)
            if derived_sv is not None:
                break
        if derived_sv is not None:
            break
    if derived_sv is not None:
        row["Space_Velocity"] = _standardize_space_velocity(derived_sv)
        row[INTERNAL_NOTES_KEY] = _append_note(
            row.get(INTERNAL_NOTES_KEY, "NR"),
            f"Space velocity derived from feed flow and catalyst mass: {derived_sv}."
        )
        return

    const_sv = constants.get("Space_Velocity", "NR")
    if const_sv not in ("NR", None, "") and not _space_velocity_string_is_ambiguous(const_sv):
        row["Space_Velocity"] = _standardize_space_velocity(const_sv)


def _finalize_precursor_and_synthesis_fields(row: Dict[str, Any], constants: Dict[str, Any], context_text: str) -> None:
    if row.get("Precursor_Salts", "NR") in ("NR", "", None):
        const_prec = _normalize_precursor_salts(constants.get("Precursor_Salts", "NR"))
        if const_prec != "NR":
            row["Precursor_Salts"] = const_prec
            if _get_row_field_source(row, "Precursor_Salts") == "NR":
                row[_field_source_key("Precursor_Salts")] = "paper"
    else:
        row["Precursor_Salts"] = _normalize_precursor_salts(row.get("Precursor_Salts", "NR"))

    pre_before = _display_norm(row.get("Precursor_Salts", "NR"))
    pre_after = _filter_precursor_salts_for_row(pre_before, row)
    if pre_after != pre_before and pre_after != "NR":
        row["Precursor_Salts"] = pre_after
        row[INTERNAL_NOTES_KEY] = _append_note(
            row.get(INTERNAL_NOTES_KEY, "NR"),
            f"Precursor_Salts narrowed to catalyst-specific entries: {pre_before} -> {pre_after}.",
        )
    else:
        row["Precursor_Salts"] = pre_after

    for fld in ["Drying_Temperature_C_or_K", "Calcination_Temperature_C_or_K", "Reduction_Temperature_C_or_K"]:
        row[fld] = _normalize_multistep_temperature_field(_strip_atmosphere_from_process_value(row.get(fld, "NR")))
    for fld in ["Drying_Time_hr_min_s", "Calcination_Time_hr_min_s", "Reduction_Time_hr_min_s"]:
        row[fld] = _normalize_multistep_time_field(_strip_atmosphere_from_process_value(row.get(fld, "NR")))
    for fld in ["Calcination_Atmosphere", "Reduction_Atmosphere"]:
        row[fld] = _normalize_atmosphere_field(row.get(fld, "NR"))
    row["Calcination_Temperature_Ramp_Rate_C_or_K_per_min"] = _normalize_temperature_ramp_rate_field(row.get("Calcination_Temperature_Ramp_Rate_C_or_K_per_min", "NR"))
    row["Reduction_Temperature_Ramp_Rate_C_or_K_per_min"] = _normalize_temperature_ramp_rate_field(row.get("Reduction_Temperature_Ramp_Rate_C_or_K_per_min", "NR"))
    row["Cooling_Rate_C_or_K_per_min"] = _normalize_cooling_rate_field(row.get("Cooling_Rate_C_or_K_per_min", "NR"))

    for process_name, temp_field, time_field, atm_field in [
        ("Drying", "Drying_Temperature_C_or_K", "Drying_Time_hr_min_s", None),
        ("Calcination", "Calcination_Temperature_C_or_K", "Calcination_Time_hr_min_s", "Calcination_Atmosphere"),
        ("Reduction", "Reduction_Temperature_C_or_K", "Reduction_Time_hr_min_s", "Reduction_Atmosphere"),
    ]:
        temp_before = _display_norm(row.get(temp_field, "NR"))
        time_before = _display_norm(row.get(time_field, "NR"))
        atm_before = _display_norm(row.get(atm_field, "NR")) if atm_field else "NR"
        temp_after, time_after, atm_after = _filter_process_values_for_row(
            row, context_text, process_name, temp_before, time_before, atm_before
        )
        if temp_after != temp_before:
            row[temp_field] = temp_after
            row[INTERNAL_NOTES_KEY] = _append_note(
                row.get(INTERNAL_NOTES_KEY, "NR"),
                f"{temp_field} narrowed to the strategy relevant for this datapoint: {temp_before} -> {temp_after}.",
            )
        else:
            row[temp_field] = temp_after
        if time_after != time_before:
            row[time_field] = time_after
            row[INTERNAL_NOTES_KEY] = _append_note(
                row.get(INTERNAL_NOTES_KEY, "NR"),
                f"{time_field} narrowed to the strategy relevant for this datapoint: {time_before} -> {time_after}.",
            )
        else:
            row[time_field] = time_after
        if atm_field:
            if atm_after != atm_before:
                row[atm_field] = atm_after
                row[INTERNAL_NOTES_KEY] = _append_note(
                    row.get(INTERNAL_NOTES_KEY, "NR"),
                    f"{atm_field} narrowed to the strategy relevant for this datapoint: {atm_before} -> {atm_after}.",
                )
            else:
                row[atm_field] = atm_after

    row["Catalyst_Class"] = infer_catalyst_class(row, context_text)

    if _is_direct_use_catalyst(row, context_text):
        for fld in [
            "Precursor_Salts",
            "Drying_Temperature_C_or_K",
            "Drying_Time_hr_min_s",
            "Calcination_Temperature_C_or_K",
            "Calcination_Temperature_Ramp_Rate_C_or_K_per_min",
            "Calcination_Time_hr_min_s",
            "Calcination_Atmosphere",
            "Reduction_Temperature_C_or_K",
            "Reduction_Temperature_Ramp_Rate_C_or_K_per_min",
            "Reduction_Time_hr_min_s",
            "Reduction_Atmosphere",
            "Cooling_Rate_C_or_K_per_min",
        ]:
            row[fld] = "NR"
        row[INTERNAL_NOTES_KEY] = _append_note(
            row.get(INTERNAL_NOTES_KEY, "NR"),
            "Direct-use catalyst detected (biochar/carbon-only catalyst or nickel foam); wet-impregnation synthesis fields set to NR.",
        )
        return

    reduction_state = _infer_reduction_state(row, context_text)
    if reduction_state == "unreduced":
        row["Reduction_Temperature_C_or_K"] = "NR"
        row["Reduction_Temperature_Ramp_Rate_C_or_K_per_min"] = "NR"
        row["Reduction_Time_hr_min_s"] = "0 h"
        row["Reduction_Atmosphere"] = "NR"
        row[INTERNAL_NOTES_KEY] = _append_note(
            row.get(INTERNAL_NOTES_KEY, "NR"),
            "Unreduced catalyst variant detected; reduction fields cleared for this datapoint.",
        )
        return

    if _is_oxide_catalyst_label(row.get("Catalyst_Label", "NR")):
        red_time = _display_norm(row.get("Reduction_Time_hr_min_s", "NR"))
        red_temp = _display_norm(row.get("Reduction_Temperature_C_or_K", "NR"))
        if reduction_state != "reduced" and red_time == "NR" and red_temp == "NR":
            row["Reduction_Time_hr_min_s"] = "0 h"
            row["Reduction_Temperature_Ramp_Rate_C_or_K_per_min"] = "NR"
            row["Reduction_Atmosphere"] = "NR"
            row[INTERNAL_NOTES_KEY] = _append_note(
                row.get(INTERNAL_NOTES_KEY, "NR"),
                "Oxide catalyst detected and no synthesis reduction step reported; Reduction_Time_hr_min_s set to 0 h.",
            )


def create_placeholder_row(paper_name: str, constants: Dict[str, Any], note: str) -> Dict[str, Any]:
    row = build_row_template(paper_name)
    merge_constants_into_row(row, constants)
    row[INTERNAL_NOTES_KEY] = _append_note(constants.get(INTERNAL_NOTES_KEY, "NR"), note)
    return row


def create_rows_for_paper(paper_name: str, parsed_xml: Dict[str, Any], image_paths: List[str], constants: Dict[str, Any]) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []

    if not image_paths:
        return [create_placeholder_row(paper_name, constants, "No image files found in the images folder.")]

    for img_path in image_paths:
        fig_filename = os.path.basename(img_path)
        print(f"    [pass 2] figure: {fig_filename}")

        recs = extract_records_from_figure(img_path, parsed_xml, constants)
        if not recs:
            continue

        recs = _dedupe_and_index_records(recs)

        for rec in recs:
            row = build_row_template(paper_name)
            merge_constants_into_row(row, constants)

            fig_notes = normalize_blank_to_NR(rec.get(INTERNAL_NOTES_KEY, "NR"))
            if "_DataPoint_Index" in rec:
                row["DataPoint_Index"] = rec.get("_DataPoint_Index")

            for k, v_raw in rec.items():
                if k.startswith("_") or k == INTERNAL_NOTES_KEY:
                    continue
                if k in row:
                    v = normalize_blank_to_NR(v_raw)
                    if v != "NR":
                        row[k] = v

            _set_row_field_sources(row, rec, constants, CATALYST_PROFILE_FIELDS)

            row["Figure_Sources"] = fig_filename
            row["Matched_XML_Figure_Label"] = normalize_blank_to_NR(rec.get("_Matched_XML_Figure_Label", "NR"))
            row["Matched_XML_Figure_Order"] = normalize_blank_to_NR(rec.get("_Matched_XML_Figure_Order", "NR"))
            row["Caption_Match_Method"] = normalize_blank_to_NR(rec.get("_Caption_Match_Method", "NR"))

            context_text = _safe_str(rec.get("_Caption_Context", ""))
            row["_Caption_Context"] = context_text
            if row.get("Series_Label", "NR") in ("NR", "", None) and row.get("Catalyst_Label", "NR") not in ("NR", "", None):
                row["Series_Label"] = row.get("Catalyst_Label", "NR")
            if _is_foam_variant_label(row.get("Series_Label", "NR")) and row.get("Catalyst_Label", "NR") in ("NR", "", None):
                row["Catalyst_Label"] = _canonicalize_foam_label(_safe_str(row.get("Series_Label", "NR")))

            _finalize_temperature_fields(row, constants, context_text)
            _finalize_temperature_provenance_columns(row)
            _finalize_reaction_time_field(row)
            _finalize_ch4_conversion_fields(row, context_text)
            _finalize_ch4_flow_fields(row, constants, context_text)
            _finalize_feed_gas_mixture(row, constants, context_text)
            _finalize_space_velocity_fields(row, constants, context_text)
            _finalize_pressure_fields(row, constants, context_text)
            _finalize_heating_mode(row, constants, context_text)
            _finalize_marker_color(row)
            if "Marker_Shape" in row:
                row["Marker_Shape"] = _standardize_marker_shape(row.get("Marker_Shape", "NR"))
            row = enrich_row_from_label(row, context_text)

            inferred_support = _infer_support_from_context(context_text)
            if inferred_support != "NR" and row.get("Support_Type", "NR") in ("NR", "", None):
                row["Support_Type"] = inferred_support
            if row.get("Support_Type", "NR") not in ("NR", "", None):
                row["Support_Type"] = _normalize_support_type_field(
                    row.get("Support_Type", "NR"),
                    " ".join([row.get("_Caption_Context", ""), row.get("Catalyst_Label", ""), row.get("Series_Label", "")]),
                )
            if row.get("Support_Surface_Area", "NR") not in ("NR", "", None):
                row["Support_Surface_Area"] = _normalize_surface_area_field(row.get("Support_Surface_Area", "NR"))
            if row.get("Support_Description", "NR") in ("NR", "", None):
                row["Support_Description"] = classify_support_description(_safe_str(row.get("Support_Type", "NR")))

            catalyst_profile = _get_catalyst_specific_synthesis_profile(paper_name, parsed_xml, row)
            catalyst_profile = _filter_catalyst_profile_for_row(catalyst_profile, row, context_text)
            _apply_catalyst_specific_synthesis_profile(row, catalyst_profile)
            _finalize_precursor_and_synthesis_fields(row, constants, context_text)
            _finalize_reaction_time_field(row)
            _finalize_temperature_provenance_columns(row)
            row["Support_Description"] = classify_support_description(_safe_str(row.get("Support_Type", "NR")))

            const_notes = constants.get(INTERNAL_NOTES_KEY, "NR")
            combined = []
            if const_notes not in ("NR", "", None):
                combined.append(f"[xml-text] {const_notes}")
            if fig_notes not in ("NR", "", None):
                combined.append(f"[figure:{fig_filename}] {fig_notes}")

            existing_notes = row.get(INTERNAL_NOTES_KEY, "NR")
            all_notes: List[str] = []
            if existing_notes not in ("NR", "", None):
                all_notes.append(existing_notes)
            all_notes.extend(combined)
            row[INTERNAL_NOTES_KEY] = " | ".join([n for n in all_notes if n]) if all_notes else "NR"
            rows.append(row)

    if not rows:
        rows.append(create_placeholder_row(paper_name, constants, "No datapoints extracted from the manually curated figures."))

    return rows


def _build_llm_usage_summary_df(mapping_records: List[Dict[str, str]], all_rows: List[Dict[str, Any]]) -> pd.DataFrame:
    del all_rows
    rows: List[Dict[str, Any]] = []
    totals = _empty_usage_bucket()

    def make_summary_row(paper_label: str, usage: Dict[str, Any]) -> Dict[str, Any]:
        usage_fixed = _empty_usage_bucket()
        usage_fixed.update(usage or {})
        return {
            "Paper #": paper_label,
            "# of calls": int(usage_fixed.get("llm_calls", 0) or 0),
            "# of input tokens": int(usage_fixed.get("input_tokens", 0) or 0),
            "# of output tokens": int(usage_fixed.get("output_tokens", 0) or 0),
        }

    for rec in mapping_records:
        paper = rec["Paper"]
        usage = _empty_usage_bucket()
        usage.update(LLM_USAGE_BY_PAPER.get(paper, {}) or {})
        rows.append(make_summary_row(paper, usage))
        for key in USAGE_COUNTER_KEYS:
            totals[key] = int(totals.get(key, 0) or 0) + int(usage.get(key, 0) or 0)

    if rows:
        rows.append(make_summary_row("TOTAL", totals))
    return pd.DataFrame(rows)


def write_llm_usage_summary_xlsx(summary_df: pd.DataFrame, out_path: str) -> None:
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        summary_df.to_excel(writer, index=False, sheet_name="LLM Usage Summary")
    wb = load_workbook(out_path)
    ws = wb["LLM Usage Summary"]
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    if ws.max_row >= 1:
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF", size=10, name="Arial")
            cell.fill = PatternFill("solid", start_color="4472C4")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border
        ws.row_dimensions[1].height = 42

    header_by_col = {idx: _safe_str(ws.cell(row=1, column=idx).value) for idx in range(1, ws.max_column + 1)}
    for row in ws.iter_rows(min_row=2):
        is_total_row = _safe_str(row[0].value).strip().upper() == "TOTAL"
        for cell in row:
            header = header_by_col.get(cell.column, "")
            cell.font = Font(size=10, name="Arial", bold=is_total_row)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = border
            if is_total_row:
                cell.fill = PatternFill("solid", start_color="D9EAF7")
            elif cell.row % 2 == 0:
                cell.fill = PatternFill("solid", start_color="F2F2F2")
            header_l = header.lower()
            if isinstance(cell.value, (int, float)) and ("tokens" in header_l or "# of" in header_l):
                cell.number_format = "#,##0"

    width_by_header = {
        "Paper #": 18,
        "# of calls": 18,
        "# of input tokens": 20,
        "# of output tokens": 20,
    }
    for idx in range(1, ws.max_column + 1):
        header = header_by_col.get(idx, "")
        ws.column_dimensions[get_column_letter(idx)].width = width_by_header.get(header, 20)
    ws.freeze_panes = "A2"
    wb.save(out_path)


SECTION_COLORS = {
    "Paper": "4472C4",
    "Catalyst_Label": "4472C4",
    "Series_Label": "4472C4",
    "Marker_Color": "4472C4",
    "Marker_Shape": "4472C4",
    "DataPoint_Index": "4472C4",
    "Heating_Mode": "4472C4",
    "CH4": "ED7D31",
    "Temperature": "ED7D31",
    "Pressure": "ED7D31",
    "Reaction_Time": "ED7D31",
    "Feed_Flow": "ED7D31",
    "Feed_Gas_Mixture": "ED7D31",
    "Space_Velocity": "ED7D31",
    "Catalyst_Mass": "ED7D31",
    "Active_Metal": "70AD47",
    "Metal_Weight": "70AD47",
    "Catalyst_Class": "70AD47",
    "Promoter": "70AD47",
    "Support": "FFC000",
    "Precursor": "9DC3E6",
    "Drying": "9DC3E6",
    "Calcination": "9DC3E6",
    "Reduction": "9DC3E6",
    "Cooling": "9DC3E6",
    "X_Axis_Label": "D9D9D9",
    "Y_Axis_Label": "D9D9D9",
    "Figure_Sources": "D9D9D9",
    "Matched_XML_Figure_Label": "D9D9D9",
    "Matched_XML_Figure_Order": "D9D9D9",
    "Caption_Match_Method": "D9D9D9",
}


def col_color(col_name: str) -> str:
    for prefix, color in SECTION_COLORS.items():
        if _safe_str(col_name).startswith(prefix):
            return color
    return "D9D9D9"


def format_xlsx(xlsx_path: str) -> None:
    wb = load_workbook(xlsx_path)
    ws = wb.active
    ws.title = "TCD Extraction"

    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    if ws.max_row >= 1:
        for cell in ws[1]:
            hex_color = col_color(cell.value or "")
            cell.font = Font(bold=True, color="FFFFFF", size=10, name="Arial")
            cell.fill = PatternFill("solid", start_color=hex_color)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border
        ws.row_dimensions[1].height = 42

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.font = Font(size=10, name="Arial")
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = border
            if cell.row % 2 == 0:
                cell.fill = PatternFill("solid", start_color="F2F2F2")

    width_map = {
        "Paper": 14,
        "Catalyst_Label": 22,
        "Series_Label": 22,
        "Marker_Color": 14,
        "Marker_Shape": 16,
        "DataPoint_Index": 14,
        "Heating_Mode": 18,
        "CH4_Conversion_%": 18,
        "Temperature_C_or_K": 18,
        "Temperature_source": 22,
        "Temperature_snippet": 32,
        "Pressure_bar_or_atm_or_Pa": 18,
        "X_Axis_Label": 22,
        "Reaction_Time_hr_min_s": 18,
        "Feed_Flow_Rate_mL_or_L_per_min": 20,
        "CH4_Feed_Flow_Rate_mL_or_L_per_min": 22,
        "Feed_Gas_Mixture": 24,
        "Space_Velocity": 22,
        "Catalyst_Mass_g_or_mg": 18,
        "Active_Metal": 16,
        "Metal_Weight_%": 18,
        "Catalyst_Class": 20,
        "Promoter_Type": 20,
        "Support_Type": 26,
        "Support_Surface_Area": 18,
        "Support_Description": 22,
        "Precursor_Salts": 30,
        "Drying_Temperature_C_or_K": 18,
        "Drying_Time_hr_min_s": 16,
        "Calcination_Temperature_C_or_K": 18,
        "Calcination_Temperature_Ramp_Rate_C_or_K_per_min": 24,
        "Calcination_Time_hr_min_s": 16,
        "Calcination_Atmosphere": 18,
        "Reduction_Temperature_C_or_K": 18,
        "Reduction_Temperature_Ramp_Rate_C_or_K_per_min": 24,
        "Reduction_Time_hr_min_s": 16,
        "Reduction_Atmosphere": 18,
        "Cooling_Rate_C_or_K_per_min": 20,
        "Y_Axis_Label": 28,
        "Figure_Sources": 22,
        "Matched_XML_Figure_Label": 22,
        "Matched_XML_Figure_Order": 18,
        "Caption_Match_Method": 20,
    }
    for col_idx, cell in enumerate(ws[1], start=1):
        col_letter = get_column_letter(col_idx)
        col_name = cell.value or ""
        ws.column_dimensions[col_letter].width = width_map.get(col_name, 22)

    ws.freeze_panes = "B2"
    wb.save(xlsx_path)
    print(f"  Formatting applied to {xlsx_path}")



def _finalize_units_and_text_fields(row: Dict[str, Any]) -> None:
    """Final deterministic cleanup for units, pressure, atmospheres, SV, and text artifacts."""
    if not isinstance(row, dict):
        return

    for fld in ["Temperature_C_or_K", "Drying_Temperature_C_or_K", "Calcination_Temperature_C_or_K", "Reduction_Temperature_C_or_K"]:
        if fld in row:
            row[fld] = _normalize_multistep_temperature_field(_strip_atmosphere_from_process_value(row.get(fld, "NR")))

    for fld in ["Drying_Time_hr_min_s", "Calcination_Time_hr_min_s", "Reduction_Time_hr_min_s"]:
        if fld in row:
            row[fld] = _normalize_multistep_time_field(_strip_atmosphere_from_process_value(row.get(fld, "NR")))

    for fld in ["Calcination_Temperature_Ramp_Rate_C_or_K_per_min", "Reduction_Temperature_Ramp_Rate_C_or_K_per_min", "Cooling_Rate_C_or_K_per_min"]:
        if fld in row:
            row[fld] = _normalize_temperature_ramp_rate_field(row.get(fld, "NR"))

    for fld in ["Calcination_Atmosphere", "Reduction_Atmosphere"]:
        if fld in row:
            row[fld] = _normalize_atmosphere_field(row.get(fld, "NR"))

    if "Pressure_bar_or_atm_or_Pa" in row:
        row["Pressure_bar_or_atm_or_Pa"] = _normalize_pressure_field(row.get("Pressure_bar_or_atm_or_Pa", "NR"))

    if _display_norm(row.get("Space_Velocity", "NR")) == "NR":
        derived_sv = _derive_space_velocity_from_flow_and_mass(
            row.get("Feed_Flow_Rate_mL_or_L_per_min", "NR"),
            row.get("Catalyst_Mass_g_or_mg", "NR"),
        )
        if derived_sv:
            row["Space_Velocity"] = _standardize_space_velocity(derived_sv)
            row[INTERNAL_NOTES_KEY] = _append_note(
                row.get(INTERNAL_NOTES_KEY, "NR"),
                f"Space velocity derived at final cleanup from feed flow and catalyst mass: {derived_sv}.",
            )
    elif "Space_Velocity" in row:
        row["Space_Velocity"] = _standardize_space_velocity(row.get("Space_Velocity", "NR"))

    for key, value in list(row.items()):
        if isinstance(value, str):
            row[key] = _clean_text_artifacts(value).strip() or "NR"


def _finalize_output_rows(all_rows: List[Dict[str, Any]]) -> None:
    """Apply deterministic final row cleanup before writing CSV/XLSX.

    This intentionally repeats CH4 conversion normalization because resumed
    progress can contain rows produced before a logic patch. Freshly finalized
    rows carry CH4_CONVERSION_FINALIZED_KEY and are skipped.
    """
    for row in all_rows:
        if not isinstance(row, dict):
            continue
        _finalize_ch4_conversion_fields(row, row.get("_Caption_Context", ""))
        _finalize_reaction_time_field(row)
        _finalize_temperature_provenance_columns(row)
        _finalize_units_and_text_fields(row)
        if row.get("Support_Type", "NR") not in ("NR", "", None):
            row["Support_Type"] = _normalize_support_type_field(
                row.get("Support_Type", "NR"),
                " ".join([row.get("_Caption_Context", ""), row.get("Catalyst_Label", ""), row.get("Series_Label", "")]),
            )
        row["Support_Description"] = classify_support_description(_safe_str(row.get("Support_Type", "NR")))


def _select_output_columns(rows: List[Dict[str, Any]]) -> List[str]:
    """Drop temperature provenance columns unless at least one row actually uses them."""
    keep_temperature_provenance = any(
        _display_norm(row.get("Temperature_source", "NR")) != "NR"
        or _display_norm(row.get("Temperature_snippet", "NR")) != "NR"
        for row in rows
        if isinstance(row, dict)
    )
    cols: List[str] = []
    seen = set()
    for col in COLUMNS:
        if col in seen:
            continue
        seen.add(col)
        if col in {"Temperature_source", "Temperature_snippet"} and not keep_temperature_provenance:
            continue
        cols.append(col)
    return cols



def _paper_number_from_label(label: Any) -> Optional[int]:
    m = re.search(r"\d+", _safe_str(label))
    if not m:
        return None
    try:
        return int(m.group(0))
    except Exception:
        return None


def run_extraction_pipeline(
    mapping_xlsx: str,
    parsed_xml_dir: str,
    pdf_base_dir: str,
    output_prefix: str,
    fresh_start: bool = False,
    paper_filters: Optional[List[str]] = None,
    resume_from_paper_number: Optional[int] = None,
    send_all_figures_to_llm: bool = False,
    save_audit_reports: bool = True,
    audit_output_root: Optional[str] = None,
    figure_source: str = "auto",
) -> Dict[str, Any]:
    _set_send_all_figures_to_llm(send_all_figures_to_llm)
    output_dir = os.path.dirname(output_prefix) or "."
    output_stem = os.path.basename(output_prefix)
    os.makedirs(output_dir, exist_ok=True)

    output_csv = os.path.join(output_dir, f"{output_stem}.csv")
    output_xlsx = os.path.join(output_dir, f"{output_stem}.xlsx")
    output_usage_xlsx = os.path.join(output_dir, f"{output_stem}_llm_usage_summary.xlsx")
    progress_path = _progress_path(os.path.join(output_dir, output_stem))
    audit_output_root = audit_output_root or os.path.join(output_dir, f"{output_stem}_audit_reports")
    if fresh_start and os.path.exists(progress_path):
        os.remove(progress_path)

    mapping_records = read_paper_mapping(mapping_xlsx)
    if paper_filters:
        wanted = {_norm_text(p) for p in paper_filters if _safe_str(p).strip()}
        all_mapping_names = {_norm_text(rec.get("Paper", "")) for rec in mapping_records}
        missing = [p for p in paper_filters if _norm_text(p) not in all_mapping_names]
        if missing:
            print(f"[warning] Requested paper name(s) not found in mapping workbook: {', '.join(missing)}")
        mapping_records = [rec for rec in mapping_records if _norm_text(rec.get("Paper", "")) in wanted]
    if resume_from_paper_number is not None:
        filtered: List[Dict[str, str]] = []
        for idx, rec in enumerate(mapping_records, start=1):
            detected_num = _paper_number_from_label(rec.get("Paper", ""))
            effective_num = detected_num if detected_num is not None else idx
            if effective_num >= int(resume_from_paper_number):
                filtered.append(rec)
        print(f"[resume-from-paper-number] Keeping {len(filtered)} paper(s) with paper number/order >= {resume_from_paper_number}")
        mapping_records = filtered
    if send_all_figures_to_llm:
        print("[figure-skip] Pre-LLM hard skip is disabled; all curated figure images will be sent to the LLM.")
    if not mapping_records:
        print("No rows found in mapping workbook after filtering.")
        return {}

    all_rows: List[Dict[str, Any]] = []
    completed_papers: List[str] = []
    LLM_USAGE_BY_PAPER.clear()
    PASS1_CONSTANTS_BY_PAPER.clear()
    PANEL_DISCOVERY_AUDIT_BY_PAPER.clear()
    AUDIT_IMAGE_PATHS_BY_PAPER.clear()
    CATALYST_SYNTHESIS_PROFILE_CACHE.clear()

    if os.path.exists(progress_path):
        progress = _load_progress(progress_path)
        completed_papers = list(progress.get("completed_papers", []))
        all_rows = list(progress.get("all_rows", []))
        for paper_name, usage in (progress.get("llm_usage_by_paper", {}) or {}).items():
            bucket = _empty_usage_bucket()
            for key in USAGE_COUNTER_KEYS:
                bucket[key] = int(usage.get(key, 0) or 0)
            LLM_USAGE_BY_PAPER[paper_name] = bucket
        for paper_name, constants_payload in (progress.get("pass1_constants_by_paper", {}) or {}).items():
            if isinstance(constants_payload, dict):
                PASS1_CONSTANTS_BY_PAPER[paper_name] = dict(constants_payload)
        for paper_name, discovery_payload in (progress.get("panel_discovery_audit_by_paper", {}) or {}).items():
            if isinstance(discovery_payload, list):
                PANEL_DISCOVERY_AUDIT_BY_PAPER[paper_name] = [dict(x) for x in discovery_payload if isinstance(x, dict)]
        for paper_name, image_payload in (progress.get("audit_image_paths_by_paper", {}) or {}).items():
            if isinstance(image_payload, list):
                AUDIT_IMAGE_PATHS_BY_PAPER[paper_name] = [_safe_str(x) for x in image_payload]
        if completed_papers:
            print(f"[resume] Loaded progress for {len(completed_papers)} completed paper(s) from {progress_path}")

    _register_progress_state_for_rate_limit(progress_path, completed_papers, all_rows)

    for rec in mapping_records:
        paper = rec["Paper"]
        if paper in completed_papers:
            print(f"[resume] Skipping completed paper: {paper}")
            continue
        _set_current_paper_for_usage(paper)
        folder_path = os.path.join(pdf_base_dir, paper)
        images_dir = os.path.join(folder_path, "images")
        pdf_image_paths = get_image_paths(images_dir)
        html_paper_dir = os.path.join(parsed_xml_dir, paper)
        html_image_paths = get_rsc_html_image_paths(html_paper_dir)
        source_mode = _norm_text(figure_source or "auto")
        if source_mode == "pdf":
            image_paths = pdf_image_paths
        elif source_mode == "html":
            image_paths = html_image_paths
        elif source_mode == "both":
            seen_paths = set()
            image_paths = []
            for p in pdf_image_paths + html_image_paths:
                key = os.path.abspath(p)
                if key not in seen_paths:
                    seen_paths.add(key)
                    image_paths.append(p)
        else:
            image_paths = pdf_image_paths if pdf_image_paths else html_image_paths

        print(f"\n{'=' * 70}")
        print(f"Processing: {paper}")
        print(f"  PDF curated images : {len(pdf_image_paths)}")
        print(f"  HTML figure images : {len(html_image_paths)}")
        print(f"  Using images       : {len(image_paths)} ({source_mode or 'auto'})")
        AUDIT_IMAGE_PATHS_BY_PAPER[paper] = list(image_paths)

        try:
            parsed_xml = load_parsed_xml(parsed_xml_dir, paper)
        except Exception as e:
            if isinstance(e, RateLimitReachedError):
                raise
            print(f"  [warning] RSC HTML artifacts missing for {paper}: {e}")
            parsed_xml = {
                "paper": paper,
                "title": "NR",
                "doi": rec.get("DOI", "NR"),
                "paragraphs": [],
                "full_text": "",
                "figures": [],
            }

        parsed_xml["paper"] = paper
        if _display_norm(parsed_xml.get("doi", "NR")) == "NR":
            parsed_xml["doi"] = rec.get("DOI", "NR")

        print("  [pass 1] Extracting paper constants from RSC HTML article text ...")
        try:
            if parsed_xml.get("full_text", ""):
                constants = extract_paper_constants_from_xml(parsed_xml)
            else:
                constants = {k: "NR" for k in TEXT_CONSTANT_FIELDS}
                constants["Title"] = parsed_xml.get("title", "NR")
                constants["DOI"] = parsed_xml.get("doi", rec.get("DOI", "NR"))
                constants[INTERNAL_NOTES_KEY] = "No parsed HTML text available."
        except Exception as e:
            if isinstance(e, RateLimitReachedError):
                raise
            print(f"  [error] constants extraction failed: {e}")
            constants = {k: "NR" for k in TEXT_CONSTANT_FIELDS}
            constants["Title"] = parsed_xml.get("title", "NR")
            constants["DOI"] = parsed_xml.get("doi", rec.get("DOI", "NR"))
            constants[INTERNAL_NOTES_KEY] = f"Constants extraction error: {e}"

        print(f"    Temp : {constants.get('Temperature_C_or_K', 'NR')}")
        print(f"    Heat : {constants.get('Heating_Mode', 'NR')}")
        print(f"    Total flow : {constants.get('Feed_Flow_Rate_mL_or_L_per_min', 'NR')}")
        print(f"    CH4 flow   : {constants.get('CH4_Feed_Flow_Rate_mL_or_L_per_min', 'NR')}")
        print(f"    Rxn time   : {constants.get('Reaction_Time_hr_min_s', 'NR')}")
        PASS1_CONSTANTS_BY_PAPER[paper] = dict(constants)

        print("  [pass 2] Extracting records from manually curated figures ...")
        try:
            paper_rows = create_rows_for_paper(
                paper_name=paper,
                parsed_xml=parsed_xml,
                image_paths=image_paths,
                constants=constants,
            )
            _finalize_output_rows(paper_rows)
            all_rows.extend(paper_rows)
            print(f"  Extracted {len(paper_rows)} row(s) from {paper}")
        except Exception as e:
            if isinstance(e, RateLimitReachedError):
                raise
            print(f"  [error] figure extraction failed: {e}")
            row = {col: "ERROR" for col in COLUMNS}
            row["Paper"] = paper
            row[INTERNAL_NOTES_KEY] = str(e)
            all_rows.append(row)
        completed_papers.append(paper)
        _finalize_output_rows(all_rows)
        _save_progress(progress_path, completed_papers, all_rows)

    if not all_rows:
        print("\nNo data extracted.")
        return {}

    _finalize_output_rows(all_rows)
    output_columns = _select_output_columns(all_rows)
    df = pd.DataFrame(all_rows, columns=output_columns).fillna("NR")
    df.to_csv(output_csv, index=False)
    print(f"\nCSV saved  : {output_csv}")

    df.to_excel(output_xlsx, index=False)
    format_xlsx(output_xlsx)
    print(f"XLSX saved : {output_xlsx}")

    usage_df = _build_llm_usage_summary_df(mapping_records, all_rows)
    write_llm_usage_summary_xlsx(usage_df, output_usage_xlsx)
    print(f"LLM usage summary saved : {output_usage_xlsx}")

    audit_paths: Dict[str, str] = {}
    if save_audit_reports:
        try:
            from extraction_audit_reports import write_audit_report_set
            audit_papers = [rec.get("Paper", "") for rec in mapping_records if _safe_str(rec.get("Paper", "")).strip()]
            audit_paths = write_audit_report_set(
                output_root=audit_output_root,
                papers=audit_papers,
                parsed_root=parsed_xml_dir,
                pdf_base_dir=pdf_base_dir,
                all_rows=all_rows,
                pass1_by_paper=PASS1_CONSTANTS_BY_PAPER,
                panel_discovery_by_paper=PANEL_DISCOVERY_AUDIT_BY_PAPER,
                image_paths_by_paper=AUDIT_IMAGE_PATHS_BY_PAPER,
                load_parsed_article_fn=load_parsed_xml,
                match_image_to_figure_fn=match_image_to_xml_figure_exact,
                extract_caption_text_fn=_extract_figure_caption_text,
                hard_skip_fn=_should_hard_skip_non_target_figure,
            )
            print(f"Audit reports saved under : {audit_output_root}")
        except Exception as audit_exc:
            print(f"[warning] Audit report generation failed: {audit_exc}")

    if os.path.exists(progress_path):
        os.remove(progress_path)

    print(f"\n{'=' * 70}")
    print(f"Done. {len(df)} row(s) written across {len(mapping_records)} paper(s).")

    return {"data_csv": output_csv, "data_xlsx": output_xlsx, "llm_usage_xlsx": output_usage_xlsx, "audit_reports": audit_paths}



def _print_run_configuration() -> None:
    print("\n" + "=" * 70)
    print("Run configuration")
    print("-" * 70)
    print(f"OpenAI model          : {MODEL_ID}")
    print(f"OpenAI base URL       : {OPENAI_BASE_URL}")
    print(f"Figure schema         : {FIGURE_SCHEMA_MODE}")
    print(f"Max figure tokens     : {FIGURE_MAX_OUTPUT_TOKENS}")
    print(f"LLM call delay        : {LLM_CALL_DELAY_SECONDS} s")
    print(f"Throttle retries      : {LLM_THROTTLE_MAX_RETRIES}")
    print("=" * 70 + "\n")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Use RSC HTML text/caption artifacts or raw RSC HTML plus figure images "
            "to extract methane-decomposition datapoints. This script accepts curated images "
            "inside PDFs/<Paper>/images/ and can also use RSC HTML assets such as "
            "HTMLs/<Paper>/<Paper>_files/*-f1.gif, *-f2.gif, etc."
        )
    )
    parser.add_argument("--mapping-xlsx", default="RSC_DOIs.xlsx", help="Excel workbook with columns 'Paper' and 'DOIs'.")
    parser.add_argument("--html-root", default="./HTMLs", help="Root directory containing one folder per paper with RSC HTML-derived artifacts, raw .html/.htm files, and optional <Paper>_files/ figure GIFs.")
    parser.add_argument("--pdf-base-dir", default="./PDFs", help="Base directory containing one folder per paper with curated images inside PDFs/<Paper>/images/.")
    parser.add_argument("--figure-source", choices=["auto", "pdf", "html", "both"], default="auto", help="Which figure images to send. auto uses PDFs/<Paper>/images if present, otherwise HTMLs/<Paper>/<Paper>_files/*-fN.gif. html uses only HTML asset figures; both combines both sources.")
    parser.add_argument("--output-prefix", default="./rsc_extraction_results", help="Output path prefix without extension.")
    parser.add_argument("--paper", action="append", default=None, help="Process only this Paper value from the mapping workbook. Repeat for multiple papers.")
    parser.add_argument("--fresh-start", action="store_true", help="Ignore and remove any saved progress file before starting.")
    parser.add_argument("--prompts-dir", default=os.environ.get("PROMPTS_DIR", "./prompts"), help="Directory containing prompt template text files.")
    parser.add_argument("--resume-from-paper-number", "--resume-from-paper", type=int, default=None, help="Process only papers whose numeric Paper value, or workbook order if no number is present, is >= this number.")
    parser.add_argument("--send-all-figures-to-llm", "--no-figure-skip", "--disable-figure-skip", action="store_true", help="Disable the pre-LLM caption-based hard skip so every selected figure image is sent to the LLM.")
    parser.add_argument("--audit-output-dir", default=None, help="Directory where post-run audit reports are written. Default: <output-prefix>_audit_reports.")
    parser.add_argument("--no-audit-reports", action="store_true", help="Do not generate the post-run audit report workbooks.")
    parser.add_argument("--figure-max-output-tokens", type=int, default=FIGURE_MAX_OUTPUT_TOKENS, help="Maximum output tokens for figure/panel/single-series extraction calls.")
    parser.add_argument("--figure-schema-mode", choices=["compact", "full"], default=FIGURE_SCHEMA_MODE, help="Use compact point-only figure schema or full legacy schema.")

    parser.add_argument("--openai-model", default=os.environ.get("OPENAI_MODEL_ID", MODEL_ID), help="OpenAI model ID, e.g. gpt-5.5 or any model slug available in your account.")
    parser.add_argument("--openai-base-url", default=os.environ.get("OPENAI_BASE_URL", OPENAI_BASE_URL), help="OpenAI API base URL. Use https://api.openai.com/v1 for the direct OpenAI API.")
    parser.add_argument("--openai-reasoning-effort", default=os.environ.get("OPENAI_REASONING_EFFORT", OPENAI_REASONING_EFFORT), help="Reasoning effort for GPT-5 series models: low, medium, high, or xhigh where supported.")
    parser.add_argument("--openai-text-verbosity", default=os.environ.get("OPENAI_TEXT_VERBOSITY", OPENAI_TEXT_VERBOSITY), help="Text verbosity for GPT-5 series models: low, medium, or high where supported.")

    parser.add_argument("--llm-call-delay-seconds", type=float, default=LLM_CALL_DELAY_SECONDS, help="Sleep this many seconds before every OpenAI Responses API call. Use a small positive value to reduce throttling; set 0 to disable.")
    parser.add_argument("--throttle-max-retries", type=int, default=LLM_THROTTLE_MAX_RETRIES, help="Additional application-level retries after OpenAI rate/quota limiting retries are exhausted.")
    parser.add_argument("--throttle-base-delay-seconds", type=float, default=LLM_THROTTLE_BASE_DELAY_SECONDS, help="Initial sleep before retrying a throttled OpenAI request; doubles after each throttled attempt.")
    parser.add_argument("--throttle-max-delay-seconds", type=float, default=LLM_THROTTLE_MAX_DELAY_SECONDS, help="Maximum sleep between application-level throttling retries.")
    parser.add_argument("--throttle-jitter-seconds", type=float, default=LLM_THROTTLE_JITTER_SECONDS, help="Random jitter added to throttling backoff sleeps.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    _set_prompts_dir(args.prompts_dir)
    _set_openai_settings(
        model_id=args.openai_model,
        base_url=args.openai_base_url,
        reasoning_effort=args.openai_reasoning_effort,
        text_verbosity=args.openai_text_verbosity,
    )

    _set_figure_extraction_settings(
        figure_max_output_tokens=args.figure_max_output_tokens,
        figure_schema_mode=args.figure_schema_mode,
    )

    _set_llm_rate_limit_settings(
        call_delay_seconds=args.llm_call_delay_seconds,
        throttle_max_retries=args.throttle_max_retries,
        throttle_base_delay_seconds=args.throttle_base_delay_seconds,
        throttle_max_delay_seconds=args.throttle_max_delay_seconds,
        throttle_jitter_seconds=args.throttle_jitter_seconds,
    )

    if selected_provider() == "openai" and openai_client is None:
        raise RuntimeError(
            "OpenAI client is not configured. Set OPENAI_API_KEY. Optionally set OPENAI_BASE_URL and OPENAI_MODEL_ID."
        )

    _print_run_configuration()

    run_extraction_pipeline(
        mapping_xlsx=args.mapping_xlsx,
        parsed_xml_dir=args.html_root,
        pdf_base_dir=args.pdf_base_dir,
        output_prefix=args.output_prefix,
        fresh_start=args.fresh_start,
        paper_filters=args.paper,
        resume_from_paper_number=args.resume_from_paper_number,
        send_all_figures_to_llm=args.send_all_figures_to_llm,
        save_audit_reports=not args.no_audit_reports,
        audit_output_root=args.audit_output_dir,
        figure_source=args.figure_source,
    )


if __name__ == "__main__":
    main()
