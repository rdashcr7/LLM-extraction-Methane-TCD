# Methane-Decomposition Data Extractor

This command-line pipeline extracts paper-level reaction constants from Elsevier XML-derived text and methane-decomposition datapoints from curated figure images. It writes a consolidated CSV, an Excel workbook, and an LLM-usage summary.

## Repository layout

```text
main.py       # Command-line entry point and runtime configuration
helper.py     # Extraction, normalization, API, and output-writing logic
prompts/      # External prompt templates loaded at runtime
```

## Requirements

Use Python 3.10 or newer, then install the required packages:

```bash
pip install openai boto3 pandas python-dotenv openpyxl pillow
```

Create a `.env` file next to `main.py`:

```dotenv
OPENAI_API_KEY=your_api_key_here
OPENAI_MODEL_ID=gpt-5.5
```

Do not commit `.env` or any file containing an API key.

## Expected input layout

```text
Elsevier dois.xlsx
XMLs/
  Paper 1/
    fulltext.txt
    sentences.txt
    figure_captions.csv
    tables.json
    ...
PDFs/
  Paper 1/
    images/
      figure_1.png
      figure_2.png
```

The mapping workbook must contain `Paper` and `DOIs` columns. The loader also supports the earlier XML artifact filenames used by the original pipeline.

## Run

From the repository folder:

```bash
python main.py \
  --mapping-xlsx "Elsevier dois.xlsx" \
  --xml-root ./XMLs \
  --pdf-base-dir ./PDFs \
  --output-prefix ./elsevier_extraction_results
```

Useful options:

- `--paper "Paper 3"` processes one paper; repeat the option for more papers.
- `--fresh-start` discards saved progress from a previous run.
- `--resume-from-paper-number 12` starts at paper 12.
- `--send-all-figures-to-llm` disables caption-based figure skipping.
- `--no-audit-reports` skips optional audit-report generation.
- `--prompts-dir /path/to/prompts` uses another prompt-template directory.

The run produces `<output-prefix>.csv`, `<output-prefix>.xlsx`, and `<output-prefix>_llm_usage_summary.xlsx`. Progress is saved after each paper and is removed after a successful run.

## Prompt templates

All model instructions are in `prompts/`, not embedded in the Python source. Keep the supplied filenames unchanged unless you also update `PROMPT_FILES` in `helper.py`.

## Optional audit reports

If `extraction_audit_reports.py` is in the repository, the pipeline will use it to generate post-run audit workbooks. If it is absent, extraction still completes and prints a warning instead.
