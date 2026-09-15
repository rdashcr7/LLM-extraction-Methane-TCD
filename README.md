# Multi-Journal Methane-Decomposition Data Extractors

Publisher-specific pipelines for extracting paper-level reaction constants from article text and methane-conversion datapoints from curated figure images.

## Included extractors

| Publisher | OpenAI | AWS Bedrock |
|---|---:|---:|
| Elsevier | `helper.py` | `extractors/elsevier_bedrock.py` |
| MDPI | `extractors/mdpi_openai.py` | Not included |
| RSC | `extractors/rsc_openai.py` | Not included |
| Wiley | `extractors/wiley_openai.py` | Not included |

The repository includes the exact publisher-specific scripts supplied for each combination. Do not use the OpenAI scripts with Bedrock credentials, or the Bedrock script with an OpenAI API key.

## Install

Use Python 3.10 or newer:

```bash
pip install openai boto3 pandas python-dotenv openpyxl pillow
```

## Configure a model provider

Create a `.env` file in the repository root. Never commit it.

### OpenAI

```dotenv
OPENAI_API_KEY=your_api_key
OPENAI_MODEL_ID=gpt-5.5
OPENAI_BASE_URL=https://api.openai.com/v1
OPENAI_REASONING_EFFORT=low
OPENAI_TEXT_VERBOSITY=low
```

### AWS Bedrock

```dotenv
AWS_REGION=us-east-1
AWS_ACCESS_KEY_ID=your_access_key_id
AWS_SECRET_ACCESS_KEY=your_secret_access_key
MODEL_ID=your_bedrock_model_or_inference_profile_id
```

`BEDROCK_MODEL_ID` may be used instead of `MODEL_ID`. `AWS_SESSION_TOKEN` is required when using temporary AWS credentials. The Bedrock script supports Claude and Meta models available to your AWS account and region, including the relevant Claude 4.6, Sonnet 4.6, Scout, and Maverick model IDs or inference profiles.

## Run

Use `main.py` to select the publisher and provider. All other options are forwarded directly to the chosen extractor:

```bash
python main.py --publisher elsevier --provider bedrock \
  --mapping-xlsx "Elsevier dois.xlsx" \
  --xml-root ./XMLs \
  --pdf-base-dir ./PDFs \
  --output-prefix ./elsevier_results
```

Example OpenAI run for RSC:

```bash
python main.py --publisher rsc --provider openai \
  --mapping-xlsx "RSC_DOIs.xlsx" \
  --html-root ./HTMLs \
  --pdf-base-dir ./PDFs \
  --output-prefix ./rsc_results
```

Run `python main.py --list` to display the bundled combinations. You can also run any individual extractor directly.

## Token limits

Equivalent OpenAI and Bedrock calls use the same limits: 2,200 tokens for panel/series discovery, 1,400 for constant extraction, 1,100 for catalyst-profile enrichment, and 12,000 by default for figure extraction. Override the figure budget with `--figure-max-output-tokens` or `FIGURE_MAX_OUTPUT_TOKENS`.

## Prompts

The `prompts/` folder contains all model instructions. The dispatcher supplies this directory automatically. Keep the prompt filenames unchanged unless you also update the corresponding extractor's `PROMPT_FILES` mapping.

## Inputs and outputs

Each publisher script documents its expected input layout and default mapping workbook name in its module header and command-line help. A run writes CSV, XLSX, LLM usage, and (when available) audit-report outputs beside the selected output prefix.
