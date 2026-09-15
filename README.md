# Knowledge extraction from articles on TCD of Methane using Large Language Models

Publisher-specific pipelines for extracting structured data from article text metadata and curated figure images along with audit reports for benchmarking.

## Included extractors

| Publisher | Extractor |
|---|---|
| Elsevier | `extractors/elsevier_extractor.py` |
| MDPI | `extractors/mdpi_extractor.py` |
| RSC | `extractors/rsc_extractor.py` |
| Wiley | `extractors/wiley_extractor.py` |

Every publisher extractor supports both `openai` and `bedrock`. The journal-specific logic selects and loads the correct article metadata; the shared AWS Bedrock Converse transport handles Claude, Sonnet, Scout, Maverick, and other Bedrock models available to your account.

## Install

Use Python 3.10 or newer and the following packages as well:

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
```

### AWS Bedrock

```dotenv
AWS_REGION=us-east-1
AWS_ACCESS_KEY_ID=your_access_key_id
AWS_SECRET_ACCESS_KEY=your_secret_access_key
MODEL_ID=your_bedrock_model_or_inference_profile_id
```
Model IDs for openAI can be found at 'https://developers.openai.com/api/docs/models/all'. These are the model IDs that allowed us to invoke the AWS models:

| Model | Model ID |
|---|---|
| Llama 4 Scout 17B | us.meta.llama4-scout-17b-instruct-v1:0 |
| Llama 4 Maverick 17B | us.meta.llama4-maverick-17b-instruct-v1:0 |
| Claude Sonnet 4.6 | us.anthropic.claude-sonnet-4-6 |
| Claude Opus 4.6 | us.anthropic.claude-opus-4-6-v1 |

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

Run `python main.py --list` to display the supported publisher/provider combinations. You can also run any individual extractor directly after setting `LLM_PROVIDER` in `.env`.

## Token limits

Equivalent OpenAI and Bedrock calls use the same limits: 2,200 tokens for panel/series discovery, 1,400 for constant extraction, 1,100 for catalyst-profile enrichment, and 12,000 by default for figure extraction. Override the figure budget with `--figure-max-output-tokens` or `FIGURE_MAX_OUTPUT_TOKENS`.

## Prompts

The `prompts/` folder contains all model instructions. The dispatcher supplies this directory automatically. Keep the prompt filenames unchanged unless you also update the corresponding extractor's `PROMPT_FILES` mapping.

## Inputs and outputs

Each publisher script documents its expected input layout and default mapping workbook name in its module header and command-line help. A run writes CSV, XLSX, LLM usage, and audit-report outputs beside the selected output prefix.
