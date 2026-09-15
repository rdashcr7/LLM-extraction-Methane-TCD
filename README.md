# Knowledge extraction from articles on TCD of Methane using Large Language Models

This repository contains files for knowledge extraction from articles on TCD of Methane from metadata files (HTML/XML) using large language models.

## Preparing Input files

These scripts were used to extract the meta data in XML/HTML format of the articles and images from the PDF format of the articles. 

| Publisher | Script |
|---|---|
| Elsevier | `meta data input/elsevier.py` |
| MDPI | `meta data input/mdpi.py` |
| RSC | `meta data input/rsc.py` |
| Wiley | `meta data input/wiley.py` |

## Extractor files

These are the extractor files that main.py relies on for extracting from different providers' articles.

| Publisher | Extractor |
|---|---|
| Elsevier | `extractors/elsevier_extractor.py` |
| MDPI | `extractors/mdpi_extractor.py` |
| RSC | `extractors/rsc_extractor.py` |
| Wiley | `extractors/wiley_extractor.py` |

## Install

Use Python 3.10 or newer and install the following packages as well:

```bash
pip install openai boto3 pandas python-dotenv openpyxl pillow beautifulsoup4 fitz pymupdf requests chemdataextractor
```

## Configure a model provider

Create a `.env` file in the repository root.

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
Model IDs for openAI can be found at 'https://developers.openai.com/api/docs/models/all'. These are the model IDs that allowed us to invoke the AWS models (When we used the model IDs on AWS website, the models could not be invoked):

| Model | Model ID |
|---|---|
| Llama 4 Scout 17B | us.meta.llama4-scout-17b-instruct-v1:0 |
| Llama 4 Maverick 17B | us.meta.llama4-maverick-17b-instruct-v1:0 |
| Claude Sonnet 4.6 | us.anthropic.claude-sonnet-4-6 |
| Claude Opus 4.6 | us.anthropic.claude-opus-4-6-v1 |

## Run

Use `main.py` to select the publisher and provider. All other options are forwarded directly to the chosen extractor:

```bash
python main.py --publisher elsevier --provider bedrock --mapping-xlsx "Elsevier dois.xlsx" --xml-root ./XMLs --pdf-base-dir ./PDFs --output-prefix ./elsevier_results
```

Example OpenAI run for RSC:

```bash
python main.py --publisher rsc --provider openai --mapping-xlsx "RSC_DOIs.xlsx" --html-root ./HTMLs --pdf-base-dir ./PDFs --output-prefix ./rsc_results  
```

Run `python main.py --list` to display the supported publisher/provider combinations. Individual extractors can be run directly after setting `LLM_PROVIDER` in `.env`.


## Prompts

The `prompts/` folder contains all model instructions. By default, the extractor accesses the prompts in this directory.

## Inputs and outputs

Each publisher script documents its expected input layout and default mapping workbook name in its module header and command-line help.
