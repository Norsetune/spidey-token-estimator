# Spidey Context Budget Estimator

A small local tool for estimating whether a Spidey turn is likely to exceed a 100,000-token model context limit after files are parsed.

The key metric is **extracted text tokens**, not file size in MB. This tool extracts text from common file types and estimates tokens per file and for the whole turn.

## Supported files

- PDF (`.pdf`) via PyMuPDF
- Word (`.docx`) via python-docx
- Excel (`.xlsx`) via openpyxl
- CSV (`.csv`)
- Text/Markdown (`.txt`, `.md`)

OCR is intentionally not included in v1. Image-only PDFs may therefore be underestimated.

## Install

```bash
cd spidey_token_estimator
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Run the Streamlit app

```bash
streamlit run streamlit_app.py
```

Then open the local URL shown in Terminal and upload files.

## Run from command line

Estimate a folder of files:

```bash
python token_estimator.py /path/to/turn_files
```

Include a prompt file:

```bash
python token_estimator.py /path/to/turn_files --prompt-file prompt.txt
```

Write a JSON report:

```bash
python token_estimator.py /path/to/turn_files --prompt-file prompt.txt --json-out report.json
```

## Output example

```text
Estimated extracted-token budget
===============================================================================================
    Tokens   %Limit      Words       MB  Risk                    File
-----------------------------------------------------------------------------------------------
   315,000    31.5%    170,000    17.50  SAFE                    NASA_SOA_2024_full.pdf
    72,000     7.2%     42,000     4.20  SAFE                    NASA_CubeSat_101.pdf
     9,200     0.9%      6,500     0.06  SAFE                    prompt.txt
-----------------------------------------------------------------------------------------------
TOTAL: 396,200 / 1,000,000 tokens (39.6%)
STATUS: SAFE
```

## Notes and caveats

- This is an estimate, not the exact Claude tokenizer.
- If `tiktoken` is installed, the tool uses `cl100k_base` as a proxy tokenizer.
- If `tiktoken` is not installed, it falls back to a conservative character/word heuristic.
- Real task systems may add hidden overhead for file metadata, extracted tables, images, tool wrappers, and prompt templates.
- Treat anything above ~850k as risky.
