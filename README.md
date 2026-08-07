# Performance Breaches Extractor

This is the second tool, designed to extract `Performance Breaches` chunks from a `.docx`
monitoring memo and turn them into row-wise records.

It is intentionally built to run **without installation** and uses only the Python standard
library.

## Run directly

```bash
python3 extract_breaches.py input.docx --output-dir out
```

Quick smoke test:

```bash
python3 extract_breaches.py smoke_input.docx --output-dir smoke_out
```

Outputs:

- `out/breach_records.csv`
- `out/breach_records.json`

## What one record means

One `Severity of Breach Identified:` chunk becomes one output record.

## Current schema

- `group_name`
- `chunk_number`
- `severity_raw`
- `record_status`
- `model_risk_rating`
- `model_ids`
- `model_names`
- `explanation_of_breach`
- `first_breach_identified`
- `second_breach_identified`
- `third_breach_identified`
- `plan_to_address`
- `remediation_date`

## Current rules

- Only `.docx` is supported in this first version
- Extraction starts after the `Performance Breaches:` heading
- Each `Severity of Breach Identified:` heading starts a new chunk
- Multiple model IDs are stored in one cell joined by ` | `
- Multiple model names are stored in one cell joined by ` | `
- Placeholder values like `Select date...` or `Click or tap to enter a date...` are converted to blanks
- `record_status` is determined from the title color:
  - black / automatic => `active`
  - anything else => `decommissioned`

## Current limitation

This version only works on a **real text-based DOCX** where Word paragraphs and tables are
present in the document XML.

It does **not** extract data from:

- screenshots pasted into Word
- image-only PDF exports converted into DOCX
- photos of a screen or monitor

For those cases, OCR must happen first, and then the OCR output has to be turned into a real
editable DOCX before this extractor can read it.
