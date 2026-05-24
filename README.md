# XLSX to JSON/CSV Extractor (POC)

This proof-of-concept extracts structured data from one or more XLSX files using a JSON configuration file.

## Why this approach

- **Dynamic file detection** by filename pattern (`filename_contains`) so changing numeric prefixes still work.
- **Dynamic row extraction**: tables stop only at true empty rows, so added/removed rows are handled automatically.
- **Config-driven extraction**: logic for file/sheet/table detection lives in `extractor_config.json`, not hard-coded in Python.

## Setup

```bash
python -m venv .venv
source .venv/bin/activate
pip install openpyxl
```

## Usage

Put source files in `input/` then run:

```bash
python extractor.py \
  --input-dir ./input \
  --config ./extractor_config.json \
  --output-json ./output/extracted.json \
  --output-csv-dir ./output/csv
```

## Current configured file type

- Matches files containing `PLI_Amadeus.xlsx` in the name (e.g., `329988_PLI_Amadeus.xlsx`).
- Uses sheet `Overview PLI Results`.
- Extracts:
  - key-values from column `C` (key) and `E` (value) until key is empty.
  - table with headers:
    `No., Company Code, Data Source, Current Status, IQR, Company Name, Net Cost Plus`.
  - table whose header starts with:
    `Summary Ranges - Percentile (Spreadsheet) Method`.

## Notes / best practices

- Keep extraction rules in config, not code, to simplify adding new file types.
- Validate config with tests as it grows.
- Prefer deterministic matching (`filename_contains`, fixed sheet names, explicit headers) over fuzzy heuristics.
