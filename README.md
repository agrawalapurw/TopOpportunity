# Contact to Company Enrichment

This project reads Excel sheets that contain website/domain columns (or contact_email), runs the same DIVE query logic used in domain enrichment, and writes back a best matched company name.

## Expected input columns

- `website` or `domain` (recommended)
- `contact_email` (optional, used for exact-email match when present)
- `ship_to_company_name` (optional, used to improve company choice)

Column name matching is case-insensitive and supports common variants. The script also auto-detects header rows when headers are not on the first row.

## Setup

From this folder:

```powershell
pip install -r requirements.txt
```

Optional credentials via environment variables:

```powershell
$env:DIVE_UID="your.uid@INFINEON.COM"
$env:DIVE_PWD="your_password"
```

## Configure

Copy `config.example.json` to `config.json` and edit paths/settings.

Suggested input folder structure:

- `input/` for your source workbooks
- `output/` for generated files

## Run

```powershell
python run_contact_company_enrichment.py
```

When run without `--workbook`, the script shows a terminal prompt so you can choose which file(s) to process.
If run in a non-interactive context (stdin EOF), the script will not default to all files; use one or more `--workbook` arguments in that case.

Optional arguments:

```powershell
python run_contact_company_enrichment.py `
  --input-dir "./input" `
  --glob "*.xlsx" `
  --workbook "OnlinESampling.xlsx" `
  --workbook "*July*.xlsx" `
  --chunk-size 1000
```

## Outputs

For each workbook:

- `<workbook>_company_enriched.xlsx`
- `<workbook>_unmatched_contacts.csv`

The enriched workbook preserves the original workbook's sheet formatting/layout and appends enrichment columns to the right.
It also includes a `Data_All Matches` tab with detailed row-level input-to-DIVE domain match records.

Added output columns per row:

- `input_domain`
- `matched_company_name`
- `company_match_method`
- `company_match_score`
- `is_company`

`company_match_method` values:

- `exact_email`
- `domain_plus_ship_to`
- `domain_frequency`
- `no_domain_match`
- `no_input_domain`

## Operations Documents

- Cloud server setup and scheduling: `docs/CLOUD_SERVER_SCHEDULING_GUIDE.md`
- End-user upload-only SOP: `docs/END_USER_UPLOAD_GUIDE.md`
