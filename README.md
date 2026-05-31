# Data Quality Pipeline

A production-grade ETL pipeline that ingests structurally inconsistent CSV files, validates every row against a strict schema using Pydantic, and exports clean records to CSV, Excel, and JSON — with a full exceptions log for any rows that fail.

## Features

- Fuzzy column name normalisation — maps `CUST_NM`, `customer name`, `Customer_Name` to the same target field
- Mixed date format parsing — `01/15/2024`, `2024-01-15`, `Jan 15 2024` all resolve to ISO `2024-01-15`
- Currency string coercion — `$1,234.56` → `1234.56`
- Boolean variant handling — `Y / N`, `yes / no`, `1 / 0`, `True / False`
- Status fuzzy-mapping with enum enforcement — `"pend"`, `"PEND"`, `"Pending"` → `PENDING`
- Row-level isolation — one bad row never aborts the run
- Exceptions log captures only the failed fields (not the full row) to limit PII exposure
- CSV formula injection sanitization on all string output columns
- Paths anchored to `__file__` — runs correctly from any working directory

## Target Schema

| Field | Type | Constraints |
|---|---|---|
| `order_id` | integer | required, unique |
| `customer_name` | string | required, non-blank |
| `order_date` | string (ISO date) | required |
| `delivery_date` | string (ISO date) | optional |
| `quantity` | integer | required, ≥ 1 |
| `unit_price` | float | required, ≥ 0 |
| `total_amount` | float | required, ≥ 0 |
| `status` | string (enum) | `PENDING` \| `SHIPPED` \| `DELIVERED` \| `CANCELLED` |
| `is_priority` | boolean | required |
| `region` | string | required, non-blank |

## Project Structure

```
data-quality-pipeline/
├── pipeline.py               # ETL engine
├── requirements.txt
├── data/
│   └── raw_sales_orders.csv  # Demo input (synthetic data)
└── output/                   # Generated on run — gitignored
    ├── clean_sales_orders.csv
    ├── clean_sales_orders.xlsx
    ├── clean_sales_orders.json
    └── exceptions_log.csv
```

## Quickstart

```bash
pip install -r requirements.txt
python pipeline.py
```

Sample output:

```
13:38:00  INFO   Loaded 15 raw rows, 10 columns
13:38:00  INFO   Columns after normalisation: ['order_id', 'customer_name', ...]
13:38:00  WARNING  Row 3 FAILED — quantity: Value error, quantity must be >= 1, got: 0
13:38:00  WARNING  Row 4 FAILED — customer_name: Field is required and cannot be null
13:38:00  INFO   RESULTS  Total=15  Passed=11  Failed=4  Pass-rate=73.3%
13:38:00  INFO   Clean CSV  -> output/clean_sales_orders.csv
13:38:00  INFO   Clean XLSX -> output/clean_sales_orders.xlsx
13:38:00  INFO   Clean JSON -> output/clean_sales_orders.json
13:38:00  INFO   Exceptions -> output/exceptions_log.csv  (4 rows)
```

## Exceptions Log

Failed rows are written to `output/exceptions_log.csv` with three columns:

| Column | Description |
|---|---|
| `source_row_index` | Original row number in the input file |
| `failed_fields` | Comma-separated list of fields that failed validation |
| `failed_field_values` | JSON object of only the failed field values |
| `validation_errors` | Full Pydantic error messages |

Only the values of fields that actually failed are stored — not the entire row — to minimise PII exposure in the error log.

## Adapting to Your Own Data

1. **Add new column mappings** — extend `COLUMN_MAP` in `pipeline.py` with any raw header variants your source uses.
2. **Add new date formats** — append to the `DATE_FORMATS` list.
3. **Extend the status map** — add abbreviation → canonical value entries to `STATUS_MAP`.
4. **Change the input file** — pass a different path to `run_pipeline(Path("your/file.csv"))`.

## Security Notes

- Real data files placed in `data/` are gitignored (only the synthetic demo file is committed).
- The `output/` directory is gitignored entirely.
- All string output columns are sanitized against CSV formula injection before export.
- File size is capped at 500 MB before loading into memory.

## Requirements

- Python 3.9+
- pandas ≥ 2.0
- pydantic ≥ 2.0
- openpyxl ≥ 3.1
