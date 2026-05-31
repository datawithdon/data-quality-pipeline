"""
Data Quality Pipeline — Sales Orders
-------------------------------------
Ingests a messy CSV, validates every row against a strict Pydantic schema,
exports clean records to CSV / Excel / JSON, and writes a full exceptions log.
"""

import re
import json
import logging
import pandas as pd
from pathlib import Path
from datetime import datetime
from typing import Optional

from pydantic import BaseModel, field_validator, ValidationError

# ---------------------------------------------------------------------------
# Paths — anchored to this file so the script works from any working directory
# ---------------------------------------------------------------------------
BASE_DIR   = Path(__file__).resolve().parent
RAW_FILE   = BASE_DIR / "data" / "raw_sales_orders.csv"
OUT_DIR    = BASE_DIR / "output"
OUT_CSV    = OUT_DIR / "clean_sales_orders.csv"
OUT_XLSX   = OUT_DIR / "clean_sales_orders.xlsx"
OUT_JSON   = OUT_DIR / "clean_sales_orders.json"
OUT_ERR    = OUT_DIR / "exceptions_log.csv"

OUT_DIR.mkdir(parents=True, exist_ok=True)

MAX_FILE_SIZE_MB = 500

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
VALID_STATUSES = {"PENDING", "SHIPPED", "DELIVERED", "CANCELLED"}

# Fuzzy status normalisation map — covers abbreviations / typos found in data
STATUS_MAP = {
    "pend":      "PENDING",
    "pending":   "PENDING",
    "ship":      "SHIPPED",
    "shipped":   "SHIPPED",
    "shippd":    "SHIPPED",
    "dlvd":      "DELIVERED",
    "delivered": "DELIVERED",
    "cancld":    "CANCELLED",
    "cancelled": "CANCELLED",
    "canceled":  "CANCELLED",
}

# All date formats that appear in the wild
DATE_FORMATS = [
    "%m/%d/%Y",   # 01/15/2024
    "%Y-%m-%d",   # 2024-01-15
    "%b %d %Y",   # Jan 15 2024
    "%B %d %Y",   # January 15 2024
]

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _parse_currency(raw) -> Optional[float]:
    """Strip $ / commas and coerce to float. Returns None for blanks."""
    if pd.isna(raw) or str(raw).strip() == "":
        return None
    cleaned = re.sub(r"[\$,\s]", "", str(raw))
    try:
        return float(cleaned)
    except ValueError:
        return None


def _parse_date(raw) -> Optional[str]:
    """Try every known date format; return ISO string or None."""
    if pd.isna(raw) or str(raw).strip() == "":
        return None
    s = str(raw).strip()
    # Collapse multiple spaces (e.g. "Jan  5 2024")
    s = re.sub(r"\s+", " ", s)
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt).date().isoformat()
        except ValueError:
            continue
    return None  # unparseable — Pydantic will raise


def _parse_bool(raw) -> Optional[bool]:
    """Map Y/N, yes/no, 1/0, True/False to bool."""
    if pd.isna(raw) or str(raw).strip() == "":
        return None
    s = str(raw).strip().lower()
    if s in {"y", "yes", "1", "true"}:
        return True
    if s in {"n", "no", "0", "false"}:
        return False
    return None


def _normalize_status(raw) -> Optional[str]:
    """Upper-case, strip, then apply fuzzy map."""
    if pd.isna(raw) or str(raw).strip() == "":
        return None
    key = str(raw).strip().lower()
    return STATUS_MAP.get(key, str(raw).strip().upper())


_FORMULA_CHARS = frozenset("=+-@\t\r")

def _sanitize_cell(value):
    """Prefix cells that start with formula trigger characters to prevent CSV injection."""
    if isinstance(value, str) and value and value[0] in _FORMULA_CHARS:
        return "'" + value
    return value

# ---------------------------------------------------------------------------
# Pydantic Model
# ---------------------------------------------------------------------------

class SalesOrder(BaseModel):
    order_id:      int
    customer_name: str
    order_date:    str          # stored as ISO date string
    delivery_date: Optional[str] = None
    quantity:      int
    unit_price:    float
    total_amount:  float
    status:        str
    is_priority:   bool
    region:        str

    # --- field-level validators ---

    @field_validator("customer_name", "region", mode="before")
    @classmethod
    def strip_and_require(cls, v):
        if v is None or (isinstance(v, float) and pd.isna(v)):
            raise ValueError("Field is required and cannot be null")
        s = str(v).strip()
        if not s:
            raise ValueError("Field cannot be blank after stripping whitespace")
        return s

    @field_validator("order_date", mode="before")
    @classmethod
    def validate_order_date(cls, v):
        parsed = _parse_date(v)
        if parsed is None:
            raise ValueError(f"Cannot parse order_date: {v!r}")
        return parsed

    @field_validator("delivery_date", mode="before")
    @classmethod
    def validate_delivery_date(cls, v):
        if v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == "":
            return None
        parsed = _parse_date(v)
        if parsed is None:
            raise ValueError(f"Cannot parse delivery_date: {v!r}")
        return parsed

    @field_validator("quantity", mode="before")
    @classmethod
    def validate_quantity(cls, v):
        try:
            qty = int(float(str(v)))
        except (ValueError, TypeError):
            raise ValueError(f"quantity must be an integer, got: {v!r}")
        if qty < 1:
            raise ValueError(f"quantity must be >= 1, got: {qty}")
        return qty

    @field_validator("unit_price", "total_amount", mode="before")
    @classmethod
    def validate_currency(cls, v):
        result = _parse_currency(v)
        if result is None:
            raise ValueError(f"Cannot parse currency value: {v!r}")
        if result < 0:
            raise ValueError(f"Currency value cannot be negative: {result}")
        return result

    @field_validator("status", mode="before")
    @classmethod
    def validate_status(cls, v):
        normalized = _normalize_status(v)
        if normalized not in VALID_STATUSES:
            raise ValueError(
                f"Invalid status {v!r} (normalized: {normalized!r}). "
                f"Expected one of {sorted(VALID_STATUSES)}"
            )
        return normalized

    @field_validator("is_priority", mode="before")
    @classmethod
    def validate_priority(cls, v):
        result = _parse_bool(v)
        if result is None:
            raise ValueError(f"Cannot interpret is_priority: {v!r}")
        return result

    @field_validator("order_id", mode="before")
    @classmethod
    def validate_order_id(cls, v):
        if v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == "":
            raise ValueError("order_id is required")
        try:
            return int(float(str(v)))
        except (ValueError, TypeError):
            raise ValueError(f"order_id must be an integer, got: {v!r}")

# ---------------------------------------------------------------------------
# Column name normalisation
# ---------------------------------------------------------------------------
COLUMN_MAP = {
    "order id":        "order_id",
    "orderid":         "order_id",
    "order_id":        "order_id",
    "cust_nm":         "customer_name",
    "customer name":   "customer_name",
    "customer_name":   "customer_name",
    "ord date":        "order_date",
    "order date":      "order_date",
    "order_date":      "order_date",
    "deliv_date":      "delivery_date",
    "delivery date":   "delivery_date",
    "delivery_date":   "delivery_date",
    "qty":             "quantity",
    "quantity":        "quantity",
    "unit price ($)":  "unit_price",
    "unit price":      "unit_price",
    "unit_price":      "unit_price",
    "total amt":       "total_amount",
    "total amount":    "total_amount",
    "total_amount":    "total_amount",
    "status":          "status",
    "priority flag":   "is_priority",
    "is_priority":     "is_priority",
    "region":          "region",
}


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [COLUMN_MAP.get(c.strip().lower(), c.strip().lower()) for c in df.columns]
    return df

# ---------------------------------------------------------------------------
# Main pipeline
# ---------------------------------------------------------------------------

def run_pipeline(raw_file: Path = RAW_FILE) -> None:
    log.info("=" * 60)
    log.info("DATA QUALITY PIPELINE — SALES ORDERS")
    log.info("=" * 60)

    # 1. Load raw CSV
    if not raw_file.exists():
        raise FileNotFoundError(f"Input file not found: {raw_file}")

    file_size_mb = raw_file.stat().st_size / (1024 * 1024)
    if file_size_mb > MAX_FILE_SIZE_MB:
        raise ValueError(f"Input file too large: {file_size_mb:.1f} MB (limit: {MAX_FILE_SIZE_MB} MB)")

    log.info(f"Loading raw file: {raw_file}  ({file_size_mb:.2f} MB)")
    df = pd.read_csv(raw_file, dtype=str, keep_default_na=False)
    df.replace("", pd.NA, inplace=True)
    log.info(f"Loaded {len(df)} raw rows, {len(df.columns)} columns")

    # 2. Normalise column names
    df = normalize_columns(df)
    log.info(f"Columns after normalisation: {list(df.columns)}")

    # 3. Validate row-by-row
    clean_rows   = []
    error_rows   = []

    for idx, row in df.iterrows():
        record = row.to_dict()
        try:
            validated = SalesOrder(**record)
            clean_rows.append(validated.model_dump())
        except ValidationError as e:
            failed_fields = [err["loc"][0] if err["loc"] else "row" for err in e.errors()]
            error_summary = "; ".join(
                f"{err['loc'][0] if err['loc'] else 'row'}: {err['msg']}"
                for err in e.errors()
            )
            # Only expose the values of fields that actually failed — not the full row —
            # to limit PII surface in the exceptions log.
            failed_values = json.dumps({f: record.get(f) for f in failed_fields})
            error_rows.append({
                "source_row_index":  idx,
                "failed_fields":     ", ".join(str(f) for f in failed_fields),
                "failed_field_values": failed_values,
                "validation_errors": error_summary,
            })
            log.warning(f"Row {idx} FAILED — {error_summary}")

    total     = len(df)
    passed    = len(clean_rows)
    failed    = len(error_rows)
    pass_rate = (passed / total * 100) if total else 0

    log.info("-" * 60)
    log.info(f"RESULTS  Total={total}  Passed={passed}  Failed={failed}  Pass-rate={pass_rate:.1f}%")
    log.info("-" * 60)

    # 4. Export clean data
    if clean_rows:
        clean_df = pd.DataFrame(clean_rows)

        # Sanitize string columns against CSV formula injection before any export
        str_cols = clean_df.select_dtypes(include="object").columns
        clean_df[str_cols] = clean_df[str_cols].map(_sanitize_cell)

        clean_df.to_csv(OUT_CSV, index=False)
        log.info(f"Clean CSV  -> {OUT_CSV}")

        clean_df.to_excel(OUT_XLSX, index=False, engine="openpyxl")
        log.info(f"Clean XLSX -> {OUT_XLSX}")

        clean_df.to_json(OUT_JSON, orient="records", indent=2, date_format="iso")
        log.info(f"Clean JSON -> {OUT_JSON}")
    else:
        log.warning("No clean rows to export.")

    # 5. Export exceptions log
    if error_rows:
        err_df = pd.DataFrame(error_rows)
        err_df.to_csv(OUT_ERR, index=False)
        log.info(f"Exceptions -> {OUT_ERR}  ({failed} rows)")
    else:
        log.info("No validation errors — exceptions log not written.")

    log.info("Pipeline complete.")


if __name__ == "__main__":
    run_pipeline()
