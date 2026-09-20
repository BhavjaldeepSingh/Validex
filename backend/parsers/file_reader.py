import os
import uuid
import csv
import re
import openpyxl
import pandas as pd
from datetime import datetime
from typing import Tuple, List, Optional, Any, Dict

from config import (
    TEMP_DIR,
    DEFAULT_PAYMENT_MODES,
    DEFAULT_PAYMENT_STATUSES,
    DEFAULT_DATE_FORMATS,
    SEMANTIC_KEYWORDS
)
from schemas import (
    DataTypeEnum,
    SemanticRoleEnum,
    ColumnRuleDefinition,
    DetectedColumn
)


def get_file_extension(filename: str) -> str:
    parts = filename.rsplit(".", 1)
    return parts[1].lower() if len(parts) > 1 else ""


def save_temp_file(file_bytes: bytes, filename: str) -> Tuple[str, str, str]:
    file_token = str(uuid.uuid4())
    ext = get_file_extension(filename)
    safe_filename = re.sub(r"[^a-zA-Z0-9_.-]", "_", filename)
    target_path = os.path.join(TEMP_DIR, f"{file_token}_{safe_filename}")
    with open(target_path, "wb") as f:
        f.write(file_bytes)
    return file_token, target_path, ext


def find_temp_file(file_token: str) -> Optional[Tuple[str, str, str]]:
    if not os.path.exists(TEMP_DIR):
        return None
    for fname in os.listdir(TEMP_DIR):
        if fname.startswith(f"{file_token}_"):
            full_path = os.path.join(TEMP_DIR, fname)
            original_filename = fname.split("_", 1)[1] if "_" in fname else fname
            ext = get_file_extension(original_filename)
            return full_path, original_filename, ext
    return None


def read_file_to_dataframe(
    file_path: str,
    file_format: str,
    sheet_name: Optional[str] = None
) -> Tuple[pd.DataFrame, List[str], Optional[str]]:
    """
    Reads a file into a DataFrame regardless of format (CSV, TSV, XLSX, XLS, JSON).
    Returns (df, available_sheets, active_sheet).
    """
    available_sheets: List[str] = []
    active_sheet: Optional[str] = None

    if file_format in ["xlsx", "xls"]:
        wb = openpyxl.load_workbook(file_path, read_only=True)
        available_sheets = wb.sheetnames
        active_sheet = wb.active.title if wb.active else available_sheets[0]
        wb.close()

        target_sheet = sheet_name if (sheet_name and sheet_name in available_sheets) else active_sheet
        df = pd.read_excel(file_path, sheet_name=target_sheet)
        return df, available_sheets, target_sheet

    elif file_format in ["csv", "tsv"]:
        sep = "\t" if file_format == "tsv" else None
        # Try multiple encodings
        encodings = ["utf-8", "utf-8-sig", "latin-1", "cp1252"]
        last_error = None
        df = None

        for enc in encodings:
            try:
                if sep:
                    df = pd.read_csv(file_path, encoding=enc, sep=sep)
                else:
                    # Sniff delimiter if CSV
                    with open(file_path, "r", encoding=enc, errors="ignore") as f:
                        sample = f.read(4096)
                        try:
                            dialect = csv.Sniffer().sniff(sample, delimiters=",\t;|")
                            delimiter = dialect.delimiter
                        except Exception:
                            delimiter = ","
                    df = pd.read_csv(file_path, encoding=enc, sep=delimiter)
                break
            except Exception as e:
                last_error = e
                continue

        if df is None:
            raise ValueError(f"Unable to parse CSV/TSV file: {last_error}")

        return df, [], None

    elif file_format == "json":
        try:
            df = pd.read_json(file_path)
        except Exception:
            try:
                df = pd.read_json(file_path, lines=True)
            except Exception as e:
                raise ValueError(f"Unable to parse JSON file: {e}")
        return df, [], None

    else:
        raise ValueError(f"Unsupported file format: .{file_format}. Supported: csv, tsv, xlsx, xls, json")


def infer_column_rules(df: pd.DataFrame) -> List[DetectedColumn]:
    """
    Examines DataFrame columns, infers primitive data types and semantic roles,
    and produces sensible suggested validation constraints.
    """
    detected_columns: List[DetectedColumn] = []
    country_col_name: Optional[str] = None

    # First pass: look for a country column to use as reference for phone validation
    for col in df.columns:
        c_lower = str(col).strip().lower()
        if any(kw in c_lower for kw in SEMANTIC_KEYWORDS["country_code"]):
            country_col_name = str(col)
            break

    for col in df.columns:
        col_str = str(col)
        col_lower = col_str.strip().lower()
        series = df[col].dropna()
        sample_vals = series.head(5).tolist()
        sample_vals_cleaned = [
            str(v) if not isinstance(v, (int, float, bool)) else v
            for v in sample_vals
        ]

        inferred_type = DataTypeEnum.string
        inferred_role = SemanticRoleEnum.general
        is_unique = False
        is_required = True
        min_val = None
        max_val = None
        allowed_vals = None
        date_formats = None
        country_ref = None

        # Sample string representation
        str_samples = [str(x).strip() for x in series.head(50)]

        # Check unique ID - match word tokens to prevent 'amount_paid' matching 'id'
        tokens = re.split(r"[\s_\-]+", col_lower)
        is_id_match = (
            col_lower in ["id", "uuid", "transaction_id", "txn_id", "order_id", "invoice_id", "invoice_no", "ref_no", "reference_id"]
            or (any(prefix in tokens for prefix in ["transaction", "order", "invoice", "receipt", "txn"]) and "id" in tokens)
            or (col_lower.endswith("_id") and not any(skip in tokens for skip in ["product", "item", "category", "status", "country", "customer", "user", "paid"]))
            or (tokens == ["id"])
        )
        if is_id_match:
            inferred_role = SemanticRoleEnum.unique_id
            is_unique = True
            is_required = True

        # Check Phone
        elif any(kw in col_lower for kw in SEMANTIC_KEYWORDS["phone_number"]):
            inferred_role = SemanticRoleEnum.phone_number
            country_ref = country_col_name
            is_required = True

        # Check Country
        elif any(kw in col_lower for kw in SEMANTIC_KEYWORDS["country_code"]):
            inferred_role = SemanticRoleEnum.country_code
            is_required = True

        # Check Email
        elif any(kw in col_lower for kw in SEMANTIC_KEYWORDS["email"]):
            inferred_role = SemanticRoleEnum.email
            is_required = True

        # Check Payment Mode
        elif any(kw in col_lower for kw in SEMANTIC_KEYWORDS["payment_mode"]):
            inferred_role = SemanticRoleEnum.payment_mode
            allowed_vals = DEFAULT_PAYMENT_MODES
            is_required = True

        # Check Payment Status
        elif any(kw in col_lower for kw in SEMANTIC_KEYWORDS["payment_status"]):
            inferred_role = SemanticRoleEnum.payment_status
            allowed_vals = DEFAULT_PAYMENT_STATUSES
            is_required = True

        # Check Date
        elif any(kw in col_lower for kw in SEMANTIC_KEYWORDS["date"]):
            inferred_role = SemanticRoleEnum.date
            inferred_type = DataTypeEnum.date
            date_formats = DEFAULT_DATE_FORMATS
            is_required = True

        # Check Time
        elif any(kw in col_lower for kw in SEMANTIC_KEYWORDS["time"]):
            inferred_role = SemanticRoleEnum.time
            inferred_type = DataTypeEnum.time
            is_required = False

        # Check Quantity
        elif any(kw in col_lower for kw in SEMANTIC_KEYWORDS["quantity"]):
            inferred_role = SemanticRoleEnum.quantity
            inferred_type = DataTypeEnum.integer
            min_val = 1.0
            is_required = True

        # Check Amount / Price
        elif any(kw in col_lower for kw in SEMANTIC_KEYWORDS["amount"]):
            inferred_role = SemanticRoleEnum.amount
            inferred_type = DataTypeEnum.float
            min_val = 0.0
            is_required = True

        # Check Customer Name
        elif any(kw in col_lower for kw in SEMANTIC_KEYWORDS["customer_name"]):
            inferred_role = SemanticRoleEnum.customer_name
            inferred_type = DataTypeEnum.string
            is_required = True

        # If not matched by keyword, infer by data inspection
        else:
            # Check if numeric
            try:
                pd.to_numeric(series.head(50))
                # Check if integers or floats
                if all(isinstance(x, (int, complex)) or (isinstance(x, float) and x.is_integer()) for x in series.head(50)):
                    inferred_type = DataTypeEnum.integer
                else:
                    inferred_type = DataTypeEnum.float
                min_val = 0.0
            except Exception:
                # Check if dates
                date_match = True
                for s in str_samples[:10]:
                    if not s or not re.search(r"\d{2,4}[-/\.]\d{1,2}[-/\.]\d{1,4}", s):
                        date_match = False
                        break
                if date_match and len(str_samples) > 0:
                    inferred_type = DataTypeEnum.date
                    inferred_role = SemanticRoleEnum.date
                    date_formats = DEFAULT_DATE_FORMATS
                else:
                    inferred_type = DataTypeEnum.string

        suggested_rule = ColumnRuleDefinition(
            column_name=col_str,
            data_type=inferred_type,
            semantic_role=inferred_role,
            is_required=is_required,
            is_unique=is_unique,
            min_value=min_val,
            max_value=max_val,
            allowed_values=allowed_vals,
            date_formats=date_formats,
            country_column_ref=country_ref,
            default_country="IN"
        )

        detected_columns.append(
            DetectedColumn(
                name=col_str,
                inferred_data_type=inferred_type,
                inferred_semantic_role=inferred_role,
                sample_values=sample_vals_cleaned,
                suggested_rule=suggested_rule
            )
        )

    return detected_columns
