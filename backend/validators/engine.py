import re
import pandas as pd
from datetime import datetime
from typing import List, Dict, Tuple, Any, Optional

from config import (
    DEFAULT_PHONE_RULES,
    DEFAULT_DATE_FORMATS,
    DEFAULT_TIME_FORMATS
)
from schemas import (
    ColumnRuleDefinition,
    DataTypeEnum,
    SemanticRoleEnum,
    ValidationErrorItem
)


class TypeDrivenValidator:
    """
    Validates any dataset based on dynamic column rules, inferred types,
    and semantic roles, completely independent of fixed column headers.
    """

    def __init__(self, rules: List[ColumnRuleDefinition]):
        self.rules = rules

    def validate(self, df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame, List[ValidationErrorItem], Dict[str, int]]:
        clean_records: List[Dict[str, Any]] = []
        error_records: List[Dict[str, Any]] = []
        detailed_errors: List[ValidationErrorItem] = []
        error_breakdown: Dict[str, int] = {
            "Missing Required Field": 0,
            "Duplicate Unique Key": 0,
            "Type / Parsing Error": 0,
            "Numeric Boundary Violation": 0,
            "Invalid Phone Number": 0,
            "Invalid Date / Time": 0,
            "Invalid Enum Value": 0,
            "Format / Regex Failure": 0,
        }

        # Track seen values for uniqueness per unique column
        seen_unique_values: Dict[str, set] = {
            r.column_name: set() for r in self.rules if r.is_unique
        }

        # Find primary ID column to use as reference identifier
        id_col = None
        for r in self.rules:
            if r.semantic_role == SemanticRoleEnum.unique_id or r.is_unique:
                id_col = r.column_name
                break
        if not id_col and len(df.columns) > 0:
            id_col = df.columns[0]

        total_rows = len(df)

        for index, row in df.iterrows():
            row_number = index + 2  # 1-based index including header
            row_dict = row.to_dict()
            identifier_val = str(row_dict.get(id_col, f"Row_{row_number}")).strip()
            row_errors: List[str] = []

            for rule in self.rules:
                col_name = rule.column_name
                val = row_dict.get(col_name)
                val_is_empty = pd.isna(val) or str(val).strip() == "" or str(val).strip().lower() == "nan"

                # 1. Null / Required check
                if val_is_empty:
                    if rule.is_required:
                        err = f"Column '{col_name}' is required but empty."
                        row_errors.append(err)
                        error_breakdown["Missing Required Field"] += 1
                        detailed_errors.append(
                            ValidationErrorItem(
                                row_number=row_number,
                                identifier=identifier_val,
                                failing_column=col_name,
                                invalid_value="<EMPTY>",
                                reason=err
                            )
                        )
                    continue

                val_str = str(val).strip()

                # 2. Uniqueness check
                if rule.is_unique:
                    if val_str in seen_unique_values[col_name]:
                        err = f"Duplicate value '{val_str}' in unique column '{col_name}'."
                        row_errors.append(err)
                        error_breakdown["Duplicate Unique Key"] += 1
                        detailed_errors.append(
                            ValidationErrorItem(
                                row_number=row_number,
                                identifier=identifier_val,
                                failing_column=col_name,
                                invalid_value=val_str,
                                reason=err
                            )
                        )
                    else:
                        seen_unique_values[col_name].add(val_str)

                # 3. Data Type & Bounds Checking
                numeric_val = None
                if rule.data_type in [DataTypeEnum.integer, DataTypeEnum.float] or rule.semantic_role in [SemanticRoleEnum.amount, SemanticRoleEnum.quantity]:
                    try:
                        # Clean currency symbols or commas if present
                        cleaned_num_str = re.sub(r"[^\d.-]", "", val_str)
                        if rule.data_type == DataTypeEnum.integer or rule.semantic_role == SemanticRoleEnum.quantity:
                            numeric_val = int(float(cleaned_num_str))
                        else:
                            numeric_val = float(cleaned_num_str)
                    except (ValueError, TypeError):
                        err = f"Column '{col_name}' expects {rule.data_type.value}, but got '{val_str}'."
                        row_errors.append(err)
                        error_breakdown["Type / Parsing Error"] += 1
                        detailed_errors.append(
                            ValidationErrorItem(
                                row_number=row_number,
                                identifier=identifier_val,
                                failing_column=col_name,
                                invalid_value=val_str,
                                reason=err
                            )
                        )

                    if numeric_val is not None:
                        if rule.min_value is not None and numeric_val < rule.min_value:
                            err = f"Column '{col_name}' value {numeric_val} is less than minimum allowed ({rule.min_value})."
                            row_errors.append(err)
                            error_breakdown["Numeric Boundary Violation"] += 1
                            detailed_errors.append(
                                ValidationErrorItem(
                                    row_number=row_number,
                                    identifier=identifier_val,
                                    failing_column=col_name,
                                    invalid_value=val_str,
                                    reason=err
                                )
                            )
                        if rule.max_value is not None and numeric_val > rule.max_value:
                            err = f"Column '{col_name}' value {numeric_val} exceeds maximum allowed ({rule.max_value})."
                            row_errors.append(err)
                            error_breakdown["Numeric Boundary Violation"] += 1
                            detailed_errors.append(
                                ValidationErrorItem(
                                    row_number=row_number,
                                    identifier=identifier_val,
                                    failing_column=col_name,
                                    invalid_value=val_str,
                                    reason=err
                                )
                            )

                # 4. Date & Time Validation
                if rule.data_type == DataTypeEnum.date or rule.semantic_role == SemanticRoleEnum.date:
                    allowed_formats = rule.date_formats or DEFAULT_DATE_FORMATS
                    date_valid = False
                    for fmt in allowed_formats:
                        try:
                            datetime.strptime(val_str, fmt)
                            date_valid = True
                            break
                        except ValueError:
                            continue
                    if not date_valid:
                        err = f"Date '{val_str}' in column '{col_name}' does not match allowed format(s): {', '.join(allowed_formats[:3])}."
                        row_errors.append(err)
                        error_breakdown["Invalid Date / Time"] += 1
                        detailed_errors.append(
                            ValidationErrorItem(
                                row_number=row_number,
                                identifier=identifier_val,
                                failing_column=col_name,
                                invalid_value=val_str,
                                reason=err
                            )
                        )

                if rule.data_type == DataTypeEnum.time or rule.semantic_role == SemanticRoleEnum.time:
                    allowed_time_formats = DEFAULT_TIME_FORMATS
                    time_valid = False
                    for fmt in allowed_time_formats:
                        try:
                            datetime.strptime(val_str, fmt)
                            time_valid = True
                            break
                        except ValueError:
                            continue
                    if not time_valid:
                        err = f"Time '{val_str}' in column '{col_name}' does not match standard HH:MM:SS format."
                        row_errors.append(err)
                        error_breakdown["Invalid Date / Time"] += 1
                        detailed_errors.append(
                            ValidationErrorItem(
                                row_number=row_number,
                                identifier=identifier_val,
                                failing_column=col_name,
                                invalid_value=val_str,
                                reason=err
                            )
                        )

                # 5. Semantic Role Specific Checks: Phone Number
                if rule.semantic_role == SemanticRoleEnum.phone_number:
                    # Determine country
                    country = rule.default_country or "IN"
                    if rule.country_column_ref and rule.country_column_ref in row_dict:
                        country_raw = row_dict.get(rule.country_column_ref)
                        if not pd.isna(country_raw) and str(country_raw).strip():
                            country = str(country_raw).strip().upper()

                    phone_digits = re.sub(r"\D", "", val_str)

                    if country in DEFAULT_PHONE_RULES:
                        expected_rule = DEFAULT_PHONE_RULES[country]
                        exp_prefix = expected_rule["prefix"]
                        exp_len = expected_rule["length"]
                        if len(phone_digits) != exp_len or not phone_digits.startswith(exp_prefix):
                            err = f"Phone '{val_str}' is invalid for country '{country}'. Expected {exp_len} digits starting with prefix '{exp_prefix}'."
                            row_errors.append(err)
                            error_breakdown["Invalid Phone Number"] += 1
                            detailed_errors.append(
                                ValidationErrorItem(
                                    row_number=row_number,
                                    identifier=identifier_val,
                                    failing_column=col_name,
                                    invalid_value=val_str,
                                    reason=err
                                )
                            )
                    else:
                        # Generic international phone length check (8 to 15 digits)
                        if not (8 <= len(phone_digits) <= 15):
                            err = f"Phone '{val_str}' has invalid length ({len(phone_digits)} digits). Expected 8-15 digits."
                            row_errors.append(err)
                            error_breakdown["Invalid Phone Number"] += 1
                            detailed_errors.append(
                                ValidationErrorItem(
                                    row_number=row_number,
                                    identifier=identifier_val,
                                    failing_column=col_name,
                                    invalid_value=val_str,
                                    reason=err
                                )
                            )

                # 6. Email Validation
                if rule.semantic_role == SemanticRoleEnum.email:
                    if not re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", val_str):
                        err = f"Invalid email format '{val_str}' in column '{col_name}'."
                        row_errors.append(err)
                        error_breakdown["Format / Regex Failure"] += 1
                        detailed_errors.append(
                            ValidationErrorItem(
                                row_number=row_number,
                                identifier=identifier_val,
                                failing_column=col_name,
                                invalid_value=val_str,
                                reason=err
                            )
                        )

                # 7. Allowed Values / Enum Whitelist
                if rule.allowed_values and len(rule.allowed_values) > 0:
                    upper_allowed = [str(x).strip().upper() for x in rule.allowed_values]
                    if val_str.upper() not in upper_allowed:
                        err = f"Value '{val_str}' in column '{col_name}' not allowed. Permitted: {', '.join(rule.allowed_values[:5])}."
                        row_errors.append(err)
                        error_breakdown["Invalid Enum Value"] += 1
                        detailed_errors.append(
                            ValidationErrorItem(
                                row_number=row_number,
                                identifier=identifier_val,
                                failing_column=col_name,
                                invalid_value=val_str,
                                reason=err
                            )
                        )

                # 8. Custom Regex
                if rule.regex_pattern:
                    if not re.search(rule.regex_pattern, val_str):
                        err = f"Value '{val_str}' in column '{col_name}' does not match pattern '{rule.regex_pattern}'."
                        row_errors.append(err)
                        error_breakdown["Format / Regex Failure"] += 1
                        detailed_errors.append(
                            ValidationErrorItem(
                                row_number=row_number,
                                identifier=identifier_val,
                                failing_column=col_name,
                                invalid_value=val_str,
                                reason=err
                            )
                        )

            # Record Classification
            if len(row_errors) == 0:
                clean_records.append(row_dict)
            else:
                error_record = dict(row_dict)
                error_record["__error_row_number__"] = row_number
                error_record["__validation_errors__"] = " | ".join(row_errors)
                error_records.append(error_record)

        clean_df = pd.DataFrame(clean_records) if clean_records else pd.DataFrame(columns=df.columns)
        error_cols = list(df.columns) + ["__error_row_number__", "__validation_errors__"]
        error_df = pd.DataFrame(error_records) if error_records else pd.DataFrame(columns=error_cols)

        return clean_df, error_df, detailed_errors, error_breakdown
