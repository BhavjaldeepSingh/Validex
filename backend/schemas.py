from pydantic import BaseModel, Field
from typing import List, Optional, Dict, Any
from enum import Enum


class DataTypeEnum(str, Enum):
    string = "string"
    integer = "integer"
    float = "float"
    boolean = "boolean"
    date = "date"
    time = "time"
    datetime = "datetime"


class SemanticRoleEnum(str, Enum):
    general = "general"
    unique_id = "unique_id"
    phone_number = "phone_number"
    country_code = "country_code"
    date = "date"
    time = "time"
    amount = "amount"
    quantity = "quantity"
    payment_mode = "payment_mode"
    payment_status = "payment_status"
    email = "email"
    customer_name = "customer_name"


class ColumnRuleDefinition(BaseModel):
    column_name: str
    data_type: DataTypeEnum = DataTypeEnum.string
    semantic_role: SemanticRoleEnum = SemanticRoleEnum.general
    is_required: bool = False
    is_unique: bool = False
    min_value: Optional[float] = None
    max_value: Optional[float] = None
    allowed_values: Optional[List[str]] = None
    date_formats: Optional[List[str]] = None
    country_column_ref: Optional[str] = None
    default_country: Optional[str] = "IN"
    regex_pattern: Optional[str] = None


class ChunkSettings(BaseModel):
    enabled: bool = True
    chunk_size: int = Field(default=100, ge=1, le=10000)


class DetectedColumn(BaseModel):
    name: str
    inferred_data_type: DataTypeEnum
    inferred_semantic_role: SemanticRoleEnum
    sample_values: List[Any]
    suggested_rule: ColumnRuleDefinition


class InspectResponse(BaseModel):
    file_token: str
    filename: str
    format: str
    available_sheets: List[str] = []
    active_sheet: Optional[str] = None
    total_rows_sampled: int
    columns: List[DetectedColumn]


class InspectSheetRequest(BaseModel):
    file_token: str
    sheet_name: str


class ValidateRequest(BaseModel):
    file_token: str
    selected_sheet: Optional[str] = None
    column_rules: Optional[List[ColumnRuleDefinition]] = None
    chunk_settings: Optional[ChunkSettings] = ChunkSettings()


class ValidationErrorItem(BaseModel):
    row_number: int
    identifier: Optional[str] = None
    failing_column: str
    invalid_value: Optional[str] = None
    reason: str


class ValidationResponse(BaseModel):
    message: str = "Validation completed"
    upload_id: int
    upload_uuid: str
    filename: str
    file_format: str
    sheet_name: Optional[str] = None
    total_rows: int
    valid_rows: int
    invalid_rows: int
    error_breakdown: Dict[str, int] = {}
    chunk_count: int = 0
    export_formats: List[str] = []


class UploadHistoryItem(BaseModel):
    id: int
    upload_uuid: str
    filename: str
    file_format: str
    sheet_name: Optional[str] = None
    total_rows: int
    valid_rows: int
    invalid_rows: int
    chunk_count: int
    uploaded_at: str
