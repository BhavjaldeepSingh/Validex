# Developer Technical Documentation & Architecture Guide (DEVS ONLY)

This document provides a comprehensive technical breakdown of the ** Datova - Transaction Validation Platform** architecture, data flow, component internals, algorithms, and persistence models. It is written exclusively for engineers contributing to or maintaining this codebase.

---

## 1. High-Level Architecture & Operating Modes

The application is engineered with a **Dual-Mode Architecture**:

```
                                  +---------------------------------------+
                                  |            CLIENT BROWSER             |
                                  |                                       |
    [CSV / TSV / JSON / XLSX] --->|  File Parsing (PapaParse / SheetJS)   |
                                  |                   |                   |
                                  |  Dynamic Schema & Role Inference      |
                                  |                   |                   |
                                  |  Custom Constraints Config Studio     |
                                  |                   |                   |
                                  |  In-Browser Validation Engine (PureJS)|
                                  |         /                   \         |
                                  |        v                     v        |
                                  |  [localStorage]        [Interactive]  |
                                  |  Persistence           Results Tables |
                                  +---------------------------------------+
                                                     |
                                                     | (Optional Full-Stack Mode)
                                                     v
                                  +---------------------------------------+
                                  |            FASTAPI BACKEND            |
                                  |                                       |
                                  |  POST /inspect, POST /validate        |
                                  |  TypeDrivenValidator (Python/Pandas)  |
                                  |  Multi-Format Exporter (XLSX/ZIP/XML) |
                                  |                   |                   |
                                  |        +----------+----------+        |
                                  |        |                     |        |
                                  |        v                     v        |
                                  |   [SQLite / MySQL]   [storage/uploads]|
                                  |   Metadata Database  Isolated Files   |
                                  +---------------------------------------+
```

### Mode A: Standalone Browser Engine (Default & Zero-Server)
- Operates entirely inside the client’s web browser without requiring a running backend.
- Files (`.csv`, `.tsv`, `.json`, `.xlsx`) are read directly into browser memory via HTML5 File API.
- All schema discovery, custom constraint rule evaluation, cell-level diagnostics, and partitioning (chunking) execute using client-side JavaScript.
- All results, metrics, and audit sessions persist in the browser's `localStorage` under `tv_validation_history_v2`.
- **Zero data is transmitted externally or forced to download.**

### Mode B: Full-Stack REST API Mode (FastAPI + SQLite / MySQL)
- Used for automated batch processing, headless validation pipelines, or backend persistence.
- Provides endpoints for inspection (`/inspect`), sheet re-inspection (`/inspect/sheet`), execution (`/validate`), direct validation (`/upload-and-validate`), and multi-format exports (`/uploads/{id}/export`).
- Stores upload records, validation metrics, and rule profiles in an SQLite database (migratable to MySQL via `SQL DATABASE.sql`).
- Generates isolated disk storage per upload session under `storage/uploads/{upload_id}/`.

---

## 2. Frontend Architecture Deep Dive

### 2.1 Technology Stack & Design System
- **Core**: Vanilla ECMAScript (ES6+) with zero build tools or bundlers required.
- **Parsers**:
  - **PapaParse** (`v5.4.1`): High-performance streaming CSV parser with delimiter auto-detection.
  - **SheetJS** (`xlsx.full.min.js v0.18.5`): Binary workbook parser for `.xlsx` and `.xls`.
  - **Native RFC 4180 Fallback**: A custom character-by-character CSV parser embedded in `app.js` guaranteeing parser functionality even in offline environments without CDN access.
- **UI / UX Style**: "Blue-on-White Glassification"
  - Frosted glass containers (`backdrop-filter: blur(20px)`, `background: rgba(255, 255, 255, 0.78)`).
  - Multi-layer sapphire/sky blue radial ambient gradients.
  - Typography: `Plus Jakarta Sans` for UI labels and `JetBrains Mono` for tabular identifiers and data values.

### 2.2 File Ingestion & Parsing Workflow
1. User drops a file or clicks "Browse Files" / "Load Sample Data".
2. `processSelectedFile(file)` inspects the extension:
   - **Excel (`.xlsx`, `.xls`)**: Read as `ArrayBuffer`, parsed into a SheetJS `workbook`. Active sheets are populated into `#excelSheetSelect`. The active worksheet is converted to JSON records via `XLSX.utils.sheet_to_json(worksheet, { defval: "" })`.
   - **JSON (`.json`)**: Parsed via `JSON.parse`. Arrays of objects are directly loaded; single objects are wrapped into an array.
   - **CSV / TSV**: Streamed as text. Evaluated via `Papa.parse()` with `header: true, skipEmptyLines: true`. If CDN is unreachable, the fallback `parseCsvFallback()` executes.
3. Once records are loaded, `onDatasetLoaded(records, filename)` is triggered:
   - Scans the first 50 rows to discover all distinct column headers.
   - Updates the Active File Banner (`#activeFileName`, `#activeFileMeta`).
   - Invokes `renderCustomConstraintsStudio()` to build constraint cards.

### 2.3 Heuristic Schema Discovery & Role Inference
The inference engine (`inferColumnRule` in `frontend/app.js` and `infer_column_rules` in `backend/parsers/file_reader.py`) analyzes column names and sample data:

```javascript
// Word-boundary tokenization avoids substring false positives:
const tokens = colLower.split(/[\s_\-]+/);
```

#### Classification Rules:
- **`unique_id`**: Matches when the column is named `id`, `uuid`, `order_id`, `transaction_id`, `txn_id`, `invoice_id`, `invoice_no`, `ref_no`, or ends with `_id` (specifically excluding `product`, `item`, `category`, `status`, `country`, `customer`, `user`, and `paid`).
  - *Fix Note*: In early implementations, substring matching on `"id"` matched `amount_paid` (because `pa-id` contains `id`). Token-aware boundary checks permanently prevent this bug.
- **`phone_number`**: Matches `phone`, `mobile`, `contact`, `cell`, `tel`, `whatsapp`. Default country rule assigned.
- **`country_code`**: Matches `country`, `nation`, `geo`, `region`.
- **`payment_mode`**: Matches `payment_mode`, `pay_mode`, `gateway`, `channel`. Automatically populates allowed whitelist with standard modes (`CARD, UPI, COD, NETBANKING, etc.`).
- **`payment_status`**: Matches `payment_status`, `txn_status`, `order_status`. Populates whitelist (`COMPLETED, PENDING, FAILED, etc.`).
- **`date`**: Matches `date`, `dob`, `created_at`, `order_date`. Sets data type to `date` and default format `DD-MM-YYYY`.
- **`time`**: Matches `time`, `hour`, `slot`. Sets data type to `time`.
- **`amount`**: Matches `amount`, `price`, `total`, `cost`, `fee`, `unit_price`, `balance`. Sets type to `float`, `min_val = 0.0`.
- **`quantity`**: Matches `quantity`, `qty`, `units`, `count`, `items`. Sets type to `integer`, `min_val = 1`.
- **`email`**: Matches `email`, `mail`. Flags for RFC 5322 regex validation.

### 2.4 Custom Constraints Evaluation Pipeline
When the user clicks "Run In-Browser Validation", `performValidation(dataset, rules)` executes row by row:

1. **Uniqueness Sets**: Initializes a `Set()` per column marked `is_unique`.
2. **Row Evaluation Loop**:
   - **Required Check**: Flags values that are null, undefined, empty string, `"nan"`, or `"null"`.
   - **Duplicate Check**: Verifies if `valStr` exists in `seenUniqueSets[col]`. If seen, records an error; if new, adds to the Set.
   - **Numeric Bounds**: Strips non-numeric characters (except `-` and `.`), parses float/int, checks `min_value` and `max_value`.
   - **Date / Time Formatting**: Validates against chosen format (`DD-MM-YYYY`, `YYYY-MM-DD`, `MM/DD/YYYY`, or standard parseable dates).
   - **Phone Rule Check**: Cleans all non-digit characters. Resolves country either from `default_country` or row-level `country_column_ref`. Evaluates country-specific prefix and expected digit lengths (e.g., India: 10 local digits or 12 with 91 prefix; US: 10 local or 11 with 1 prefix).
   - **Allowed Values (Enums)**: Normalizes to uppercase and checks membership in configured array.
   - **Custom Regex**: Evaluates arbitrary regex provided in the UI card.
3. **Diagnostics Compilation**:
   - Clean rows are pushed to `cleanRecords`.
   - Failing rows are tagged with `__error_row_number__`, `__validation_errors__` (summary text), and `__cell_errors__` (dictionary mapping column names to `{ reason, type, val }` objects).
   - Error categories are tallied across 8 categories.

### 2.5 In-Memory Chunking Engine
If `#enableChunking` is checked:
- Partitions the `cleanRecords` array into subsets of `#chunkSizeInput` rows (default: 100).
- Chunks are stored as objects: `{ chunk_index, name: 'chunk_X.csv', start_row, end_row, records }`.
- Chunks can be browsed and inspected directly within the "In-Browser Chunks" tab without downloading.

### 2.6 LocalStorage Persistence Model
Key: `tv_validation_history_v2`
```json
[
  {
    "id": "run_1726667890123",
    "filename": "transactions.csv",
    "timestamp": "2026-09-18T14:15:00.000Z",
    "total_rows": 500,
    "valid_rows": 480,
    "invalid_rows": 20,
    "health_score": 96,
    "error_breakdown": {
      "Missing Required Field": 4,
      "Duplicate Unique Key": 2,
      "Numeric Boundary Violation": 14
    },
    "columns": ["order_id", "amount", "customer_name"],
    "clean_records": [...],
    "error_records": [...],
    "chunks": [...]
  }
]
```
- **Quota Management**: Maintains the latest 15 runs. Calculates approximate byte consumption: `totalBytes = sum(length * 2)` across keys, displaying a live gauge (`#storageUsageText`, `#storageMeterFill`).
- **Inspection & Restore**: Clicking "Inspect" on any historical run restores the entire dashboard, error tables, and chunk views for that run.

---

## 3. Backend Architecture Deep Dive

### 3.1 Backend Directory Structure
```
backend/
├── __init__.py           # Path configuration ensuring submodules resolve cleanly
├── config.py             # System constants, phone rules, export formats, directory paths
├── database.py           # SQLAlchemy database engine, SessionLocal, Base, get_db()
├── models.py             # SQLAlchemy models (Upload table)
├── schemas.py            # Pydantic models for request/response validation
├── main.py               # FastAPI application, middleware, route handlers
├── parsers/
│   └── file_reader.py    # Multi-format ingestion, encoding detection, schema discovery
├── validators/
│   └── engine.py         # TypeDrivenValidator (header-agnostic dynamic validation)
└── exporters/
    └── file_writer.py    # Multi-format file generation (CSV, XLSX, XML, HTML, ZIP)
```

### 3.2 Ingestion & Parsing Engine (`backend/parsers/file_reader.py`)
- **`save_temp_file(file_bytes, filename)`**:
  - Generates a UUID token (`file_token`).
  - Sanitizes filename (`re.sub(r"[^a-zA-Z0-9_.-]", "_", filename)`).
  - Persists file temporarily in `storage/temp/`.
- **`read_file_to_dataframe(file_path, file_format, sheet_name)`**:
  - **Excel (`.xlsx`, `.xls`)**: Reads workbook via `openpyxl(read_only=True)`. Extracts sheet names and active sheet. Loads target sheet into Pandas DataFrame.
  - **CSV / TSV**: Attempts delimiter sniffing via `csv.Sniffer()`. Cycles through encodings (`utf-8`, `utf-8-sig`, `latin-1`, `cp1252`) to ensure legacy encoding compatibility.
  - **JSON**: Parses standard array JSON or newline-delimited JSON (`lines=True`).

### 3.3 Dynamic Validation Engine (`backend/validators/engine.py`)
Class: `TypeDrivenValidator(rules: List[ColumnRuleDefinition])`

- Operates dynamically on DataFrame columns using provided Pydantic rule definitions.
- Header-agnostic: Rules specify `column_name`, `data_type`, `semantic_role`, `is_required`, `is_unique`, `min_value`, `max_value`, `allowed_values`, `date_formats`, and `country_column_ref`.
- **Cell Diagnostics**: Produces `clean_df`, `error_df`, `detailed_errors` (list of `ValidationErrorItem`), and categorized error breakdown counts.
- Appends `__error_row_number__` and `__validation_errors__` to `error_df`.

### 3.4 Multi-Format Exporter (`backend/exporters/file_writer.py`)
- **`export_dataframe(df, format, base_name)`**:
  - `csv`: Returns UTF-8 encoded CSV bytes.
  - `tsv`: Returns tab-separated values.
  - `xlsx`: Generates formatted OpenPyxl workbook with bold navy headers, soft gray alternating row fills, and auto-fitted column widths.
  - `json`: Returns formatted JSON records array.
  - `xml`: Builds a structured XML tree (`<dataset><record><col>val</col></record></dataset>`).
  - `html`: Produces a standalone, styled HTML document containing responsive tables and glassmorphic badge styling.
  - `zip`: Creates a zip archive containing clean data, validation error report, and individual chunk CSVs.
- **Isolated Session Storage**:
  - Saves artifacts under `storage/uploads/{upload_id}/`:
    - `clean_transactions.csv`
    - `validation_errors.csv`
    - `chunks/chunk_1.csv`, `chunks/chunk_2.csv`, etc.

---

## 4. Database Architecture & Schema

The platform uses **SQLAlchemy ORM** configured by default with SQLite, with production DDL available for MySQL in `SQL DATABASE.sql`.

### Table: `uploads`
| Column | Type | Description |
|---|---|---|
| `id` | Integer (PK, AutoIncrement) | Unique sequential upload job ID |
| `upload_uuid` | String(36), Indexed | UUIDv4 identifier for the session |
| `filename` | String(255) | Original name of uploaded file |
| `file_format` | String(10) | Extension (`csv`, `tsv`, `xlsx`, `json`) |
| `sheet_name` | String(100), Nullable | Selected Excel worksheet |
| `total_rows` | Integer | Total records evaluated |
| `valid_rows` | Integer | Clean records passing all checks |
| `invalid_rows` | Integer | Records failing one or more constraints |
| `chunk_count` | Integer | Number of partition chunks generated |
| `error_breakdown` | Text (JSON String) | Categorized count of error types |
| `rule_config` | Text (JSON String) | Serialized list of ColumnRuleDefinitions used |
| `storage_path` | String(500), Nullable | Filesystem path to artifacts folder |
| `uploaded_at` | DateTime | Timestamp of validation execution |

### Database Connection (`backend/database.py`)
```python
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "transaction_validator.db")
DATABASE_URL = f"sqlite:///{DB_PATH}"

engine = create_engine(DATABASE_URL, connect_args={"check_same_thread": False})
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
```

For MySQL deployments, switch `DATABASE_URL` to:
`mysql+pymysql://<user>:<password>@<host>:<port>/transaction_validator`

---

## 5. Development & Execution Instructions

### 5.1 Standalone Frontend Execution
To run the in-browser engine on port 5500:
```bash
# From workspace root:
python -m http.server 5500 --directory frontend
```
Navigate to `http://localhost:5500`.

### 5.2 FastAPI Backend Execution
To run the FastAPI backend on port 3000:
```bash
# Activate virtual environment
.\venv\Scripts\activate

# Run Uvicorn
python -m uvicorn backend.main:app --host 127.0.0.1 --port 3000 --reload
```
Interactive OpenAPI documentation will be accessible at `http://127.0.0.1:3000/docs`.

### 5.3 Backend Import Resolution Pattern
`backend/__init__.py` and `backend/main.py` explicitly inject `backend/` into `sys.path`. This guarantees that internal imports (`from config import ...`, `from database import ...`) resolve without `ModuleNotFoundError`, whether invoked from the project root or within the `backend/` directory.

---

## 6. Testing & Debugging Guidelines

- **Client-Side Testing**: Open DevTools Console in `http://localhost:5500`. The `currentDataset`, `currentValidationResult`, and `tv_validation_history_v2` objects are accessible for inspection.
- **Backend Testing**:
  ```python
  # Verification of schema inference
  from backend.parsers.file_reader import infer_column_rules
  import pandas as pd
  df = pd.DataFrame({"order_id": ["ORD1"], "amount_paid": [10.5]})
  rules = infer_column_rules(df)
  # Assert order_id.suggested_rule.is_unique == True
  # Assert amount_paid.suggested_rule.is_unique == False
  ```
- **Clearing State**:
  - Client state: Click "Clear History" in the UI or call `localStorage.clear()` in the browser console.
  - Backend state: `DELETE /uploads` clears database records and purges `storage/uploads/`.
