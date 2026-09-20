# ⚡Datova - Transaction Validation & Audit Platform

> An enterprise-grade, header-agnostic transaction data auditing platform featuring an **in-browser local storage engine**, an interactive **custom constraints studio**, and a modern **blue-on-white glassmorphic interface**.

[![Python Version](https://img.shields.io/badge/python-3.10%2B-blue.svg)](https://www.python.org/)
[![FastAPI](https://img.shields.io/badge/FastAPI-0.115%2B-009688.svg?logo=fastapi)](https://fastapi.tiangolo.com/)
[![Pandas](https://img.shields.io/badge/pandas-2.2%2B-150458.svg?logo=pandas)](https://pandas.pydata.org/)
[![Architecture](https://img.shields.io/badge/Architecture-In--Browser%20%2B%20REST%20API-6366f1.svg)](#architecture)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](#license)

---

## 🌟 Key Highlights

- 🖥️ **100% In-Browser Execution**: All file parsing, validation, chunking, and auditing execute client-side directly in memory.
- 📦 **Zero Mandatory Downloads**: Clean transactions, cell-level diagnostics, chunks, and history are viewable on screen. No automatic file downloads required.
- 💾 **Local Storage Persistence**: Past validation runs automatically persist across browser refreshes using `localStorage` with a real-time storage quota meter.
- 🎛️ **Interactive Custom Constraints Studio**: Configure per-column validation constraints on the fly:
  - **Data Types**: `string`, `integer`, `float`, `date`, `time`, `boolean`.
  - **Semantic Roles**: `unique_id`, `phone_number`, `amount`, `quantity`, `payment_mode`, `payment_status`, `date`, `time`, `email`, `country_code`.
  - **Rule Toggles**: Enforce uniqueness across all rows, mark fields as required, set numeric minimums/maximums, define allowed enums, configure date formats, and test custom regex patterns.
- 🔍 **Cell-Level Error Auditing**: Erroneous table cells are highlighted in soft red with tooltip badges displaying the exact constraint that failed.
- 🧩 **In-Browser Chunking**: Partition large datasets into configurable chunks (default: 100 rows) with instant in-browser chunk inspection.
- 💎 **Blue-on-White Glassmorphism UI**: High-contrast, frosted glass cards (`backdrop-filter: blur(20px)`), ambient sapphire/ice-blue lighting, and clean typography.
- 🔄 **Dual-Mode Operation**: Run standalone in any browser with zero dependencies, or connect to the included **FastAPI + SQLite/MySQL** backend for automated pipelines.

---

## 📁 Supported Ingestion & Export Formats

| Format | Extension | In-Browser Engine | FastAPI Backend |
|---|---|:---:|:---:|
| **CSV** (Comma Separated Values) | `.csv` | ✅ | ✅ |
| **TSV** (Tab Separated Values) | `.tsv` | ✅ | ✅ |
| **JSON** (Array / NDJSON) | `.json` | ✅ | ✅ |
| **Excel** (Workbooks & Multi-Sheet) | `.xlsx`, `.xls` | ✅ | ✅ |
| **XML** (Structured Dataset) | `.xml` | N/A | ✅ (Export) |
| **HTML Report** (Standalone Styled Table) | `.html` | N/A | ✅ (Export) |
| **ZIP Bundle** (Clean + Errors + Chunks) | `.zip` | N/A | ✅ (Export) |

---

## 🚀 Quickstart Guide

### Option 1: Standalone Browser Mode (Zero Backend Setup)

Run a simple local web server to serve the frontend on port `5500`:

```bash
# Python built-in HTTP server
python -m http.server 5500 --directory frontend
```

Open your browser and navigate to:
👉 **`http://localhost:5500`**

*You can immediately click **"Load Sample Data"** to test validation and error diagnostics without uploading a file!*

---

### Option 2: Full-Stack Mode (Frontend + FastAPI Backend)

#### 1. Setup Python Environment & Dependencies
```bash
# Clone the repository
git clone https://github.com/your-username/TransactionValidator.git
cd TransactionValidator

# Create and activate virtual environment
python -m venv venv

# Windows:
.\venv\Scripts\activate
# macOS / Linux:
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt
```

#### 2. Start the FastAPI Backend (Port 3000)
```bash
python -m uvicorn backend.main:app --host 127.0.0.1 --port 3000 --reload
```
API Documentation will be live at: **`http://127.0.0.1:3000/docs`**

#### 3. Start the Frontend (Port 5500)
In a separate terminal window:
```bash
python -m http.server 5500 --directory frontend
```
Navigate to: **`http://localhost:5500`**

---

## 🛠️ Step-by-Step Usage Workflow

```
[1. INGEST]            [2. CONFIGURE]          [3. VALIDATE]           [4. AUDIT & STORE]
Drag & Drop file   --> Set Column Rules    --> Click Validate     --> Inspect Red Cells
(CSV, TSV, XLSX)       (Unique, Min, Enums)    (Pure JS / Memory)      Browse Chunks & History
```

1. **Step 1: File Ingestion & Discovery**:
   - Drag and drop your `.csv`, `.tsv`, `.json`, or `.xlsx` file into the frosted glass dropzone.
   - For Excel workbooks, select your target sheet from the dropdown.
   - Or click **"Load Sample Data"** for an immediate demo dataset with deliberate errors.

2. **Step 2: Custom Constraints Studio**:
   - The platform samples your columns and automatically infers smart rules (e.g. recognizing `order_id` as unique, `amount_paid` as numeric currency, `phone_number` as phone, and `order_date` as date).
   - Customize rules per column:
     - Toggle **Required** or **Unique**.
     - Set **Min / Max** values for price or quantities.
     - Enter comma-separated whitelist items for **Allowed Values** (e.g., `CARD, UPI, COD`).
     - Pick **Date Formats** (`DD-MM-YYYY`, `YYYY-MM-DD`, `MM/DD/YYYY`).
     - Choose country phone validation rules (`IN`, `US`, `UK`, `SG`, `AE`, `AU`).
   - Use quick presets: **Auto-Infer Rules**, **Strict Financial**, or **Select All**.

3. **Step 3: Validation Execution**:
   - Click **"Run In-Browser Validation"**.
   - Processing completes in milliseconds inside browser memory.

4. **Step 4: Interactive Audit & Local Storage History**:
   - Review 4 high-level KPI cards: **Total Rows**, **Clean Records**, **Invalid Records**, and **Quality Score**.
   - Inspect categorized breakdown chips (e.g. `Duplicate Unique Key`, `Invalid Phone Number`).
   - Explore the **Validation Errors** table with failing cells highlighted in soft red with failure tooltips.
   - Switch tabs to inspect **Clean Transactions** or explore **In-Browser Chunks**.
   - All runs are saved to **Local Storage History**—click **"Inspect"** anytime to restore past runs.

---

## 🔌 REST API Reference (FastAPI Backend)

When running the backend server on `http://127.0.0.1:3000`:

| Method | Endpoint | Description |
|---|---|---|
| `GET` | `/` | API status, version, and feature capabilities |
| `POST` | `/inspect` | Ingests file, returns sheet names, sample values, and suggested column rules |
| `POST` | `/inspect/sheet` | Re-inspects a specific Excel sheet using the file session token |
| `POST` | `/validate` | Executes validation using user-supplied column rules and chunk settings |
| `POST` | `/upload-and-validate` | 1-step direct endpoint: auto-infers schema and validates immediately |
| `GET` | `/uploads` | Lists all historical validation sessions from SQLite database |
| `GET` | `/uploads/{id}` | Returns detailed report and rule configurations for a specific job |
| `GET` | `/uploads/{id}/export` | Exports clean or error data (`?type=clean\|error&format=csv\|xlsx\|json\|xml\|html\|zip`) |
| `GET` | `/uploads/{id}/chunks` | Lists generated chunk files for a given upload |
| `GET` | `/uploads/{id}/chunks/{name}` | Downloads a specific chunk file |
| `GET` | `/uploads/{id}/chunks-zip` | Downloads all generated chunks bundled as a single `.zip` |
| `DELETE` | `/uploads/{id}` | Deletes a specific upload record and its storage files |
| `DELETE` | `/uploads` | Purges all upload records and disk storage |

---

## 📂 Repository Structure

```
TransactionValidator-main/
├── index.html               # Main application interface (Glassmorphism design)
├── style.css                # CSS design system (tokens, glass cards, animations)
├── app.js                   # Client-side validation engine & localStorage manager
├── frontend/                # Frontend directory for modular serving
│   ├── index.html           # In-browser validation studio
│   ├── style.css            # Stylesheet with embedded glass tokens
│   └── app.js               # Browser engine logic
├── backend/                 # FastAPI backend service
│   ├── __init__.py          # Module path initialization
│   ├── config.py            # Phone rules, keywords, supported formats
│   ├── database.py          # SQLAlchemy SQLite connection & session generator
│   ├── models.py            # SQLAlchemy database models (Upload table)
│   ├── schemas.py           # Pydantic request & response models
│   ├── main.py              # FastAPI endpoints, CORS, file routing
│   ├── parsers/
│   │   └── file_reader.py   # Multi-format ingestion & token-aware inference
│   ├── validators/
│   │   └── engine.py        # Dynamic TypeDrivenValidator engine
│   └── exporters/
│       └── file_writer.py   # Multi-format exporter (XLSX, XML, HTML, ZIP)
├── DEVS.md                  # Comprehensive technical documentation for engineers
├── requirements.txt         # Python dependencies
├── SQL DATABASE.sql         # Production MySQL database schema
└── README.md                # Project documentation (this file)
```

---

## ⚙️ Configuration & Phone Validation Rules

Country phone rules are built into both the frontend and backend engines:

```javascript
const PHONE_RULES = {
    "IN": { prefix: "91", length: 12, localLength: 10 }, // India
    "US": { prefix: "1",  length: 11, localLength: 10 }, // United States / Canada
    "UK": { prefix: "44", length: 12, localLength: 10 }, // United Kingdom
    "SG": { prefix: "65", length: 10, localLength: 8 },  // Singapore
    "AE": { prefix: "971",length: 12, localLength: 9 },  // UAE
    "AU": { prefix: "61", length: 11, localLength: 9 }   // Australia
};
```

---

## 🤝 Contributing

Contributions are welcome! To contribute:

1. Fork the repository.
2. Create your feature branch (`git checkout -b feature/AmazingFeature`).
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`).
4. Push to the branch (`git push origin feature/AmazingFeature`).
5. Open a Pull Request.

For architecture and internals, read [DEVS.md](DEVS.md).

---

## 📄 License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
