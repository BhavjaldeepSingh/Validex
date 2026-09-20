import os
from typing import Dict, List, Any

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(BASE_DIR)
STORAGE_DIR = os.path.join(PROJECT_ROOT, "storage")
UPLOADS_DIR = os.path.join(STORAGE_DIR, "uploads")
TEMP_DIR = os.path.join(STORAGE_DIR, "temp")

os.makedirs(UPLOADS_DIR, exist_ok=True)
os.makedirs(TEMP_DIR, exist_ok=True)

# Default chunking settings
DEFAULT_ENABLE_CHUNKING = True
DEFAULT_CHUNK_SIZE = 100

# Supported export formats
SUPPORTED_EXPORT_FORMATS = ["csv", "tsv", "xlsx", "json", "xml", "html", "zip"]

# Built-in phone validation rules: {country_code: {"prefix": str, "length": int}}
DEFAULT_PHONE_RULES: Dict[str, Dict[str, Any]] = {
    "IN": {"prefix": "91", "length": 12},
    "SG": {"prefix": "65", "length": 10},
    "US": {"prefix": "1", "length": 11},
    "CA": {"prefix": "1", "length": 11},
    "UK": {"prefix": "44", "length": 12},
    "GB": {"prefix": "44", "length": 12},
    "AU": {"prefix": "61", "length": 11},
    "AE": {"prefix": "971", "length": 12},
    "DE": {"prefix": "49", "length": 12},
    "FR": {"prefix": "33", "length": 11},
}

# Default allowed payment modes & statuses
DEFAULT_PAYMENT_MODES: List[str] = [
    "CARD", "COD", "PAYPAL", "NETBANKING", "WALLET", "UPI", "BANK_TRANSFER", "CRYPTO", "STRIPE"
]

DEFAULT_PAYMENT_STATUSES: List[str] = [
    "COMPLETED", "SUCCESS", "PENDING", "FAILED", "REFUNDED", "CANCELLED", "PROCESSING"
]

# Default accepted date & time patterns
DEFAULT_DATE_FORMATS: List[str] = [
    "%d-%m-%Y",
    "%Y-%m-%d",
    "%d/%m/%Y",
    "%m/%d/%Y",
    "%Y/%m/%d",
    "%d.%m.%Y"
]

DEFAULT_TIME_FORMATS: List[str] = [
    "%H:%M:%S",
    "%H:%M",
    "%I:%M:%S %p",
    "%I:%M %p"
]

# Keywords used for heuristic semantic role detection
SEMANTIC_KEYWORDS = {
    "unique_id": ["id", "uuid", "order_id", "transaction_id", "txn_id", "invoice_no", "ref_no", "reference"],
    "phone_number": ["phone", "mobile", "contact", "cell", "tel", "whatsapp", "ph_no"],
    "country_code": ["country", "nation", "country_code", "geo", "region"],
    "date": ["date", "dob", "created_at", "order_date", "timestamp", "dt"],
    "time": ["time", "order_time", "hour", "slot"],
    "amount": ["amount", "price", "total", "cost", "fee", "amt", "paid", "balance", "unit_price", "revenue"],
    "quantity": ["quantity", "qty", "units", "count", "items", "volume"],
    "payment_mode": ["payment_mode", "payment_method", "pay_mode", "mode", "gateway", "channel"],
    "payment_status": ["payment_status", "status", "txn_status", "order_status", "state"],
    "email": ["email", "mail", "e_mail", "email_address"],
    "customer_name": ["customer", "name", "buyer", "client", "user", "customer_name", "full_name"]
}
