import os
import sys
import json
import shutil
import uuid
from datetime import datetime
from typing import Optional, List, Dict, Any

# Ensure backend directory is in python path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from fastapi import FastAPI, UploadFile, File, Form, Query, HTTPException, Depends, Response
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from sqlalchemy.orm import Session
import pandas as pd

from config import (
    UPLOADS_DIR,
    TEMP_DIR,
    DEFAULT_CHUNK_SIZE,
    SUPPORTED_EXPORT_FORMATS,
    DEFAULT_PAYMENT_MODES,
    DEFAULT_PAYMENT_STATUSES,
    DEFAULT_PHONE_RULES
)
from database import engine, Base, get_db
from models import Upload
from schemas import (
    InspectResponse,
    InspectSheetRequest,
    ValidateRequest,
    ValidationResponse,
    UploadHistoryItem,
    ChunkSettings,
    ColumnRuleDefinition
)
from parsers.file_reader import (
    save_temp_file,
    find_temp_file,
    read_file_to_dataframe,
    infer_column_rules,
    get_file_extension
)
from validators.engine import TypeDrivenValidator
from exporters.file_writer import (
    export_dataframe,
    save_upload_artifacts,
    create_chunks_zip,
    create_full_bundle_zip
)

# Ensure required directories exist
os.makedirs("uploads", exist_ok=True)
os.makedirs("outputs", exist_ok=True)
os.makedirs("errors", exist_ok=True)
os.makedirs("chunks", exist_ok=True)
os.makedirs(UPLOADS_DIR, exist_ok=True)
os.makedirs(TEMP_DIR, exist_ok=True)

# Create/update database tables
Base.metadata.create_all(bind=engine)

app = FastAPI(
    title="Transaction Validation & Audit Platform API",
    description="Header-agnostic, data-type driven transaction validation engine supporting CSV, TSV, Excel, and JSON.",
    version="2.0.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/")
def home():
    return {
        "message": "Transaction Validation & Audit Platform Running",
        "version": "2.0.0",
        "supported_inputs": ["csv", "tsv", "xlsx", "xls", "json"],
        "supported_exports": SUPPORTED_EXPORT_FORMATS,
        "features": [
            "Header-agnostic validation via data-type and semantic roles",
            "Interactive schema discovery & inspection",
            "Multi-sheet Excel selection",
            "Customizable and optional chunking",
            "Cell-level error diagnostics with exact reasons",
            "Multi-format exports (CSV, TSV, XLSX, JSON, XML, HTML, ZIP)"
        ]
    }


# ==========================================
# STAGE 1: INSPECTION & SCHEMA DISCOVERY
# ==========================================

@app.post("/inspect", response_model=InspectResponse)
async def inspect_uploaded_file(file: UploadFile = File(...)):
    """
    Stage 1: Receives file, caches it temporarily, inspects sheets (if Excel),
    samples rows, and infers column data types and suggested validation constraints.
    """
    contents = await file.read()
    if not contents:
        raise HTTPException(status_code=400, detail="Uploaded file is empty.")

    file_token, saved_path, file_format = save_temp_file(contents, file.filename)

    try:
        df, available_sheets, active_sheet = read_file_to_dataframe(
            file_path=saved_path,
            file_format=file_format
        )
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Failed to read file: {str(e)}")

    detected_columns = infer_column_rules(df)

    return InspectResponse(
        file_token=file_token,
        filename=file.filename,
        format=file_format,
        available_sheets=available_sheets,
        active_sheet=active_sheet,
        total_rows_sampled=len(df),
        columns=detected_columns
    )


@app.post("/inspect/sheet", response_model=InspectResponse)
def inspect_excel_sheet(request: InspectSheetRequest):
    """
    Re-inspects a specific worksheet of a previously inspected Excel file.
    """
    temp_info = find_temp_file(request.file_token)
    if not temp_info:
        raise HTTPException(status_code=404, detail="File session expired or not found. Please upload again.")

    saved_path, original_filename, file_format = temp_info
    if file_format not in ["xlsx", "xls"]:
        raise HTTPException(status_code=400, detail="Sheet inspection is only available for Excel files.")

    try:
        df, available_sheets, active_sheet = read_file_to_dataframe(
            file_path=saved_path,
            file_format=file_format,
            sheet_name=request.sheet_name
        )
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Failed to read sheet '{request.sheet_name}': {str(e)}")

    detected_columns = infer_column_rules(df)

    return InspectResponse(
        file_token=request.file_token,
        filename=original_filename,
        format=file_format,
        available_sheets=available_sheets,
        active_sheet=request.sheet_name,
        total_rows_sampled=len(df),
        columns=detected_columns
    )


# ==========================================
# STAGE 2: VALIDATION EXECUTION
# ==========================================

@app.post("/validate", response_model=ValidationResponse)
def validate_file(request: ValidateRequest, db: Session = Depends(get_db)):
    """
    Stage 2: Executes validation on the inspected file using dynamic column constraints
    and chunk settings chosen by the user.
    """
    temp_info = find_temp_file(request.file_token)
    if not temp_info:
        raise HTTPException(status_code=404, detail="File session expired or not found. Please upload again.")

    saved_path, original_filename, file_format = temp_info

    try:
        df, available_sheets, active_sheet = read_file_to_dataframe(
            file_path=saved_path,
            file_format=file_format,
            sheet_name=request.selected_sheet
        )
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Failed to load dataset: {str(e)}")

    # Use user-supplied rules or fallback to auto-inferred rules
    if request.column_rules and len(request.column_rules) > 0:
        rules = request.column_rules
    else:
        detected = infer_column_rules(df)
        rules = [d.suggested_rule for d in detected]

    validator = TypeDrivenValidator(rules)
    clean_df, error_df, detailed_errors, error_breakdown = validator.validate(df)

    total_rows = len(df)
    valid_rows = len(clean_df)
    invalid_rows = len(error_df)

    upload_uuid = str(uuid.uuid4())
    chunk_settings = request.chunk_settings or ChunkSettings()

    # Create DB entry first to obtain upload.id
    new_upload = Upload(
        upload_uuid=upload_uuid,
        filename=original_filename,
        file_format=file_format,
        sheet_name=request.selected_sheet or active_sheet,
        total_rows=total_rows,
        valid_rows=valid_rows,
        invalid_rows=invalid_rows,
        chunk_count=0,
        error_breakdown=json.dumps(error_breakdown),
        rule_config=json.dumps([r.model_dump() for r in rules]),
        uploaded_at=datetime.utcnow()
    )
    db.add(new_upload)
    db.commit()
    db.refresh(new_upload)

    # Save isolated artifacts and generate chunks
    upload_storage_dir, chunk_count = save_upload_artifacts(
        upload_id=new_upload.id,
        clean_df=clean_df,
        error_df=error_df,
        chunk_settings=chunk_settings
    )

    new_upload.chunk_count = chunk_count
    new_upload.storage_path = upload_storage_dir
    db.commit()

    # Maintain legacy fallback files for backward compatibility
    try:
        clean_df.to_csv("outputs/clean_transactions.csv", index=False)
        error_df.to_csv("errors/validation_errors.csv", index=False)
        # Refresh root chunks folder
        for old_f in os.listdir("chunks"):
            f_p = os.path.join("chunks", old_f)
            if os.path.isfile(f_p):
                os.remove(f_p)
        if chunk_count > 0:
            iso_chunks_dir = os.path.join(upload_storage_dir, "chunks")
            for chunk_file in os.listdir(iso_chunks_dir):
                shutil.copy(os.path.join(iso_chunks_dir, chunk_file), os.path.join("chunks", chunk_file))
    except Exception as e:
        print(f"Warning: Legacy file sync error: {e}")

    return ValidationResponse(
        message="Validation completed successfully",
        upload_id=new_upload.id,
        upload_uuid=upload_uuid,
        filename=original_filename,
        file_format=file_format,
        sheet_name=request.selected_sheet or active_sheet,
        total_rows=total_rows,
        valid_rows=valid_rows,
        invalid_rows=invalid_rows,
        error_breakdown=error_breakdown,
        chunk_count=chunk_count,
        export_formats=SUPPORTED_EXPORT_FORMATS
    )


# ==========================================
# DIRECT 1-STEP UPLOAD & VALIDATE ENDPOINT
# ==========================================

@app.post("/upload-and-validate", response_model=ValidationResponse)
async def upload_and_validate_direct(
    file: UploadFile = File(...),
    sheet_name: Optional[str] = Form(None),
    enable_chunking: bool = Form(True),
    chunk_size: int = Form(100),
    db: Session = Depends(get_db)
):
    """
    1-Step direct validation: uploads file, auto-detects column types/semantics,
    validates data, generates chunks and isolated storage, and returns results.
    """
    contents = await file.read()
    if not contents:
        raise HTTPException(status_code=400, detail="Uploaded file is empty.")

    file_token, saved_path, file_format = save_temp_file(contents, file.filename)

    try:
        df, available_sheets, active_sheet = read_file_to_dataframe(
            file_path=saved_path,
            file_format=file_format,
            sheet_name=sheet_name
        )
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Failed to read file: {str(e)}")

    # Auto-infer rules
    detected = infer_column_rules(df)
    rules = [d.suggested_rule for d in detected]

    validator = TypeDrivenValidator(rules)
    clean_df, error_df, detailed_errors, error_breakdown = validator.validate(df)

    upload_uuid = str(uuid.uuid4())
    chunk_settings = ChunkSettings(enabled=enable_chunking, chunk_size=chunk_size)

    new_upload = Upload(
        upload_uuid=upload_uuid,
        filename=file.filename,
        file_format=file_format,
        sheet_name=sheet_name or active_sheet,
        total_rows=len(df),
        valid_rows=len(clean_df),
        invalid_rows=len(error_df),
        chunk_count=0,
        error_breakdown=json.dumps(error_breakdown),
        rule_config=json.dumps([r.model_dump() for r in rules]),
        uploaded_at=datetime.utcnow()
    )
    db.add(new_upload)
    db.commit()
    db.refresh(new_upload)

    upload_storage_dir, chunk_count = save_upload_artifacts(
        upload_id=new_upload.id,
        clean_df=clean_df,
        error_df=error_df,
        chunk_settings=chunk_settings
    )

    new_upload.chunk_count = chunk_count
    new_upload.storage_path = upload_storage_dir
    db.commit()

    # Sync legacy files
    try:
        clean_df.to_csv("outputs/clean_transactions.csv", index=False)
        error_df.to_csv("errors/validation_errors.csv", index=False)
        for old_f in os.listdir("chunks"):
            f_p = os.path.join("chunks", old_f)
            if os.path.isfile(f_p):
                os.remove(f_p)
        if chunk_count > 0:
            iso_chunks_dir = os.path.join(upload_storage_dir, "chunks")
            for chunk_file in os.listdir(iso_chunks_dir):
                shutil.copy(os.path.join(iso_chunks_dir, chunk_file), os.path.join("chunks", chunk_file))
    except Exception as e:
        print(f"Warning: Legacy file sync error: {e}")

    return ValidationResponse(
        message="Validation completed successfully",
        upload_id=new_upload.id,
        upload_uuid=upload_uuid,
        filename=file.filename,
        file_format=file_format,
        sheet_name=sheet_name or active_sheet,
        total_rows=len(df),
        valid_rows=len(clean_df),
        invalid_rows=len(error_df),
        error_breakdown=error_breakdown,
        chunk_count=chunk_count,
        export_formats=SUPPORTED_EXPORT_FORMATS
    )


# ==========================================
# LEGACY /upload ENDPOINT (Upgraded)
# ==========================================

@app.post("/upload")
async def upload_file_legacy(file: UploadFile = File(...), db: Session = Depends(get_db)):
    """
    Maintains backward compatibility with original frontend while applying
    the resilient type-driven engine and saving isolated session history.
    """
    res = await upload_and_validate_direct(
        file=file,
        sheet_name=None,
        enable_chunking=True,
        chunk_size=50,
        db=db
    )

    return {
        "message": "File processed successfully",
        "upload_id": res.upload_id,
        "filename": res.filename,
        "total_rows": res.total_rows,
        "valid_rows": res.valid_rows,
        "invalid_rows": res.invalid_rows,
        "error_breakdown": res.error_breakdown,
        "clean_file": "outputs/clean_transactions.csv",
        "error_file": "errors/validation_errors.csv",
        "chunks_folder": "chunks/",
        "export_endpoints": {
            fmt: f"/uploads/{res.upload_id}/export?type=clean&format={fmt}"
            for fmt in SUPPORTED_EXPORT_FORMATS
        }
    }


# ==========================================
# MULTI-FORMAT EXPORT & PREVIEW ENDPOINTS
# ==========================================

def safe_read_csv(file_path: str) -> pd.DataFrame:
    if os.path.exists(file_path) and os.path.getsize(file_path) > 0:
        try:
            return pd.read_csv(file_path)
        except Exception:
            return pd.DataFrame()
    return pd.DataFrame()


@app.get("/uploads/{upload_id}/export")
def export_upload_data(
    upload_id: int,
    type: str = Query("clean", pattern="^(clean|error|bundle)$"),
    format: str = Query("csv", pattern="^(csv|tsv|xlsx|json|xml|html|zip)$"),
    db: Session = Depends(get_db)
):
    """
    Exports clean or error data in any requested format:
    csv, tsv, xlsx, json, xml, html, or zip.
    """
    upload = db.query(Upload).filter(Upload.id == upload_id).first()
    if not upload or not upload.storage_path or not os.path.exists(upload.storage_path):
        raise HTTPException(status_code=404, detail=f"Upload #{upload_id} records not found.")

    clean_path = os.path.join(upload.storage_path, "clean_transactions.csv")
    error_path = os.path.join(upload.storage_path, "validation_errors.csv")

    clean_df = safe_read_csv(clean_path)
    error_df = safe_read_csv(error_path)

    base_name = f"{os.path.splitext(upload.filename)[0]}_{type}"

    # Handle full bundle zip
    if format == "zip" or type == "bundle":
        zip_bytes = create_full_bundle_zip(upload_id, clean_df, error_df)
        return Response(
            content=zip_bytes,
            media_type="application/zip",
            headers={"Content-Disposition": f'attachment; filename="{base_name}_bundle.zip"'}
        )

    target_df = clean_df if type == "clean" else error_df

    try:
        file_bytes, media_type, filename = export_dataframe(target_df, format, base_name=base_name)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Export generation failed: {str(e)}")

    return Response(
        content=file_bytes,
        media_type=media_type,
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )


@app.get("/uploads/{upload_id}/preview")
def preview_upload_data(upload_id: int, limit: int = 20, db: Session = Depends(get_db)):
    """Returns sample rows of clean and error data for in-browser inspection."""
    upload = db.query(Upload).filter(Upload.id == upload_id).first()
    if not upload or not upload.storage_path or not os.path.exists(upload.storage_path):
        raise HTTPException(status_code=404, detail=f"Upload #{upload_id} not found.")

    clean_path = os.path.join(upload.storage_path, "clean_transactions.csv")
    error_path = os.path.join(upload.storage_path, "validation_errors.csv")

    clean_df = safe_read_csv(clean_path)
    error_df = safe_read_csv(error_path)

    return {
        "upload_id": upload.id,
        "filename": upload.filename,
        "total_rows": upload.total_rows,
        "valid_rows": upload.valid_rows,
        "invalid_rows": upload.invalid_rows,
        "error_breakdown": json.loads(upload.error_breakdown or "{}"),
        "clean_preview": clean_df.head(limit).to_dict(orient="records"),
        "error_preview": error_df.head(limit).to_dict(orient="records")
    }


# ==========================================
# CHUNK MANAGEMENT ENDPOINTS
# ==========================================

@app.get("/uploads/{upload_id}/chunks")
def list_upload_chunks(upload_id: int, db: Session = Depends(get_db)):
    """Lists chunk files generated for a specific upload."""
    upload = db.query(Upload).filter(Upload.id == upload_id).first()
    if not upload or not upload.storage_path:
        raise HTTPException(status_code=404, detail="Upload not found.")

    chunks_dir = os.path.join(upload.storage_path, "chunks")
    if not os.path.exists(chunks_dir):
        return []

    return sorted([f for f in os.listdir(chunks_dir) if f.endswith(".csv")])


@app.get("/uploads/{upload_id}/chunks/{chunk_name}")
def download_upload_chunk(upload_id: int, chunk_name: str, db: Session = Depends(get_db)):
    """Downloads a specific chunk file from an upload."""
    upload = db.query(Upload).filter(Upload.id == upload_id).first()
    if not upload or not upload.storage_path:
        raise HTTPException(status_code=404, detail="Upload not found.")

    chunk_path = os.path.join(upload.storage_path, "chunks", chunk_name)
    if not os.path.exists(chunk_path):
        raise HTTPException(status_code=404, detail=f"Chunk {chunk_name} not found.")

    return FileResponse(path=chunk_path, filename=chunk_name, media_type="text/csv")


@app.get("/uploads/{upload_id}/chunks-zip")
def download_all_chunks_zip(upload_id: int, db: Session = Depends(get_db)):
    """Downloads all chunks for an upload packaged into a single ZIP archive."""
    zip_bytes = create_chunks_zip(upload_id)
    if not zip_bytes:
        raise HTTPException(status_code=404, detail="No chunks found for this upload.")

    return Response(
        content=zip_bytes,
        media_type="application/zip",
        headers={"Content-Disposition": f'attachment; filename="upload_{upload_id}_chunks.zip"'}
    )


# ==========================================
# UPLOAD HISTORY & RECORD MANAGEMENT
# ==========================================

@app.get("/uploads")
def get_uploads(db: Session = Depends(get_db)):
    """Returns list of all uploads ordered by upload date."""
    uploads = db.query(Upload).order_by(Upload.uploaded_at.desc()).all()
    result = []
    for upload in uploads:
        result.append({
            "id": upload.id,
            "upload_uuid": upload.upload_uuid,
            "filename": upload.filename,
            "file_format": upload.file_format,
            "sheet_name": upload.sheet_name,
            "total_rows": upload.total_rows,
            "valid_rows": upload.valid_rows,
            "invalid_rows": upload.invalid_rows,
            "chunk_count": upload.chunk_count,
            "error_breakdown": json.loads(upload.error_breakdown or "{}"),
            "uploaded_at": upload.uploaded_at.isoformat() if upload.uploaded_at else None
        })
    return result


@app.get("/uploads/{upload_id}")
def get_upload_detail(upload_id: int, db: Session = Depends(get_db)):
    """Returns detailed statistics and rule profile for a specific upload."""
    upload = db.query(Upload).filter(Upload.id == upload_id).first()
    if not upload:
        raise HTTPException(status_code=404, detail="Upload not found.")

    return {
        "id": upload.id,
        "upload_uuid": upload.upload_uuid,
        "filename": upload.filename,
        "file_format": upload.file_format,
        "sheet_name": upload.sheet_name,
        "total_rows": upload.total_rows,
        "valid_rows": upload.valid_rows,
        "invalid_rows": upload.invalid_rows,
        "chunk_count": upload.chunk_count,
        "error_breakdown": json.loads(upload.error_breakdown or "{}"),
        "rule_config": json.loads(upload.rule_config or "[]"),
        "uploaded_at": upload.uploaded_at.isoformat() if upload.uploaded_at else None
    }


@app.delete("/uploads/{upload_id}")
def delete_single_upload(upload_id: int, db: Session = Depends(get_db)):
    """Deletes an upload record and cleans up its storage folder."""
    upload = db.query(Upload).filter(Upload.id == upload_id).first()
    if not upload:
        raise HTTPException(status_code=404, detail="Upload not found.")

    if upload.storage_path and os.path.exists(upload.storage_path):
        shutil.rmtree(upload.storage_path, ignore_errors=True)

    db.delete(upload)
    db.commit()
    return {"message": f"Upload #{upload_id} deleted successfully"}


@app.delete("/uploads")
def delete_uploads(db: Session = Depends(get_db)):
    """Deletes all uploads from history and disk."""
    db.query(Upload).delete()
    db.commit()
    if os.path.exists(UPLOADS_DIR):
        for item in os.listdir(UPLOADS_DIR):
            item_path = os.path.join(UPLOADS_DIR, item)
            if os.path.isdir(item_path):
                shutil.rmtree(item_path, ignore_errors=True)
    return {"message": "Upload history deleted"}


# ==========================================
# BACKWARD COMPATIBLE LEGACY DOWNLOAD ROUTES
# ==========================================

@app.get("/download/clean")
def download_clean():
    clean_path = "outputs/clean_transactions.csv"
    if not os.path.exists(clean_path):
        raise HTTPException(status_code=404, detail="Clean transactions file not found.")
    return FileResponse(clean_path, filename="clean_transactions.csv", media_type="text/csv")


@app.get("/download/error")
def download_error():
    error_path = "errors/validation_errors.csv"
    if not os.path.exists(error_path):
        raise HTTPException(status_code=404, detail="Validation errors file not found.")
    return FileResponse(error_path, filename="validation_errors.csv", media_type="text/csv")


@app.get("/chunks")
def get_chunks():
    chunk_folder = "chunks"
    if not os.path.exists(chunk_folder):
        return []
    return [f for f in os.listdir(chunk_folder) if f.endswith(".csv")]


@app.get("/download/chunk/{filename}")
def download_chunk(filename: str):
    file_path = f"chunks/{filename}"
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail=f"Chunk {filename} not found.")
    return FileResponse(path=file_path, filename=filename, media_type="text/csv")