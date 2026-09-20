import os
import io
import re
import zipfile
import html
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd
from typing import Tuple, List, Dict, Any, Optional

from config import UPLOADS_DIR, DEFAULT_CHUNK_SIZE
from schemas import ChunkSettings


def sanitize_xml_tag(tag: str) -> str:
    cleaned = re.sub(r"[^a-zA-Z0-9_]", "_", str(tag))
    if cleaned and cleaned[0].isdigit():
        cleaned = "_" + cleaned
    return cleaned or "field"


def dataframe_to_xml(df: pd.DataFrame) -> bytes:
    xml_lines = ['<?xml version="1.0" encoding="UTF-8"?>', f'<dataset total_records="{len(df)}">']
    for _, row in df.iterrows():
        xml_lines.append("  <record>")
        for col, val in row.items():
            tag = sanitize_xml_tag(col)
            val_str = html.escape(str(val)) if not pd.isna(val) else ""
            xml_lines.append(f"    <{tag}>{val_str}</{tag}>")
        xml_lines.append("  </record>")
    xml_lines.append("</dataset>")
    return "\n".join(xml_lines).encode("utf-8")


def dataframe_to_styled_excel(df: pd.DataFrame, sheet_name: str = "Transactions") -> bytes:
    if df.empty and len(df.columns) == 0:
        df = pd.DataFrame({"Status": ["No records"]})

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        # Header styling
        header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
        thin_border = Border(
            left=Side(style='thin', color='CBD5E1'),
            right=Side(style='thin', color='CBD5E1'),
            top=Side(style='thin', color='CBD5E1'),
            bottom=Side(style='thin', color='CBD5E1')
        )
        zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
        regular_font = Font(name="Segoe UI", size=10)

        for col_idx, col_name in enumerate(df.columns, 1):
            cell = worksheet.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border

        # Alternate row coloring & borders
        for row_idx in range(2, len(df) + 2):
            fill = zebra_fill if row_idx % 2 == 0 else PatternFill(fill_type=None)
            for col_idx in range(1, len(df.columns) + 1):
                cell = worksheet.cell(row=row_idx, column=col_idx)
                if row_idx % 2 == 0:
                    cell.fill = fill
                cell.font = regular_font
                cell.border = thin_border

        # Auto-fit column widths
        for col_idx, col in enumerate(worksheet.columns, 1):
            cell_lengths = [len(str(cell.value or '')) for cell in col]
            max_len = max(cell_lengths) if cell_lengths else 10
            col_letter = get_column_letter(col_idx)
            worksheet.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 45)

        # Freeze top header row
        worksheet.freeze_panes = "A2"

    return output.getvalue()


def dataframe_to_styled_html(df: pd.DataFrame, title: str = "Transaction Report") -> bytes:
    table_html = df.to_html(classes="styled-table", index=False, escape=True)
    html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>{html.escape(title)}</title>
    <style>
        body {{
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            background-color: #f1f5f9;
            margin: 0;
            padding: 30px;
            color: #1e293b;
        }}
        .report-card {{
            background: white;
            border-radius: 12px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.06);
            padding: 25px;
            max-width: 1300px;
            margin: 0 auto;
        }}
        h2 {{
            color: #0f172a;
            margin-top: 0;
        }}
        .meta {{
            color: #64748b;
            font-size: 14px;
            margin-bottom: 20px;
        }}
        .table-responsive {{
            overflow-x: auto;
        }}
        .styled-table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 14px;
        }}
        .styled-table th {{
            background-color: #1e3a8a;
            color: white;
            text-align: left;
            padding: 12px 16px;
            position: sticky;
            top: 0;
        }}
        .styled-table td {{
            padding: 10px 16px;
            border-bottom: 1px solid #e2e8f0;
        }}
        .styled-table tr:nth-child(even) {{
            background-color: #f8fafc;
        }}
        .styled-table tr:hover {{
            background-color: #eff6ff;
        }}
    </style>
</head>
<body>
    <div class="report-card">
        <h2>{html.escape(title)}</h2>
        <div class="meta">Total Records: <strong>{len(df)}</strong></div>
        <div class="table-responsive">
            {table_html}
        </div>
    </div>
</body>
</html>
"""
    return html_content.encode("utf-8")


def export_dataframe(
    df: pd.DataFrame,
    export_format: str,
    base_name: str = "transactions"
) -> Tuple[bytes, str, str]:
    """
    Exports a DataFrame into the target format.
    Returns: (file_bytes, media_type, file_name)
    """
    fmt = export_format.lower().strip()

    if fmt == "csv":
        return df.to_csv(index=False).encode("utf-8"), "text/csv", f"{base_name}.csv"
    elif fmt == "tsv":
        return df.to_csv(index=False, sep="\t").encode("utf-8"), "text/tab-separated-values", f"{base_name}.tsv"
    elif fmt in ["xlsx", "excel"]:
        return dataframe_to_styled_excel(df, sheet_name=base_name[:30]), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", f"{base_name}.xlsx"
    elif fmt == "json":
        return df.to_json(orient="records", indent=2).encode("utf-8"), "application/json", f"{base_name}.json"
    elif fmt == "xml":
        return dataframe_to_xml(df), "application/xml", f"{base_name}.xml"
    elif fmt == "html":
        return dataframe_to_styled_html(df, title=base_name.replace('_', ' ').title()), "text/html", f"{base_name}.html"
    else:
        raise ValueError(f"Unsupported export format '{fmt}'. Supported: csv, tsv, xlsx, json, xml, html, zip")


def save_upload_artifacts(
    upload_id: int,
    clean_df: pd.DataFrame,
    error_df: pd.DataFrame,
    chunk_settings: Optional[ChunkSettings] = None
) -> Tuple[str, int]:
    """
    Saves clean and error dataframes into an isolated folder for the given upload_id.
    Optionally generates chunks according to chunk_settings.
    Returns (upload_storage_dir, chunk_count).
    """
    upload_dir = os.path.join(UPLOADS_DIR, str(upload_id))
    os.makedirs(upload_dir, exist_ok=True)

    # Save master clean & error files
    clean_path = os.path.join(upload_dir, "clean_transactions.csv")
    error_path = os.path.join(upload_dir, "validation_errors.csv")

    clean_df.to_csv(clean_path, index=False)
    error_df.to_csv(error_path, index=False)

    # Manage chunks
    chunk_count = 0
    chunks_dir = os.path.join(upload_dir, "chunks")

    if chunk_settings and chunk_settings.enabled and len(clean_df) > 0:
        os.makedirs(chunks_dir, exist_ok=True)
        chunk_size = chunk_settings.chunk_size or DEFAULT_CHUNK_SIZE
        total_rows = len(clean_df)

        for i in range(0, total_rows, chunk_size):
            chunk_df = clean_df.iloc[i:i + chunk_size]
            chunk_num = (i // chunk_size) + 1
            chunk_file = os.path.join(chunks_dir, f"chunk_{chunk_num}.csv")
            chunk_df.to_csv(chunk_file, index=False)
            chunk_count += 1

    return upload_dir, chunk_count


def create_chunks_zip(upload_id: int) -> Optional[bytes]:
    """Creates a ZIP of all chunks for a specific upload."""
    chunks_dir = os.path.join(UPLOADS_DIR, str(upload_id), "chunks")
    if not os.path.exists(chunks_dir):
        return None

    chunk_files = [f for f in os.listdir(chunks_dir) if f.endswith(".csv")]
    if not chunk_files:
        return None

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for fname in sorted(chunk_files):
            fpath = os.path.join(chunks_dir, fname)
            zf.write(fpath, arcname=fname)

    return zip_buffer.getvalue()


def create_full_bundle_zip(upload_id: int, clean_df: pd.DataFrame, error_df: pd.DataFrame) -> bytes:
    """Creates a complete ZIP bundle containing clean, error, and all chunks."""
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("clean_transactions.csv", clean_df.to_csv(index=False))
        zf.writestr("validation_errors.csv", error_df.to_csv(index=False))

        chunks_dir = os.path.join(UPLOADS_DIR, str(upload_id), "chunks")
        if os.path.exists(chunks_dir):
            for fname in sorted(os.listdir(chunks_dir)):
                if fname.endswith(".csv"):
                    fpath = os.path.join(chunks_dir, fname)
                    zf.write(fpath, arcname=f"chunks/{fname}")

    return zip_buffer.getvalue()
