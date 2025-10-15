from fastapi import FastAPI, UploadFile, File, HTTPException, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pathlib import Path
import pandas as pd
import numpy as np
import tempfile
import json
import io
import re

app = FastAPI()

origins = [
    "http://localhost:3000",
    "http://127.0.0.1:3000",
    "https://bsc-fe-traveloka.vercel.app"
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    # allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/")
def root():
    return {"message": "FastAPI backend is running!"}

def parse_cell_ref(ref: str):
    if not ref or not isinstance(ref, str):
        return None, None
    m = re.match(r"\[?([A-Z]+)(\d+)\]?", ref.strip())
    if not m:
        return None, None
    col_letters, row_str = m.groups()
    col_index = sum((ord(c) - 64) * (26 ** i) for i, c in enumerate(reversed(col_letters))) - 1
    row_index = int(row_str) - 1
    return col_index, row_index

def handle_excel(contents: bytes, filename: str, mapping: dict):
    try:
        try:
            df = pd.read_excel(io.BytesIO(contents), header=None, engine=None)
        except Exception:
            df = pd.read_html(io.BytesIO(contents))[0]

        # Parse Untuk Header (1)
        _, header_row = parse_cell_ref(mapping.get("Date [Header] *"))
        if header_row is None:
            raise HTTPException(status_code=400, detail="Invalid mapping: missing or invalid Date [Header] * cell reference")

        # Ambil Value Dari Column Date
        header_values = df.iloc[header_row]
        df = df.iloc[header_row + 1:].reset_index(drop=True)
        df.columns = header_values
        df = df.dropna(how="all")
        df.replace({np.nan: None, np.inf: None, -np.inf: None}, inplace=True)

        # Simpan Nama" Header
        col_map = {}
        for key, ref in mapping.items():
            if not ref or not isinstance(ref, str) or "[" not in ref:
                continue
            col_idx, _ = parse_cell_ref(ref)
            if col_idx is not None and col_idx < len(df.columns):
                col_map[key] = df.columns[col_idx]

        # Ubah Nama Header Sesuai Dengan Nama Label Yang Ada Di FE
        df = df.rename(columns={v: k for k, v in col_map.items() if v in df.columns})

        # Ubah Format Tanggal
        if "Date [Header] *" in df.columns:
            df["Transaction Date"] = pd.to_datetime(df["Date [Header] *"], errors="coerce", dayfirst=True)

        # Buat Format Amount Menjadi Siap Di Kalkulasi
        for key in ["Debit Amount *", "Credit Amount *"]:
            if key in df.columns:
                df[key] = (
                    pd.to_numeric(df[key].astype(str).str.replace(",", ""), errors="coerce").fillna(0)
                )

        # Mapping Value Static
        account_number = mapping.get("Account ID *")
        currency = mapping.get("Account Currency *")
        balance = mapping.get("Opening balance amount *")
        statement_id = mapping.get("Statement ID *")
        opening_balance = float(str(balance).replace(",", "").replace(".00", ""))

        # Group By Date
        df["DateOnly"] = df["Transaction Date"].dt.date
        grouped = df.groupby("DateOnly", sort=True)
        output_lines = []

        def fmt_amount(x):
            return f"{int(x)}" if float(x).is_integer() else f"{x:.2f}"

        for day, group in grouped:
            total_debit = group["Debit Amount *"].sum()
            total_credit = group["Credit Amount *"].sum()
            closing_balance = opening_balance + total_credit - total_debit

            opening_direction = "D" if opening_balance < 0 else "C"
            closing_direction = "D" if closing_balance < 0 else "C"

            opening_balance_str = fmt_amount(abs(opening_balance))
            closing_balance_str = fmt_amount(abs(closing_balance))
            date_str = pd.to_datetime(day).strftime("%Y%m%d")

            header_line = (
                f"1;{account_number};{date_str};{opening_direction};{opening_balance_str};"
                f"{date_str};{closing_direction};{closing_balance_str};{currency};{statement_id};"
            )
            output_lines.append(header_line)

            for _, row in group.iterrows():
                transaction_date = row["Transaction Date"].strftime("%Y%m%d")

                amount = ""
                direction = ""
                if row.get("Debit Amount *", 0) != 0:
                    direction = "D"
                    amount = fmt_amount(abs(row["Debit Amount *"]))
                elif row.get("Credit Amount *", 0) != 0:
                    direction = "C"
                    amount = fmt_amount(abs(row["Credit Amount *"]))

                desc_col = mapping.get("Description")
                ref_col = mapping.get("Reference")

                description = ""
                reference = ""

                if desc_col:
                    col_idx, _ = parse_cell_ref(desc_col)
                    if col_idx is not None and col_idx < len(df.columns):
                        description = str(row.get("Description", "")).strip().replace(";", ".")
                if ref_col:
                    col_idx, _ = parse_cell_ref(ref_col)
                    if col_idx is not None and col_idx < len(df.columns):
                        reference = str(row.get("Reference", "")).strip()

                line = (
                    f"2;NTRF;;{transaction_date};{transaction_date};{direction};{amount};{currency};"
                    f"{description};{reference};;;"
                )
                output_lines.append(line)

            opening_balance = closing_balance

        with tempfile.NamedTemporaryFile(delete=False, suffix=".txt", mode="w", encoding="utf-8") as out:
            out.write("\n".join(output_lines))
            txt_path = Path(out.name)

        return FileResponse(
            txt_path,
            media_type="text/plain",
            filename=f"{Path(filename).stem}_converted_excel.txt"
        )

    except Exception as e:
        import traceback; traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Excel conversion failed: {str(e)}")

def handle_csv(contents: bytes, filename: str, mapping: dict):
    encodings_to_try = ["utf-8", "latin-1", "iso-8859-1", "cp1252"]
    df = None
    used_encoding = None
    last_error = None

    for enc in encodings_to_try:
        try:
            df = pd.read_csv(io.BytesIO(contents), encoding=enc, on_bad_lines="skip", sep=None, engine="python", header=None)
            used_encoding = enc
            break
        except Exception as e:
            last_error = e

    if df is None:
        raise HTTPException(status_code=400, detail=f"Failed to read CSV file: {last_error}")

    try:
        # === Same logic as handle_excel ===
        _, header_row = parse_cell_ref(mapping.get("Date [Header] *"))
        if header_row is None:
            raise HTTPException(status_code=400, detail="Invalid mapping: missing or invalid Date [Header] * cell reference")

        header_values = df.iloc[header_row]
        df = df.iloc[header_row + 1:].reset_index(drop=True)
        df.columns = header_values
        df = df.dropna(how="all")
        df.replace({np.nan: None, np.inf: None, -np.inf: None}, inplace=True)

        # Build header mapping
        col_map = {}
        for key, ref in mapping.items():
            if not ref or not isinstance(ref, str) or "[" not in ref:
                continue
            col_idx, _ = parse_cell_ref(ref)
            if col_idx is not None and col_idx < len(df.columns):
                col_map[key] = df.columns[col_idx]

        # Rename columns to frontend mapping names
        df = df.rename(columns={v: k for k, v in col_map.items() if v in df.columns})

        # Convert date column
        if "Date [Header] *" in df.columns:
            df["Transaction Date"] = pd.to_datetime(df["Date [Header] *"], errors="coerce", dayfirst=True)

        # Prepare numeric columns
        for key in ["Debit Amount *", "Credit Amount *"]:
            if key in df.columns:
                df[key] = pd.to_numeric(df[key].astype(str).str.replace(",", ""), errors="coerce").fillna(0)

        # Static mappings
        account_number = mapping.get("Account ID *")
        currency = mapping.get("Account Currency *")
        balance = mapping.get("Opening balance amount *")
        statement_id = mapping.get("Statement ID *")
        opening_balance = float(str(balance).replace(",", "").replace(".00", ""))

        # Group by date
        df["DateOnly"] = df["Transaction Date"].dt.date
        grouped = df.groupby("DateOnly", sort=True)
        output_lines = []

        def fmt_amount(x):
            return f"{int(x)}" if float(x).is_integer() else f"{x:.2f}"

        for day, group in grouped:
            total_debit = group["Debit Amount *"].sum()
            total_credit = group["Credit Amount *"].sum()
            closing_balance = opening_balance + total_credit - total_debit

            opening_direction = "D" if opening_balance < 0 else "C"
            closing_direction = "D" if closing_balance < 0 else "C"

            opening_balance_str = fmt_amount(abs(opening_balance))
            closing_balance_str = fmt_amount(abs(closing_balance))
            date_str = pd.to_datetime(day).strftime("%Y%m%d")

            header_line = (
                f"1;{account_number};{date_str};{opening_direction};{opening_balance_str};"
                f"{date_str};{closing_direction};{closing_balance_str};{currency};{statement_id};"
            )
            output_lines.append(header_line)

            for _, row in group.iterrows():
                transaction_date = row["Transaction Date"].strftime("%Y%m%d")

                amount = ""
                direction = ""
                if row.get("Debit Amount *", 0) != 0:
                    direction = "D"
                    amount = fmt_amount(abs(row["Debit Amount *"]))
                elif row.get("Credit Amount *", 0) != 0:
                    direction = "C"
                    amount = fmt_amount(abs(row["Credit Amount *"]))

                desc_col = mapping.get("Description")
                ref_col = mapping.get("Reference")

                description = ""
                reference = ""

                if desc_col:
                    col_idx, _ = parse_cell_ref(desc_col)
                    if col_idx is not None and col_idx < len(df.columns):
                        description = str(row.get("Description", "")).strip().replace(";", ".")
                if ref_col:
                    col_idx, _ = parse_cell_ref(ref_col)
                    if col_idx is not None and col_idx < len(df.columns):
                        reference = str(row.get("Reference", "")).strip()

                line = (
                    f"2;NTRF;;{transaction_date};{transaction_date};{direction};{amount};{currency};"
                    f"{description};{reference};;;"
                )
                output_lines.append(line)

            opening_balance = closing_balance

        with tempfile.NamedTemporaryFile(delete=False, suffix=".txt", mode="w", encoding="utf-8") as out:
            out.write("\n".join(output_lines))
            txt_path = Path(out.name)

        return FileResponse(
            txt_path,
            media_type="text/plain",
            filename=f"{Path(filename).stem}_converted_csv.txt"
        )

    except Exception as e:
        import traceback; traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"CSV conversion failed: {str(e)}")

@app.post("/convert")
async def read_file(file: UploadFile = File(...), mapping: str = Form(...)):
    filename = file.filename
    suffix = Path(filename).suffix.lower()
    contents = await file.read()
    mapping = json.loads(mapping) 
    print(mapping)
    try:
        if suffix in [".csv", ".txt"]:
            return handle_csv(contents, filename, mapping)
        elif suffix in [".xls", ".xlsx", ".xlsm", ".ods"]:
            return handle_excel(contents, filename, mapping)
        else:
            raise HTTPException(status_code=400, detail=f"Unsupported file type: {suffix}")
    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Failed to read file: {e}")
