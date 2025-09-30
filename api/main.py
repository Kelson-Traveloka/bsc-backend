from fastapi import FastAPI, HTTPException, UploadFile, File
from fastapi.responses import FileResponse
import pandas as pd
import re
from datetime import datetime
from pathlib import Path
import tempfile
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI()

origins = [
    "http://localhost:3000",
    "http://127.0.0.1:3000",
    "bsc-fe-traveloka.vercel.app"
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,        # or ["*"] to allow all
    allow_credentials=True,
    allow_methods=["*"],          # GET, POST, PUT, etc.
    allow_headers=["*"],          # Any headers
)

@app.get("/")
def root():
    return {"Hello": "World"}

@app.post("/convert")
async def convert_xls(file: UploadFile = File(...)): 
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xls") as tmp:
        contents = await file.read()
        tmp.write(contents)
        tmp_path = Path(tmp.name)
 
    tables = pd.read_html(tmp_path)
    meta = tables[0]

    account_number  = meta.iloc[3, 1]
    currency        = meta.iloc[6, 1]
    period          = meta.iloc[7, 1]
    carried_balance = meta.iloc[9, 1]
    balance         = meta.iloc[9, 3]

    carried_balance_num = float(str(carried_balance).replace(",", "").replace(".00", ""))
    balance_num         = float(str(balance).replace(",", "").replace(".00", ""))

    opening_direction = "C" if carried_balance_num < 0 else "D"
    closing_direction = "C" if balance_num < 0 else "D"

    carried_balance_str = f"{int(carried_balance_num)}" if carried_balance_num.is_integer() else f"{carried_balance_num:.2f}"
    balance_str         = f"{int(balance_num)}" if balance_num.is_integer() else f"{balance_num:.2f}"

    opening_date = ""
    closing_date = ""
    match = re.search(r"From:\s*(\d{2}/\d{2}/\d{4})\s*To:\s*(\d{2}/\d{2}/\d{4})", period)
    if match:
        from_date_str = match.group(1)
        to_date_str   = match.group(2)
        opening_date = datetime.strptime(from_date_str, "%d/%m/%Y").strftime("%Y%m%d")
        closing_date = datetime.strptime(to_date_str, "%d/%m/%Y").strftime("%Y%m%d")

    header_line = (
        f"1;{account_number};{opening_date};{opening_direction};{carried_balance_str};"
        f"{closing_date};{closing_direction};{balance_str};{currency};{account_number}"
    )

    df = pd.read_html(tmp_path, skiprows=11)[0]
    df.columns = ["Transaction Date", "Reference", "Debit Amount", "Credit Amount", "Description"]
    df = df.iloc[:840]

    output_lines = []
    for _, row in df.iterrows():
        date_val = pd.to_datetime(row["Transaction Date"], errors='coerce', dayfirst=True)
        transaction_date = date_val.strftime("%Y%m%d") if pd.notna(date_val) else ""

        debit_val = pd.to_numeric(str(row["Debit Amount"]).replace(",", ""), errors='coerce')
        credit_val = pd.to_numeric(str(row["Credit Amount"]).replace(",", ""), errors='coerce')

        amount = ""
        direction = ""
        if pd.notna(debit_val) and debit_val != 0:
            direction = "D"
            amount = f"{int(debit_val)}" if debit_val.is_integer() else f"{debit_val:.2f}"
        elif pd.notna(credit_val) and credit_val != 0:
            direction = "C"
            amount = f"{int(credit_val)}" if credit_val.is_integer() else f"{credit_val:.2f}"

        description = str(row["Description"]).strip() if pd.notna(row["Description"]) else ""
        reference   = str(row["Reference"]).strip() if pd.notna(row["Reference"]) else ""

        line = (
            f"2;[BankTransactionCode];[InternalBankTransactionCode];"
            f"{transaction_date};{transaction_date};{direction};{amount};{currency};"
            f"{description};{reference};N/A;N/A;"
        )
        output_lines.append(line)
 
    with tempfile.NamedTemporaryFile(delete=False, suffix=".txt", mode="w", encoding="utf-8") as out:
        out.write(header_line + "\n" + "\n".join(output_lines))
        txt_path = Path(out.name)
 
    return FileResponse(
        txt_path,
        media_type="text/plain",
        filename="converted_output.txt"  
    )
