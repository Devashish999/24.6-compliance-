import pandas as pd
import tkinter as tk
from tkinter import filedialog
import re
import datetime as dt
import calendar

# ---------------- Pre-compiled regex patterns ----------------
PHASE_PATTERN = re.compile(r"MW:\s*([\d.]+).*?Date:\s*([0-9./-]+)")
MW_DATE_PATTERN = re.compile(r"(\d+)\s*MW\s*([0-9]{8}|\d{1,2}[./-]\d{1,2}[./-]\d{2,4}|[A-Za-z]{3,}-\d{2,4}|\d{4}-\d{2}-\d{2})")
DATE_PATTERN = re.compile(r"\d{4}-\d{2}-\d{2}|\d{1,2}[./-][A-Za-z]{3,}[./-]\d{2,4}|\d{1,2}[./-]\d{1,2}[./-]\d{2,4}|\d{8}")

# ---------------- File picker ----------------
def pick_excel_file():
    root = tk.Tk()
    root.withdraw()
    filetypes = [("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")]
    path = filedialog.askopenfilename(title="Select Excel file", filetypes=filetypes)
    if not path:
        print("No file selected.")
        return None, None
    try:
        df = pd.read_excel(path, sheet_name="PMG Generator Details", engine="openpyxl")
        print(f"Loaded {len(df)} rows from: {path}")
        return df, path
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None, None

# ---------------- Phasewise split ----------------
def split_phasewise_text(value, base_row):
    new_rows = []
    phases = str(value).splitlines()
    base_dict = base_row.to_dict()
    for phase in phases:
        mw_match = re.search(r"MW:\s*([\d.]+)", phase)
        date_match = re.search(r"Date:\s*([0-9./-]+)", phase)
        new_row = base_dict.copy()
        new_row["Phase Info"] = phase.strip()
        new_row["Phase MW"] = float(mw_match.group(1)) if mw_match else None
        new_row["Phase Date"] = pd.to_datetime(date_match.group(1), errors="coerce").date() if date_match else None
        new_rows.append(new_row)
    return new_rows

# ---------------- Month-year parser ----------------
def parse_month_year(text):
    try:
        parsed = pd.to_datetime(text, errors="coerce")
        if not pd.isna(parsed):
            last_day = calendar.monthrange(parsed.year, parsed.month)[1]
            return dt.date(parsed.year, parsed.month, last_day)
    except Exception:
        return None
    return None

# ---------------- SCOD processor ----------------
def process_scod(value, base_row):
    base_dict = base_row.to_dict()
    text = str(value).strip()
    new_rows = []

    # Case 0: Phase blocks
    if "Phase" in text and "MW" in text and "Date" in text:
        return split_phasewise_text(text, base_row)

    # Case 1: MW + Date pairs
    pattern = MW_DATE_PATTERN.findall(text)
    if pattern:
        for mw, date_str in pattern:
            if re.fullmatch(r"\d{8}", date_str):
                parsed = pd.to_datetime(date_str, format="%d%m%Y", errors="coerce")
            else:
                parsed = pd.to_datetime(date_str, errors="coerce", dayfirst=True)
                if pd.isna(parsed):
                    parsed = parse_month_year(date_str)
            new_row = base_dict.copy()
            new_row["Phase Info"] = f"{mw} MW {date_str}"
            new_row["Phase MW"] = float(mw)
            new_row["Phase Date"] = parsed.date() if parsed is not None and not pd.isna(parsed) else None
            new_rows.append(new_row)
        return new_rows

    # Case 2: Standalone dates
    candidates = DATE_PATTERN.findall(text)
    if not candidates:
        parsed = pd.to_datetime(text, errors="coerce", dayfirst=True)
        if not pd.isna(parsed):
            candidates = [text]
    for date_str in candidates:
        if re.fullmatch(r"\d{8}", date_str):
            parsed = pd.to_datetime(date_str, format="%d%m%Y", errors="coerce")
        else:
            parsed = pd.to_datetime(date_str, errors="coerce", dayfirst=True)
            if pd.isna(parsed):
                parsed = parse_month_year(date_str)
        new_row = base_dict.copy()
        new_row["Phase Info"] = date_str
        new_row["Phase MW"] = None
        new_row["Phase Date"] = parsed.date() if parsed is not None and not pd.isna(parsed) else None
        new_rows.append(new_row)
    return new_rows

# ---------------- Expand Part Effective ----------------
def expand_part_effective(df):
    expanded = []
    for _, row in df.iterrows():
        row_copy = row.copy()
        if str(row_copy["Connectivity Status"]).strip().upper() == "PART EFFECTIVE":
            effective_row = row_copy.copy()
            effective_row["Connectivity Status"] = "Effective"
            effective_row["24.6 Compliance due date"] = compute_max_date(effective_row)
            not_effective_row = row_copy.copy()
            not_effective_row["Connectivity Status"] = "Not Effective"
            not_effective_row["24.6 Compliance due date"] = compute_max_date(not_effective_row)
            expanded.extend([effective_row, not_effective_row])
        else:
            row_copy["24.6 Compliance due date"] = compute_max_date(row_copy)
            expanded.append(row_copy)
    return pd.DataFrame(expanded)

# ---------------- Compliance date calculator ----------------
def compute_max_date(row):
    status = str(row.get("Connectivity Status", "")).strip().upper()
    criterion = str(row.get("Revised Criterion", "")).strip().upper()
    dates = []
    if status == "NOT EFFECTIVE":
        if criterion == "LOA OR PPA":
            for col in ["Connectivity GNA Start Date Firm","Phase Date","Updated Revised SCOD Generation Project",
                        "Maximum Delay Permitted By REIA DL","Anticipated COD Of Terminal Bay",
                        "Anticipated COD/Charging Date Of Transmission System Last Element"]:
                val = pd.to_datetime(row.get(col), errors="coerce")
                if not pd.isna(val): dates.append(val)
        elif criterion in ("L&A","L&FC","LAND BG ROUTE","LAND ROUTE"):
            conn = pd.to_datetime(row.get("Connectivity GNA Start Date Firm"), errors="coerce")
            scod = pd.to_datetime(row.get("Phase Date"), errors="coerce")
            if not pd.isna(conn): dates.append(conn + pd.DateOffset(months=6))
            if not pd.isna(scod): dates.append(scod + pd.DateOffset(months=6))
            for col in ["Anticipated COD Of Terminal Bay","Anticipated COD/Charging Date Of Transmission System Last Element"]:
                val = pd.to_datetime(row.get(col), errors="coerce")
                if not pd.isna(val): dates.append(val)
    elif status == "EFFECTIVE":
        if criterion == "LOA OR PPA":
            for col in ["Connectivity GNA Start Date Firm","Phase Date","Updated Revised SCOD Generation Project",
                        "Maximum Delay Permitted By REIA DL","Start Date Connectivity GNA Made Effective"]:
                val = pd.to_datetime(row.get(col), errors="coerce")
                if not pd.isna(val): dates.append(val)
        elif criterion in ("L&A","L&FC","LAND BG ROUTE","LAND ROUTE"):
            conn = pd.to_datetime(row.get("Connectivity GNA Start Date Firm"), errors="coerce")
            scod = pd.to_datetime(row.get("Phase Date"), errors="coerce")
            if not pd.isna(conn): dates.append(conn + pd.DateOffset(months=6))
            if not pd.isna(scod): dates.append(scod + pd.DateOffset(months=6))
            val = pd.to_datetime(row.get("Start Date Connectivity GNA Made Effective"), errors="coerce")
            if not pd.isna(val): dates.append(val)
    return max(dates) if dates else pd.NaT

# ---------------- Safe max date ----------------
def safe_max_date(values):
    valid = [v for v in values if not pd.isna(v)]
    return max(valid).date() if valid else None

# ---------------- Next 3 months sheet ----------------
def create_next_three_months_sheet(df_expanded):
    today = dt.date.today()
    start_month = dt.date(today.year, today.month, 1)
    end_month = (start_month + pd.DateOffset(months=3)).date()
    mask = (df_expanded["24.6 Compliance due date"] >= start_month) & (df_expanded["24.6 Compliance due date"] < end_month)
    df_three_months_raw = df_expanded.loc[mask]
    mapped_rows = []
    for _, row in df_three_months_raw.iterrows():
        if str(row