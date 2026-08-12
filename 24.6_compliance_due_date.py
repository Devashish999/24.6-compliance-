import pandas as pd
import tkinter as tk
from tkinter import filedialog
import re
import datetime as dt
import calendar

def pick_excel_file():
    # Hide the root Tk window
    root = tk.Tk()
    root.withdraw()

    # Ask user to select an Excel file
    filetypes = [("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")]
    path = filedialog.askopenfilename(title="Select Excel file", filetypes=filetypes)

    if not path:
        print("No file selected.")
        return None, None

    # Read the Excel file into a DataFrame
    try:
        df = pd.read_excel(path, sheet_name="PMG Generator Details",engine="openpyxl")
        print(f"Loaded {len(df)} rows from: {path}")
        return df, path
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None, None

def split_phasewise_text(value, base_row):
    """Split a single cell containing phase info into multiple rows."""
    new_rows = []
    phases = str(value).splitlines()
    base_dict = base_row.to_dict()   # convert Series to dict
    for phase in phases:
        mw_match = re.search(r"MW:\s*([\d.]+)", phase)
        date_match = re.search(r"Date:\s*([\d-]+)", phase)
        new_row = base_dict.copy()
        new_row["Phase Info"] = phase.strip()
        new_row["Phase MW"] = float(mw_match.group(1)) if mw_match else None
        new_row["Phase Date"] = pd.to_datetime(date_match.group(1), errors="coerce").date() if date_match else None
        new_rows.append(new_row)
    return new_rows

def parse_month_year(text):
    """Parse strings like 'Aug-26' or 'Mar-2027' into the last day of that month."""
    try:
        # Try parsing with pandas first
        parsed = pd.to_datetime(text, errors="coerce")
        if not pd.isna(parsed):
            # Move to last day of that month
            last_day = calendar.monthrange(parsed.year, parsed.month)[1]
            return dt.date(parsed.year, parsed.month, last_day)
    except Exception:
        return None
    return None

def process_scod(value, base_row):
    """Process SCOD cell: handle Phase blocks, MW+Date pairs, or standalone dates."""
    base_dict = base_row.to_dict()
    text = str(value).strip()
    new_rows = []

    # --- Case 0: Phase blocks ---
    if "Phase" in text and "MW" in text and "Date" in text:
        return split_phasewise_text(text, base_row)

    # --- Case 1: MW + Date pairs ---
    pattern = re.findall(
        r"(\d+)\s*MW\s*([0-9]{8}|\d{1,2}[./-]\d{1,2}[./-]\d{2,4}|[A-Za-z]{3,}-\d{2,4}|\d{4}-\d{2}-\d{2})",
        text
    )
    if pattern:
        for mw, date_str in pattern:
            if re.fullmatch(r"\d{8}", date_str):  # DDMMYYYY
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

    # --- Case 2: Standalone dates ---
    candidates = re.findall(
        r"\d{4}-\d{2}-\d{2}|\d{1,2}[./-][A-Za-z]{3,}[./-]\d{2,4}|\d{1,2}[./-]\d{1,2}[./-]\d{2,4}|\d{8}",
        text
    )
    if not candidates:
        parsed = pd.to_datetime(text, errors="coerce", dayfirst=True)
        if not pd.isna(parsed):
            candidates = [text]

    for date_str in candidates:
        if re.fullmatch(r"\d{8}", date_str):  # DDMMYYYY
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

def expand_part_effective(df):

    new_rows = []
    for idx, row in df.iterrows():
        row_copy = row.copy()

        if str(row_copy["Connectivity Status"]).strip().upper() == "PART EFFECTIVE":
            # --- Original row becomes Effective ---
            effective_row = row_copy.copy()
            effective_row["Connectivity Status"] = "Effective"
            if "Connectivity GNA To Be Made Effective MW" in df.columns:
                effective_row["Connectivity GNA To Be Made Effective MW"] = None
            effective_row["24.6 Compliance due date"] = compute_max_date(effective_row)

            # --- Create one extra row for Not Effective ---
            not_effective_row = row_copy.copy()
            not_effective_row["Connectivity Status"] = "Not Effective"
            if "Connectivity GNA Made Effective MW" in df.columns:
                not_effective_row["Connectivity GNA Made Effective MW"] = None
            if "Start Date Connectivity GNA Made Effective" in df.columns:
                not_effective_row["Start Date Connectivity GNA Made Effective"] = None
            not_effective_row["24.6 Compliance due date"] = compute_max_date(not_effective_row)

            new_rows.append(effective_row)
            new_rows.append(not_effective_row)
        else:
            row_copy["24.6 Compliance due date"] = compute_max_date(row_copy)
            new_rows.append(row_copy)

    return pd.DataFrame(new_rows)

def compute_max_date(row):
    """
    Compute maximum date based on Connectivity Status and Revised Criterion.
    Excel must have these columns exactly:
      - Connectivity Status (Effective / Not Effective)
      - Revised Criterion (LOA/PPA, L&A, L&FC, Land BG Route, Land Route)
      - Connectivity GNA Start Date Firm
      - Phase Date
      - Updated Revised SCOD Generation Project
      - Maximum Delay Permitted By REIA DL
      - Anticipated COD of Terminal Bay
      - Anticipated COD/Charging Date of Transmission System Last Element
      - Start Date Connectivity GNA Made Effective
    """

    status = str(row.get("Connectivity Status", "")).strip().upper()
    criterion = str(row.get("Revised Criterion", "")).strip().upper()
    dates = []

    if status == "NOT EFFECTIVE":
        if criterion == "LOA OR PPA":
            for col in [
                "Connectivity GNA Start Date Firm",
                "Phase Date",
                "Updated Revised SCOD Generation Project",
                "Maximum Delay Permitted By REIA DL",
                "Anticipated COD Of Terminal Bay",
                "Anticipated COD/Charging Date Of Transmission System Last Element"
            ]:
                val = pd.to_datetime(row.get(col), errors="coerce")
                if not pd.isna(val):
                    dates.append(val)

        elif criterion in ("L&A", "L&FC", "LAND BG ROUTE", "LAND ROUTE"):
            conn = pd.to_datetime(row.get("Connectivity GNA Start Date Firm"), errors="coerce")
            scod = pd.to_datetime(row.get("Phase Date"), errors="coerce")
            if not pd.isna(conn):
                dates.append(conn + pd.DateOffset(months=6))
            if not pd.isna(scod):
                dates.append(scod + pd.DateOffset(months=6))
            for col in [
                "Anticipated COD Of Terminal Bay",
                "Anticipated COD/Charging Date Of Transmission System Last Element"
            ]:
                val = pd.to_datetime(row.get(col), errors="coerce")
                if not pd.isna(val):
                    dates.append(val)

    elif status == "EFFECTIVE":
        if criterion == "LOA OR PPA":
            for col in [
                "Connectivity GNA Start Date Firm",
                "Phase Date",
                "Updated Revised SCOD Generation Project",
                "Maximum Delay Permitted By REIA DL",
                "Start Date Connectivity GNA Made Effective"
            ]:
                val = pd.to_datetime(row.get(col), errors="coerce")
                if not pd.isna(val):
                    dates.append(val)

        elif criterion in ("L&A", "L&FC", "LAND BG ROUTE", "LAND ROUTE"):
            conn = pd.to_datetime(row.get("Connectivity GNA Start Date Firm"), errors="coerce")
            scod = pd.to_datetime(row.get("Phase Date"), errors="coerce")
            if not pd.isna(conn):
                dates.append(conn + pd.DateOffset(months=6))
            if not pd.isna(scod):
                dates.append(scod + pd.DateOffset(months=6))
            val = pd.to_datetime(row.get("Start Date Connectivity GNA Made Effective"), errors="coerce")
            if not pd.isna(val):
                dates.append(val)

    if not dates:
        return pd.NaT
    return max(dates)

def safe_max_date(values):
    """Return the maximum valid date from a list, ignoring NaT/None, as a pure date."""
    valid = [v for v in values if not pd.isna(v)]
    if not valid:
        return None
    return max(valid).date()   # ensure only date part is returned


def create_next_three_months_sheet(df_expanded):
    today = dt.date.today()
    start_month = dt.date(today.year, today.month, 1)
    end_month = (start_month + pd.DateOffset(months=3)).date()

    mask = (df_expanded["24.6 Compliance due date"] >= start_month) & (df_expanded["24.6 Compliance due date"] < end_month)
    df_three_months_raw = df_expanded.loc[mask]

    mapped_rows = []
    for _, row in df_three_months_raw.iterrows():
        # Handle special Effective/Not Effective logic
        if str(row["Connectivity Status"]).strip().upper() == "EFFECTIVE":
            val = pd.to_datetime(row.get("Start Date Connectivity GNA Made Effective"), errors="coerce")
            start_date_val = val.date() if not pd.isna(val) else None
        else:
            start_date_val = safe_max_date([
                pd.to_datetime(row.get("Anticipated COD Of Terminal Bay"), errors="coerce"),
                pd.to_datetime(row.get("Anticipated COD/Charging Date Of Transmission System Last Element"), errors="coerce"),
                pd.to_datetime(row.get("Connectivity GNA Start Date Firm"), errors="coerce")
            ])

        new_row = {
            "Application ID": row["Application ID"],
            "Applicant Name": row["Applicant Name"],
            "Region": row["Region"],
            "Revised Criterion": row["Revised Criterion"],
            "Type Of Project Installed Capacity (MW)": f"{row['Type Of Project']} {row['Installed Capacity (MW)']}",
            "Solar": row["Solar"],
            "Wind": row["Wind"],
            "ESS": row["ESS"],
            "Present Connectivity Deemed GNA": row["Present Connectivity Deemed GNA"],
            "Substation Generation Connected": row["Substation Generation Connected"],
            "Connectivity GNA Start Date Firm": pd.to_datetime(row["Connectivity GNA Start Date Firm"],
                                                               errors="coerce").date() if not pd.isna(row["Connectivity GNA Start Date Firm"]) else None,
            "SCOD Phasewise Date": pd.to_datetime(row["Phase Date"], errors="coerce").date() if not pd.isna(row["Phase Date"]) else None,
            "SCOD Phasewise MW": row["Phase MW"],
            "Connectivity GNA Made Effective MW": row["Connectivity GNA Made Effective MW"],
            "Start Date Connectivity GNA Made Effective/ To be made Effective": start_date_val,
            "Connectivity GNA To Be Made Effective MW": row["Connectivity GNA To Be Made Effective MW"],
            "Revised SCOD Generation Project": row["Updated Revised SCOD Generation Project"],
            "Maximum Delay Permitted By REIA DL": row["Maximum Delay Permitted By REIA DL"],
            "Anticipated COD Of Terminal Bay": pd.to_datetime(row["Anticipated COD Of Terminal Bay"],
                                                              errors="coerce").date() if not pd.isna(row["Anticipated COD Of Terminal Bay"]) else None,
            "Anticipated COD/Charging Date Of Transmission System Last Element": pd.to_datetime(
                row["Anticipated COD/Charging Date Of Transmission System Last Element"], errors="coerce").date() if not pd.isna(row["Anticipated COD/Charging Date Of Transmission System Last Element"]) else None,
            "Connectivity Status": row["Connectivity Status"],
            "Generation Commissioning Status": row["Generation Commissioning Status"],
            "CoD Declared For Total Quantum": row["CoD Declared For Total Quantum"],
            "Quantum consider for 24.6 compliance of GNA Reg.": (
                row["Present Connectivity Deemed GNA"] - row["CoD Declared For Total Quantum"]
            ),
            "24.6 Compliance due date": pd.to_datetime(row["24.6 Compliance due date"], errors="coerce").date() if not pd.isna(row["24.6 Compliance due date"]) else None,
        }
        mapped_rows.append(new_row)

    return pd.DataFrame(mapped_rows)

def normalize_date_columns(df, date_columns):
    """Ensure all specified columns are converted to pure date format."""
    for col in date_columns:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce").dt.date
    return df


# ---------------- MAIN ----------------
if __name__ == "__main__":
    df, path = pick_excel_file()
    if df is not None:
        allowed_projects = ["Hybrid", "Solar", "Wind", "Standalone ESS", "Hybrid (RHGS)"]
        df = df[
            df["Type Of Project"].isin(allowed_projects)
            & (df["Application Status"].str.strip().str.upper() == "GRANTED")
            & (df["Date Connectivity Intimation Final"].fillna(0) != 0)
            & (df["Generation Commissioning Status"].str.strip().str.upper() != "COMMISSIONED")
        ]

        '''print(df["SCOD As per Application Phasewise"].head(10))'''

        rows = []
        for _, row in df.iterrows():
            processed = process_scod(row["SCOD As per Application Phasewise"], row)
            rows.extend(processed)
        clean_df = pd.DataFrame(rows)

        if clean_df.empty:
            print("No SCOD rows parsed. Check input data and regex in process_scod.")
        else:
            date_cols = [
                "Phase Date","Connectivity GNA Start Date Firm","Start Date Connectivity GNA Made Effective",
                "Updated Revised SCOD Generation Project","Maximum Delay Permitted By REIA DL",
                "Anticipated COD Of Terminal Bay","Anticipated COD/Charging Date Of Transmission System Last Element",
                "24.6 Compliance due date"
            ]
            clean_df = normalize_date_columns(clean_df, date_cols)

            df_expanded = expand_part_effective(clean_df)
            df_expanded = normalize_date_columns(df_expanded, date_cols)

            df_expanded = df_expanded.rename(columns={
                "24.6 Compliacne due date": "24.6 Compliance due date"
            })

            print("Expanded columns:", df_expanded.columns.tolist())

            df_three_months = create_next_three_months_sheet(df_expanded)

            with pd.ExcelWriter("processed_phases.xlsx", engine="openpyxl") as writer:
                df_expanded.to_excel(writer, sheet_name="Processed Data", index=False)
                df_three_months.to_excel(writer, sheet_name="Next 3 Months", index=False)

            print("Saved Processed Data and Next 3 Months sheets into processed_phases.xlsx")
