import pandas as pd
import tkinter as tk
from tkinter import filedialog

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
        df = pd.read_excel(path, sheet_name="Sheet1",engine="openpyxl")
        print(f"Loaded {len(df)} rows from: {path}")
        return df, path
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None, None


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
      - SCOD Phasewise Date
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
                "SCOD Phasewise Date",
                "Updated Revised SCOD Generation Project",
                "Maximum Delay Permitted By REIA DL",
                "Anticipated COD of Terminal Bay",
                "Anticipated COD/Charging Date Of Transmission System Last Element"
            ]:
                val = pd.to_datetime(row.get(col), errors="coerce")
                if not pd.isna(val):
                    dates.append(val)

        elif criterion in ("L&A", "L&FC", "LAND BG ROUTE", "LAND ROUTE"):
            conn = pd.to_datetime(row.get("Connectivity GNA Start Date Firm"), errors="coerce")
            scod = pd.to_datetime(row.get("SCOD Phasewise Date"), errors="coerce")
            if not pd.isna(conn):
                dates.append(conn + pd.DateOffset(months=6))
            if not pd.isna(scod):
                dates.append(scod + pd.DateOffset(months=6))
            for col in [
                "Anticipated COD of Terminal Bay",
                "Anticipated COD/Charging Date Of Transmission System Last Element"
            ]:
                val = pd.to_datetime(row.get(col), errors="coerce")
                if not pd.isna(val):
                    dates.append(val)

    elif status == "EFFECTIVE":
        if criterion == "LOA OR PPA":
            for col in [
                "Connectivity GNA Start Date Firm",
                "SCOD Phasewise Date",
                "Updated Revised SCOD Generation Project",
                "Maximum Delay Permitted By REIA DL",
                "Start Date Connectivity GNA Made Effective"
            ]:
                val = pd.to_datetime(row.get(col), errors="coerce")
                if not pd.isna(val):
                    dates.append(val)

        elif criterion in ("L&A", "L&FC", "LAND BG ROUTE", "LAND ROUTE"):
            conn = pd.to_datetime(row.get("Connectivity GNA Start Date Firm"), errors="coerce")
            scod = pd.to_datetime(row.get("SCOD Phasewise Date"), errors="coerce")
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

if __name__ == "__main__":
    df, path = pick_excel_file()
    if df is not None:
        # Expand Part Effective rows
        df_expanded = expand_part_effective(df)

        # Ensure date-only output
        df_expanded["24.6 Compliance due date"] = pd.to_datetime(df_expanded["24.6 Compliance due date"], errors="coerce").dt.date

        # Save results back to Excel
        df_expanded.to_excel("processed.xlsx", index=False, engine="openpyxl")
        print("Part Effective rows expanded: original → Effective, extra row → Not Effective. Results saved to processed.xlsx")



