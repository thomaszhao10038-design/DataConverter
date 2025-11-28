import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
import datetime as dt


st.set_page_config(page_title="Excel Data Converter", layout="wide")

st.title("📊 Excel Data Converter (10-min Interval Power Data)")


# --------------------------------------------------------------------
# Helper: Round timestamp to LOWER 10-min bucket
# --------------------------------------------------------------------
def round_to_10min(ts: pd.Timestamp):
    return ts.floor("10min")


# --------------------------------------------------------------------
# Build Excel output bytes WITHOUT deprecated functions
# --------------------------------------------------------------------
def build_excel_bytes(data_by_sheet):
    wb = Workbook()
    default = wb.active
    wb.remove(default)

    for sheet_name, date_map in data_by_sheet.items():
        ws = wb.create_sheet(title=sheet_name[:31])

        dates = sorted(date_map.keys())

        subheaders = ["UTC Offset (minutes)", "Local Time Stamp", "Active Power (W)", "kW"]

        # ------------------ HEADER ROWS ------------------
        col_index = 1
        for d in dates:
            start_col = col_index
            end_col = col_index + 3
            start_letter = get_column_letter(start_col)
            end_letter = get_column_letter(end_col)

            # merge date header
            ws.merge_cells(f"{start_letter}1:{end_letter}1")
            cell = ws[f"{start_letter}1"]
            cell.value = d.strftime("%Y-%m-%d")
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")

            # write subheaders
            for i, h in enumerate(subheaders):
                c = ws.cell(row=2, column=col_index + i, value=h)
                c.font = Font(bold=True)

            col_index += 4

        # ------------------ TIME ROWS (144 rows = 24h × 6 intervals/hr) ------------------
        times = []
        base = dt.datetime.combine(dt.date.today(), dt.time(0, 0))

        for _ in range(144):
            times.append(base.time())
            base += dt.timedelta(minutes=10)

        start_row = 3
        for r, t in enumerate(times):
            col_index = 1
            for d in dates:
                # UTC Offset → show the date
                ws.cell(row=start_row + r, column=col_index, value=d.strftime("%Y-%m-%d"))

                # Local timestamp
                ts = dt.datetime.combine(d, t).strftime("%Y-%m-%d %H:%M:%S")
                ws.cell(row=start_row + r, column=col_index + 1, value=ts)

                # Active Power (W)
                val = date_map[d].get(t, None)
                if val is not None:
                    ws.cell(row=start_row + r, column=col_index + 2, value=float(val))

                    # kW (absolute W → kW)
                    ws.cell(row=start_row + r, column=col_index + 3,
                            value=round(abs(float(val)) / 1000, 6))

                col_index += 4

        # Auto column width
        max_col = len(dates) * 4
        for c in range(1, max_col + 1):
            ws.column_dimensions[get_column_letter(c)].width = 18

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()


# --------------------------------------------------------------------
# MAIN PROCESSING LOGIC
# --------------------------------------------------------------------
uploaded = st.file_uploader("Upload your Excel file (.xlsx)", type=["xlsx"])

if uploaded:
    xls = pd.ExcelFile(uploaded)
    sheet_names = xls.sheet_names
    st.success(f"Detected sheets: {sheet_names}")

    data_by_sheet = {}

    for sheet in sheet_names:
        df = pd.read_excel(uploaded, sheet_name=sheet)

        # Check required column
        if "PSum (W)" not in df.columns:
            st.error(f"Sheet '{sheet}' missing required column: PSum (W)")
            continue

        # Standardise timestamp column: auto-detect first datetime-like column
        ts_col = None
        for col in df.columns:
            if np.issubdtype(df[col].dtype, np.datetime64):
                ts_col = col
                break

        if ts_col is None:
            st.error(f"Sheet '{sheet}' has no datetime column.")
            continue

        df = df[[ts_col, "PSum (W)"]].dropna()

        # Round timestamp → 10 min bucket
        df["Rounded"] = df[ts_col].appl_]()
