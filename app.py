# app.py
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime, timedelta, time
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment
from openpyxl.writer.excel import save_virtual_workbook

st.set_page_config(page_title="DataConverter - Excel 10-min aggregator", layout="wide")

st.title("DataConverter — Excel → 10-min daily layout")
st.caption("Takes an input .xlsx (one or more sheets). For each sheet it produces a sheet where each day occupies 4 columns: "
           "`UTC Offset (minutes)`, `Local Time Stamp`, `Active Power (W)`, `kW`. Dates are shown as merged headers above each day's 4 columns.")

st.markdown("""
**Behavior & assumptions**
- The app tries to detect a datetime column (common names: timestamp, datetime, date, time). If multiple parseable columns exist, the first parseable column is used.
- The app tries to detect the power column by name (contains `'psum'` or `'p_sum'` or `'power'`) — case-insensitive. If not found, the user may select it manually from a dropdown.
- 10-minute bins: timestamps are floored to nearest **lower** 10-minute (e.g. 12:12:01 → 12:10:00). Values in the same bin are **summed**.
- For each day (calendar day, local timestamps), the sheet will have 144 rows (00:00 to 23:50, every 10 minutes).
- `UTC Offset (minutes)` column will contain the date string (YYYY-MM-DD) extracted from the input file (per your instruction).
""")

uploaded_file = st.file_uploader("Upload input Excel (.xlsx)", type=["xlsx"], accept_multiple_files=False)

def guess_datetime_column(df: pd.DataFrame):
    # Try common name matches first
    candidates = []
    names = list(df.columns)
    lowers = [c.lower() for c in names]
    for want in ["timestamp", "date time", "datetime", "date", "time", "local timestamp", "local time stamp", "localtime", "ts"]:
        for i, c in enumerate(lowers):
            if want in c:
                candidates.append(names[i])
    # fallback: any column that can parse as datetime
    for col in names:
        if col in candidates:
            return col
    # try parsing each column quickly
    for col in names:
        try:
            parsed = pd.to_datetime(df[col], errors='coerce')
            if parsed.notna().sum() > 0.5 * len(parsed):  # >50% parseable
                return col
        except Exception:
            continue
    return None

def guess_power_column(df: pd.DataFrame):
    names = list(df.columns)
    lowers = [c.lower() for c in names]
    for i,c in enumerate(lowers):
        if "psum" in c or "p_sum" in c or "p sum" in c or ("power" in c and "active" in c) or "kw" in c and "w" in c:
            return names[i]
    # if none match, choose numeric column with most non-null numeric values
    numeric_cols = []
    for col in names:
        # try convert to numeric
        coerced = pd.to_numeric(df[col], errors='coerce')
        nonnull = coerced.notna().sum()
        if nonnull > 0:
            numeric_cols.append((col, nonnull))
    if numeric_cols:
        # return the column with the highest numeric count
        return sorted(numeric_cols, key=lambda x: x[1], reverse=True)[0][0]
    return None

def floor_to_10min(ts):
    # pandas floor function is easiest, but accept datetime
    return pd.to_datetime(ts).floor('10T')

def process_sheet(df: pd.DataFrame, datetime_col: str, power_col: str):
    df = df.copy()
    # parse datetime
    df['__dt_parsed'] = pd.to_datetime(df[datetime_col], errors='coerce')
    df = df.dropna(subset=['__dt_parsed'])
    # floor to 10-minute
    df['__dt_floor'] = df['__dt_parsed'].dt.floor('10T')
    # date key (calendar date)
    df['__date'] = df['__dt_floor'].dt.date
    # numeric power
    df['__power_w'] = pd.to_numeric(df[power_col], errors='coerce')
    # sum power in same floored timestamp
    grouped = df.groupby(['__date', '__dt_floor'], as_index=False)['__power_w'].sum()

    # build dict: date -> { time -> sum }
    data_by_date = {}
    for date, g in grouped.groupby('__date'):
        # create mapping from time only (HH:MM:SS) or full datetime?
        # We'll map to times (time) for easier insertion into 00:00..23:50
        times = pd.to_datetime(g['__dt_floor']).dt.time
        sums = g['__power_w'].values
        data_by_date[date] = {t: s for t, s in zip(times, sums)}

    return data_by_date

def build_excel_bytes(data_by_sheet):
    """
    data_by_sheet: dict of sheet_name -> dict(date -> {time->value})
    """
    wb = Workbook()
    # remove the default sheet created by Workbook
    default = wb.active
    wb.remove(default)

    for sheet_name, date_map in data_by_sheet.items():
        ws = wb.create_sheet(title=sheet_name[:31])  # Excel sheet name limit

        # sort dates ascending
        dates = sorted(date_map.keys())

        # header row 1: per-day merged date header (merge over 4 columns each)
        # header row 2: subheaders for each day's 4 columns
        subheaders = ["UTC Offset (minutes)", "Local Time Stamp", "Active Power (W)", "kW"]
        # write merged date headers across 4 columns each
        col_index = 1
        for d in dates:
            start_col = col_index
            end_col = col_index + 3
            start_letter = get_column_letter(start_col)
            end_letter = get_column_letter(end_col)
            # Merge and write date string (use YYYY-MM-DD)
            date_str = d.isoformat()
            ws.merge_cells(f"{start_letter}1:{end_letter}1")
            cell = ws[f"{start_letter}1"]
            cell.value = date_str
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            # write subheaders in row 2
            for i, sh in enumerate(subheaders):
                ws.cell(row=2, column=col_index + i, value=sh).font = Font(bold=True)
            col_index += 4

        # Now rows: times 00:00:00 to 23:50:00, step 10 minutes -- 144 rows
        times = []
        cur = datetime.combine(datetime.today(), time(0,0))
        for i in range(144):
            times.append((cur.time()))
            cur = cur + timedelta(minutes=10)

        # Starting from row 3, we write each time row
        start_row = 3
        for row_idx, t in enumerate(times):
            ws.row_dimensions[start_row + row_idx].height = 18
            col_index = 1
            for d in dates:
                # UTC Offset (minutes) column: user asked "this shows the date (search and extract from the input file**)"
                # We'll place the date string (YYYY-MM-DD) in this column (repeated per row). The date is already merged header but user requested date - keep for clarity.
                utc_cell = ws.cell(row=start_row + row_idx, column=col_index)
                utc_cell.value = d.isoformat()    # follows instruction to show date; if you want offset change later
                # Local Time Stamp column: full local timestamp (YYYY-MM-DD HH:MM:SS)
                ts_cell = ws.cell(row=start_row + row_idx, column=col_index + 1)
                local_dt = datetime.combine(d, t)
                # Format as ISO-like string without timezone
                ts_cell.value = local_dt.strftime("%Y-%m-%d %H:%M:%S")
                # Active Power (W)
                ap_cell = ws.cell(row=start_row + row_idx, column=col_index + 2)
                # lookup in data_by_sheet
                v = date_map.get(d, {}).get(t, None)
                if v is None or (isinstance(v, float) and np.isnan(v)):
                    ap_cell.value = None
                else:
                    # if value is nearly integer, cast to int, else keep float
                    if float(v).is_integer():
                        ap_cell.value = int(v)
                    else:
                        ap_cell.value = float(round(v, 6))
                # kW: absolute value / 1000
                kw_cell = ws.cell(row=start_row + row_idx, column=col_index + 3)
                if ap_cell.value is None:
                    kw_cell.value = None
                else:
                    try:
                        kw_value = abs(float(ap_cell.value)) / 1000.0
                        # write with 6 decimal places if needed
                        kw_cell.value = float(round(kw_value, 6))
                    except Exception:
                        kw_cell.value = None
                col_index += 4

        # optional: auto-width (simple)
        max_col = (len(dates) * 4)
        for col in range(1, max_col+1):
            ws.column_dimensions[get_column_letter(col)].width = 18

    # return bytes
    return save_virtual_workbook(wb)

if uploaded_file is not None:
    try:
        # read excel with pandas - preserve sheet names
        excel = pd.ExcelFile(uploaded_file)
        sheet_names = excel.sheet_names

        st.write(f"Found sheets: {sheet_names}")

        # Let user optionally pick columns if automatic detection fails
        user_confirm_cols = {}

        data_by_sheet = {}
        for sheet in sheet_names:
            st.subheader(f"Sheet: {sheet}")
            df = pd.read_excel(excel, sheet_name=sheet)
            if df.empty:
                st.warning(f"Sheet '{sheet}' is empty — skipping.")
                continue

            # try guess
            dt_guess = guess_datetime_column(df)
            pw_guess = guess_power_column(df)

            cols_display = df.columns.tolist()
            st.write("Columns detected:", cols_display)

            # Show guesses and allow manual override
            st.write(f"Auto-detected datetime column: `{dt_guess}`")
            st.write(f"Auto-detected power column: `{pw_guess}`")

            dt_col = dt_guess
            pw_col = pw_guess
            # If any guess is None, allow user to select
            if dt_col is None or pw_col is None:
                with st.form(key=f"manual_cols_{sheet}"):
                    dt_col = st.selectbox("Choose datetime column", options=[None] + cols_display, index=0 if dt_col is None else cols_display.index(dt_col)+1)
                    pw_col = st.selectbox("Choose power column (PSum W)", options=[None] + cols_display, index=0 if pw_col is None else cols_display.index(pw_col)+1)
                    submitted = st.form_submit_button("Use these columns")
                    if submitted:
                        pass  # proceed
            # final check
            if dt_col is None or pw_col is None:
                st.error(f"Could not determine datetime or power column for sheet `{sheet}` — skipping this sheet. Please ensure the sheet contains a datetime column and a numeric power column.")
                continue

            # Process
            try:
                date_map = process_sheet(df, datetime_col=dt_col, power_col=pw_col)
                if len(date_map) == 0:
                    st.warning(f"No parseable timestamps found in sheet `{sheet}`. Skipping.")
                    continue
                data_by_sheet[sheet] = date_map
                st.success(f"Processed sheet `{sheet}` — found {len(date_map)} distinct dates.")
            except Exception as e:
                st.error(f"Error processing sheet `{sheet}`: {e}")
                continue

        if len(data_by_sheet) == 0:
            st.warning("No sheets processed. Upload a valid .xlsx and ensure each sheet has a datetime column and a PSum (W) column.")
        else:
            st.info("Building output Excel...")
            out_bytes = build_excel_bytes(data_by_sheet)
            st.success("Output ready — click to download.")

            bname = uploaded_file.name.replace(".xlsx", "_converted.xlsx")
            st.download_button(label="Download converted Excel", data=out_bytes, file_name=bname, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"Failed to read/process the uploaded Excel: {e}")

else:
    st.info("Upload an .xlsx file to begin.")
