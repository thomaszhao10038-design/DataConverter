import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment

# -----------------------------
# ROUND TIMESTAMP TO 10 MIN
# -----------------------------
def round_to_10min(ts):
    if pd.isna(ts):
        return ts
    ts = pd.to_datetime(ts)
    m = ts.minute
    r = m % 10
    if r < 5:
        new_m = m - r
    else:
        new_m = m + (10 - r)
    if new_m == 60:
        ts = ts.replace(minute=0) + pd.Timedelta(hours=1)
    else:
        ts = ts.replace(minute=new_m)
    return ts.replace(second=0, microsecond=0)

# -----------------------------
# PROCESS SINGLE SHEET
# -----------------------------
def process_sheet(df, timestamp_col, psum_col):
    df[timestamp_col] = pd.to_datetime(df[timestamp_col], errors="coerce")
    df["Rounded"] = df[timestamp_col].apply(round_to_10min)

    # Extract date and time
    df["Date"] = df["Rounded"].dt.date
    df["Time"] = df["Rounded"].dt.strftime("%H:%M:%S")

    # Create all possible 10-min intervals for each date
    all_days = df["Date"].unique()
    all_intervals = pd.date_range("00:00", "23:50", freq="10min").time
    rows = []
    for d in all_days:
        day_data = df[df["Date"] == d].set_index("Time")
        for t in all_intervals:
            t_str = t.strftime("%H:%M:%S")
            val = day_data[psum_col].get(t_str, 0) if t_str in day_data.index else 0
            rows.append({"Date": d, "Time": t_str, "PSum (W)": val})

    grouped = pd.DataFrame(rows)
    return grouped

# -----------------------------
# BUILD EXCEL FORMAT
# -----------------------------
def build_output_excel(sheets_dict):
    wb = Workbook()
    wb.remove(wb.active)

    for sheet_name, df in sheets_dict.items():
        ws = wb.create_sheet(sheet_name)
        dates = sorted(df["Date"].unique())

        col_start = 1
        for date in dates:
            # Merge date header
            ws.merge_cells(start_row=1, start_column=col_start, end_row=1, end_column=col_start+3)
            ws.cell(row=1, column=col_start, value=str(date))
            ws.cell(row=1, column=col_start).alignment = Alignment(horizontal="center")

            # Sub-headers
            ws.cell(row=2, column=col_start, value="UTC Offset (minutes)")
            ws.cell(row=2, column=col_start+1, value="Local Time Stamp")
            ws.cell(row=2, column=col_start+2, value="Active Power (W)")
            ws.cell(row=2, column=col_start+3, value="kW")

            # Fill 10-min rows
            day_data = df[df["Date"] == date].sort_values("Time")
            for idx, r in enumerate(day_data.itertuples(), start=3):
                ws.cell(row=idx, column=col_start, value=str(date))
                ws.cell(row=idx, column=col_start+1, value=r.Time)
                ws.cell(row=idx, column=col_start+2, value=r._3)  # PSum (W)
                ws.cell(row=idx, column=col_start+3, value=r._3/1000)

            col_start += 4  # next day block

    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    return stream

# -----------------------------
# STREAMLIT UI
# -----------------------------
st.title("📊 Excel 10-Minute Electricity Data Converter")
st.write("Upload an Excel file. Each sheet will be processed separately.")

uploaded = st.file_uploader("Upload .xlsx file", type=["xlsx"])

if uploaded:
    xls = pd.ExcelFile(uploaded)
    result_sheets = {}

    for sheet_name in xls.sheet_names:
        st.write(f"Processing sheet: **{sheet_name}**")
        df = pd.read_excel(uploaded, sheet_name=sheet_name)

        # Auto-detect timestamp column
        possible_time_cols = [
            "Date & Time", "Date&Time", "Date_Time",
            "Timestamp", "TimeStamp", "DateTime", "Date Time",
            "LocalTime", "Local Time", "TIME", "time", "datetime",
            "Date", "date", "ts"
        ]
        timestamp_col = next((col for col in df.columns if col.strip() in possible_time_cols), None)
        if timestamp_col is None:
            st.error(f"No valid timestamp column in sheet {sheet_name}. Columns: {list(df.columns)}")
            continue

        # Auto-detect PSum column
        possible_psum_cols = [
            "PSum (W)", "Psum (W)", "psum", "PSum", "Psum",
            "Power", "Active Power", "ActivePower", "P (W)"
        ]
        psum_col = next((col for col in df.columns if col.strip() in possible_psum_cols), None)
        if psum_col is None:
            st.error(f"No valid PSum column in sheet {sheet_name}. Columns: {list(df.columns)}")
            continue

        processed = process_sheet(df, timestamp_col, psum_col)
        result_sheets[sheet_name] = processed

    if result_sheets:
        output_stream = build_output_excel(result_sheets)
        st.download_button(
            label="📥 Download Converted Excel",
            data=output_stream,
            file_name="Converted_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
