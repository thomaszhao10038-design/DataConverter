import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side, numbers
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.series import Series

POWER_COL_OUT = 'PSumW'

def process_sheet(df, timestamp_col, psum_col):
    df.columns = df.columns.astype(str).str.strip()
    df[timestamp_col] = pd.to_datetime(df[timestamp_col], errors="coerce", dayfirst=True)
    df[psum_col] = pd.to_numeric(df[psum_col].astype(str).str.replace(',', '.', regex=False), errors='coerce')
    df = df.dropna(subset=[timestamp_col, psum_col])
    if df.empty:
        return pd.DataFrame()

    df_indexed = df.set_index(timestamp_col)
    df_indexed.index = df_indexed.index.floor('10min')
    resampled = df_indexed[psum_col].groupby(level=0).sum().reset_index()
    resampled.columns = ['Rounded', POWER_COL_OUT]

    if resampled.empty:
        return pd.DataFrame()

    original_dates = set(resampled['Rounded'].dt.date)
    full_range = pd.date_range(
        start=resampled['Rounded'].min().floor('D'),
        end=resampled['Rounded'].max().ceil('D'),
        freq='10min',
        inclusive='left'
    )
    padded = resampled.set_index('Rounded')[POWER_COL_OUT].reindex(full_range, fill_value=0)
    result = padded.reset_index()
    result.columns = ['Rounded', POWER_COL_OUT]
    result["Date"] = result["Rounded"].dt.date
    result["Time"] = result["Rounded"].dt.strftime("%H:%M")
    result = result[result["Date"].isin(original_dates)]
    return result

def build_output_excel(sheets_dict):
    wb = Workbook()
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])

    header_fill = PatternFill(start_color='ADD8E6', end_color='ADD8E6', fill_type='solid')
    title_font = Font(bold=True, size=12)
    header_font = Font(bold=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))
    alt_fill = PatternFill(start_color='F0F8FF', end_color='F0F8FF', fill_type='solid')

    for sheet_name, df in sheets_dict.items():
        ws = wb.create_sheet(title=sheet_name)
        dates = sorted(df["Date"].unique())
        col_start = 1
        daily_max_summary = []
        max_row_used = 0
        chart_categories_ref = None
        series_list = []  # Only store the Series objects

        for date in dates:
            date_str_full = date.strftime('%Y-%m-%d')
            date_str_short = date.strftime('%d-%b')
            day_data = df[df["Date"] == date].sort_values("Time")
            n_rows = len(day_data)
            if n_rows == 0:
                continue

            row_start = 3
            row_end = row_start + n_rows - 1

            # Header
            ws.merge_cells(start_row=1, start_column=col_start, end_row=1, end_column=col_start+3)
            ws.cell(1, col_start, date_str_full).alignment = Alignment(horizontal="center", vertical="center")

            # Sub-headers
            for c, txt in enumerate(["UTC Offset (minutes)", "Local Time Stamp", "Active Power (W)", "kW"], col_start):
                ws.cell(2, c, txt)

            # UTC column merged
            ws.merge_cells(start_row=row_start, start_column=col_start, end_row=row_end, end_column=col_start)
            ws.cell(row_start, col_start, date_str_full).alignment = Alignment(horizontal="center", vertical="center")

            # Data rows
            for idx, row in enumerate(day_data.itertuples(), row_start):
                ws.cell(idx, col_start + 1, value=row.Time)
                power_w = getattr(row, POWER_COL_OUT)
                ws.cell(idx, col_start + 2, value=power_w)
                ws.cell(idx, col_start + 3, value=abs(power_w) / 1000)

            # Summary stats
            sum_kw = day_data[POWER_COL_OUT].abs().sum() / 1000
            avg_kw = day_data[POWER_COL_OUT].abs().mean() / 1000
            max_kw = day_data[POWER_COL_OUT].abs().max() / 1000

            stat_row = row_end + 2
            ws.cell(stat_row,     col_start + 1, "Total");   ws.cell(stat_row,     col_start + 3, sum_kw)
            ws.cell(stat_row + 1, col_start + 1, "Average"); ws.cell(stat_row + 1, col_start + 3, avg_kw)
            ws.cell(stat_row + 2, col_start + 1, "Max");     ws.cell(stat_row + 2, col_start + 3, max_kw)

            max_row_used = max(max_row_used, stat_row + 2)
            daily_max_summary.append((date_str_short, round(max_kw, 2)))

            # === CHART PREPARATION (100% compatible with openpyxl ≥ 3.1) ===
            if chart_categories_ref is None:
                chart_categories_ref = Reference(ws, min_col=col_start + 1, min_row=row_start, max_row=row_end)

            values_ref = Reference(ws, min_col=col_start + 3, min_row=row_start, max_row=row_end)
            title_ref  = Reference(ws, min_col=col_start, min_row=1, max_col=col_start + 3)

            series = Series(values_ref)           # ← Correct: only values
            series.title = title_ref              # Title comes from the merged date cell
            series_list.append(series)            # ← openpyxl will auto-assign idx when appended to chart

            col_start += 4

        # === ADD CHART ===
        if series_list:
            chart = LineChart()
            chart.title = f"Daily 10-Minute Power Profile – {sheet_name}"
            chart.style = 10
            chart.y_axis.title = "Power (kW)"
            chart.x_axis.title = "Time of Day"

            for ser in series_list:
                chart.append(ser)                 # ← This auto-assigns correct idx
            chart.set_categories(chart_categories_ref)

            ws.add_chart(chart, f"G{max_row_used + 5}")
            max_row_used += 30

        # === DAILY MAX SUMMARY TABLE ===
        if daily_max_summary:
            r = max_row_used + 5
            title_cell = ws.cell(r, 1, "Daily Max Power (kW) Summary")
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
            title_cell.font = title_font
            title_cell.alignment = Alignment(horizontal="center")

            r += 2
            for c, h in enumerate(["Day", "Max (kW)"], 1):
                cell = ws.cell(r, c, h)
                cell.fill = header_fill
                cell.font = header_font
                cell.border = thin_border

            for i, (day_str, max_val) in enumerate(daily_max_summary):
                row = r + 1 + i
                fill = alt_fill if i % 2 == 0 else PatternFill(fill_type=None)
                ws.cell(row, 1, day_str).fill = fill
                cell = ws.cell(row, 2, max_val)
                cell.number_format = "0.00"
                cell.fill = fill

            ws.column_dimensions['A'].width = 12
            ws.column_dimensions['B'].width = 15

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ==============================
# STREAMLIT APP
# ==============================
def app():
    st.title("10-Minute Electricity Data Converter")
    st.markdown("Upload an Excel file → get clean 10-minute intervals with charts & summary")

    uploaded = st.file_uploader("Choose .xlsx file", type="xlsx")
    if uploaded:
        xls = pd.ExcelFile(uploaded)
        result_sheets = {}

        for sheet in xls.sheet_names:
            st.info(f"Processing **{sheet}**...")
            df = pd.read_excel(uploaded, sheet_name=sheet)
            df.columns = df.columns.astype(str).str.strip()

            time_cols = ["Date & Time", "Date&Time", "Timestamp", "DateTime", "Local Time", "TIME", "ts"]
            power_cols = ["PSum (W)", "Psum (W)", "PSum", "P (W)", "Power"]

            ts_col = next((c for c in df.columns if c in time_cols), None)
            pw_col = next((c for c in df.columns if c in power_cols), None)

            if not ts_col or not pw_col:
                st.error(f"Missing columns in **{sheet}**")
                continue

            processed = process_sheet(df, ts_col, pw_col)
            if not processed.empty:
                result_sheets[sheet] = processed
                st.success(f"**{sheet}** → {len(processed['Date'].unique())} days")
            else:
                st.warning(f"**{sheet}** → no valid data")

        if result_sheets:
            out = build_output_excel(result_sheets)
            st.success("Conversion complete!")
            st.download_button(
                "Download Converted Excel",
                data=out,
                file_name="10min_Power_Converted.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("No sheets processed successfully.")

if __name__ == "__main__":
    app()
