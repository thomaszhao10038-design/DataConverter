import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side, numbers
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.axis import DateAxis, NumericAxis

# --- Configuration ---
POWER_COL_OUT = 'PSumW'

# -----------------------------
# PROCESS SINGLE SHEET
# -----------------------------
def process_sheet(df, timestamp_col, psum_col):
    df.columns = df.columns.astype(str).str.strip()
    df[timestamp_col] = pd.to_datetime(df[timestamp_col], errors="coerce", dayfirst=True)
    
    power_series = df[psum_col].astype(str).str.strip()
    power_series = power_series.str.replace(',', '.', regex=False)
    df[psum_col] = pd.to_numeric(power_series, errors='coerce')
    
    df = df.dropna(subset=[timestamp_col, psum_col])
    if df.empty:
        return pd.DataFrame()
    
    df_indexed = df.set_index(timestamp_col)
    df_indexed.index = df_indexed.index.floor('10min')
    resampled_data = df_indexed[psum_col].groupby(level=0).sum()
    df_out = resampled_data.reset_index()
    df_out.columns = ['Rounded', POWER_COL_OUT] 
    
    if df_out.empty or df_out['Rounded'].isna().all():
        return pd.DataFrame()
    
    original_dates = set(df_out['Rounded'].dt.date)
    min_dt = df_out['Rounded'].min().floor('D')
    max_dt_exclusive = df_out['Rounded'].max().ceil('D') 
    if min_dt >= max_dt_exclusive:
        return pd.DataFrame()
    
    full_time_index = pd.date_range(
        start=min_dt.to_pydatetime(),
        end=max_dt_exclusive.to_pydatetime(),
        freq='10min',
        inclusive='left'
    )
    
    df_indexed_for_reindex = df_out.set_index('Rounded')
    df_padded_series = df_indexed_for_reindex[POWER_COL_OUT].reindex(full_time_index, fill_value=0)
    
    grouped = df_padded_series.reset_index().rename(columns={'index': 'Rounded'})
    grouped.columns = ['Rounded', POWER_COL_OUT]
    grouped["Date"] = grouped["Rounded"].dt.date
    grouped["Time"] = grouped["Rounded"].dt.strftime("%H:%M") 
    grouped = grouped[grouped["Date"].isin(original_dates)]
    
    grouped['kW'] = grouped[POWER_COL_OUT].abs() / 1000
    
    return grouped

# -----------------------------
# BUILD EXCEL
# -----------------------------
def build_output_excel(sheets_dict):
    wb = Workbook()
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])

    header_fill = PatternFill(start_color='ADD8E6', end_color='ADD8E6', fill_type='solid')
    title_font = Font(bold=True, size=12)
    header_font = Font(bold=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))
    data_fill_alt = PatternFill(start_color='F0F8FF', end_color='F0F8FF', fill_type='solid')

    for sheet_name, df in sheets_dict.items():
        ws = wb.create_sheet(sheet_name)
        dates = sorted(df["Date"].unique())
        col_start = 1
        max_row_used = 0
        daily_max_summary = []

        day_intervals = []

        for date in dates:
            day_data = df[df["Date"] == date].sort_values("Time")
            n_rows = len(day_data)
            day_intervals.append(n_rows)
            merge_start = 3
            merge_end = merge_start + n_rows - 1

            date_str_full = date.strftime('%Y-%m-%d')
            date_str_short = date.strftime('%d-%b')

            ws.merge_cells(start_row=1, start_column=col_start, end_row=1, end_column=col_start+3)
            ws.cell(row=1, column=col_start, value=date_str_full).alignment = Alignment(horizontal="center")

            ws.cell(row=2, column=col_start, value="UTC Offset (minutes)")
            ws.cell(row=2, column=col_start+1, value="Local Time Stamp")
            ws.cell(row=2, column=col_start+2, value="Active Power (W)")
            ws.cell(row=2, column=col_start+3, value="kW")

            ws.merge_cells(start_row=merge_start, start_column=col_start, end_row=merge_end, end_column=col_start)
            ws.cell(row=merge_start, column=col_start, value=date_str_full).alignment = Alignment(horizontal="center", vertical="center")

            for idx, r in enumerate(day_data.itertuples(), start=merge_start):
                ws.cell(row=idx, column=col_start+1, value=r.Time)
                ws.cell(row=idx, column=col_start+2, value=getattr(r, POWER_COL_OUT))
                ws.cell(row=idx, column=col_start+3, value=r.kW)

            stats_row_start = merge_end + 1
            sum_w = day_data[POWER_COL_OUT].sum()
            mean_w = day_data[POWER_COL_OUT].mean()
            max_w = day_data[POWER_COL_OUT].max()
            sum_kw = day_data['kW'].sum()
            mean_kw = day_data['kW'].mean()
            max_kw = day_data['kW'].max()

            ws.cell(row=stats_row_start, column=col_start+1, value="Total")
            ws.cell(row=stats_row_start, column=col_start+2, value=sum_w)
            ws.cell(row=stats_row_start, column=col_start+3, value=sum_kw)
            ws.cell(row=stats_row_start+1, column=col_start+1, value="Average")
            ws.cell(row=stats_row_start+1, column=col_start+2, value=mean_w)
            ws.cell(row=stats_row_start+1, column=col_start+3, value=mean_kw)
            ws.cell(row=stats_row_start+2, column=col_start+1, value="Max")
            ws.cell(row=stats_row_start+2, column=col_start+2, value=max_w)
            ws.cell(row=stats_row_start+2, column=col_start+3, value=max_kw)

            max_row_used = max(max_row_used, stats_row_start+2)
            daily_max_summary.append((date_str_short, max_kw))

            col_start += 4

        # -----------------------------
        # Line Chart (with custom axis)
        # -----------------------------
        if dates:
            chart = LineChart()
            chart.title = f"{sheet_name} - power consumption"
            chart.x_axis.title = "Power (kW)"
            chart.y_axis.title = "Time"
            chart.y_axis.majorUnit = 2  # approximate, label every 2 rows (adjust later)
            chart.y_axis.scaling.orientation = "minMax"

            max_rows = max(day_intervals)
            categories_ref = Reference(ws, min_col=2, min_row=3, max_row=2+max_rows)

            col_start = 1
            for n_rows, date in zip(day_intervals, dates):
                data_ref = Reference(ws, min_col=col_start+3, min_row=3, max_row=2+n_rows)
                chart.add_data(data_ref, titles_from_data=False)
                col_start += 4

            chart.set_categories(categories_ref)
            ws.add_chart(chart, f'G{max_row_used+2}')

        # -----------------------------
        # Daily Max Summary Table
        # -----------------------------
        if daily_max_summary:
            start_row = max_row_used + 22
            ws.cell(row=start_row, column=1, value="Daily Max Power (kW) Summary").font = Font(bold=True, size=12)
            ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=2)
            start_row += 1

            ws.cell(row=start_row, column=1, value="Day").fill = header_fill
            ws.cell(row=start_row, column=2, value="Max (kW)").fill = header_fill

            for d, (date_str, max_kw) in enumerate(daily_max_summary):
                row = start_row+1+d
                ws.cell(row=row, column=1, value=date_str)
                ws.cell(row=row, column=2, value=max_kw).number_format = numbers.FORMAT_NUMBER_00

            ws.column_dimensions['A'].width = 15
            ws.column_dimensions['B'].width = 15

    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    return stream

# -----------------------------
# STREAMLIT APP
# -----------------------------
def app():
    st.title("📊 Excel 10-Minute Electricity Data Converter")
    st.markdown("""
        Upload an **Excel file (.xlsx)** with time-series data.
        Each sheet is processed to calculate total absolute power (W) in 10-minute intervals.
        The output Excel file includes a **line chart** and a **Max Power Summary table**.
    """)

    uploaded = st.file_uploader("Upload .xlsx file", type=["xlsx"])

    if uploaded:
        xls = pd.ExcelFile(uploaded)
        result_sheets = {}

        for sheet_name in xls.sheet_names:
            st.info(f"Processing sheet: **{sheet_name}**")
            try:
                df = pd.read_excel(uploaded, sheet_name=sheet_name)
            except Exception as e:
                st.error(f"Error reading sheet '{sheet_name}': {e}")
                continue

            df.columns = df.columns.astype(str).str.strip()

            timestamp_col = next((c for c in df.columns if c in ["Date & Time","Date&Time","Timestamp","DateTime","Local Time","TIME","ts"]), None)
            if not timestamp_col:
                st.error(f"No valid timestamp column in sheet '{sheet_name}'")
                continue

            psum_col = next((c for c in df.columns if c in ["PSum (W)","Psum (W)","PSum","P (W)","Power"]), None)
            if not psum_col:
                st.error(f"No valid PSum column in sheet '{sheet_name}'")
                continue

            processed = process_sheet(df, timestamp_col, psum_col)
            if not processed.empty:
                result_sheets[sheet_name] = processed
                st.success(f"Sheet '{sheet_name}' processed successfully with {len(processed['Date'].unique())} days of data.")
            else:
                st.warning(f"Sheet '{sheet_name}' had no usable data.")

        if result_sheets:
            output_stream = build_output_excel(result_sheets)
            st.download_button(
                label="📥 Download Converted Excel",
                data=output_stream,
                file_name="Converted_Output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    app()
