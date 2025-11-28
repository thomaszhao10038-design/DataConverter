import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side, numbers
from openpyxl.chart import LineChart, Reference, Series

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
    
    non_zero_indices = df[df[psum_col].abs() != 0].index
    if non_zero_indices.empty:
        return pd.DataFrame()
        
    first_valid_idx = non_zero_indices.min()
    last_valid_idx = non_zero_indices.max()
    df = df.loc[first_valid_idx:last_valid_idx].copy()

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
    df_padded_series = df_indexed_for_reindex[POWER_COL_OUT].reindex(full_time_index)
    
    grouped = df_padded_series.reset_index().rename(columns={'index': 'Rounded'})
    grouped.columns = ['Rounded', POWER_COL_OUT]
    grouped[POWER_COL_OUT] = grouped[POWER_COL_OUT].astype(float) 
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
    total_header_fill = PatternFill(start_color='90EE90', end_color='90EE90', fill_type='solid')
    title_font = Font(bold=True, size=12)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))

    # Dictionary to hold daily max power per sheet
    total_data = {}

    # -----------------------------
    # Create individual sheets
    for sheet_name, df in sheets_dict.items():
        ws = wb.create_sheet(sheet_name)
        dates = sorted(df["Date"].unique())
        col_start = 1
        max_row_used = 0
        daily_max_summary = []

        for date in dates:
            day_data_full = df[df["Date"] == date].sort_values("Time")
            day_data_active = day_data_full.dropna(subset=[POWER_COL_OUT])
            
            n_rows = len(day_data_full)
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

            for idx, r in enumerate(day_data_full.itertuples(), start=merge_start):
                ws.cell(row=idx, column=col_start+1, value=r.Time)
                ws.cell(row=idx, column=col_start+2, value=getattr(r, POWER_COL_OUT))
                ws.cell(row=idx, column=col_start+3, value=r.kW)

            stats_row_start = merge_end + 1
            max_kw = day_data_active['kW'].max()
            total_data.setdefault(date_str_full, {})[sheet_name] = max_kw

            max_row_used = max(max_row_used, stats_row_start+2)
            col_start += 4

    # -----------------------------
    # Create "Total" sheet
    ws_total = wb.create_sheet("Total")
    all_dates = sorted(total_data.keys())
    sheet_names = list(sheets_dict.keys())

    # Header
    ws_total.cell(row=1, column=1, value="Date").font = title_font
    for i, sheet_name in enumerate(sheet_names):
        ws_total.cell(row=1, column=2+i, value=sheet_name).font = title_font
    ws_total.cell(row=1, column=2+len(sheet_names), value="Total Load").font = title_font

    # Fill data rows
    for r_idx, date_val in enumerate(all_dates, start=2):
        ws_total.cell(row=r_idx, column=1, value=date_val)
        total_load = 0
        for c_idx, sheet_name in enumerate(sheet_names, start=2):
            value = total_data[date_val].get(sheet_name, 0)
            ws_total.cell(row=r_idx, column=c_idx, value=value)
            total_load += value
        ws_total.cell(row=r_idx, column=2+len(sheet_names), value=total_load)

    # Beautify table
    max_col = 2 + len(sheet_names)
    for r in range(1, len(all_dates)+2):
        for c in range(1, max_col+1):
            cell = ws_total.cell(row=r, column=c)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if r == 1:
                cell.fill = total_header_fill
            if isinstance(cell.value, float):
                cell.number_format = numbers.FORMAT_NUMBER_00

    # Column widths
    for col in range(1, max_col+1):
        ws_total.column_dimensions[chr(64+col)].width = 18

    # Line chart from Total sheet
    chart_total = LineChart()
    chart_total.title = "Daily Max Power per Sheet and Total Load"
    chart_total.y_axis.title = "kW"
    chart_total.x_axis.title = "Date"
    chart_total.height = 12
    chart_total.width = 30

    categories = Reference(ws_total, min_col=1, min_row=2, max_row=1+len(all_dates))
    for c_idx, sheet_name in enumerate(sheet_names, start=2):
        data = Reference(ws_total, min_col=c_idx, min_row=2, max_row=1+len(all_dates))
        chart_total.series.append(Series(data, title=sheet_name))
    # Total load
    data_total = Reference(ws_total, min_col=2+len(sheet_names), min_row=2, max_row=1+len(all_dates))
    chart_total.series.append(Series(data_total, title="Total Load"))

    chart_total.set_categories(categories)
    ws_total.add_chart(chart_total, "A10")

    # Save workbook
    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    return stream

# -----------------------------
# STREAMLIT APP
def app():
    st.set_page_config(layout="wide", page_title="Electricity Data Converter")
    st.title("📊 Excel 10-Minute Electricity Data Converter")

    uploaded = st.file_uploader("Upload .xlsx file", type=["xlsx"])
    if uploaded:
        xls = pd.ExcelFile(uploaded)
        result_sheets = {}
        st.write("---")
        for sheet_name in xls.sheet_names:
            try:
                df = pd.read_excel(uploaded, sheet_name=sheet_name)
            except:
                continue
            df.columns = df.columns.astype(str).str.strip()
            timestamp_col = next((c for c in df.columns if c in ["Date & Time","Date&Time","Timestamp","DateTime","Local Time","TIME","ts"]), None)
            psum_col = next((c for c in df.columns if c in ["PSum (W)","Psum (W)","PSum","P (W)","Power"]), None)
            if timestamp_col and psum_col:
                processed = process_sheet(df, timestamp_col, psum_col)
                if not processed.empty:
                    result_sheets[sheet_name] = processed

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
