import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side, numbers
from openpyxl.chart import LineChart, Reference

# --- Configuration ---
POWER_COL_OUT = 'PSumW'

# -----------------------------
# PROCESS SINGLE SHEET
# -----------------------------
def process_sheet(df, timestamp_col, psum_col):
    """
    Cleans, converts, and resamples power data to 10-minute intervals.
    Filters out leading and trailing zero/missing readings.
    """
    df.columns = df.columns.astype(str).str.strip()
    df[timestamp_col] = pd.to_datetime(df[timestamp_col], errors="coerce", dayfirst=True)
    
    # Handle potential string commas in power data
    power_series = df[psum_col].astype(str).str.strip()
    power_series = power_series.str.replace(',', '.', regex=False)
    df[psum_col] = pd.to_numeric(power_series, errors='coerce')
    
    df = df.dropna(subset=[timestamp_col, psum_col])
    if df.empty:
        return pd.DataFrame()
    
    # Identify the active data period by non-zero power values
    non_zero_indices = df[df[psum_col].abs() != 0].index
    if non_zero_indices.empty:
        return pd.DataFrame()
        
    first_valid_idx = non_zero_indices.min()
    last_valid_idx = non_zero_indices.max()
    df = df.loc[first_valid_idx:last_valid_idx].copy()

    # Resample to 10-minute intervals (summing power values)
    df_indexed = df.set_index(timestamp_col)
    df_indexed.index = df_indexed.index.floor('10min')
    resampled_data = df_indexed[psum_col].groupby(level=0).sum()
    
    df_out = resampled_data.reset_index()
    df_out.columns = ['Rounded', POWER_COL_OUT] 
    if df_out.empty or df_out['Rounded'].isna().all():
        return pd.DataFrame()
    
    # Create a full 10-minute index for all recorded dates
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
    
    # Reindex and fill missing 10-min slots (NaN)
    df_indexed_for_reindex = df_out.set_index('Rounded')
    df_padded_series = df_indexed_for_reindex[POWER_COL_OUT].reindex(full_time_index)
    
    grouped = df_padded_series.reset_index().rename(columns={'index': 'Rounded'})
    grouped.columns = ['Rounded', POWER_COL_OUT]
    grouped[POWER_COL_OUT] = grouped[POWER_COL_OUT].astype(float) 
    grouped["Date"] = grouped["Rounded"].dt.date
    grouped["Time"] = grouped["Rounded"].dt.strftime("%H:%M") 
    
    # Filter back to only the dates that had original data
    grouped = grouped[grouped["Date"].isin(original_dates)]
    
    # Calculate kW (absolute value for consumption)
    grouped['kW'] = grouped[POWER_COL_OUT].abs() / 1000
    
    return grouped

# -----------------------------
# BUILD EXCEL
# -----------------------------
def build_output_excel(sheets_dict):
    """
    Generates the final Excel workbook with formatted sheets and a Total summary.
    """
    wb = Workbook()
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])

    # --- Styles for Individual Sheets ---
    header_fill = PatternFill(start_color='ADD8E6', end_color='ADD8E6', fill_type='solid') # Light Blue
    summary_fill = PatternFill(start_color='F0F8FF', end_color='F0F8FF', fill_type='solid') # Alice Blue
    title_font = Font(bold=True, size=12)
    summary_font = Font(bold=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))
    power_format = numbers.FORMAT_NUMBER_00 # 0.00

    # Dictionary to hold daily max power per sheet for Total sheet
    total_data = {}

    for sheet_name, df in sheets_dict.items():
        ws = wb.create_sheet(sheet_name)
        dates = sorted(df["Date"].unique())
        col_start = 1
        max_row_used = 0
        daily_max_summary = []
        day_intervals = []

        for date in dates:
            day_data_full = df[df["Date"] == date].sort_values("Time")
            day_data_active = day_data_full.dropna(subset=[POWER_COL_OUT])
            
            n_rows = len(day_data_full)
            day_intervals.append(n_rows)
            merge_start = 3
            merge_end = merge_start + n_rows - 1

            date_str_full = date.strftime('%Y-%m-%d')
            date_str_short = date.strftime('%d-%b')

            # --- Headers ---
            # Merge date header (Row 1)
            ws.merge_cells(start_row=1, start_column=col_start, end_row=1, end_column=col_start+3)
            date_cell = ws.cell(row=1, column=col_start, value=date_str_full)
            date_cell.alignment = Alignment(horizontal="center")
            date_cell.fill = header_fill
            date_cell.border = thin_border

            # Sub-headers (Row 2)
            headers_row_2 = ["UTC Offset (minutes)", "Local Time Stamp", "Active Power (W)", "kW"]
            for i, header in enumerate(headers_row_2):
                cell = ws.cell(row=2, column=col_start+i, value=header)
                cell.fill = header_fill
                cell.font = title_font
                cell.border = thin_border
                
            # --- Data Rows ---
            # Merge UTC column
            ws.merge_cells(start_row=merge_start, start_column=col_start, end_row=merge_end, end_column=col_start)
            utc_cell = ws.cell(row=merge_start, column=col_start, value=date_str_full)
            utc_cell.alignment = Alignment(horizontal="center", vertical="center")
            utc_cell.border = thin_border

            # Fill data
            for idx, r in enumerate(day_data_full.itertuples(), start=merge_start):
                time_cell = ws.cell(row=idx, column=col_start+1, value=r.Time)
                power_w_cell = ws.cell(row=idx, column=col_start+2, value=getattr(r, POWER_COL_OUT))
                power_kw_cell = ws.cell(row=idx, column=col_start+3, value=r.kW)
                
                # Apply formatting
                time_cell.border = thin_border
                power_w_cell.border = thin_border
                power_w_cell.number_format = numbers.FORMAT_NUMBER
                power_kw_cell.border = thin_border
                power_kw_cell.number_format = power_format

            # --- Summary Stats ---
            stats_row_start = merge_end + 1
            sum_w = day_data_active[POWER_COL_OUT].sum()
            mean_w = day_data_active[POWER_COL_OUT].mean()
            max_w = day_data_active[POWER_COL_OUT].max()
            sum_kw = day_data_active['kW'].sum()
            mean_kw = day_data_active['kW'].mean()
            max_kw = day_data_active['kW'].max()

            stats_labels = ["Total", "Average", "Max"]
            stats_w = [sum_w, mean_w, max_w]
            stats_kw = [sum_kw, mean_kw, max_kw]
            
            for i, label in enumerate(stats_labels):
                r = stats_row_start + i
                ws.cell(row=r, column=col_start+1, value=label).font = summary_font
                
                cell_w = ws.cell(row=r, column=col_start+2, value=stats_w[i])
                cell_kw = ws.cell(row=r, column=col_start+3, value=stats_kw[i])
                
                # Apply summary formatting
                for c in range(col_start + 1, col_start + 4):
                    ws.cell(row=r, column=c).fill = summary_fill
                    ws.cell(row=r, column=c).border = thin_border
                    
                cell_w.number_format = numbers.FORMAT_NUMBER
                cell_kw.number_format = power_format
                
            max_row_used = max(max_row_used, stats_row_start+2)
            daily_max_summary.append((date_str_short, max_kw))

            # Save max kw for Total sheet
            total_data.setdefault(sheet_name, {})
            total_data[sheet_name][date_str_short] = max_kw

            col_start += 4

        # Line chart for individual sheet
        if dates:
            chart = LineChart()
            chart.title = f"{sheet_name} - Power Consumption"
            chart.y_axis.title = "Power (kW)"
            chart.x_axis.title = "Time"
            
            max_rows = max(day_intervals)
            # Use Local Time Stamp as categories
            categories_ref = Reference(ws, min_col=col_start - 3, min_row=3, max_row=2+max_rows) 
            
            col_start_data = 1
            for n_rows in day_intervals:
                # Use kW column for data
                data_ref = Reference(ws, min_col=col_start_data+3, min_row=3, max_row=2+n_rows)
                chart.add_data(data_ref, titles_from_data=False)
                col_start_data += 4
                
            chart.set_categories(categories_ref)
            ws.add_chart(chart, f'A{max_row_used+5}')
            
            # Format columns width for all columns in the sheet
            for col_idx in range(1, col_start):
                ws.column_dimensions[chr(64 + col_idx)].width = 15


    # -----------------------------
    # Add Total Sheet (Formatted)
    # -----------------------------
    ws_total = wb.create_sheet("Total")
    
    # Prepare data for easy writing (dates as rows, sheets as columns)
    all_dates = sorted(list(set(date_key for sheet_dict in total_data.values() for date_key in sheet_dict.keys())))
    sheet_names = list(sheets_dict.keys())
    
    # --- Styles for Total Sheet ---
    total_header_fill = PatternFill(start_color='6495ED', end_color='6495ED', fill_type='solid') # Cornflower Blue
    total_header_font = Font(bold=True, color="FFFFFF")
    data_format = numbers.FORMAT_NUMBER_00

    # 1. Write and format Headers (Row 1)
    headers = ["Date"] + sheet_names + ["Total Load"]
    for c_idx, header in enumerate(headers, start=1):
        cell = ws_total.cell(row=1, column=c_idx, value=header)
        cell.font = total_header_font
        cell.fill = total_header_fill
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # 2. Write Data (Start Row 2)
    max_data_row = 1
    for r_idx, date_val in enumerate(all_dates, start=2):
        max_data_row = r_idx
        # Date column (Column 1)
        ws_total.cell(row=r_idx, column=1, value=date_val).border = thin_border
        
        # Individual Sheet Power columns (Columns 2 to N)
        total_load_formula_parts = []
        for c_idx, sheet_name in enumerate(sheet_names, start=2):
            value = total_data.get(sheet_name, {}).get(date_val, 0)
            cell = ws_total.cell(row=r_idx, column=c_idx, value=value)
            cell.number_format = data_format
            cell.border = thin_border
            total_load_formula_parts.append(cell.coordinate)
            
        # Total Load column (Column N+1) - Using SUM formula for better Excel behavior
        total_col_idx = 2 + len(sheet_names)
        
        # Check if there are any sheets to sum
        if total_load_formula_parts:
            formula = f"=SUM({','.join(total_load_formula_parts)})"
            cell_total = ws_total.cell(row=r_idx, column=total_col_idx, value=formula)
        else:
             cell_total = ws_total.cell(row=r_idx, column=total_col_idx, value=0)
             
        cell_total.font = Font(bold=True)
        cell_total.number_format = data_format
        cell_total.border = thin_border


    # 3. Format columns width
    for col in range(1, 3 + len(sheet_names)):
        ws_total.column_dimensions[chr(64 + col)].width = 15

    # 4. Add Line Chart to Total Sheet
    if all_dates:
        chart = LineChart()
        chart.title = "Daily Max Load Comparison (kW)"
        chart.y_axis.title = "Max Power (kW)"
        chart.x_axis.title = "Date"
        chart.style = 10 # A clean, modern style

        # Data series (All columns from B to Total Load column, Row 1 (titles) to max_data_row)
        data_end_col = 2 + len(sheet_names)
        data_ref = Reference(ws_total, min_col=2, min_row=1, max_col=data_end_col, max_row=max_data_row)
        
        # Categories (Date column, Row 2 to max_data_row)
        categories_ref = Reference(ws_total, min_col=1, min_row=2, max_row=max_data_row)

        # titles_from_data=True uses Row 1 as titles
        chart.add_data(data_ref, titles_from_data=True, from_rows=False) 
        chart.set_categories(categories_ref)
        
        # Position the chart
        ws_total.add_chart(chart, f'A{max_data_row + 4}')

    # Save workbook to BytesIO
    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    return stream

# -----------------------------
# STREAMLIT APP
# -----------------------------
def app():
    st.set_page_config(layout="wide", page_title="Electricity Data Converter")
    st.title("📊 Excel 10-Minute Electricity Data Converter")
    st.markdown("""
        Upload an **Excel file (.xlsx)** with time-series data. Each sheet is processed to calculate total absolute power (W) in 10-minute intervals. 
        
        Leading and trailing zero values (representing missing readings) are filtered out and appear blank, but zero values *within* the active recording period are kept.
        
        The output Excel file includes a **line chart** per sheet, a **Max Power Summary table**, and a highly visual **Total sheet** with a summary table and a **comparison chart** of daily max power for all sheets.
    """)

    uploaded = st.file_uploader("Upload .xlsx file", type=["xlsx"])
    if uploaded:
        xls = pd.ExcelFile(uploaded)
        result_sheets = {}
        st.write("---")

        for sheet_name in xls.sheet_names:
            st.markdown(f"**Processing sheet: `{sheet_name}`**")
            try:
                df = pd.read_excel(uploaded, sheet_name=sheet_name)
            except Exception as e:
                st.error(f"Error reading sheet '{sheet_name}': {e}")
                continue

            df.columns = df.columns.astype(str).str.strip()
            # Try to infer timestamp and power columns based on common names
            timestamp_col = next((c for c in df.columns if c.lower() in ["date & time", "date&time", "timestamp", "datetime", "local time", "time", "ts"]), None)
            psum_col = next((c for c in df.columns if c.lower() in ["psum (w)", "psum(w)", "psum", "p (w)", "power"]), None)
            
            if not timestamp_col or not psum_col:
                st.error(f"Sheet '{sheet_name}' missing required columns. Please ensure a timestamp column (e.g., 'Date & Time') and a power column (e.g., 'PSum (W)') exist.")
                continue

            processed = process_sheet(df, timestamp_col, psum_col)
            if not processed.empty:
                result_sheets[sheet_name] = processed
                st.success(f"Sheet '{sheet_name}' processed successfully with {len(processed['Date'].unique())} day(s) of data.")
            else:
                st.warning(f"Sheet '{sheet_name}' had no usable data.")

        if result_sheets:
            output_stream = build_output_excel(result_sheets)
            st.download_button(
                label="📥 Download Converted Excel (Converted_Output.xlsx)",
                data=output_stream,
                file_name="Converted_Output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    app()
