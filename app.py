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
    Processes a single sheet DataFrame: cleans data, rounds timestamps to 10-min, 
    sums power in intervals, pads missing time slots, and calculates kW.
    Removes leading/trailing zero-value time periods.
    """
    df.columns = df.columns.astype(str).str.strip()
    # Convert timestamp column, allowing for dayfirst format inference
    df[timestamp_col] = pd.to_datetime(df[timestamp_col], errors="coerce", dayfirst=True)
    
    # Clean and convert power column
    power_series = df[psum_col].astype(str).str.strip()
    power_series = power_series.str.replace(',', '.', regex=False)
    df[psum_col] = pd.to_numeric(power_series, errors='coerce')
    
    df = df.dropna(subset=[timestamp_col, psum_col])
    if df.empty:
        return pd.DataFrame()
    
    # Filter out leading and trailing zero power periods
    non_zero_indices = df[df[psum_col].abs() != 0].index
    if non_zero_indices.empty:
        return pd.DataFrame()
        
    first_valid_idx = non_zero_indices.min()
    last_valid_idx = non_zero_indices.max()
    df = df.loc[first_valid_idx:last_valid_idx].copy()

    # Resample to 10-minute intervals and sum (total W for that interval)
    df_indexed = df.set_index(timestamp_col)
    df_indexed.index = df_indexed.index.floor('10min')
    resampled_data = df_indexed[psum_col].groupby(level=0).sum()
    
    df_out = resampled_data.reset_index()
    df_out.columns = ['Rounded', POWER_COL_OUT] 
    if df_out.empty or df_out['Rounded'].isna().all():
        return pd.DataFrame()
    
    # Get original dates for final filtering
    original_dates = set(df_out['Rounded'].dt.date)
    
    # Create a full 10-minute time index to ensure all slots exist
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
    
    # Reindex and fill missing 10-min slots with NaN
    df_indexed_for_reindex = df_out.set_index('Rounded')
    df_padded_series = df_indexed_for_reindex[POWER_COL_OUT].reindex(full_time_index)
    
    # Final cleanup and calculations
    grouped = df_padded_series.reset_index().rename(columns={'index': 'Rounded'})
    grouped.columns = ['Rounded', POWER_COL_OUT]
    grouped[POWER_COL_OUT] = grouped[POWER_COL_OUT].astype(float) 
    
    # Extract Date and Time columns
    grouped["Date"] = grouped["Rounded"].dt.date
    grouped["Time"] = grouped["Rounded"].dt.strftime("%H:%M") 
    
    # Only keep data that corresponds to the original dates in the data
    grouped = grouped[grouped["Date"].isin(original_dates)]
    
    # Calculate kW (absolute value)
    grouped['kW'] = grouped[POWER_COL_OUT].abs() / 1000
    
    return grouped

# -----------------------------
# BUILD EXCEL
# -----------------------------
def build_output_excel(sheets_dict):
    """
    Generates the final Excel workbook with per-day data sheets and a summary "Total" sheet.
    """
    wb = Workbook()
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])

    header_fill = PatternFill(start_color='ADD8E6', end_color='ADD8E6', fill_type='solid')
    title_font = Font(bold=True, size=12)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))

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
            # Use full date string (YYYY-MM-DD) for Total sheet key
            date_key_for_total = date_str_full 

            # Date Header
            ws.merge_cells(start_row=1, start_column=col_start, end_row=1, end_column=col_start+3)
            ws.cell(row=1, column=col_start, value=date_str_full).alignment = Alignment(horizontal="center")

            # Column Headers
            ws.cell(row=2, column=col_start, value="UTC Offset (minutes)")
            ws.cell(row=2, column=col_start+1, value="Local Time Stamp")
            ws.cell(row=2, column=col_start+2, value="Active Power (W)")
            ws.cell(row=2, column=col_start+3, value="kW")

            # UTC Offset column (merged)
            ws.merge_cells(start_row=merge_start, start_column=col_start, end_row=merge_end, end_column=col_start)
            ws.cell(row=merge_start, column=col_start, value=date_str_full).alignment = Alignment(horizontal="center", vertical="center")

            # Data rows
            for idx, r in enumerate(day_data_full.itertuples(), start=merge_start):
                ws.cell(row=idx, column=col_start+1, value=r.Time)
                ws.cell(row=idx, column=col_start+2, value=getattr(r, POWER_COL_OUT))
                ws.cell(row=idx, column=col_start+3, value=r.kW)

            # Daily Statistics
            stats_row_start = merge_end + 1
            sum_w = day_data_active[POWER_COL_OUT].sum()
            mean_w = day_data_active[POWER_COL_OUT].mean()
            max_w = day_data_active[POWER_COL_OUT].max()
            sum_kw = day_data_active['kW'].sum()
            mean_kw = day_data_active['kW'].mean()
            max_kw = day_data_active['kW'].max()

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
            daily_max_summary.append((date_key_for_total, max_kw))

            # Save max kw for Total sheet (using the full date key)
            for date_val, max_kw_val in daily_max_summary:
                total_data.setdefault(date_val, {})[sheet_name] = max_kw_val

            col_start += 4

        # Per-Sheet Chart (Max power over the day)
        if dates:
            chart = LineChart()
            chart.title = f"{sheet_name} - Power Consumption"
            chart.y_axis.title = "Power (kW)"
            chart.x_axis.title = "Time"
            max_rows = max(day_intervals)
            # Use the Time column (column B relative to col_start=1) as categories
            categories_ref = Reference(ws, min_col=2, min_row=3, max_row=2+max_rows)
            
            col_start = 1
            for n_rows in day_intervals:
                # Use the kW column (column D relative to col_start=1) as data
                data_ref = Reference(ws, min_col=col_start+3, min_row=3, max_row=2+n_rows)
                chart.add_data(data_ref, titles_from_data=False)
                col_start += 4
            chart.set_categories(categories_ref)
            ws.add_chart(chart, f'G{max_row_used+2}')

    # -----------------------------
    # Total Sheet (Daily Max Power Summary)
    # -----------------------------
    ws_total = wb.create_sheet("Total")
    all_dates = sorted(total_data.keys()) # These are now YYYY-MM-DD strings
    sheet_names = list(sheets_dict.keys())

    # Write Headers
    header_fill_total = PatternFill(start_color='90EE90', end_color='90EE90', fill_type='solid')
    header_cols = ["Date"] + sheet_names + ["Total Load"]
    for col_idx, col_name in enumerate(header_cols, start=1):
        cell = ws_total.cell(row=1, column=col_idx, value=col_name)
        cell.font = Font(bold=True)
        cell.fill = header_fill_total
        cell.alignment = Alignment(horizontal="center")

    # Write Data Rows
    for r_idx, date_val in enumerate(all_dates, start=2):
        # Date column value includes year (e.g., 2024-11-20)
        ws_total.cell(row=r_idx, column=1, value=date_val)
        total_load = 0
        for c_idx, sheet_name in enumerate(sheet_names, start=2):
            value = total_data[date_val].get(sheet_name, 0)
            ws_total.cell(row=r_idx, column=c_idx, value=value)
            total_load += value
        # Total Load column
        ws_total.cell(row=r_idx, column=2+len(sheet_names), value=total_load)

    # Formatting for Total sheet
    for r in ws_total.iter_rows(min_row=1, max_row=1+len(all_dates), min_col=1, max_col=2+len(sheet_names)):
        for cell in r:
            cell.border = thin_border
            # Apply number format to floats (data columns)
            if isinstance(cell.value, (float, int)) and cell.column > 1:
                cell.number_format = numbers.FORMAT_NUMBER_00

    # Adjust column widths
    for col in range(1, 3+len(sheet_names)):
        ws_total.column_dimensions[chr(64+col)].width = 15

    # -----------------------------
    # Total Sheet Line Chart (Daily Max Power)
    # -----------------------------
    chart_total = LineChart()
    # Updated title to match the requested style
    chart_total.title = "Total Building Load" 
    # Updated Y-axis title to match the requested style
    chart_total.y_axis.title = "Power, kW" 
    chart_total.x_axis.title = "Date"

    # Add each sheet as separate series, reading data directly from the Total table
    for c_idx, sheet_name in enumerate(sheet_names, start=2):
        # Data starts at row 1 (header) to get title, down to the last data row
        data_ref = Reference(ws_total, min_col=c_idx, min_row=1, max_row=1+len(all_dates))
        chart_total.add_data(data_ref, titles_from_data=True)
    
    # Add Total Load as last series
    col_idx_total = 2+len(sheet_names)
    data_ref_total = Reference(ws_total, min_col=col_idx_total, min_row=1, max_row=1+len(all_dates))
    chart_total.add_data(data_ref_total, titles_from_data=True)

    # Set Categories (X-axis labels) using the Date column (min_col=1)
    # This uses the YYYY-MM-DD dates in the table for the chart labels.
    dates_ref = Reference(ws_total, min_col=1, min_row=2, max_row=1+len(all_dates))
    chart_total.set_categories(dates_ref)
    
    chart_total.height = 12
    chart_total.width = 30
    ws_total.add_chart(chart_total, "A10")

    # Save to stream
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
        
        Leading and trailing zero values are filtered out and appear blank, zero values within the active recording period are kept.
        
        The output Excel file includes a **line chart**, a **Max Power Summary table**, and a **Total sheet** with daily max power for all sheets and a line chart of all sheets + total load.
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
            # Heuristics for column detection
            timestamp_col = next((c for c in df.columns if c.lower() in ["date & time","date&time","timestamp","datetime","local time","time","ts"]), None)
            psum_col = next((c for c in df.columns if c.lower() in ["psum (w)","psum (w)","psum","p (w)","power"]), None)
            
            if not timestamp_col or not psum_col:
                st.error(f"Sheet '{sheet_name}' is missing required columns. Looked for Timestamp-like (e.g., 'Date & Time') and Power-like (e.g., 'PSum (W)').")
                continue

            processed = process_sheet(df, timestamp_col, psum_col)
            if not processed.empty:
                result_sheets[sheet_name] = processed
                st.success(f"Sheet '{sheet_name}' processed successfully with {len(processed['Date'].unique())} day(s) of data.")
            else:
                st.warning(f"Sheet '{sheet_name}' had no usable data (might be all zeros or contain no valid dates/power values).")

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
