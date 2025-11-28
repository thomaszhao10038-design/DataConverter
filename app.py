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
    Processes a single DataFrame sheet: cleans data, rounds timestamps to 10-minute intervals,
    filters out leading/trailing zero periods, and prepares data for Excel output.
    Periods outside the first and last non-zero reading are filled with NaN (blank) upon re-indexing.
    """
    df.columns = df.columns.astype(str).str.strip()
    # Convert timestamp column, handling various date formats
    df[timestamp_col] = pd.to_datetime(df[timestamp_col], errors="coerce", dayfirst=True)
    
    # Clean and convert power column (handle commas as decimal separators)
    power_series = df[psum_col].astype(str).str.strip()
    power_series = power_series.str.replace(',', '.', regex=False)
    df[psum_col] = pd.to_numeric(power_series, errors='coerce')
    
    df = df.dropna(subset=[timestamp_col, psum_col])
    if df.empty:
        return pd.DataFrame()
    
    # --- CORE LOGIC: FILTER LEADING AND TRAILING ZEROS ---
    
    # Identify indices where the absolute power reading is non-zero
    non_zero_indices = df[df[psum_col].abs() != 0].index
    
    if non_zero_indices.empty:
        # If all valid readings are zero, return an empty DataFrame (no usable period)
        return pd.DataFrame() 
        
    # Get the index of the first and last non-zero reading
    first_valid_idx = non_zero_indices.min()
    last_valid_idx = non_zero_indices.max()
    
    # Slice the DataFrame to keep data between the first and last active reading.
    # This preserves internal zero readings but removes periods before and after activity.
    df = df.loc[first_valid_idx:last_valid_idx].copy()

    # ----------------------------------------------------
    
    # Resample data to 10-minute intervals
    df_indexed = df.set_index(timestamp_col)
    df_indexed.index = df_indexed.index.floor('10min')
    # Sum the power values within each 10-minute slot
    resampled_data = df_indexed[psum_col].groupby(level=0).sum()
    
    df_out = resampled_data.reset_index()
    df_out.columns = ['Rounded', POWER_COL_OUT] 
    
    if df_out.empty or df_out['Rounded'].isna().all():
        return pd.DataFrame()
    
    # Get the original dates present in the processed data
    original_dates = set(df_out['Rounded'].dt.date)
    
    # Create a full 10-minute index from the start of the first day to the end of the last day
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
    
    # Reindex with the full index, filling missing slots with NaN (blank) instead of 0.
    # This ensures periods before the first recorded activity and after the last recorded 
    # activity are blank, while any legitimate 0s within the active period remain 0.
    df_indexed_for_reindex = df_out.set_index('Rounded')
    df_padded_series = df_indexed_for_reindex[POWER_COL_OUT].reindex(full_time_index) # Removed fill_value=0
    
    grouped = df_padded_series.reset_index().rename(columns={'index': 'Rounded'})
    grouped.columns = ['Rounded', POWER_COL_OUT]
    
    # Ensure the column is float type to correctly hold NaN values
    grouped[POWER_COL_OUT] = grouped[POWER_COL_OUT].astype(float) 

    grouped["Date"] = grouped["Rounded"].dt.date
    grouped["Time"] = grouped["Rounded"].dt.strftime("%H:%M") 
    
    # Filter back to only the dates originally present in the file
    grouped = grouped[grouped["Date"].isin(original_dates)]
    
    # Add kW column (absolute value). Since NaN * 1000 = NaN, this works fine.
    grouped['kW'] = grouped[POWER_COL_OUT].abs() / 1000
    
    return grouped

# -----------------------------
# BUILD EXCEL
# -----------------------------
def build_output_excel(sheets_dict):
    """Creates the final formatted Excel file with data, charts, and summary."""
    wb = Workbook()
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])

    header_fill = PatternFill(start_color='ADD8E6', end_color='ADD8E6', fill_type='solid')
    title_font = Font(bold=True, size=12)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                             top=Side(style='thin'), bottom=Side(style='thin'))

    for sheet_name, df in sheets_dict.items():
        ws = wb.create_sheet(sheet_name)
        dates = sorted(df["Date"].unique())
        col_start = 1
        max_row_used = 0
        daily_max_summary = []

        day_intervals = []

        for date in dates:
            # Get all data for the day (including NaNs for missing periods)
            day_data_full = df[df["Date"] == date].sort_values("Time")
            
            # Data used for calculations (excluding the new NaNs from outside the active period)
            # This is crucial for correct statistical calculation (Total, Average, Max)
            day_data_active = day_data_full.dropna(subset=[POWER_COL_OUT])
            
            n_rows = len(day_data_full) # Use full count for row structure
            day_intervals.append(n_rows)
            merge_start = 3
            merge_end = merge_start + n_rows - 1

            date_str_full = date.strftime('%Y-%m-%d')
            date_str_short = date.strftime('%d-%b')

            # Merge date header
            ws.merge_cells(start_row=1, start_column=col_start, end_row=1, end_column=col_start+3)
            ws.cell(row=1, column=col_start, value=date_str_full).alignment = Alignment(horizontal="center")

            # Sub-headers
            ws.cell(row=2, column=col_start, value="UTC Offset (minutes)")
            ws.cell(row=2, column=col_start+1, value="Local Time Stamp")
            ws.cell(row=2, column=col_start+2, value="Active Power (W)")
            ws.cell(row=2, column=col_start+3, value="kW")

            # Merge UTC column
            ws.merge_cells(start_row=merge_start, start_column=col_start, end_row=merge_end, end_column=col_start)
            ws.cell(row=merge_start, column=col_start, value=date_str_full).alignment = Alignment(horizontal="center", vertical="center")

            # Fill data (using day_data_full to include NaNs)
            for idx, r in enumerate(day_data_full.itertuples(), start=merge_start):
                # .itertuples() preserves NaN, which openpyxl writes as blank
                ws.cell(row=idx, column=col_start+1, value=r.Time)
                # Power (W) column
                ws.cell(row=idx, column=col_start+2, value=getattr(r, POWER_COL_OUT)) 
                # kW column
                ws.cell(row=idx, column=col_start+3, value=r.kW)

            # Summary stats (using day_data_active to exclude NaNs)
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
            daily_max_summary.append((date_str_short, max_kw))

            col_start += 4

        # Add Line Chart
        if dates:
            chart = LineChart()
            chart.title = f"Daily 10-Minute Absolute Power Profile - {sheet_name}"
            
            # --- Chart Size Adjustment (Approx. 7.37" H x 19.68" W) ---
            # Openpyxl uses unitless floats for chart dimensions.
            chart.height = 7.37 
            chart.width = 19.68 

            # --- Y-Axis Configuration (kW) ---
            chart.y_axis.title = "kW"
            # Start Y-axis at zero
            chart.y_axis.scaling.min = 0 
            # Removed fixed majorUnit=10: openpyxl will auto-determine the interval
            
            # --- X-Axis Configuration (Time) ---
            chart.x_axis.title = "Time"
            # Removed fixed majorUnit: openpyxl will auto-determine the interval
            # Display time in H:MM format
            chart.x_axis.numberFormat = 'h:mm' 

            max_rows = max(day_intervals)
            # Find the starting column for the first date's time stamps
            first_time_col = 2
            categories_ref = Reference(ws, min_col=first_time_col, min_row=3, max_row=2+max_rows)

            col_start = 1
            for n_rows in day_intervals:
                # Reference for data (kW column)
                data_ref = Reference(ws, min_col=col_start+3, min_row=3, max_row=2+n_rows)
                chart.add_data(data_ref, titles_from_data=False)
                col_start += 4

            chart.set_categories(categories_ref)
            ws.add_chart(chart, f'G{max_row_used+2}')

        # Add Daily Max Summary Table
        if daily_max_summary:
            start_row = max_row_used + 22
            ws.cell(row=start_row, column=1, value="Daily Max Power (kW) Summary").font = title_font
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
    # Remove the default empty sheet created automatically if it's still there
    if 'Sheet' in wb.sheetnames and len(wb.sheetnames) > len(sheets_dict):
        wb.remove(wb['Sheet'])
        
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
        
        **New Feature:** Leading and trailing zero values (representing missing readings) are now filtered out and appear blank, but zero values *within* the active recording period are kept.
        
        The output Excel file includes a **line chart** and a **Max Power Summary table** for each day.
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

            timestamp_col = next((c for c in df.columns if c in ["Date & Time","Date&Time","Timestamp","DateTime","Local Time","TIME","ts"]), None)
            if not timestamp_col:
                st.error(f"No valid timestamp column in sheet '{sheet_name}' (expected: Date & Time, Timestamp, etc.)")
                continue

            psum_col = next((c for c in df.columns if c in ["PSum (W)","Psum (W)","PSum","P (W)","Power"]), None)
            if not psum_col:
                st.error(f"No valid PSum column in sheet '{sheet_name}' (expected: PSum (W), Power, etc.)")
                continue

            processed = process_sheet(df, timestamp_col, psum_col)
            if not processed.empty:
                result_sheets[sheet_name] = processed
                st.success(f"Sheet '{sheet_name}' processed successfully with {len(processed['Date'].unique())} day(s) of data.")
            else:
                st.warning(f"Sheet '{sheet_name}' had no usable data (or all readings were zero/missing).")
        
        st.write("---")
        if result_sheets:
            st.balloons()
            st.success("All usable sheets processed. Generating Excel output...")
            output_stream = build_output_excel(result_sheets)
            st.download_button(
                label="📥 Download Converted Excel (Converted_Output.xlsx)",
                data=output_stream,
                file_name="Converted_Output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        elif uploaded:
            st.error("No data could be processed from the uploaded file.")

if __name__ == "__main__":
    app()
