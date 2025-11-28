import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side, numbers
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.series import Series

# --- Configuration ---
POWER_COL_OUT = 'PSumW'

# -----------------------------
# PROCESS SINGLE SHEET
# -----------------------------
def process_sheet(df, timestamp_col, psum_col):
    """
    Processes a single sheet DataFrame:
    1. Standardizes columns and converts types.
    2. Resamples data to 10-minute intervals by summing power values.
    3. Pads the resulting series with zeros for missing 10-minute slots.
    4. Calculates power in kW (absolute value).
    """
    df.columns = df.columns.astype(str).str.strip()
    # Convert timestamp column, allowing for day-first format
    df[timestamp_col] = pd.to_datetime(df[timestamp_col], errors="coerce", dayfirst=True)
    
    # Convert power column, handling comma as decimal separator
    df[psum_col] = pd.to_numeric(df[psum_col].astype(str).str.replace(',', '.', regex=False), errors='coerce')
    df = df.dropna(subset=[timestamp_col, psum_col])
    if df.empty:
        return pd.DataFrame()
    
    # Resample to 10-minute intervals
    df_indexed = df.set_index(timestamp_col)
    df_indexed.index = df_indexed.index.floor('10min')
    # Sum the power values within each 10-minute slot
    resampled_data = df_indexed[psum_col].groupby(level=0).sum()
    
    df_out = resampled_data.reset_index()
    df_out.columns = ['Rounded', POWER_COL_OUT]
    
    if df_out.empty or df_out['Rounded'].isna().all():
        return pd.DataFrame()
    
    # Get the original dates present in the data for filtering later
    original_dates = set(df_out['Rounded'].dt.date)
    min_dt = df_out['Rounded'].min().floor('D')
    max_dt_exclusive = df_out['Rounded'].max().ceil('D') 
    if min_dt >= max_dt_exclusive:
        return pd.DataFrame()
    
    # Create a full 10-minute time index for the date range
    full_time_index = pd.date_range(
        start=min_dt.to_pydatetime(),
        end=max_dt_exclusive.to_pydatetime(),
        freq='10min',
        inclusive='left'
    )
    
    # Reindex and fill missing 10-minute slots with 0 (padding)
    df_indexed_for_reindex = df_out.set_index('Rounded')
    df_padded_series = df_indexed_for_reindex[POWER_COL_OUT].reindex(full_time_index, fill_value=0)
    
    grouped = df_padded_series.reset_index().rename(columns={'index': 'Rounded'})
    grouped.columns = ['Rounded', POWER_COL_OUT]
    
    # Extract Date and Time components
    grouped["Date"] = grouped["Rounded"].dt.date
    grouped["Time"] = grouped["Rounded"].dt.strftime("%H:%M")
    
    # Filter back to only include the dates that were originally present
    grouped = grouped[grouped["Date"].isin(original_dates)]
    
    # Calculate power in kW (absolute value)
    grouped['kW'] = grouped[POWER_COL_OUT].abs() / 1000
    return grouped

# -----------------------------
# BUILD EXCEL
# -----------------------------
def build_output_excel(sheets_dict):
    """
    Creates an Excel workbook with processed data formatted for each date,
    including daily statistics, a multi-series line chart, and a daily summary table.
    """
    wb = Workbook()
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])
    
    # Styles setup
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
        
        chart_series_list = []
        categories_ref = None # Reference for Time (X-axis)
        
        for date in dates:
            day_data = df[df["Date"] == date].sort_values("Time")
            n_rows = len(day_data)
            # Data starts at row 3 (after the two header rows)
            merge_start = 3 
            merge_end = merge_start + n_rows - 1

            date_str_full = date.strftime('%Y-%m-%d')
            date_str_short = date.strftime('%d-%b')

            # --- Data Table Generation ---
            
            # Merge header (Date)
            ws.merge_cells(start_row=1, start_column=col_start, end_row=1, end_column=col_start+3)
            date_cell = ws.cell(row=1, column=col_start, value=date_str_full)
            date_cell.alignment = Alignment(horizontal="center")

            # Sub-headers
            ws.cell(row=2, column=col_start, value="UTC Offset (minutes)")
            ws.cell(row=2, column=col_start+1, value="Local Time Stamp")
            ws.cell(row=2, column=col_start+2, value="Active Power (W)")
            ws.cell(row=2, column=col_start+3, value="kW")

            # Merge UTC Offset (placeholder column, merged down the rows)
            ws.merge_cells(start_row=merge_start, start_column=col_start, end_row=merge_end, end_column=col_start)
            ws.cell(row=merge_start, column=col_start, value=date_str_full).alignment = Alignment(horizontal="center", vertical="center")

            # Fill data
            for idx, r in enumerate(day_data.itertuples(), start=merge_start):
                ws.cell(row=idx, column=col_start+1, value=r.Time) # Local Time Stamp (Time)
                ws.cell(row=idx, column=col_start+2, value=getattr(r, POWER_COL_OUT)) # Active Power (W)
                ws.cell(row=idx, column=col_start+3, value=r.kW).number_format = numbers.FORMAT_NUMBER_00 # kW

            # Summary stats
            stats_row_start = merge_end + 1
            sum_kw = day_data['kW'].sum()
            mean_kw = day_data['kW'].mean()
            max_kw = day_data['kW'].max()
            
            ws.cell(row=stats_row_start, column=col_start+1, value="Total")
            ws.cell(row=stats_row_start, column=col_start+3, value=sum_kw).number_format = numbers.FORMAT_NUMBER_00
            ws.cell(row=stats_row_start+1, column=col_start+1, value="Average")
            ws.cell(row=stats_row_start+1, column=col_start+3, value=mean_kw).number_format = numbers.FORMAT_NUMBER_00
            ws.cell(row=stats_row_start+2, column=col_start+1, value="Max")
            ws.cell(row=stats_row_start+2, column=col_start+3, value=max_kw).number_format = numbers.FORMAT_NUMBER_00
            
            max_row_used = max(max_row_used, stats_row_start+2)
            daily_max_summary.append((date_str_short, max_kw))
            
            # --- Chart Series References ---
            
            # The X-axis categories are the Time stamps (column +1)
            if categories_ref is None:
                categories_ref = Reference(ws, min_col=col_start+1, min_row=merge_start, max_row=merge_end)
            
            # The data series (Y values) are the kW values (column +3)
            data_ref = Reference(ws, min_col=col_start+3, min_row=merge_start, max_row=merge_end)
            
            # --- FIX: Use a Reference object for the series title ---
            # The title is the date string, which is in the cell starting at (1, col_start)
            title_ref = Reference(ws, min_col=col_start, min_row=1)
            
            # Create a Series object for the chart, setting the title using the Reference
            series = Series(data_ref, title=title_ref)
            chart_series_list.append(series)
            
            col_start += 4 # Move to the next set of 4 columns for the next day
        
        # --- Create Chart ---
        if chart_series_list:
            chart = LineChart()
            chart.title = f"{sheet_name} - power consumption"
            
            # CORRECTED AXIS TITLES: Time on X, Power on Y
            chart.x_axis.title = "Time"
            chart.y_axis.title = "Power (kW)"
            
            # This line ensures the category axis uses the Time labels
            chart.set_categories(categories_ref) 
            
            # Add each date series to the chart
            for s in chart_series_list:
                chart.series.append(s)
            
            # Position the chart below the data tables
            ws.add_chart(chart, f'G{max_row_used+2}')
        
        # --- Final Daily Max Summary Table ---
        if daily_max_summary:
            # Position summary table
            start_row = max_row_used + 22 
            ws.cell(row=start_row, column=1, value="Daily Max Power (kW) Summary").font = title_font
            ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=2)
            
            start_row += 1
            ws.cell(row=start_row, column=1, value="Day").fill = header_fill
            ws.cell(row=start_row, column=2, value="Max (kW)").fill = header_fill
            
            for d, (date_str, max_kw) in enumerate(daily_max_summary):
                row = start_row + 1 + d
                ws.cell(row=row, column=1, value=date_str)
                ws.cell(row=row, column=2, value=max_kw).number_format = numbers.FORMAT_NUMBER_00
                
            # Set column widths for summary table
            ws.column_dimensions['A'].width = 15
            ws.column_dimensions['B'].width = 15

    # Save workbook to BytesIO stream
    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    return stream

# -----------------------------
# STREAMLIT APP
# -----------------------------
def app():
    """
    Main Streamlit application interface.
    """
    st.set_page_config(page_title="Electricity Data Converter", layout="wide")
    st.title("⚡ Excel 10-Minute Electricity Data Converter")
    st.markdown("Upload your Excel file containing high-resolution power data. The app will resample it into 10-minute intervals, pad missing intervals with zero, and generate a new Excel file with daily tables, charts, and summaries.")
    
    uploaded = st.file_uploader("Upload .xlsx file", type=["xlsx"])
    
    if uploaded:
        try:
            xls = pd.ExcelFile(uploaded)
            result_sheets = {}
            
            # Automatically detect column names across all sheets
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(uploaded, sheet_name=sheet_name)
                df.columns = df.columns.astype(str).str.strip()
                
                # List of common names for Timestamp and Power columns
                timestamp_col_candidates = ["Date & Time","Date&Time","Timestamp","DateTime","Local Time","TIME","ts", "Rounded"]
                psum_col_candidates = ["PSum (W)","Psum (W)","PSum","P (W)","Power", POWER_COL_OUT]

                timestamp_col = next((c for c in df.columns if c in timestamp_col_candidates), None)
                psum_col = next((c for c in df.columns if c in psum_col_candidates), None)
                
                if timestamp_col and psum_col:
                    processed = process_sheet(df, timestamp_col, psum_col)
                    if not processed.empty:
                        result_sheets[sheet_name] = processed
                    else:
                        st.warning(f"Sheet '{sheet_name}' processed but resulted in empty data (possible date/power parsing failure).")
                else:
                    st.error(f"Sheet '{sheet_name}' is missing required columns. Looked for Timestamp candidates: {timestamp_col_candidates} and Power candidates: {psum_col_candidates}")

            if result_sheets:
                output_stream = build_output_excel(result_sheets)
                st.success("Data successfully processed and Excel file is ready for download!")
                
                st.download_button(
                    "📥 Download Converted Excel", 
                    output_stream,
                    file_name=f"Converted_{uploaded.name.replace('.xlsx', '')}_10min.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("No valid data was found across any sheets for conversion.")

        except Exception as e:
            st.error(f"An error occurred during file processing: {e}")
            st.code(f"Error details: {e}", language='text')

if __name__ == "__main__":
    app()
