import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
# Import PatternFill, Font, numbers, and the necessary chart components
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side, numbers
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.series import Series # Crucial for chart data

# --- Configuration ---
POWER_COL_OUT = 'PSumW'

# -----------------------------
# PROCESS SINGLE SHEET (FINAL ROBUST VERSION)
# -----------------------------
def process_sheet(df, timestamp_col, psum_col):
    """
    Processes a single sheet, including robust date handling and full padding.
    """
    # 1. Cleaning and Preparation
    
    df.columns = df.columns.astype(str).str.strip()
    df[timestamp_col] = pd.to_datetime(df[timestamp_col], errors="coerce", dayfirst=True)
    
    power_series = df[psum_col].astype(str).str.strip()
    power_series = power_series.str.replace(',', '.', regex=False)
    df[psum_col] = pd.to_numeric(power_series, errors='coerce')
    
    df = df.dropna(subset=[timestamp_col, psum_col])
    
    if df.empty:
        return pd.DataFrame()
    
    # 2. Aggregate Data (Fixed Logic)
    
    # df[psum_col] = df[psum_col].abs() is REMOVED to preserve the sign
    
    df_indexed = df.set_index(timestamp_col)
    
    # Floor the index to the nearest 10 minutes
    df_indexed.index = df_indexed.index.floor('10min')
    
    # Group by the new, rounded index and sum the power values
    resampled_data = df_indexed[psum_col].groupby(level=0).sum()
    
    df_out = resampled_data.reset_index()
    df_out.columns = ['Rounded', POWER_COL_OUT] 
    
    # 3. Robust Padding (Ensuring all 10-min slots for all days are present)
    
    if df_out.empty or df_out['Rounded'].isna().all():
        return pd.DataFrame()
    
    # Store the set of original valid dates to filter the final padded range
    original_dates = set(df_out['Rounded'].dt.date)

    # Use floor('D') and ceil('D') for robust date range calculation
    min_dt = df_out['Rounded'].min().floor('D')
    max_dt_exclusive = df_out['Rounded'].max().ceil('D') 
    
    # Ensure the date range is valid (start < end)
    if min_dt >= max_dt_exclusive:
        return pd.DataFrame()

    # Create a continuous 10-minute index covering the entire data range
    full_time_index = pd.date_range(
        # Convert to native Python datetime objects for compatibility
        start=min_dt.to_pydatetime(), 
        end=max_dt_exclusive.to_pydatetime(),
        freq='10min',
        # FIX: Replaced 'closed' with 'inclusive' for Pandas version compatibility
        inclusive='left' 
    )

    # Re-index the resampled data onto the full time index, filling missing intervals with 0
    df_indexed_for_reindex = df_out.set_index('Rounded')
    df_padded_series = df_indexed_for_reindex[POWER_COL_OUT].reindex(full_time_index, fill_value=0)
    
    # Convert back to DataFrame and clean up
    grouped = df_padded_series.reset_index().rename(columns={'index': 'Rounded'})
    grouped.columns = ['Rounded', POWER_COL_OUT]
    
    # 4. Final Formatting
    
    grouped["Date"] = grouped["Rounded"].dt.date
    grouped["Time"] = grouped["Rounded"].dt.strftime("%H:%M") 

    # Filter the result to only include dates that were present in the original data.
    grouped = grouped[grouped["Date"].isin(original_dates)]
    
    return grouped

# -----------------------------
# BUILD EXCEL FORMAT
# -----------------------------
def build_output_excel(sheets_dict):
    """Builds the final Excel workbook with the wide, merged column format and includes a daily max kW chart."""
    wb = Workbook()
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])

    # Define styles for the final summary table
    header_fill = PatternFill(start_color='ADD8E6', end_color='ADD8E6', fill_type='solid') # Light Blue
    title_font = Font(bold=True, size=12)
    header_font = Font(bold=True)
    # Define a thin black border
    thin_border = Border(left=Side(style='thin'), 
                          right=Side(style='thin'), 
                          top=Side(style='thin'), 
                          bottom=Side(style='thin'))
    # Alternating row color (AliceBlue)
    data_fill_alt = PatternFill(start_color='F0F8FF', end_color='F0F8FF', fill_type='solid')


    for sheet_name, df in sheets_dict.items():
        ws = wb.create_sheet(sheet_name)
        dates = sorted(df["Date"].unique())

        col_start = 1
        daily_max_summary = [] # List to store max kW for the final summary table
        max_row_used = 0 # Track the lowest row written to across all columns
        
        # Define containers to store column references for charting
        chart_categories_ref = None # Will store the reference to the first 'Local Time Stamp' column
        chart_data_refs = [] # Will store Series objects for each day's kW data


        for date in dates:
            # 1. Update date format to DD-Mon (e.g., 12-Nov) for the summary table
            date_str_short = date.strftime('%d-%b') 
            
            # Use original date string for main table header
            date_str_full = date.strftime('%Y-%m-%d')
            
            day_data = df[df["Date"] == date].sort_values("Time")
            data_rows_count = len(day_data)
            merge_start_row = 3
            merge_end_row = 2 + data_rows_count
            
            # 1. Merge date header (Row 1, columns 1 to 4)
            ws.merge_cells(start_row=1, start_column=col_start, end_row=1, end_column=col_start+3)
            ws.cell(row=1, column=col_start, value=date_str_full) # Use full date for column headers
            ws.cell(row=1, column=col_start).alignment = Alignment(horizontal="center", vertical="center")

            # 2. Sub-headers (Row 2)
            ws.cell(row=2, column=col_start, value="UTC Offset (minutes)")
            ws.cell(row=2, column=col_start+1, value="Local Time Stamp")
            ws.cell(row=2, column=col_start+2, value="Active Power (W)")
            ws.cell(row=2, column=col_start+3, value="kW")

            # 3. Handle UTC Offset (minutes) Column Merging (FIXED)
            if data_rows_count > 0:
                # Merge the UTC Offset column for all data rows of this day
                ws.merge_cells(start_row=merge_start_row, 
                                 start_column=col_start, 
                                 end_row=merge_end_row, 
                                 end_column=col_start)
                
                # Set the UTC Offset value to the DATE STRING and center alignment
                utc_cell = ws.cell(row=merge_start_row, column=col_start, value=date_str_full)
                utc_cell.alignment = Alignment(horizontal="center", vertical="center")


            # 4. Fill 10-min rows (Data starts on Row 3)
            for idx, r in enumerate(day_data.itertuples(), start=3):
                # Column 1 (col_start) is now skipped as its value is set in the merged block above.
                
                # Column 2: Local Time Stamp (col_start + 1)
                ws.cell(row=idx, column=col_start+1, value=r.Time) 
                
                # Column 3: Active Power (W) (col_start + 2) - Retains sign (NO ABSOLUTE VALUE).
                power_w = getattr(r, POWER_COL_OUT)
                ws.cell(row=idx, column=col_start+2, value=power_w)
                
                # Column 4: kW (W / 1000) (col_start + 3) - Applies absolute value (ABS).
                ws.cell(row=idx, column=col_start+3, value=abs(power_w) / 1000)

            
            # 5. Add summary statistics (Total, Average, Max)
            if data_rows_count > 0:
                # Calculations
                sum_w = day_data[POWER_COL_OUT].sum()
                mean_w = day_data[POWER_COL_OUT].mean()
                max_w = day_data[POWER_COL_OUT].max()
                
                # kW stats are calculated on the absolute values, as per the kW column's logic.
                sum_kw_abs = day_data[POWER_COL_OUT].abs().sum() / 1000
                mean_kw_abs = day_data[POWER_COL_OUT].abs().mean() / 1000
                max_kw_abs = day_data[POWER_COL_OUT].abs().max() / 1000
                
                # Determine starting row for summaries (1 row below the last data row)
                stats_row_start = merge_end_row + 1
                
                # TOTAL Row
                ws.cell(row=stats_row_start, column=col_start + 1, value="Total")
                ws.cell(row=stats_row_start, column=col_start + 2, value=sum_w)
                ws.cell(row=stats_row_start, column=col_start + 3, value=sum_kw_abs)
                
                # AVERAGE Row
                ws.cell(row=stats_row_start + 1, column=col_start + 1, value="Average")
                ws.cell(row=stats_row_start + 1, column=col_start + 2, value=mean_w)
                ws.cell(row=stats_row_start + 1, column=col_start + 3, value=mean_kw_abs)
                
                # MAX Row
                ws.cell(row=stats_row_start + 2, column=col_start + 1, value="Max")
                ws.cell(row=stats_row_start + 2, column=col_start + 2, value=max_w)
                ws.cell(row=stats_row_start + 2, column=col_start + 3, value=max_kw_abs)
                
                # Update max row used for final summary placement
                max_row_used = max(max_row_used, stats_row_start + 2)
                
                # Collect data for final summary table using short date format
                daily_max_summary.append((date_str_short, max_kw_abs))
                
                # --- Chart Data Tracking (FINAL ROBUST SERIES CREATION) ---
                # The first day's Local Time Stamp column provides the categories (x-axis)
                if chart_categories_ref is None:
                    # Categories start at row 3, column col_start + 1 (Local Time Stamp)
                    chart_categories_ref = Reference(ws, min_col=col_start + 1, min_row=merge_start_row, max_row=merge_end_row)

                # Data for kW column (Y-axis series)
                # Data starts at row 3, column col_start + 3 (kW)
                data_ref = Reference(ws, min_col=col_start + 3, min_row=merge_start_row, max_row=merge_end_row, max_col=col_start + 3)
                
                # Create a Reference for the series title (the merged date header in Row 1)
                title_ref = Reference(ws, min_col=col_start, min_row=1)
                
                # FIX: Instantiate Series with no arguments and assign properties later.
                # This bypasses strict constructor type-checking that was causing the TypeError.
                series_idx = len(chart_data_refs)
                
                series = Series()
                series.idx = series_idx      # Explicitly set the index (should be an integer)
                series.values = data_ref     # Assign data reference
                series.title = title_ref     # Assign title reference
                
                chart_data_refs.append(series)


            col_start += 4
            
        # 7. Add Line Chart for Daily Power Profiles
        if chart_data_refs and chart_categories_ref:
            chart = LineChart()
            chart.style = 10
            chart.title = f"Daily 10-Minute Absolute Power Profile - {sheet_name}"
            chart.y_axis.title = "Power (kW)"
            chart.x_axis.title = "Time"

            # Add all series data
            for series in chart_data_refs:
                chart.series.append(series)
                
            # Set the categories (X-axis labels)
            chart.set_categories(chart_categories_ref)

            # Position the chart below the main data blocks, starting at column G (7)
            # This ensures it doesn't overlap the final summary table starting at A1
            chart_anchor = f'G{max_row_used + 2}'
            ws.add_chart(chart, chart_anchor)
            
            # Update max_row_used to ensure the summary table starts below the chart
            max_row_used = max(max_row_used, max_row_used + 22) # Assume chart takes ~20 rows


        # 6. Add final summary table for Max kW across all days
        if daily_max_summary:
            # Start the summary table 2 rows below the end of the last day block
            final_summary_row = max_row_used + 2 
            
            # --- Summary Table Title (Merged over 2 columns) ---
            title_cell = ws.cell(row=final_summary_row, column=1, value="Daily Max Power (kW) Summary")
            ws.merge_cells(start_row=final_summary_row, start_column=1, end_row=final_summary_row, end_column=2)
            title_cell.alignment = Alignment(horizontal="center", vertical="center")
            title_cell.font = title_font

            # --- Summary Table Headers ---
            final_summary_row += 1
            header_row = final_summary_row
            
            # Header Column 1: Day
            day_header_cell = ws.cell(row=header_row, column=1, value="Day")
            day_header_cell.fill = header_fill
            day_header_cell.font = header_font
            day_header_cell.border = thin_border
            day_header_cell.alignment = Alignment(horizontal="center")
            
            # Header Column 2: Max (kW)
            max_header_cell = ws.cell(row=header_row, column=2, value="Max (kW)")
            max_header_cell.fill = header_fill
            max_header_cell.font = header_font
            max_header_cell.border = thin_border
            max_header_cell.alignment = Alignment(horizontal="center")


            # --- Write data (applying 2dp formatting and color) ---
            for date_str, max_kw in daily_max_summary:
                final_summary_row += 1
                
                # Apply alternating row color
                if (final_summary_row % 2) == 0:
                    fill_style = data_fill_alt
                else:
                    fill_style = PatternFill(fill_type=None) # No fill for odd rows
                
                # Column 1: Day (DD-Mon format)
                day_cell = ws.cell(row=final_summary_row, column=1, value=date_str)
                day_cell.border = thin_border
                day_cell.fill = fill_style
                day_cell.alignment = Alignment(horizontal="center")
                
                # Column 2: Max (kW) - Value rounded to 2dp and explicitly formatted
                max_cell = ws.cell(row=final_summary_row, column=2, value=max_kw)
                max_cell.number_format = numbers.FORMAT_NUMBER_00 # Ensures 2 decimal places (e.g., 0.00)
                max_cell.border = thin_border
                max_cell.fill = fill_style
                max_cell.alignment = Alignment(horizontal="right")
                
            # Auto-size the summary columns for readability
            ws.column_dimensions['A'].width = 15
            ws.column_dimensions['B'].width = 15


    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    return stream

# -----------------------------
# STREAMLIT UI
# -----------------------------
def app():
    st.title("📊 Excel 10-Minute Electricity Data Converter")
    st.markdown("""
        Upload an **Excel file (.xlsx)** with time-series data. Each sheet is processed 
        separately to calculate the total absolute power (W) consumed/generated 
        in fixed **10-minute intervals**. The output Excel file now includes a **line chart** showing the daily kW profiles and a **Max Power Summary table**.
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
                st.error(f"Error reading sheet '{sheet_name}'. {e}")
                continue

            df.columns = df.columns.astype(str).str.strip()

            possible_time_cols = ["Date & Time", "Date&Time", "Timestamp", "DateTime", "Local Time", "TIME", "ts"]
            timestamp_col = next((col for col in df.columns if col in possible_time_cols), None)
            
            if timestamp_col is None:
                st.error(f"No valid **Timestamp** column found in sheet **{sheet_name}**.")
                continue

            possible_psum_cols = ["PSum (W)", "Psum (W)", "PSum", "P (W)", "Power"]
            psum_col = next((col for col in df.columns if col in possible_psum_cols), None)
            
            if psum_col is None:
                st.error(f"No valid **PSum** column found in sheet **{sheet_name}**.")
                continue
            
            # Process the sheet
            try:
                processed = process_sheet(df, timestamp_col, psum_col)
            except Exception as e:
                st.error(f"A critical error occurred while processing sheet **{sheet_name}**'s dates. Error: {e}")
                processed = pd.DataFrame()
            
            if not processed.empty:
                result_sheets[sheet_name] = processed
                st.success(f"Sheet **{sheet_name}** processed successfully and contains **{len(processed['Date'].unique())}** days of data.")
            else:
                st.warning(f"Sheet **{sheet_name}** contained no usable data after cleaning or the date range was invalid.")


        if result_sheets:
            output_stream = build_output_excel(result_sheets)
            st.success(f"Conversion complete for {len(result_sheets)} sheet(s).")
            st.download_button(
                label="📥 Download Converted Excel",
                data=output_stream,
                file_name="Converted_Output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        elif uploaded and not result_sheets:
            st.error("No sheets were successfully processed. Please check the input file for correct column names and data.")

if __name__ == '__main__':
    app()
