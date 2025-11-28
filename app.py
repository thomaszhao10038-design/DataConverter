import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
# Import PatternFill, Font, numbers, and chart components for enhanced styling and graphing
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side, numbers 
from openpyxl.chart import LineChart, Reference

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
    
    # REMOVED: df[psum_col] = df[psum_col].abs() to preserve the sign (positive/negative)
    
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
    """Builds the final Excel workbook with the wide, merged column format and adds a chart."""
    wb = Workbook()
    if 'Sheet' in wb.sheetnames:
         wb.remove(wb['Sheet'])

    # Define styles for the final summary table
    header_fill = PatternFill(start_color='ADD8E6', end_color='ADD8E6', fill_type='solid') # Light Blue
    title_font = Font(bold=True, size=12)
    header_font = Font(bold=True)
    thin_border = Border(left=Side(style='thin'), 
                         right=Side(style='thin'), 
                         top=Side(style='thin'), 
                         bottom=Side(style='thin'))
    data_fill_alt = PatternFill(start_color='F0F8FF', end_color='F0F8FF', fill_type='solid')

    for sheet_name, df in sheets_dict.items():
        ws = wb.create_sheet(sheet_name)
        dates = sorted(df["Date"].unique())

        col_start = 1
        daily_max_summary = [] # List to store max kW for the final summary table
        max_row_used = 0 # Track the lowest row written to across all columns
        max_intervals = 0 # Track the maximum number of 10-minute intervals in any day
        chart_series_data = [] # Stores (kW data column index, kW label row index)

        for date in dates:
            # 1. Update date format to DD-Mon (e.g., 12-Nov) for the summary table
            date_str_short = date.strftime('%d-%b') 
            
            # Use original date string for main table header
            date_str_full = date.strftime('%Y-%m-%d')
            
            day_data = df[df["Date"] == date].sort_values("Time")
            data_rows_count = len(day_data)
            merge_start_row = 3
            merge_end_row = 2 + data_rows_count
            
            # Update max intervals for chart reference sizing
            max_intervals = max(max_intervals, data_rows_count) 
            
            # 1. Merge date header (Row 1, columns 1 to 4)
            ws.merge_cells(start_row=1, start_column=col_start, end_row=1, end_column=col_start+3)
            ws.cell(row=1, column=col_start, value=date_str_full) # Use full date for column headers
            ws.cell(row=1, column=col_start).alignment = Alignment(horizontal="center", vertical="center")

            # 2. Sub-headers (Row 2)
            ws.cell(row=2, column=col_start, value="UTC Offset (minutes)")
            ws.cell(row=2, column=col_start+1, value="Local Time Stamp")
            ws.cell(row=2, column=col_start+2, value="Active Power (W)")
            ws.cell(row=2, column=col_start+3, value="kW")

            # 3. Handle UTC Offset (minutes) Column Merging
            if data_rows_count > 0:
                ws.merge_cells(start_row=merge_start_row, 
                               start_column=col_start, 
                               end_row=merge_end_row, 
                               end_column=col_start)
                
                utc_cell = ws.cell(row=merge_start_row, column=col_start, value=date_str_full)
                utc_cell.alignment = Alignment(horizontal="center", vertical="center")

            # 4. Fill 10-min rows (Data starts on Row 3)
            for idx, r in enumerate(day_data.itertuples(), start=3):
                # Column 2: Local Time Stamp (col_start + 1)
                ws.cell(row=idx, column=col_start+1, value=r.Time) 
                
                # Column 3: Active Power (W) (col_start + 2) - Retains sign
                power_w = getattr(r, POWER_COL_OUT)
                ws.cell(row=idx, column=col_start+2, value=power_w)
                
                # Column 4: kW (W / 1000) (col_start + 3) - Applies absolute value (ABS)
                ws.cell(row=idx, column=col_start+3, value=abs(power_w) / 1000)

            
            # 5. Add summary statistics (Total, Average, Max)
            if data_rows_count > 0:
                # Calculations
                sum_w = day_data[POWER_COL_OUT].sum()
                mean_w = day_data[POWER_COL_OUT].mean()
                max_w = day_data[POWER_COL_OUT].max()
                
                # kW stats are calculated on the absolute values.
                sum_kw_abs = day_data[POWER_COL_OUT].abs().sum() / 1000
                mean_kw_abs = day_data[POWER_COL_OUT].abs().mean() / 1000
                max_kw_abs = day_data[POWER_COL_OUT].abs().max() / 1000
                
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
                
                max_row_used = max(max_row_used, stats_row_start + 2)
                
                # Collect data for final summary table
                daily_max_summary.append((date_str_short, max_kw_abs))
                
                # Store data needed for chart: (kW column, date header column)
                chart_series_data.append((col_start + 3, col_start))

            col_start += 4
            
        # Determine the starting row for the summary and chart area
        final_summary_row = max_row_used + 2 
            
        # 6. Add final summary table for Max kW across all days (Col 1-2)
        if daily_max_summary:
            
            # Summary Table Title (Merged over 2 columns)
            title_cell = ws.cell(row=final_summary_row, column=1, value="Daily Max Power (kW) Summary")
            ws.merge_cells(start_row=final_summary_row, start_column=1, end_row=final_summary_row, end_column=2)
            title_cell.alignment = Alignment(horizontal="center", vertical="center")
            title_cell.font = title_font

            # Summary Table Headers
            header_row = final_summary_row + 1
            
            day_header_cell = ws.cell(row=header_row, column=1, value="Day")
            day_header_cell.fill = header_fill
            day_header_cell.font = header_font
            day_header_cell.border = thin_border
            day_header_cell.alignment = Alignment(horizontal="center")
            
            max_header_cell = ws.cell(row=header_row, column=2, value="Max (kW)")
            max_header_cell.fill = header_fill
            max_header_cell.font = header_font
            max_header_cell.border = thin_border
            max_header_cell.alignment = Alignment(horizontal="center")


            # Write data (applying 2dp formatting and color)
            current_row = header_row
            for date_str, max_kw in daily_max_summary:
                current_row += 1
                
                if (current_row % 2) == 0:
                    fill_style = data_fill_alt
                else:
                    fill_style = PatternFill(fill_type=None) 
                
                # Column 1: Day (DD-Mon format)
                day_cell = ws.cell(row=current_row, column=1, value=date_str)
                day_cell.border = thin_border
                day_cell.fill = fill_style
                day_cell.alignment = Alignment(horizontal="center")
                
                # Column 2: Max (kW) - Value rounded to 2dp and explicitly formatted
                max_cell = ws.cell(row=current_row, column=2, value=max_kw)
                max_cell.number_format = numbers.FORMAT_NUMBER_00 # Ensures 2 decimal places (e.g., 0.00)
                max_cell.border = thin_border
                max_cell.fill = fill_style
                max_cell.alignment = Alignment(horizontal="right")
        
        # 7. Add Chart (Anchored next to the summary table)
        if chart_series_data and max_intervals > 0:
            
            chart = LineChart()
            chart.title = "Daily Absolute Power Profile (kW)"
            chart.style = 10 
            
            # Set X-axis (Categories/Time)
            # Time column is column 2 (B) of the first day's block, rows 3 up to 2 + max_intervals
            time_categories = Reference(ws, min_col=2, min_row=3, 
                                        max_col=2, max_row=2 + max_intervals)
            chart.set_categories(time_categories)
            
            # Add series one by one for each day
            for kw_col_idx, header_col_idx in chart_series_data:
                # Values reference (data starts at row 3)
                values = Reference(ws, min_col=kw_col_idx, min_row=3, 
                                   max_col=kw_col_idx, max_row=2 + max_intervals)
                
                # Name reference (The date label is in Row 1, merged cell at the header_col_idx)
                title = Reference(ws, min_col=header_col_idx, min_row=1, 
                                  max_col=header_col_idx, max_row=1)
                
                # Add series, titles_from_data is needed to use the date string as the series name
                chart.add_series(values, title_from_data=title)

            # Set Axis Titles
            chart.x_axis.title = "10-Minute Interval"
            chart.y_axis.title = "Absolute Power (kW)"
            
            # Anchor the chart (starting at Column D/4, right of the summary table in Col 1-2)
            # The anchor row is the same as the summary table title
            ws.add_chart(chart, f'D{final_summary_row}')
            
            chart.width = 18 # Increase width for better legibility
            chart.height = 12 
                

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
        in fixed **10-minute intervals**.
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
