import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side, numbers
from openpyxl.chart import LineChart, Reference
# CRUCIAL IMPORT: Explicitly import Series for robust chart construction
from openpyxl.chart.series import Series 

# --- Configuration ---
POWER_COL_OUT = 'PSumW'

# Output Column mapping (relative to col_start, 1-based index in the 4-column block)
COL_UTC_REL = 1    # UTC Offset (merged column)
COL_TIME_REL = 2   # Local Time Stamp (X-Axis Categories)
COL_W_REL = 3      # Active Power (W)
COL_KW_REL = 4     # kW (Y-Axis Data)
COL_BLOCK_WIDTH = 4

# -----------------------------
# PROCESS SINGLE SHEET (TIDY DATAFRAME GENERATION)
# -----------------------------
def process_sheet(df, timestamp_col, psum_col):
    """
    Processes a single sheet: cleans data, handles dates, aggregates to 10-min sums,
    pads the result, and calculates the absolute kW column.
    """
    # 1. Cleaning and Preparation
    df.columns = df.columns.astype(str).str.strip()
    # Use dayfirst=True to handle common European formats (dd/mm/yyyy)
    df[timestamp_col] = pd.to_datetime(df[timestamp_col], errors="coerce", dayfirst=True)
    
    # Handle commas as decimal separators before converting to numeric
    power_series = df[psum_col].astype(str).str.strip()
    power_series = power_series.str.replace(',', '.', regex=False)
    df[psum_col] = pd.to_numeric(power_series, errors='coerce')
    
    df = df.dropna(subset=[timestamp_col, psum_col])
    
    if df.empty:
        return pd.DataFrame()
    
    # 2. Aggregate Data and Resampling
    df_indexed = df.set_index(timestamp_col)
    df_indexed.index = df_indexed.index.floor('10min')
    resampled_data = df_indexed[psum_col].groupby(level=0).sum()
    df_out = resampled_data.reset_index()
    df_out.columns = ['Rounded', POWER_COL_OUT] 
    
    if df_out.empty or df_out['Rounded'].isna().all():
        return pd.DataFrame()
    
    # 3. Robust Padding
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

    df_indexed_for_reindex = df_out.set_index('Rounded')
    # Reindex and fill missing 10-min intervals with 0
    df_padded_series = df_indexed_for_reindex[POWER_COL_OUT].reindex(full_time_index, fill_value=0)
    
    grouped = df_padded_series.reset_index().rename(columns={'index': 'Rounded'})
    grouped.columns = ['Rounded', POWER_COL_OUT]
    
    # 4. Final Formatting
    grouped["Date"] = grouped["Rounded"].dt.date
    grouped["Time"] = grouped["Rounded"].dt.strftime("%H:%M") 

    # Filter to only include dates that were present in the original data (removes the padding before min/after max date)
    grouped = grouped[grouped["Date"].isin(original_dates)]
    
    # Calculate absolute kW (W / 1000)
    grouped['kW'] = grouped[POWER_COL_OUT].abs() / 1000
    
    return grouped

# -----------------------------
# BUILD EXCEL FORMAT
# -----------------------------
def build_output_excel(sheets_dict):
    """Builds the final Excel workbook with the wide, merged column format and includes a daily max kW chart."""
    wb = Workbook()
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])

    # Define styles
    header_fill = PatternFill(start_color='ADD8E6', end_color='ADD8E6', fill_type='solid') # Light Blue
    title_font = Font(bold=True, size=12)
    thin_border = Border(left=Side(style='thin'), 
                          right=Side(style='thin'), 
                          top=Side(style='thin'), 
                          bottom=Side(style='thin'))
    data_fill_alt = PatternFill(start_color='F0F8FF', end_color='F0F8FF', fill_type='solid')

    for sheet_name, df in sheets_dict.items():
        ws = wb.create_sheet(sheet_name)
        dates = sorted(df["Date"].unique())

        col_start = 1
        daily_max_summary = []
        max_row_used = 0
        
        chart_categories_ref = None 
        chart_series_list = []


        for date in dates:
            date_str_short = date.strftime('%d-%b') 
            date_str_full = date.strftime('%Y-%m-%d')
            
            day_data = df[df["Date"] == date].sort_values("Time")
            data_rows_count = len(day_data)
            
            # Data starts on Row 3
            merge_start_row = 3
            merge_end_row = 2 + data_rows_count
            
            # --- 1. Merge date header (Row 1) ---
            ws.merge_cells(start_row=1, start_column=col_start, end_row=1, end_column=col_start + COL_BLOCK_WIDTH - 1)
            # This cell holds the date string used for the chart legend
            ws.cell(row=1, column=col_start, value=date_str_full).alignment = Alignment(horizontal="center", vertical="center")

            # --- 2. Sub-headers (Row 2) ---
            ws.cell(row=2, column=col_start + COL_UTC_REL - 1, value="UTC Offset (minutes)")
            ws.cell(row=2, column=col_start + COL_TIME_REL - 1, value="Local Time Stamp")
            ws.cell(row=2, column=col_start + COL_W_REL - 1, value="Active Power (W)")
            ws.cell(row=2, column=col_start + COL_KW_REL - 1, value="kW")

            # --- 3. UTC Offset Column Merging (Col 1) ---
            if data_rows_count > 0:
                utc_col = col_start + COL_UTC_REL - 1
                ws.merge_cells(start_row=merge_start_row, start_column=utc_col, end_row=merge_end_row, end_column=utc_col)
                utc_cell = ws.cell(row=merge_start_row, column=utc_col, value=date_str_full)
                utc_cell.alignment = Alignment(horizontal="center", vertical="center")


            # --- 4. Fill 10-min rows (Data starts on Row 3) ---
            for idx, r in enumerate(day_data.itertuples(), start=merge_start_row):
                
                # Column 2: Local Time Stamp
                ws.cell(row=idx, column=col_start + COL_TIME_REL - 1, value=r.Time) 
                
                # Column 3: Active Power (W)
                ws.cell(row=idx, column=col_start + COL_W_REL - 1, value=getattr(r, POWER_COL_OUT))
                
                # Column 4: kW (W / 1000)
                ws.cell(row=idx, column=col_start + COL_KW_REL - 1, value=r.kW)

            
            # --- 5. Add summary statistics (Total, Average, Max) ---
            if data_rows_count > 0:
                # Calculations
                sum_w = day_data[POWER_COL_OUT].sum()
                mean_w = day_data[POWER_COL_OUT].mean()
                max_w = day_data[POWER_COL_OUT].max()
                
                sum_kw = day_data['kW'].sum()
                mean_kw = day_data['kW'].mean()
                max_kw = day_data['kW'].max() 
                
                stats_row_start = merge_end_row + 1
                
                # TOTAL Row
                ws.cell(row=stats_row_start, column=col_start + COL_TIME_REL - 1, value="Total")
                ws.cell(row=stats_row_start, column=col_start + COL_W_REL - 1, value=sum_w).number_format = numbers.FORMAT_NUMBER
                ws.cell(row=stats_row_start, column=col_start + COL_KW_REL - 1, value=sum_kw).number_format = numbers.FORMAT_NUMBER_00
                
                # AVERAGE Row
                ws.cell(row=stats_row_start + 1, column=col_start + COL_TIME_REL - 1, value="Average")
                ws.cell(row=stats_row_start + 1, column=col_start + COL_W_REL - 1, value=mean_w).number_format = numbers.FORMAT_NUMBER
                ws.cell(row=stats_row_start + 1, column=col_start + COL_KW_REL - 1, value=mean_kw).number_format = numbers.FORMAT_NUMBER_00
                
                # MAX Row
                ws.cell(row=stats_row_start + 2, column=col_start + COL_TIME_REL - 1, value="Max")
                ws.cell(row=stats_row_start + 2, column=col_start + COL_W_REL - 1, value=max_w).number_format = numbers.FORMAT_NUMBER
                ws.cell(row=stats_row_start + 2, column=col_start + COL_KW_REL - 1, value=max_kw).number_format = numbers.FORMAT_NUMBER_00
                
                max_row_used = max(max_row_used, stats_row_start + 2)
                
                daily_max_summary.append((date_str_short, max_kw))
                
                # --- 6. Chart Data Preparation (Explicit Series Constructor) ---
                
                # Categories (X-axis) reference (Local Time Stamp column - Col 2)
                if chart_categories_ref is None:
                    cat_col = col_start + COL_TIME_REL - 1
                    chart_categories_ref = Reference(ws, min_col=cat_col, min_row=merge_start_row, max_row=merge_end_row)

                # Data (Y-axis) reference (kW column - Col 4)
                data_col = col_start + COL_KW_REL - 1
                data_ref = Reference(ws, min_col=data_col, min_row=merge_start_row, max_row=merge_end_row, max_col=data_col)
                
                # Title reference (The merged cell at row 1, col_start contains the date string)
                title_ref = Reference(ws, min_col=col_start, min_row=1, max_col=col_start, max_row=1)
                
                # Use Series constructor with values and title reference (FIX for stability)
                series = Series(values=data_ref, title=title_ref)
                
                chart_series_list.append(series)

            col_start += COL_BLOCK_WIDTH
            
        # --- 7. Add Line Chart for Daily Power Profiles ---
        if chart_series_list and chart_categories_ref:
            chart = LineChart()
            chart.style = 10
            
            chart.title = f"{sheet_name} - power consumption (kW)"
            
            # X-axis (Time stamps)
            chart.x_axis.title = "Time" 
            # Y-axis (kW values)
            chart.y_axis.title = "Power (kW)" 

            chart.set_categories(chart_categories_ref)
            for series in chart_series_list:
                chart.series.append(series)
            
            # Set tickLblSkip to 1 for 20-minute interval labels (since data is 10-min)
            chart.x_axis.tickLblSkip = 1 

            chart_anchor = f'G{max_row_used + 2}'
            ws.add_chart(chart, chart_anchor)
            
            # Ensure max_row_used is updated to place the summary table below the chart
            max_row_used = max(max_row_used, max_row_used + 22)


        # --- 8. Add final summary table for Max kW across all days ---
        if daily_max_summary:
            final_summary_row = max_row_used + 2 
            
            # Title
            title_cell = ws.cell(row=final_summary_row, column=1, value="Daily Max Power (kW) Summary")
            ws.merge_cells(start_row=final_summary_row, start_column=1, end_row=final_summary_row, end_column=2)
            title_cell.alignment = Alignment(horizontal="center", vertical="center")
            title_cell.font = title_font

            # Headers
            final_summary_row += 1
            day_header_cell = ws.cell(row=final_summary_row, column=1, value="Day")
            day_header_cell.fill = header_fill
            day_header_cell.border = thin_border
            max_header_cell = ws.cell(row=final_summary_row, column=2, value="Max (kW)")
            max_header_cell.fill = header_fill
            max_header_cell.border = thin_border

            # Data
            for date_idx, (date_str, max_kw) in enumerate(daily_max_summary):
                row = final_summary_row + 1 + date_idx
                
                # Apply alternating row color
                fill_style = data_fill_alt if (row % 2) == 0 else PatternFill(fill_type=None)
                
                # Column 1: Day (DD-Mon format)
                day_cell = ws.cell(row=row, column=1, value=date_str)
                day_cell.border = thin_border
                day_cell.fill = fill_style
                day_cell.alignment = Alignment(horizontal="center")
                
                # Column 2: Max (kW) - Value rounded to 2dp
                max_cell = ws.cell(row=row, column=2, value=max_kw)
                max_cell.number_format = numbers.FORMAT_NUMBER_00
                max_cell.border = thin_border
                max_cell.fill = fill_style
                max_cell.alignment = Alignment(horizontal="right")
                
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
        in fixed **10-minute intervals**. The output Excel file includes a robust 
        **line chart** showing the daily kW profiles and a **Max Power Summary table**.
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
