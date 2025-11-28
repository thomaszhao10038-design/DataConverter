import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment 

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
    """Builds the final Excel workbook with the wide, merged column format."""
    wb = Workbook()
    if 'Sheet' in wb.sheetnames:
         wb.remove(wb['Sheet'])

    for sheet_name, df in sheets_dict.items():
        ws = wb.create_sheet(sheet_name)
        dates = sorted(df["Date"].unique())

        col_start = 1
        for date in dates:
            date_str = date.strftime('%Y-%m-%d')
            
            day_data = df[df["Date"] == date].sort_values("Time")
            data_rows_count = len(day_data)
            merge_start_row = 3
            merge_end_row = 2 + data_rows_count
            
            # 1. Merge date header (Row 1, columns 1 to 4)
            ws.merge_cells(start_row=1, start_column=col_start, end_row=1, end_column=col_start+3)
            ws.cell(row=1, column=col_start, value=date_str)
            ws.cell(row=1, column=col_start).alignment = Alignment(horizontal="center", vertical="center")

            # 2. Sub-headers (Row 2)
            # NOTE: The header remains "UTC Offset (minutes)" as requested by the output format,
            # but the content below it will show the date string.
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
                utc_cell = ws.cell(row=merge_start_row, column=col_start, value=date_str)
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


            col_start += 4

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
