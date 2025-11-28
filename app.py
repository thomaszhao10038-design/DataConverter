import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment 

# --- Configuration ---
# Define the output header columns that repeat for each day (4 columns total)
OUTPUT_HEADERS = [
    'UTC Offset (minutes)', 
    'Local Time Stamp', 
    'Active Power (W)', 
    'kW'
]
# Define a robust internal column name for PSum (W) aggregation
POWER_COL_OUT = 'PSumW'

# -----------------------------
# PROCESS SINGLE SHEET (IMPROVED PADDING & AGGREGATION)
# -----------------------------
def process_sheet(df, timestamp_col, psum_col):
    """
    Processes a single sheet by rounding timestamps to 10-minute intervals, 
    summing absolute power values in each interval, and ensuring all 10-minute 
    intervals for the entire time range are present using Pandas reindexing.
    """
    # 1. Cleaning and Preparation
    
    # Ensure columns are stripped of leading/trailing spaces for reliable access
    df.columns = df.columns.astype(str).str.strip()

    # Convert timestamp. dayfirst=True handles DD/MM/YYYY format.
    df[timestamp_col] = pd.to_datetime(df[timestamp_col], errors="coerce", dayfirst=True)
    
    # Aggressively clean and ensure Power column is numeric
    power_series = df[psum_col].astype(str).str.strip()
    
    # Handle the comma decimal separator replacement before conversion (Robust for various locales)
    power_series = power_series.str.replace(',', '.', regex=False)
    
    # Attempt conversion, coercing errors to NaN
    df[psum_col] = pd.to_numeric(power_series, errors='coerce')
    
    # Drop rows where essential data is missing/invalid
    initial_rows = len(df)
    
    df = df.dropna(subset=[timestamp_col, psum_col])
    valid_rows = len(df)
    
    if df.empty:
        return pd.DataFrame()
    
    # 2. Aggregate Data (Fixed Logic)
    
    # Use absolute value to correctly sum total power magnitude
    df[psum_col] = df[psum_col].abs()
    
    # Set the timestamp as index
    df_indexed = df.set_index(timestamp_col)
    
    # FIX: Floor the index to the nearest 10 minutes (e.g., 12:12:13 -> 12:10:00)
    # This correctly groups all non-uniform time stamps into the start of the 10-min interval.
    df_indexed.index = df_indexed.index.floor('10min')
    
    # Group by the new, rounded index and sum the power values for the 10-min interval
    resampled_data = df_indexed[psum_col].groupby(level=0).sum()
    
    # Convert the aggregated Series back to a DataFrame
    df_out = resampled_data.reset_index()
    # Rename the PSum column to a simple, guaranteed name
    df_out.columns = ['Rounded', POWER_COL_OUT] 
    
    # 3. Robust Padding (Ensuring all 10-min slots for all days are present)
    
    if df_out.empty or df_out['Rounded'].isna().all():
        return pd.DataFrame()
    
    # FIX: Use floor('D') and ceil('D') for robust date range calculation
    # Start of the first day (e.g., 2025-11-11 00:00:00)
    min_dt = df_out['Rounded'].min().floor('D')
    # Start of the day *after* the last day (e.g., 2025-11-27 00:00:00)
    max_dt_exclusive = df_out['Rounded'].max().ceil('D') 
    
    # Create a continuous 10-minute index covering the entire data range
    full_time_index = pd.date_range(
        start=min_dt, 
        end=max_dt_exclusive,
        freq='10min',
        closed='left' # Ensures the end time (e.g., 00:00 of the next day) is excluded
    )

    # Re-index the resampled data onto the full time index, filling missing intervals with 0
    df_indexed_for_reindex = df_out.set_index('Rounded')
    df_padded_series = df_indexed_for_reindex[POWER_COL_OUT].reindex(full_time_index, fill_value=0)
    
    # Convert back to DataFrame and clean up
    grouped = df_padded_series.reset_index().rename(columns={'index': 'Rounded'})
    grouped.columns = ['Rounded', POWER_COL_OUT]
    
    # 4. Final Formatting
    
    # Store the set of original valid dates to filter the final padded range
    original_dates = set(df_out['Rounded'].dt.date)

    # Extract date and time columns from the final padded (and now complete) time series
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
    # Remove the default sheet created by openpyxl
    if 'Sheet' in wb.sheetnames:
         wb.remove(wb['Sheet'])

    for sheet_name, df in sheets_dict.items():
        ws = wb.create_sheet(sheet_name)
        dates = sorted(df["Date"].unique())

        col_start = 1
        for date in dates:
            # Use 'YYYY-MM-DD' format for consistency
            date_str = date.strftime('%Y-%m-%d')
            
            # 1. Merge date header (Row 1, columns 1 to 4)
            ws.merge_cells(start_row=1, start_column=col_start, end_row=1, end_column=col_start+3)
            ws.cell(row=1, column=col_start, value=date_str)
            ws.cell(row=1, column=col_start).alignment = Alignment(horizontal="center", vertical="center")

            # 2. Sub-headers (Row 2)
            ws.cell(row=2, column=col_start, value="UTC Offset (minutes)")
            ws.cell(row=2, column=col_start+1, value="Local Time Stamp")
            ws.cell(row=2, column=col_start+2, value="Active Power (W)")
            ws.cell(row=2, column=col_start+3, value="kW")

            # 3. Fill 10-min rows (Data starts on Row 3)
            # Filter and ensure data is sorted by time for correct order
            day_data = df[df["Date"] == date].sort_values("Time")
            
            for idx, r in enumerate(day_data.itertuples(), start=3):
                # Column 1: UTC Offset (minutes) - Set to 0 as placeholder
                ws.cell(row=idx, column=col_start, value=0) 
                
                # Column 2: Local Time Stamp
                ws.cell(row=idx, column=col_start+1, value=r.Time) 
                
                # Column 3: Active Power (W) - The aggregated sum
                power_w = getattr(r, POWER_COL_OUT)
                ws.cell(row=idx, column=col_start+2, value=power_w)
                
                # Column 4: kW (W / 1000)
                ws.cell(row=idx, column=col_start+3, value=power_w / 1000)

            col_start += 4  # Move to the start of the next day block

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
        Upload an Excel file (.xlsx) with time-series data. Each sheet is processed 
        separately to calculate the total absolute power (W) consumed/generated 
        in fixed 10-minute intervals. The output is a wide format file suitable for analysis.
        """)

    uploaded = st.file_uploader("Upload .xlsx file", type=["xlsx"])

    if uploaded:
        xls = pd.ExcelFile(uploaded)
        result_sheets = {}

        for sheet_name in xls.sheet_names:
            st.info(f"Preparing to process sheet: **{sheet_name}**")
            try:
                # Use Pandas to read the sheet
                df = pd.read_excel(uploaded, sheet_name=sheet_name)
            except Exception as e:
                st.error(f"Error reading sheet '{sheet_name}'. {e}")
                continue

            # Clean column names for robust matching
            df.columns = df.columns.astype(str).str.strip()

            # Auto-detect timestamp column
            possible_time_cols = ["Date & Time", "Date&Time", "Timestamp", "DateTime", "Local Time", "TIME", "ts"]
            timestamp_col = next((col for col in df.columns if col in possible_time_cols), None)
            
            if timestamp_col is None:
                st.error(f"No valid timestamp column found in sheet **{sheet_name}**. (Tried: {', '.join(possible_time_cols)})")
                continue

            # Auto-detect PSum column
            possible_psum_cols = ["PSum (W)", "Psum (W)", "PSum", "P (W)", "Power"]
            psum_col = next((col for col in df.columns if col in possible_psum_cols), None)
            
            if psum_col is None:
                st.error(f"No valid PSum column found in sheet **{sheet_name}**. (Tried: {', '.join(possible_psum_cols)})")
                continue

            processed = process_sheet(df, timestamp_col, psum_col)
            
            if not processed.empty:
                result_sheets[sheet_name] = processed
                # st.success(f"Sheet **{sheet_name}** processed successfully and contains {len(processed['Date'].unique())} days of data.")
            else:
                st.warning(f"Sheet **{sheet_name}** contained no usable data after cleaning.")


        if result_sheets:
            output_stream = build_output_excel(result_sheets)
            st.success(f"All valid sheets ({len(result_sheets)}) converted to wide format.")
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
