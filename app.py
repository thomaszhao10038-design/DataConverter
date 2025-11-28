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
    # The snippet shows DD/MM/YYYY, so we use dayfirst=True.
    df[timestamp_col] = pd.to_datetime(df[timestamp_col], errors="coerce", dayfirst=True)
    
    # Aggressively clean and ensure Power column is numeric
    power_series = df[psum_col].astype(str).str.strip()
    
    # Handle the comma decimal separator replacement before conversion (Robust for various locales)
    power_series = power_series.str.replace(',', '.', regex=False)
    
    # Attempt conversion, coercing errors to NaN
    df[psum_col] = pd.to_numeric(power_series, errors='coerce')
    
    # Drop rows where essential data is missing/invalid
    initial_rows = len(df)
    
    # --- DIAGNOSTIC CHECK ---
    timestamp_nans = df[timestamp_col].isna().sum()
    power_nans = df[psum_col].isna().sum()
    # ------------------------

    df = df.dropna(subset=[timestamp_col, psum_col])
    valid_rows = len(df)
    
    if df.empty:
        # st.error(f"Sheet contained no valid data after cleaning (0 out of {initial_rows} rows kept).")
        # st.error(f"**Diagnostic:** {timestamp_nans} rows failed Timestamp conversion. {power_nans} rows failed Power conversion.")
        return pd.DataFrame()

    # st.info(f"Using {valid_rows} data points for aggregation (from initial {initial_rows} rows).")
    
    # 2. Aggregate Data (Fixed Logic)
    
    # Use absolute value to correctly sum total power magnitude
    df[psum_col] = df[psum_col].abs()
    
    # Set the timestamp as index
    df_indexed = df.set_index(timestamp_col)
    
    # --- FIX: Floor the index to the nearest 10 minutes and use groupby().sum() ---
    # This correctly assigns non-uniform time stamps (e.g., 12:12:13) to the start of their 10-min interval (12:10:00)
    df_indexed.index = df_indexed.index.floor('10min')
    
    # Group by the new, rounded index and sum the power values for the 10-min interval
    resampled_data = df_indexed[psum_col].groupby(level=0).sum()
    # --- END FIX ---
    
    # Convert the aggregated Series back to a DataFrame
    df_out = resampled_data.reset_index()
    # Rename the PSum column to a simple, guaranteed name
    df_out.columns = ['Rounded', POWER_COL_OUT] 
    
    # --- STABILIZING CHECK FOR DATE RANGE ---
    if df_out['Rounded'].empty or df_out['Rounded'].isna().all():
        # st.error("Aggregation produced no valid timestamps, likely due to a time range error. Cannot pad data.")
        return pd.DataFrame()

    # Store the set of original valid dates to filter the final padded range
    original_dates = set(df_out['Rounded'].dt.date)

    # 3. Robust Padding (Ensuring all 10-min slots for all days are present)
    
    min_date = df_out['Rounded'].min().normalize()
    max_date = df_out['Rounded'].max().normalize()
    
    # Create a continuous 10-minute index covering the entire data range
    full_time_index = pd.date_range(
        start=min_date, 
        # End at the start of the last 10-minute interval on the max_date (23:50:00)
        # We go up to 00:00 of the next day, and use closed='left'
        end=max_date + pd.Timedelta(days=1),
        freq='10min',
        closed='left' 
    )

    # Re-index the resampled data onto the full time index, filling missing intervals with 0
    df_indexed_for_reindex = df_out.set_index('Rounded')
    df_padded_series = df_indexed_for_reindex[POWER_COL_OUT].reindex(full_time_index, fill_value=0)
    
    # Convert back to DataFrame and clean up
    grouped = df_padded_series.reset_index().rename(columns={'index': 'Rounded'})
    grouped.columns = ['Rounded', POWER_COL_OUT]

    # 4. Final Formatting
    
    # Extract date and time columns from the final padded (and now complete) time series
    grouped["Date"] = grouped["Rounded"].dt.date
    # Format time as HH:MM to match the Excel output requirement "Local Time Stamp"
    grouped["Time"] = grouped["Rounded"].dt.strftime("%H:%M") 

    # Filter the result to only include dates that were present in the original data.
    grouped = grouped[grouped["Date"].isin(original_dates)]
    
    # Final check: ensure the data has all days.
    # st.info(f"Total unique days found in output data: {len(grouped['Date'].unique())}")
    
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
            
            # Double-check: ensure we have 144 rows for a full day (144 = 24 * 6)
            # The padding logic ensures this unless the original data range was less than a day
            
            for idx, r in enumerate(day_data.itertuples(), start=3):
                # Column 1: UTC Offset (minutes) - Set to 0 as placeholder since no offset is provided in source
                ws.cell(row=idx, column=col_start, value=0) 
                
                # Column 2: Local Time Stamp
                ws.cell(row=idx, column=col_start+1, value=r.Time) 
                
                # Column 3: Active Power (W) - The aggregated sum
                # Access the column by its guaranteed name (PSumW)
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
            # Using str.strip() here on the columns inside the loop is redundant since we did it before the loop.
            timestamp_col = next((col for col in df.columns if col in possible_time_cols), None)
            
            if timestamp_col is None:
                st.error(f"No valid timestamp column found in sheet **{sheet_name}**. (Tried: {', '.join(possible_time_cols)})")
                continue

            # Auto-detect PSum column
            possible_psum_cols = ["PSum (W)", "Psum (W)", "PSum", "P (W)", "Power"]
            # Using str.strip() here on the columns inside the loop is redundant since we did it before the loop.
            psum_col = next((col for col in df.columns if col in possible_psum_cols), None)
            
            if psum_col is None:
                st.error(f"No valid PSum column found in sheet **{sheet_name}**. (Tried: {', '.join(possible_psum_cols)})")
                continue

            processed = process_sheet(df, timestamp_col, psum_col)
            
            if not processed.empty:
                result_sheets[sheet_name] = processed
                st.success(f"Sheet **{sheet_name}** processed successfully and contains {len(processed['Date'].unique())} days of data.")
            else:
                st.warning(f"Sheet **{sheet_name}** contained no usable data after cleaning.")


        if result_sheets:
            output_stream = build_output_excel(result_sheets)
            st.success("All valid sheets converted to wide format.")
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
