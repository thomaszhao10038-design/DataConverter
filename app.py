import streamlit as st
import pandas as pd
from io import BytesIO
# Import openpyxl components for advanced Excel formatting
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
# ROUND TIMESTAMP TO 10 MIN
# -----------------------------
def round_to_10min(ts):
    """
    Rounds a timestamp down to the nearest 10-minute interval (e.g., 12:12:01 -> 12:10:00).
    This function is now part of the Pandas resampling process, but is kept for clarity 
    in the manual data processing step (though Pandas handles this automatically better).
    For our corrected aggregation, we will use a more standard method within Pandas.
    """
    if pd.isna(ts) or ts is None:
        return pd.NaT
    
    # We round down (floor) to the start of the 10-minute interval
    ts = pd.to_datetime(ts)
    start_of_day = ts.normalize()
    # Calculate total minutes since start of day
    minutes_since_midnight = (ts - start_of_day).total_seconds() // 60
    # Determine the floor to the nearest 10 minutes
    rounded_minutes = (minutes_since_midnight // 10) * 10
    
    return start_of_day + pd.Timedelta(minutes=rounded_minutes)

# -----------------------------
# PROCESS SINGLE SHEET (CORRECTED AGGREGATION)
# -----------------------------
def process_sheet(df, timestamp_col, psum_col):
    """
    Processes a single sheet by rounding timestamps to 10-minute intervals, 
    summing absolute power values in each interval, and ensuring all 144 
    intervals for every day are present.
    """
    # 1. Cleaning and Preparation (Improved from previous versions)
    
    # Ensure columns are stripped of leading/trailing spaces for reliable access
    df.columns = df.columns.astype(str).str.strip()

    # Convert timestamp and drop invalid rows
    df[timestamp_col] = pd.to_datetime(df[timestamp_col], errors="coerce", dayfirst=True)
    
    # Aggressively clean and ensure Power column is numeric
    power_series = df[psum_col].astype(str).str.strip().str.replace(',', '.', regex=False)
    df[psum_col] = pd.to_numeric(power_series, errors='coerce')
    
    # Drop rows where essential data is missing/invalid
    df = df.dropna(subset=[timestamp_col, psum_col])
    
    if df.empty:
        return pd.DataFrame()

    # 2. Resample and Aggregate Data
    
    # Use absolute value to correctly sum total power magnitude
    df[psum_col] = df[psum_col].abs()
    
    # Set the timestamp as index
    df_indexed = df.set_index(timestamp_col)
    
    # CRITICAL FIX: Resample the data to a 10-minute frequency, taking the SUM.
    # 'label=left' means the time 12:10:00 covers the period from 12:10:00 up to 12:19:59.
    # 'origin=start' ensures the interval begins exactly at the start of the day (00:00:00).
    resampled_data = df_indexed[psum_col].resample(
        '10min', 
        label='left', 
        origin='start'
    ).sum()
    
    # Convert back to DataFrame
    df_out = resampled_data.reset_index()
    # Rename the PSum column to a simple, guaranteed name for reliable access later
    df_out.columns = ['Rounded', POWER_COL_OUT] 
    
    # 3. Padding and Final Formatting
    
    df_out["Date"] = df_out["Rounded"].dt.date
    df_out["Time"] = df_out["Rounded"].dt.strftime("%H:%M") # Use HH:MM format

    # Create a standardized time index for padding: 00:00 to 23:50
    all_intervals_str = [pd.to_datetime(f'{i:02d}:{j:02d}', format='%H:%M').strftime('%H:%M') 
                         for i in range(24) for j in range(0, 60, 10)]
    
    # Group the aggregated data by day
    final_rows = []
    
    for date in df_out["Date"].unique():
        # Select only the power column by its standardized name
        day_data = df_out[df_out["Date"] == date].set_index("Time")[[POWER_COL_OUT]] 
        
        # Create a day-specific template DataFrame (144 rows)
        template_df = pd.DataFrame(index=all_intervals_str)
        template_df.index.name = "Time"
        
        # Join the actual data onto the template to ensure all 144 intervals are present
        padded_day_data = template_df.join(day_data, how='left')
        
        # Fill NaN (periods with no power readings) with 0, as required
        padded_day_data[POWER_COL_OUT] = padded_day_data[POWER_COL_OUT].fillna(0)

        # Prepare final output structure
        padded_day_data['Date'] = date
        padded_day_data = padded_day_data.reset_index().rename(columns={'index': 'Time'})
        
        # Append to the final list
        final_rows.append(padded_day_data)

    if not final_rows:
        return pd.DataFrame()
        
    # Concatenate all days back together
    grouped = pd.concat(final_rows, ignore_index=True)
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
                # Column 1: UTC Offset (Date)
                ws.cell(row=idx, column=col_start, value=date_str) 
                
                # Column 2: Local Time Stamp
                ws.cell(row=idx, column=col_start+1, value=r.Time) 
                
                # Column 3: Active Power (W) - The aggregated sum
                # FIX: Access the column by its guaranteed name (PSumW) instead of fragile positional index (_3)
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
                st.success(f"Sheet **{sheet_name}** processed successfully.")
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
