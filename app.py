import streamlit as st
import pandas as pd
import io
import datetime

# --- Configuration ---
# Define the expected input column names in the source Excel file
# UPDATED: Changed 'Timestamp' to 'Date & Time' to match the user's input file structure.
TIMESTAMP_COL = 'Date & Time' 
POWER_COL_IN = 'PSum (W)'

# Define the output header columns that repeat for each day (4 columns total)
OUTPUT_HEADERS = [
    'UTC Offset (minutes)', 
    'Local Time Stamp', 
    'Active Power (W)', 
    'kW'
]

def transform_sheet(df: pd.DataFrame, sheet_name: str) -> pd.DataFrame:
    """
    Transforms a single input sheet (DataFrame) into the required wide-format 
    structure (4 columns per day).

    Args:
        df: The input DataFrame containing time series data.
        sheet_name: The name of the original sheet.

    Returns:
        A new DataFrame in the required wide format, or None if essential columns are missing.
    """
    st.info(f"Processing sheet: **{sheet_name}**...")
    
    # 1. Input Validation and Preparation
    required_cols = [TIMESTAMP_COL, POWER_COL_IN]
    if not all(col in df.columns for col in required_cols):
        st.error(f"Sheet **{sheet_name}** is missing required columns. Expected: '{TIMESTAMP_COL}' and '{POWER_COL_IN}'.")
        return None

    try:
        # Convert the timestamp column to datetime objects
        # FIX APPLIED: Added format='mixed', dayfirst=True to correctly parse date formats 
        # that use Day/Month/Year structure (e.g., 13/11/2025).
        df[TIMESTAMP_COL] = pd.to_datetime(
            df[TIMESTAMP_COL],
            format='mixed', 
            dayfirst=True
        )
    except Exception as e:
        st.error(f"Error converting column '{TIMESTAMP_COL}' to datetime in sheet {sheet_name}. Error: {e}")
        return None

    # Sort data by timestamp to ensure consistent intervals
    df = df.sort_values(by=TIMESTAMP_COL).reset_index(drop=True)

    # 2. Derive new columns based on requirements
    
    # Date (for "UTC Offset" column)
    df['Date'] = df[TIMESTAMP_COL].dt.date
    
    # 10-minute interval time (for "Local Time Stamp" column)
    df['Local Time Stamp'] = df[TIMESTAMP_COL].dt.strftime('%H:%M')
    
    # Active Power (W) (from input)
    df['Active Power (W)'] = df[POWER_COL_IN]
    
    # kW: (modulus assumed to mean magnitude, then convert W to kW)
    df['kW'] = df['Active Power (W)'].abs() / 1000
    
    # 3. Restructure Data into Wide Format (Day by Day)
    
    final_df = pd.DataFrame()
    all_dates = df['Date'].unique()
    
    # Track column index for naming (A, E, I, ...)
    current_col_index = 0
    
    for date in all_dates:
        # Filter data for the current day
        day_group = df[df['Date'] == date].copy()
        
        # Select the 4 required output metrics
        day_data = day_group[['Date', 'Local Time Stamp', 'Active Power (W)', 'kW']].reset_index(drop=True)
        
        # Rename the columns internally for clarity before concatenation
        day_data.columns = OUTPUT_HEADERS
        
        # Set the required date value for the "UTC Offset (minutes)" column (first column of the block)
        # Note: The requirement is to show the date here. The column name "UTC Offset (minutes)" is misleading 
        # but the content must be the date as per the prompt.
        day_data['UTC Offset (minutes)'] = date.strftime('%Y-%m-%d')
        
        # Rename columns to maintain the repeating structure (e.g., A, B, C, D, A, B, C, D...)
        # Concatenate the current day's 4-column block to the final DataFrame
        final_df = pd.concat([final_df, day_data], axis=1)

    st.success(f"Sheet **{sheet_name}** processed successfully. Found {len(all_dates)} days of data.")
    return final_df

def app():
    """Main Streamlit application function."""
    st.set_page_config(page_title="Energy Data Converter", layout="wide")
    
    st.title("💡 Energy Data Wide-Format Converter")
    st.markdown("""
        Upload an Excel file (.xlsx) containing time-series energy data. 
        
        **Expected Input Columns:**
        1.  `Date & Time` (or equivalent column containing date and time information)
        2.  `PSum (W)` (Active Power data)
        
        The program will convert the data into a wide format where each day's data 
        occupies a repeating block of 4 columns.
        """)

    uploaded_file = st.file_uploader(
        "Choose an Excel file (.xlsx)", 
        type=["xlsx"],
        help="The file can contain multiple sheets, which will be processed independently."
    )

    if uploaded_file is not None:
        try:
            # Use ExcelFile to read all sheets
            xls = pd.ExcelFile(uploaded_file)
            sheet_names = xls.sheet_names
            
            output_buffer = io.BytesIO()
            
            # Use Pandas ExcelWriter to write processed data to multiple sheets in memory
            with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
                all_processed_successfully = True
                
                for sheet_name in sheet_names:
                    # Read the current sheet
                    try:
                        df_in = xls.parse(sheet_name)
                    except Exception as e:
                        st.error(f"Could not read sheet '{sheet_name}'. Error: {e}")
                        all_processed_successfully = False
                        continue
                        
                    # Process the data
                    df_out = transform_sheet(df_in, sheet_name)
                    
                    if df_out is not None and not df_out.empty:
                        # Write the processed DataFrame to a new sheet in the output file
                        df_out.to_excel(
                            writer, 
                            sheet_name=sheet_name, 
                            index=False,
                            header=True # 'First row is the header row'
                        )
                    elif df_out is None:
                        all_processed_successfully = False

                if all_processed_successfully and sheet_names:
                    # Prepare file for download
                    st.download_button(
                        label="Download Processed Excel File (.xlsx)",
                        data=output_buffer.getvalue(),
                        file_name="Converted_Energy_Data.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    st.balloons()
                    st.info("The processed file is ready for download above!")
                elif not sheet_names:
                    st.warning("The uploaded file appears to be empty or has no sheets.")
                else:
                    st.error("Processing failed for one or more sheets. Please check the error messages above and ensure your input file format is correct.")

        except Exception as e:
            st.error(f"An unexpected error occurred during file processing: {e}")
            st.exception(e)

if __name__ == '__main__':
    app()
