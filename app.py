import streamlit as st
import pandas as pd
import io
import datetime

# --- Configuration ---
# Define the expected input column names in the source Excel file
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
    structure (4 columns per day), enforcing fixed 10-minute intervals and 
    aggregating power data.

    Args:
        df: The input DataFrame containing time series data.
        sheet_name: The name of the original sheet.

    Returns:
        A new DataFrame in the required wide format, or an empty DataFrame if processing fails.
    """
    st.info(f"Processing sheet: **{sheet_name}**...")
    
    # 1. Input Validation and Preparation
    required_cols = [TIMESTAMP_COL, POWER_COL_IN]
    if not all(col in df.columns for col in required_cols):
        st.error(f"Sheet **{sheet_name}** is missing required columns. Expected: '{TIMESTAMP_COL}' and '{POWER_COL_IN}'.")
        return pd.DataFrame()

    try:
        # Convert the timestamp column to datetime objects, handling D/M/Y ambiguity
        df[TIMESTAMP_COL] = pd.to_datetime(
            df[TIMESTAMP_COL],
            format='mixed', 
            dayfirst=True
        )
    except Exception as e:
        st.error(f"Error converting column '{TIMESTAMP_COL}' to datetime in sheet {sheet_name}. Error: {e}")
        return pd.DataFrame()

    # Drop rows where timestamp conversion failed (resulted in NaT)
    df = df.dropna(subset=[TIMESTAMP_COL]).sort_values(by=TIMESTAMP_COL).reset_index(drop=True)
    
    # CRITICAL FIX: Aggressively clean and ensure Power column is numeric before aggregation.
    
    # 1. Convert to string, strip whitespace.
    power_series = df[POWER_COL_IN].astype(str).str.strip()
    
    # 2. Replace commas with periods to handle international decimal format.
    power_series = power_series.str.replace(',', '.', regex=False)
    
    # 3. AGGRESSIVE CLEANING: Remove everything that is NOT a digit, a period, or a minus sign.
    # This removes hidden units, currency symbols, and other non-numeric garbage.
    power_series = power_series.str.replace(r'[^0-9\.\-]', '', regex=True)

    # 4. Coerce non-numeric values to NaN.
    df[POWER_COL_IN] = pd.to_numeric(power_series, errors='coerce')
    
    # 5. Drop rows where power value is invalid (NaN)
    df = df.dropna(subset=[POWER_COL_IN])
    
    # Set the valid timestamp column as the index for resampling
    df_indexed = df.set_index(TIMESTAMP_COL)

    # 2. Resample data into fixed 10-minute intervals (144 intervals per day)
    
    transformed_data = []
    
    # Find all unique dates to iterate over. .index.normalize() gets the date component.
    all_dates = df_indexed.index.normalize().unique().date
    all_dates = pd.Series(all_dates).dropna().unique()

    # Create a template index for a full 24-hour day (144 points)
    # This is crucial for consistent horizontal alignment. Use datetime.time objects for index.
    time_only_index = pd.to_datetime([f'{i:02d}:{j:02d}' for i in range(24) for j in range(0, 60, 10)], format='%H:%M').time
    template_df = pd.DataFrame(index=time_only_index)

    for date in all_dates:
        # Define the start and end of the current day for clean slicing
        day_start = pd.Timestamp(date)
        day_end = day_start + pd.Timedelta(days=1)
        # Select data for the current day
        day_group = df_indexed.loc[day_start:day_end - pd.Timedelta(seconds=1)]

        if day_group.empty:
            continue

        # Resample the PSum (W) column to 10-minute intervals. 
        # Sum is used for aggregation, matching the requirement (00:00:00 up to 00:09:59...).
        resampled_series = day_group[POWER_COL_IN].resample(
            '10min', 
            label='left', 
            origin='start'
        ).sum()

        # Create the daily output DataFrame using the resampled data
        daily_output = pd.DataFrame({
            'Active Power (W)': resampled_series.values,
        }, index=resampled_series.index)
        
        # --- Prepare for Padding and Final Output Columns ---
        
        # Extract the time component (datetime.time) from the DatetimeIndex and set it as the new index
        daily_output.index = daily_output.index.time
        
        # Re-index against the full 144-row template to fill missing intervals with NaN
        final_daily_output = template_df.join(daily_output, how='left')
        
        # Calculate derived metrics and set the required output columns
        final_daily_output['UTC Offset (minutes)'] = date.strftime('%Y-%m-%d')
        
        # Use a list comprehension to call strftime on each element 
        # of the Index to correctly generate the time stamps.
        final_daily_output['Local Time Stamp'] = [t.strftime('%H:%M') for t in final_daily_output.index]

        # These columns should now be filled because Active Power (W) is numeric
        final_daily_output['kW'] = final_daily_output['Active Power (W)'].abs() / 1000
        
        # Final column selection and naming convention
        final_daily_output = final_daily_output[['UTC Offset (minutes)', 'Local Time Stamp', 'Active Power (W)', 'kW']]
        final_daily_output.columns = OUTPUT_HEADERS 
        
        transformed_data.append(final_daily_output.reset_index(drop=True))

    if not transformed_data:
        st.warning(f"Sheet **{sheet_name}** contained no valid time series data after processing.")
        return pd.DataFrame()

    # Concatenate all daily blocks horizontally
    final_df = pd.concat(transformed_data, axis=1)

    st.success(f"Sheet **{sheet_name}** processed successfully. Found {len(transformed_data)} days of data.")
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
        occupies a repeating block of 4 columns, and the date column will be merged 
        for a cleaner presentation.
        """)

    uploaded_file = st.file_uploader(
        "Choose an Excel file (.xlsx)", 
        type=["xlsx"],
        help="The file can contain multiple sheets, which will be processed independently."
    )

    if uploaded_file is not None:
        try:
            xls = pd.ExcelFile(uploaded_file)
            sheet_names = xls.sheet_names
            
            output_buffer = io.BytesIO()
            all_processed_successfully = True
            
            # Use Pandas ExcelWriter with xlsxwriter engine
            writer = pd.ExcelWriter(output_buffer, engine='xlsxwriter')
            
            try:
                for sheet_name in sheet_names:
                    try:
                        df_in = xls.parse(sheet_name)
                    except Exception as e:
                        st.error(f"Could not read sheet '{sheet_name}'. Error: {e}")
                        all_processed_successfully = False
                        continue
                        
                    df_out = transform_sheet(df_in, sheet_name)
                    
                    if df_out is not None and not df_out.empty:
                        # 1. Write the processed DataFrame to a new sheet
                        # Do not write the index. Header is required (first row).
                        df_out.to_excel(
                            writer, 
                            sheet_name=sheet_name, 
                            index=False,
                            header=True
                        )
                        
                        # 2. Apply Excel Formatting (Cell Merging for Date Column)
                        workbook = writer.book
                        worksheet = writer.sheets[sheet_name]
                        
                        # Define the format for the merged cell (center text)
                        merge_format = workbook.add_format({
                            'align': 'center', 
                            'valign': 'vcenter'
                        })

                        num_rows = len(df_out)
                        num_cols = len(df_out.columns)
                        num_days = num_cols // 4
                        
                        for i in range(num_days):
                            # The date column is the first in every 4-column block: 0, 4, 8, ...
                            col_index = i * 4 
                            
                            # The date value is in the first data row (Excel row 1, Python index 0)
                            # We use try/except block just in case the value is missing in the first row
                            try:
                                date_value = df_out.iloc[0, col_index]
                            except IndexError:
                                # This handles the unlikely case of a completely empty column block
                                continue 
                            
                            # Merge cells for the date column from the first data row (1) 
                            # down to the last data row (num_rows). Column indices are 0-based.
                            # range is (row_start, col_start, row_end, col_end)
                            worksheet.merge_range(
                                1,                    # Start row (1st data row, after header 0)
                                col_index,            # Start column (0, 4, 8, ...)
                                num_rows,             # End row (last data data row index)
                                col_index,            # End column (same as start)
                                date_value,           # The value to display in the merged cell
                                merge_format          # Formatting
                            )
                    elif df_out is None or df_out.empty:
                        all_processed_successfully = False

            except Exception as e:
                st.error(f"An error occurred during sheet processing: {e}")
                all_processed_successfully = False
            
            # CRITICAL: Explicitly close the writer to finalize the Excel file structure in the buffer
            writer.close()

            if all_processed_successfully and sheet_names:
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
