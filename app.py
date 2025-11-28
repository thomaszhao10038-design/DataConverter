import streamlit as st
import pandas as pd
import io

# Set up the Streamlit page configuration
st.set_page_config(
    page_title="10-Minute Interval Power Data Converter",
    layout="wide",
    initial_sidebar_state="auto"
)

# --- Core Data Processing Function ---
def process_power_data(uploaded_file):
    """
    Reads the input file, creates a full 10-minute time series,
    merges the PSum data, and calculates kW (absolute modulus).
    """
    st.info("Starting data processing...")

    # 1. Read the input file (handling both CSV and Excel, though the upload is named .xlsx)
    try:
        # Use filename extension to determine reader
        if uploaded_file.name.endswith('.csv'):
            df_raw = pd.read_csv(uploaded_file)
        elif uploaded_file.name.endswith(('.xlsx', '.xls')):
            # Assumes the data is in the first sheet for simplicity
            df_raw = pd.read_excel(uploaded_file, sheet_name=0)
        else:
            st.error("Unsupported file type. Please upload a .csv or .xlsx file.")
            return None
            
        # Standardize column names based on the example input
        # The column is 'Date & Time' and the value column is 'PSum (W)'
        df_raw.columns = [col.strip() for col in df_raw.columns]
        datetime_col = 'Date & Time' if 'Date & Time' in df_raw.columns else df_raw.columns[0]
        ps_col = 'PSum (W)' if 'PSum (W)' in df_raw.columns else df_raw.columns[1]
        
        df = df_raw.rename(columns={datetime_col: 'Timestamp', ps_col: 'PSum (W)'})
        df = df[['Timestamp', 'PSum (W)']].dropna(subset=['Timestamp'])

    except Exception as e:
        st.error(f"Error reading or preparing file: {e}")
        return None

    # 2. Convert 'Timestamp' to datetime objects and set as index
    try:
        df['Timestamp'] = pd.to_datetime(df['Timestamp'], dayfirst=True) # Assuming Day-first format from snippet
        df = df.set_index('Timestamp').sort_index()
    except Exception as e:
        st.error(f"Error converting timestamp column: {e}. Check your date format.")
        return None

    # 3. Determine the full date range
    if df.empty:
        st.warning("The input file is empty after cleaning.")
        return None
        
    start_date = df.index.min().floor('D')
    # The end date is the start of the next day after the max recorded date
    end_date = df.index.max().ceil('D')
    
    # 4. Create the full 10-minute time range for all days (24 hours)
    # The range starts at 00:00:00 of the first day and ends at 23:50:00 of the last day
    full_time_range = pd.date_range(
        start=start_date,
        end=end_date - pd.Timedelta(minutes=10), # Stop before 00:00 of the next day
        freq='10min'
    )
    
    # Create a DataFrame from the full range
    df_full = pd.DataFrame(index=full_time_range)
    
    # 5. Reindex the raw data onto the full range
    # This ensures every 10-minute interval is present, filling gaps with NaN
    df_result = df_full.merge(df, how='left', left_index=True, right_index=True)
    
    # 6. Calculate the kW column
    # kW = |PSum (W)| / 1000
    # Use .abs() for modulus and fillna(0) ensures that if PSum is NaN, kW is 0. 
    # However, keeping it as NaN (empty) might be more accurate if the value is truly missing.
    # We will only calculate kW if PSum is not NaN, and leave it NaN otherwise.
    df_result['kW'] = df_result['PSum (W)'].apply(lambda x: abs(x) / 1000 if pd.notna(x) else pd.NA)

    # 7. Final formatting and column separation
    df_result['date'] = df_result.index.strftime('%Y-%m-%d')
    df_result['local time stamp'] = df_result.index.strftime('%H:%M:%S')

    # Reorder and rename columns as requested
    df_final = df_result.reset_index(drop=True)[['date', 'local time stamp', 'PSum (W)', 'kW']]
    
    st.success("Data conversion complete! The output file now contains a continuous 10-minute interval time series.")
    
    return df_final

# --- Streamlit UI ---

st.title("DataConverter: 10-Minute Interval Generator")
st.markdown(f"""
This application converts your raw power data (like your `{uploaded_file.originalName}` example) into a standard 10-minute interval Excel file.

**Output Requirements Implemented:**
1.  **Continuous:** Generates a complete 24-hour, 10-minute timestamp for every day in the dataset, even if no data was recorded.
2.  **kW Calculation:** The `kW` column is calculated as the **modulus** (absolute value) of `PSum (W)` and converted to **kilowatts** ($W \rightarrow kW$).
""")

uploaded_file = st.file_uploader(
    "Upload your power data file (.csv or .xlsx)",
    type=['csv', 'xlsx']
)

if uploaded_file is not None:
    
    # Check if processing should start
    if st.button("Process Data and Generate Excel"):
        
        # Process data
        df_output = process_power_data(uploaded_file)
        
        if df_output is not None:
            # Create a Pandas Excel writer object for in-memory storage
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # The sheet name can be anything; '10min_Data' is descriptive
                df_output.to_excel(writer, sheet_name='10min_Data', index=False)
            
            output.seek(0)
            
            # --- Download Button ---
            st.download_button(
                label="📥 Download Processed Excel File",
                data=output,
                file_name="Processed_10min_Power_Data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Click to download the new Excel file with 10-minute intervals."
            )
            
            st.subheader("Preview of Generated Data")
            st.dataframe(df_output.head(50)) # Show the first 50 rows for verification
            st.write(f"Total rows generated: **{len(df_output)}**")
