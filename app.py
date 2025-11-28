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
            # The 'engine' parameter is not needed for CSV reading unless dealing with specific delimiters/encodings
            df_raw = pd.read_csv(uploaded_file)
        elif uploaded_file.name.endswith(('.xlsx', '.xls')):
            # Assumes the data is in the first sheet for simplicity
            # Pandas will use 'openpyxl' or 'xlrd' (if installed) to read the Excel file
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
        # Assuming Day-first format (e.g., 26/11/2025) from the snippet
        df['Timestamp'] = pd.to_datetime(df['Timestamp'], dayfirst=True) 
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
    end_date = df.index.max().ceil
