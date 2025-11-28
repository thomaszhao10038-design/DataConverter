import streamlit as st
import pandas as pd
import re
from io import BytesIO
import numpy as np
from datetime import datetime, timedelta

# Import xlsxwriter (or openpyxl) for Excel writing
try:
    import xlsxwriter 
except ImportError:
    pass
try:
    import openpyxl
except ImportError:
    pass


st.set_page_config(layout="wide")
st.title("⚡ Streamlined Electricity Data Aggregator")
st.markdown("---")
st.markdown("This tool is optimized for **multi-sheet XLSX files**. It consolidates data, aggregates it to a specified frequency, zero-fills all gaps, and provides a clean Excel export.")


# Dictionary to map user-friendly options to Pandas frequency strings and display labels
FREQ_MAP = {
    '10-minute': {'pandas_freq': '10T', 'label': '10-minute period'},
    'Hourly': {'pandas_freq': 'H', 'label': 'Hourly period'},
    'Daily': {'pandas_freq': 'D', 'label': 'Daily period'},
}

# HARDCODED: Set interpolation method to 'none' to enforce zero-filling
INTERPOLATION_METHOD = 'none' 
# Since we are removing outlier checking, max_kwh/min_kwh are fixed to defaults.
DEFAULT_MAX_KWH = 1000000 
DEFAULT_MIN_KWH = -1000000


# ---------- UI Controls and File Upload ----------

st.sidebar.header("Data Control & Analysis")
uploaded_file = st.file_uploader("Upload an Excel file (.xlsx) containing energy data across multiple sheets", type=["xlsx"])

# ---------- Frequency Selection Control ----------
st.sidebar.markdown("### Aggregation Settings")
selected_freq_key = st.sidebar.selectbox(
    "Select Aggregation Period",
    list(FREQ_MAP.keys()),
    index=0 # Default to 10-minute
)

# Get the dynamic frequency settings
freq_str = FREQ_MAP[selected_freq_key]['pandas_freq']
period_label = FREQ_MAP[selected_freq_key]['label']
# ---------------------------------------------------


# ---------- Helper Functions (Simplified) ----------

def detect_metadata(lines):
    """Attempts to infer the energy unit (e.g., kWh, Watt) from the file content."""
    unit = 'Watt' # Defaulting to Watt based on provided snippet
    unit_patterns = {
        r'(?:k|K)W[hH]|kW-h': 'kWh',
        r'(?:M|m)W[hH]|MW-h': 'MWh',
        r'(?:W|w)att|\(W\)|PSum': 'Watt', # Expanded patterns
        r'(?:J|j)oule': 'Joule'
    }
    sample_text = " ".join(lines[:200])
    for pattern, detected_unit in unit_patterns.items():
        if re.search(pattern, sample_text, re.IGNORECASE):
            unit = detected_unit
            break
    return unit

def extract_lines(file):
    """
    Extracts raw text lines from a single XLSX file. 
    Consolidates data from ALL sheets.
    """
    lines = []
    try:
        # Reads ALL sheets into a dictionary of DataFrames
        df = pd.read_excel(file, sheet_name=None, dtype=str)
        for sheet_name, sheet_df in df.items():
            # Consolidate all columns from all sheets into one list of strings
            for col in sheet_df.columns:
                lines += sheet_df[col].dropna().astype(str).tolist()
    except Exception as e:
        st.error(f"Error during file reading: {e}")
    return lines

def parse_energy_data(lines, max_kwh=DEFAULT_MAX_KWH, min_kwh=DEFAULT_MIN_KWH):
    """
    Extracts Datetime and Reading values from raw lines.
    Data validation is ignored as per new instructions.
    """
    # Common Date/Time patterns
    timestamp_patterns = [
        r"\d{4}-\d{2}-\d{2} \d{2}:\d{2}(:\d{2})?", # YYYY-MM-DD HH:MM:SS
        r"\d{2}/\d{2}/\d{4} \d{2}:\d{2}(:\d{2})?", # DD/MM/YYYY HH:MM:SS or MM/DD/YYYY HH:MM:SS
        r"\d{4}\d{2}\d{2}\s\d{4}"                # YYYYMMDD HHMM
    ]
    timestamp_regex = re.compile("|".join(timestamp_patterns))
    number_regex = re.compile(r"[-+]?(\d*\.\d+|\d+)")

    data = []
    buffered_ts_str = None

    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        ts_match = timestamp_regex.search(line)
        tokens = re.split(r"[,\s;]+", line)
        
        current_ts_str = ts_match.group().strip() if ts_match else None
        found_value = None
        
        for token in tokens:
            if number_regex.fullmatch(token):
                if current_ts_str and token in current_ts_str:
                    continue 

                try:
                    val = float(token)
                    # No validation applied; all found values are accepted
                    found_value = val
                    break
                except ValueError:
                    continue
        
        if current_ts_str and found_value is not None:
            # Timestamp and value found in the same line
            data.append({"Datetime_str": current_ts_str, "Reading": found_value})
            buffered_ts_str = None
                
        elif found_value is not None and buffered_ts_str:
            # Value found on a separate line after a timestamp
            data.append({"Datetime_str": buffered_ts_str, "Reading": found_value})
            buffered_ts_str = None
            
        elif current_ts_str:
            # Only timestamp found; buffer it
            buffered_ts_str = current_ts_str
            
    df = pd.DataFrame(data)
    
    if df.empty:
        return pd.DataFrame({'Datetime': [], 'Reading': []})

    df['Datetime'] = pd.to_datetime(df['Datetime_str'], errors='coerce')
    df.drop(columns=['Datetime_str'], inplace=True)
    
    df = df.dropna(subset=['Datetime'])
    df = df.sort_values('Datetime').drop_duplicates(subset=['Datetime'], keep='first').reset_index(drop=True)
    
    return df

def clean_energy_df(df: pd.DataFrame, freq: str) -> pd.DataFrame:
    """Sets index, enforces frequency, aggregates, and fills gaps with zero."""
    df = df.copy()
    if 'Datetime' not in df.columns or df['Datetime'].empty:
        return df

    # Use a flag 'Raw_Count' to distinguish between True Gaps (Count=0) and actual readings
    df['Raw_Count'] = 1 
    df.set_index('Datetime', inplace=True)
    
    # 1. Aggregate/Resample Data
    df_resampled = df.resample(freq).agg({
        'Reading': 'sum', 
        'Raw_Count': 'sum' 
    }).rename(columns={'Reading': 'Aggregated_Reading', 'Raw_Count': 'Raw_Count_sum'}) 

    
    # 2. Zero-Fill and Flag
    is_true_gap = (df_resampled['Raw_Count_sum'] == 0)
    
    # Fill NaN (gaps) with 0
    df_resampled['Zero_Filled_Reading'] = df_resampled['Aggregated_Reading'].fillna(0)
    
    # Assign the final flag
    df_resampled['Flag'] = 'OK (Aggregated/Binned)'
    df_resampled['Flag'] = np.where(is_true_gap, 'True Gap (Zero-Filled)', df_resampled['Flag'])
    
    # Calculate Cumulative Electricity based on the filled, continuous data
    df_resampled['Cumulative_Reading'] = df_resampled['Zero_Filled_Reading'].cumsum().fillna(0)
    
    # Final cleanup and indexing
    df_resampled.drop(columns=['Aggregated_Reading', 'Raw_Count_sum'], inplace=True)
    df_resampled.reset_index(inplace=True)
    return df_resampled

# ---------- Main Application Logic ----------
if uploaded_file:
    
    # 1. Extraction and Unit Detection
    lines = extract_lines(uploaded_file)
    unit = detect_metadata(lines) 
    
    # 2. Parsing
    df_raw = parse_energy_data(lines)
    if df_raw.empty:
        st.error("❌ No valid timestamp–reading pairs were found in the Excel file.")
        st.stop()
        
    st.success(f"Successfully extracted {df_raw.shape[0]} raw readings. Detected unit: **{unit}**.")

    with st.expander("Raw Data Status"):
        st.info(f"Raw data timestamps were inconsistent and have been consolidated. Data is now being resampled/aggregated to **{period_label}** and all gaps are Zero-Filled.")

    # 3. Resampling/Cleaning and Interpolation
    df_clean = clean_energy_df(df_raw, freq=freq_str)

    # ---------- Show Results (Minimal Display) ----------
    
    df_clean_display = df_clean.rename(columns={
        'Zero_Filled_Reading': f'Usage_Per_{period_label.replace(" ", "_").replace("-", "")}_{unit}', 
        'Cumulative_Reading': f'Cumulative_Total_{unit}' 
    })

    st.subheader("🧹 Cleaned & Resampled Data")
    st.caption(f"The table below shows the continuous time series where all gaps are Zero-Filled and data is aggregated to the {period_label}.")
    
    # Select only the relevant columns for the final display and export
    df_final_export = df_clean_display[[
        'Datetime', 
        f'Usage_Per_{period_label.replace(" ", "_").replace("-", "")}_{unit}', 
        f'Cumulative_Total_{unit}', 
        'Flag'
    ]].copy()
    
    st.dataframe(df_final_export, use_container_width=True)

    # ---------- Data Integrity Check (Simplified) ----------
    st.markdown("### Data Integrity Check")
    flag_counts = df_clean['Flag'].value_counts()
    
    zero_filled_count = flag_counts.get('True Gap (Zero-Filled)', 0)
    ok_count = flag_counts.get('OK (Aggregated/Binned)', 0)
    
    if zero_filled_count > 0:
        st.warning(f"  - **{zero_filled_count}** periods were **True Gaps** (no raw data present in the {period_label} bin) and were Zero-Filled.")
        st.info(f"Summary: A total of **{zero_filled_count}** periods required Zero-Filling to create a continuous {period_label} series.")
    else:
        st.success("The time series is continuous and contained no gaps at the selected resolution.")
        
    st.success(f"✅ **{ok_count}** periods contained aggregated raw data.")
    st.markdown("---")


    # ---------- Download Excel (DATA ONLY) ----------
    output = BytesIO()
    
    excel_engine = "openpyxl"
    if 'xlsxwriter' in globals() and xlsxwriter is not None:
        excel_engine = "xlsxwriter"
        
    try:
        with pd.ExcelWriter(output, engine=excel_engine) as writer:
            sheet_name = "Cleaned_Data"
            
            # Write the final export data to the Excel sheet
            df_final_export.to_excel(writer, index=False, sheet_name=sheet_name, startrow=0)
            
    except Exception as e:
        st.error(f"Error during Excel export: {e}")
        st.stop()

    output.seek(0)
    st.download_button(
        "Download Cleaned Data (Excel)",
        data=output,
        file_name=f"{uploaded_file.name.split('.')[0]}_cleaned_data_export_{selected_freq_key.lower().replace('-', '')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
