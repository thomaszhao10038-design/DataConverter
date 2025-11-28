import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from PyPDF2 import PdfReader
from docx import Document
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
st.title("⚡ Electricity Data Cleaner & Visualizer")
st.markdown("---")
st.markdown("Upload your energy log file (supports CSV, TXT, XLSX, DOCX, PDF) to clean, resample, and visualize your usage data.")


# Dictionary to map user-friendly options to Pandas frequency strings and display labels
FREQ_MAP = {
    '10-minute': {'pandas_freq': '10T', 'label': '10-minute period'},
    'Hourly': {'pandas_freq': 'H', 'label': 'Hourly period'},
    'Daily': {'pandas_freq': 'D', 'label': 'Daily period'},
}

# ---------- UI Controls and File Upload ----------

st.sidebar.header("Data Control & Analysis")
uploaded_file = st.file_uploader("Upload a file containing energy data (txt, csv, xlsx, docx, pdf)", type=["txt","csv","xlsx","docx","pdf"])

# Configuration Parameters for Cleaning
st.sidebar.markdown("### Cleaning Parameters")
st.sidebar.info("Missing/Invalid data (Gaps) will be **Zero-Filled** (treated as 0 kWh/W) as per current application settings.")
max_kwh = st.sidebar.slider("Max Valid Reading (Outlier Threshold)", min_value=10.0, max_value=200.0, value=50.0, step=5.0)
min_kwh = st.sidebar.slider("Min Valid Reading (Invalid Threshold)", min_value=-5.0, max_value=5.0, value=0.0, step=0.5)

# HARDCODED: Set interpolation method to 'none' to enforce zero-filling
INTERPOLATION_METHOD = 'none' 

# ---------- New Frequency Selection Control ----------
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


# ---------- Helper Functions ----------

def detect_metadata(lines):
    """Attempts to infer the energy unit (e.g., kWh, Watt) from the file content."""
    unit = 'kWh'
    unit_patterns = {
        r'(?:k|K)W[hH]|kW-h': 'kWh',
        r'(?:M|m)W[hH]|MW-h': 'MWh',
        r'(?:W|w)att|\(W\)': 'Watt', # Added (W) based on user's data snippet
        r'(?:J|j)oule': 'Joule'
    }
    sample_text = " ".join(lines[:200])
    for pattern, detected_unit in unit_patterns.items():
        if re.search(pattern, sample_text, re.IGNORECASE):
            unit = detected_unit
            break
    return unit

def extract_lines(file, file_type):
    """
    Extracts raw text lines from various file types. 
    Crucially, for XLSX, it consolidates data from ALL sheets.
    """
    lines = []
    try:
        if file_type in ["txt","csv"]:
            text = file.read().decode("utf8", errors="ignore")
            lines = [line.strip() for line in text.splitlines() if line.strip()]
        elif file_type == "xlsx":
            # Reads ALL sheets into a dictionary of DataFrames
            df = pd.read_excel(file, sheet_name=None, dtype=str)
            for sheet_name, sheet_df in df.items():
                # Consolidate all columns from all sheets into one list of strings
                for col in sheet_df.columns:
                    lines += sheet_df[col].dropna().astype(str).tolist()
        elif file_type == "docx":
            doc = Document(file)
            for para in doc.paragraphs:
                if para.text.strip():
                    lines.append(para.text.strip())
        elif file_type == "pdf":
            reader = PdfReader(file)
            for page in reader.pages:
                text = page.extract_text()
                if text:
                    lines += [line.strip() for line in text.splitlines() if line.strip()]
    except Exception as e:
        st.error(f"Error during file reading: {e}")
    return lines

def parse_energy_data(lines, max_kwh, min_kwh, unit_name='kWh'):
    """Extracts Datetime and kWh/Watt values from raw lines, accommodating various formats."""
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
        flag = None
        
        for token in tokens:
            if number_regex.fullmatch(token):
                if current_ts_str and token in current_ts_str:
                    continue 

                try:
                    val = float(token)
                    # Data validation based on user-defined thresholds
                    if min_kwh <= val <= max_kwh:
                        found_value = val
                        flag = "OK"
                        break
                    elif val > max_kwh:
                        found_value = float('nan')
                        flag = f"Outlier (>{max_kwh}{unit_name})"
                        break
                    else:
                        found_value = float('nan')
                        flag = f"Invalid (<{min_kwh}{unit_name})"
                        break
                except ValueError:
                    continue
        
        if current_ts_str:
            if found_value is not None:
                data.append({"Datetime_str": current_ts_str, "kWh": found_value, "Flag": flag})
                buffered_ts_str = None
            else:
                buffered_ts_str = current_ts_str
                
        elif found_value is not None and buffered_ts_str:
            data.append({"Datetime_str": buffered_ts_str, "kWh": found_value, "Flag": flag})
            buffered_ts_str = None
            
    df = pd.DataFrame(data)
    
    if df.empty:
        return pd.DataFrame({'Datetime': [], 'kWh': [], 'Flag': []})

    df['Datetime'] = pd.to_datetime(df['Datetime_str'], errors='coerce')
    df.drop(columns=['Datetime_str'], inplace=True)
    
    df = df.dropna(subset=['Datetime'])
    df = df.sort_values('Datetime').drop_duplicates(subset=['Datetime'], keep='first').reset_index(drop=True)
    
    return df

def clean_energy_df(df: pd.DataFrame, freq: str, interpolation_method: str = 'none') -> pd.DataFrame:
    """Sets index, enforces frequency, aggregates, and fills gaps according to interpolation_method."""
    df = df.copy()
    if 'Datetime' not in df.columns or df['Datetime'].empty:
        return df

    df['Raw_Count'] = 1 
    df.set_index('Datetime', inplace=True)
    
    def flag_aggregator_simple(x):
        """Determines if the raw data in the bin contained only invalid/outlier points."""
        valid_flags = x.dropna().tolist()
        if not valid_flags:
            return 'True Gap Placeholder'
        elif all(f.startswith('Outlier') or f.startswith('Invalid') for f in valid_flags):
            return 'Invalid/Outlier' 
        else:
            return 'OK'

    # 1. Aggregate/Resample Data and Flags using the dynamic frequency `freq`
    df_resampled = df.resample(freq).agg({
        'kWh': 'sum', 
        'Flag': flag_aggregator_simple,
        'Raw_Count': 'sum' 
    }).rename(columns={'kWh': 'Aggregated_kWh', 'Raw_Count': 'Raw_Count_sum', 'Flag': 'Aggregated_Flag_Type'}) 

    
    # 2. Gap Detection & Final Flag Assignment
    is_true_gap = (df_resampled['Raw_Count_sum'] == 0)
    is_invalid_gap = (df_resampled['Raw_Count_sum'] > 0) & (df_resampled['Aggregated_Flag_Type'] == 'Invalid/Outlier')
    is_ok = (df_resampled['Aggregated_Flag_Type'] == 'OK')
    
    
    # --- Interpolation/Zero-Fill (Hardcoded to 'none') ---
    
    if 'none' in interpolation_method:
        # Fill NaN (gaps) with 0
        df_resampled['Interpolated_kWh'] = df_resampled['Aggregated_kWh'].fillna(0)
        df_resampled['Flag'] = 'OK (Aggregated/Binned)'
        df_resampled['Flag'] = np.where(is_true_gap, 'True Gap (Zero-Filled)', df_resampled['Flag'])
        df_resampled['Flag'] = np.where(is_invalid_gap, 'Invalid/Outlier (Zero-Filled)', df_resampled['Flag'])
        df_resampled['Flag'] = np.where(is_ok, 'OK (Aggregated/Binned)', df_resampled['Flag'])
    
    # Calculate Cumulative Electricity based on the filled, continuous data
    df_resampled['Cumulative_kWh'] = df_resampled['Interpolated_kWh'].cumsum().fillna(0)
    
    # Final cleanup
    df_resampled.drop(columns=['Aggregated_Flag_Type', 'Raw_Count_sum'], inplace=True)
    df_resampled.rename(columns={'Aggregated_kWh': 'kWh'}, inplace=True)
    
    df_resampled.reset_index(inplace=True)
    return df_resampled

# ---------- Main Application Logic ----------
if uploaded_file:
    file_type = uploaded_file.name.split(".")[-1].lower()
    
    # 1. Extraction and Unit Detection
    lines = extract_lines(uploaded_file, file_type)
    unit = detect_metadata(lines) 
    
    # 2. Parsing
    df_raw = parse_energy_data(lines, max_kwh, min_kwh, unit_name=unit)
    if df_raw.empty:
        st.error("❌ No valid timestamp–reading pairs were found in the file.")
        st.stop()
        
    # --- Raw Frequency De-emphasis (as requested) ---
    raw_period_label = 'recorded interval'
    
    with st.expander("Raw Text Preview"):
        st.text_area("Preview (first 50 lines)", "\n".join(lines[:50]), height=200)
        # Informative message about inconsistent raw data being resampled
        st.info(f"Raw data timestamps were likely inconsistent. Data is being resampled/aggregated to **{period_label}**.")


    # 3. Resampling/Cleaning and Interpolation using dynamic `freq_str`
    df_clean = clean_energy_df(df_raw, freq=freq_str, interpolation_method=INTERPOLATION_METHOD)

    # --- Plot Mode Definition ---
    usage_column = 'Interpolated_kWh'
    line_style = '-'
    line_color = 'tab:blue'
    plot_mode_label = f'Zero-Filled Usage ({unit})'
    # -------------------------------------------------


    # ---------- Show Results ----------
    col1, col2 = st.columns([1, 1])

    df_raw_display = df_raw.rename(columns={'kWh': f'Usage_{unit}'})
    
    df_clean_display = df_clean.rename(columns={
        'kWh': f'Aggregated_Usage_{unit}', 
        'Interpolated_kWh': f'Zero_Filled_Usage_{unit}',
        'Cumulative_kWh': f'Cumulative_{unit}' 
    })

    with col1:
        st.subheader("📊 Raw Extracted Data")
        st.caption(f"Valid records detected before aggregation. Readings represent total consumption for the **{raw_period_label}**.")
        st.dataframe(df_raw_display, use_container_width=True)

    with col2:
        st.subheader("🧹 Cleaned & Resampled Data")
        st.caption(f"**Aggregated_Usage_{unit}** is the sum of raw readings within the {period_label} bin. Gaps are filled with zero.")
        st.dataframe(df_clean_display, use_container_width=True)

    # ---------- Alerts and Statistics ----------
    st.markdown("### Data Integrity Check")
    
    # --- Detailed Raw Data Status (Problematic data before aggregation) ---
    st.markdown("#### Raw Data Issues (Before Aggregation)")
    raw_flag_counts = df_raw['Flag'].value_counts()
    
    invalid_raw_count = raw_flag_counts.filter(regex=r'^Invalid').sum()
    outlier_raw_count = raw_flag_counts.filter(regex=r'^Outlier').sum()
        
    if invalid_raw_count > 0:
        st.error(f"🚩 **{invalid_raw_count}** raw readings flagged as **Invalid** (below {min_kwh} {unit}).")
        
    if outlier_raw_count > 0:
        st.error(f"🚩 **{outlier_raw_count}** raw readings flagged as **Outlier** (above {max_kwh} {unit}).")
        
    
    if invalid_raw_count + outlier_raw_count > 0:
        st.info(f"Total **{invalid_raw_count + outlier_raw_count}** raw readings were identified as problematic and were converted to NaN for the cleaning process.")
    else:
        st.success("🎉 No raw readings were flagged as Invalid or Outlier based on your current thresholds.")
        
    st.markdown("---")
    st.markdown(f"#### Aggregated Period Status (After Resampling to {period_label})")
    
    flag_counts = df_clean['Flag'].value_counts()
    
    zero_filled_true_gap_count = flag_counts.get('True Gap (Zero-Filled)', 0)
    zero_filled_invalid_count = flag_counts.get('Invalid/Outlier (Zero-Filled)', 0)
    zero_filled_total = zero_filled_true_gap_count + zero_filled_invalid_count
    
    if zero_filled_true_gap_count > 0:
        st.warning(f"  - **{zero_filled_true_gap_count}** periods were **True Gaps** (no raw data present in the {period_label} bin) and were Zero-Filled.")
    else:
        st.success(f"  - **0** periods were **True Gaps**.")


    if zero_filled_invalid_count > 0:
        st.warning(f"  - **{zero_filled_invalid_count}** periods contained **Invalid/Outlier** raw data and were Zero-Filled after aggregation.")

    if zero_filled_total > 0:
        st.info(f"Summary: A total of **{zero_filled_total}** periods required Zero-Filling to create a continuous {period_label} series.")
    
    ok_count = flag_counts.get('OK (Aggregated/Binned)', 0)
    if ok_count > 0:
        st.success(f"✅ **{ok_count}** periods contained valid, aggregated raw data.")
    
    if zero_filled_total == 0 and ok_count == 0:
        st.warning("No data found or processed in the aggregated time series.")
    elif zero_filled_total == 0 and ok_count > 0:
        st.success("The time series is continuous and contains no gaps at the selected resolution.")
        
    st.markdown("---")


    # ---------- Plot (Usage and Cumulative) - VISIBLE IN WEB APP ONLY ----------
    st.subheader(f"Energy Usage per {period_label} Over Time (Zero-Filled, Continuous)")
    
    fig, ax1 = plt.subplots(figsize=(10, 5)) 
    
    # --- AXIS 1: Usage per Period (Primary) ---
    color_inst = line_color
    ax1.set_xlabel("Datetime")
    ax1.set_ylabel(f"{unit} Usage per {period_label}", color=color_inst) 
    
    # 1. Plot the FILLED data for a continuous line
    inst_line, = ax1.plot(df_clean['Datetime'], df_clean[usage_column], color=color_inst, linestyle=line_style, label=plot_mode_label)
    
    # 2. Highlight points that were originally valid
    ok_df = df_clean[df_clean['Flag'].str.startswith('OK')]
    inst_scatter = ax1.scatter(ok_df['Datetime'], ok_df['kWh'], color='green', marker='o', s=20, zorder=5, label='Aggregated Valid Readings')

    # 3. Highlight points that were zero-filled (Red 'x' markers)
    filled_flags = ['True Gap (Zero-Filled)', 'Invalid/Outlier (Zero-Filled)']
    filled_points_df = df_clean[df_clean['Flag'].isin(filled_flags)]
    scatter_label_part = 'Zero-Filled'
    
    interp_scatter = ax1.scatter(filled_points_df['Datetime'], filled_points_df['Interpolated_kWh'], 
               color='red', marker='x', s=50, zorder=5, label=f'{scatter_label_part} Points')
        
    ax1.tick_params(axis='y', labelcolor=color_inst)
    
    
    # --- AXIS 2: Cumulative Total (Secondary) ---
    color_cum = 'tab:red'
    ax2 = ax1.twinx() 
    ax2.set_ylabel(f"Cumulative Total {unit}", color=color_cum)  
    
    # Plot the cumulative total line
    cum_line, = ax2.plot(df_clean['Datetime'], df_clean['Cumulative_kWh'], color=color_cum, linestyle='--', linewidth=1.5, label=f'Cumulative Total {unit}') 
    
    ax2.tick_params(axis='y', labelcolor=color_cum)

    # --- Final Touches ---
    ax1.set_title(f"Energy Usage per {period_label} and Cumulative Total Over Time")
    plt.xticks(rotation=45, ha='right')
    
    lines = [inst_line, inst_scatter, interp_scatter, cum_line]
    labels = [line.get_label() for line in lines]
    
    ax1.legend(lines, labels, loc='upper left', fontsize='small')

    plt.tight_layout()
    st.pyplot(fig)
    
    # ---------- Download Excel (DATA ONLY, NO CHART) ----------
    st.markdown("---")
    output = BytesIO()
    
    # --- PREPARE DATA FOR EXCEL EXPORT ---
    # Select only the relevant columns for export: Datetime, the zero-filled usage, cumulative total, and flag
    df_excel = df_clean_display[[
        'Datetime', 
        f'Zero_Filled_Usage_{unit}', 
        f'Cumulative_{unit}', 
        'Flag'
    ]].copy()
    
    # Rename for clearer column headers in Excel
    usage_col_name = f'Usage_Per_{period_label.replace(" ", "_").replace("-", "")}_{unit}'
    cumulative_col_name = f'Cumulative_Total_{unit}'
    
    df_excel.rename(columns={
        f'Zero_Filled_Usage_{unit}': usage_col_name,
        f'Cumulative_{unit}': cumulative_col_name
    }, inplace=True)
    
    
    # Use standard Excel writer
    excel_engine = "openpyxl"
    if 'xlsxwriter' in globals() and xlsxwriter is not None:
        excel_engine = "xlsxwriter"
        
    try:
        with pd.ExcelWriter(output, engine=excel_engine) as writer:
            sheet_name = "Cleaned_Data"
            
            # Write Data to Sheet (starting at row 0)
            df_excel.to_excel(writer, index=False, sheet_name=sheet_name, startrow=0)
            
            # NOTE: No chart generation code is included here, ensuring a data-only export.
            
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
