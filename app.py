import streamlit as st
import pandas as pd
from io import BytesIO

# --- Configuration ---
st.set_page_config(
    page_title="Data Converter",
    layout="centered"
)

def process_data_sheet(df, sheet_name):
    """Loads, processes, and prepares data from one sheet."""
    # Ensure DataFrame is not empty
    if df.empty:
        st.warning(f"Sheet '{sheet_name}' is empty.")
        return None, None

    # Assume the required columns are present: 'Date & Time' and 'PSum (W)'
    required_cols = ['Date & Time', 'PSum (W)']
    if not all(col in df.columns for col in required_cols):
        st.error(f"Sheet '{sheet_name}' is missing required columns. Found: {list(df.columns)}. Expected: {required_cols}")
        return None, None

    # Rename the PSum column and calculate kW
    df = df.rename(columns={'PSum (W)': 'Active Power (W)'})

    # Convert 'Date & Time' to datetime objects (handling the DD/MM/YYYY HH:MM:SS format)
    # Using 'coerce' turns unparseable dates into NaT (Not a Time)
    df['Date & Time'] = pd.to_datetime(df['Date & Time'], format='%d/%m/%Y %H:%M:%S', errors='coerce')
    
    # Drop rows where datetime parsing failed
    df = df.dropna(subset=['Date & Time'])

    if df.empty:
        st.warning(f"Sheet '{sheet_name}' has no valid data after processing 'Date & Time'.")
        return None, None

    # Extract Date (for header) and Time (for merging)
    # Date will be YYYY-MM-DD for consistency in the header
    date_str = df['Date & Time'].dt.strftime('%Y-%m-%d').iloc[0] 
    df['Local Time Stamp'] = df['Date & Time'].dt.strftime('%H:%M:%S')

    # Calculate kW: abs(Active Power (W)) / 1000
    df['kW'] = (df['Active Power (W)'] / 1000).abs()

    # Select columns and set 'Local Time Stamp' as index for the merge
    df_processed = df[['Local Time Stamp', 'Active Power (W)', 'kW']]
    df_processed = df_processed.set_index('Local Time Stamp')

    return df_processed, date_str

def create_excel_file(df1, date1, sheet_name1, df2, date2, sheet_name2):
    """Merges the two processed DataFrames and creates the Excel file in Book1 format."""
    
    # Merge the two DataFrames on 'Local Time Stamp' (index)
    # Using 'outer' to include all time stamps present in either sheet
    merged_df = df1.merge(df2, left_index=True, right_index=True, how='outer', 
                          suffixes=('_1', '_2'))

    # --- Create Multi-Level Header Structure ---
    
    # 1. Define the multi-level headers (as seen in Book1 sample)
    # The top level uses the sheet name and data date
    headers = [
        ('UTC Offset (minutes)', ''), 
        (f'{sheet_name1} ({date1})', 'Local Time Stamp'),
        (f'{sheet_name1} ({date1})', 'Active Power (W)'), 
        (f'{sheet_name1} ({date1})', 'kW'),
        ('', ''), # Blank Separator Column
        (f'{sheet_name2} ({date2})', 'Local Time Stamp'),
        (f'{sheet_name2} ({date2})', 'Active Power (W)'), 
        (f'{sheet_name2} ({date2})', 'kW')
    ]
    
    multi_index = pd.MultiIndex.from_tuples(headers)
    output_df = pd.DataFrame(columns=multi_index)

    # 2. Populate the DataFrame
    # Note: Column names in merged_df are 'Active Power (W)_1', 'kW_1', 'Active Power (W)_2', 'kW_2'
    for index, row in merged_df.iterrows():
        new_row = [
            '',  # UTC Offset (minutes)
            index,  # Local Time Stamp (1)
            row.get('Active Power (W)_1', ''), # Use .get() and default to '' if missing
            row.get('kW_1', ''),
            '',  # Separator
            index,  # Local Time Stamp (2)
            row.get('Active Power (W)_2', ''),
            row.get('kW_2', '')
        ]
        output_df.loc[len(output_df)] = new_row
    
    
    # 3. Create the Excel file in memory
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Note: index=False to prevent writing a default index
        output_df.to_excel(writer, sheet_name='Consolidated', index=False, startrow=0)
    
    return output.getvalue()


# --- Streamlit UI ---
st.title("💡 Data Converter")
st.markdown("Upload a single Excel file containing multiple data sheets. You will then select two sheets to consolidate and format into the required side-by-side structure.")

# File Uploader
st.subheader("1. Upload Excel File")
uploaded_excel_file = st.file_uploader(
    "Upload Consolidated Data Excel File (.xlsx)",
    type=['xlsx'],
    key='excel_upload'
)

sheet_names = []
if uploaded_excel_file:
    # Read the sheet names from the uploaded file
    try:
        xls = pd.ExcelFile(uploaded_excel_file)
        sheet_names = xls.sheet_names
    except Exception as e:
        st.error(f"Error reading Excel file sheets: {e}")
        sheet_names = []

    if sheet_names:
        st.subheader("2. Select Sheets to Compare")
        
        # Select the two sheets using two columns for clarity
        col1, col2 = st.columns(2)
        
        with col1:
            sheet_name_1 = st.selectbox(
                "Select Sheet 1 (Left Column)",
                options=sheet_names,
                key='sheet1_select'
            )

        with col2:
            sheet_name_2 = st.selectbox(
                "Select Sheet 2 (Right Column)",
                options=[name for name in sheet_names if name != sheet_name_1], # Exclude the first selected sheet
                key='sheet2_select'
            )

        if sheet_name_1 and sheet_name_2:
            st.subheader("3. Consolidate and Generate")
            
            if st.button("Generate Consolidated Excel"):
                # Load the two selected sheets into DataFrames
                try:
                    df_sheet1 = xls.parse(sheet_name_1)
                    df_sheet2 = xls.parse(sheet_name_2)
                except Exception as e:
                    st.error(f"Error parsing selected sheets: {e}")
                    st.stop()
                
                # --- Processing Logic ---
                with st.spinner(f'Processing sheets "{sheet_name_1}" and "{sheet_name_2}"...'):
                    # Process both DataFrames
                    df1, date1 = process_data_sheet(df_sheet1, sheet_name_1)
                    df2, date2 = process_data_sheet(df_sheet2, sheet_name_2)
                    
                    if df1 is not None and df2 is not None:
                        st.success(f"Processing complete: {sheet_name_1} ({date1}) vs {sheet_name_2} ({date2})")
                        
                        try:
                            # Create the final Excel file
                            excel_data = create_excel_file(df1, date1, sheet_name_1, df2, date2, sheet_name_2)
                            
                            st.subheader("4. Download Result")
                            st.download_button(
                                label="Download Formatted Excel File",
                                data=excel_data,
                                file_name=f'Consolidated_Data_{sheet_name_1}_vs_{sheet_name_2}.xlsx',
                                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                            )
                            st.markdown("Your Excel file is ready!")
                            
                        except Exception as e:
                            st.error(f"An error occurred during Excel file creation: {e}")

    elif uploaded_excel_file:
        st.warning("The uploaded file does not contain any readable sheets or is not a valid Excel file.")

elif not uploaded_excel_file:
    st.info("Please upload an Excel file to see the sheet selection options.")
