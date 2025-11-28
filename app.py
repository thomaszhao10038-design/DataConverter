import streamlit as st
import pandas as pd
from io import BytesIO

# --- Configuration ---
st.set_page_config(
    page_title="Data Converter",
    layout="centered"
)

def process_data_file(uploaded_file, name):
    """Loads, processes, and prepares data from one MSB file."""
    try:
        # Load the CSV file
        df = pd.read_csv(uploaded_file)
    except Exception as e:
        st.error(f"Error reading {name} file: {e}")
        return None, None

    # Rename the PSum column and calculate kW
    df = df.rename(columns={'PSum (W)': 'Active Power (W)'})

    # Convert 'Date & Time' to datetime objects (handling the DD/MM/YYYY format)
    df['Date & Time'] = pd.to_datetime(df['Date & Time'], format='%d/%m/%Y %H:%M:%S', errors='coerce')
    
    # Drop rows where datetime parsing failed
    df = df.dropna(subset=['Date & Time'])

    if df.empty:
        st.warning(f"{name} has no valid data after processing 'Date & Time'.")
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

def create_excel_file(df1, date1, df2, date2):
    """Merges the two processed DataFrames and creates the Excel file in Book1 format."""
    
    # Merge the two DataFrames on 'Local Time Stamp' (index)
    merged_df = df1.merge(df2, left_index=True, right_index=True, how='outer', 
                          suffixes=(f'_{date1}', f'_{date2}'))

    # --- Create Multi-Level Header Structure ---
    
    # 1. Define the multi-level headers (as seen in Book1 sample)
    headers = [
        ('UTC Offset (minutes)', ''), 
        (date1, 'Local Time Stamp'),
        (date1, 'Active Power (W)'), 
        (date1, 'kW'),
        ('', ''), # Blank Separator Column
        (date2, 'Local Time Stamp'),
        (date2, 'Active Power (W)'), 
        (date2, 'kW')
    ]
    
    multi_index = pd.MultiIndex.from_tuples(headers)
    output_df = pd.DataFrame(columns=multi_index)

    # 2. Populate the DataFrame
    # Iterate over the merged data (index is the 'Local Time Stamp' time)
    for index, row in merged_df.iterrows():
        new_row = [
            '',  # UTC Offset (minutes)
            index,  # Local Time Stamp (1)
            row.get(f'Active Power (W)_{date1}', ''), # Use .get() for merged data
            row.get(f'kW_{date1}', ''),
            '',  # Separator
            index,  # Local Time Stamp (2)
            row.get(f'Active Power (W)_{date2}', ''),
            row.get(f'kW_{date2}', '')
        ]
        output_df.loc[len(output_df)] = new_row
    
    
    # 3. Create the Excel file in memory
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Note: index=False to prevent writing a default index
        output_df.to_excel(writer, sheet_name='Sheet1', index=False, startrow=0)
    
    return output.getvalue()


# --- Streamlit UI ---
st.title("💡 Data Converter (Book1 Format)")
st.markdown("Upload your two **'raw data MSB'** files to merge and format them into the Book1 side-by-side structure.")

# File Uploaders
st.subheader("1. Upload Files")
col1, col2 = st.columns(2)

with col1:
    uploaded_file_msb1 = st.file_uploader(
        "Upload Consolidated Data MSB 1 CSV",
        type=['csv'],
        key='msb1'
    )

with col2:
    uploaded_file_msb2 = st.file_uploader(
        "Upload Consolidated Data MSB 2 CSV",
        type=['csv'],
        key='msb2'
    )

# Processing Logic
if uploaded_file_msb1 and uploaded_file_msb2:
    st.subheader("2. Processing Data")
    
    # Process both files
    df1, date1 = process_data_file(uploaded_file_msb1, "MSB 1")
    df2, date2 = process_data_file(uploaded_file_msb2, "MSB 2")
    
    if df1 is not None and df2 is not None:
        st.success(f"MSB 1 Data Date identified: **{date1}**")
        st.success(f"MSB 2 Data Date identified: **{date2}**")
        
        try:
            with st.spinner('Generating Excel file...'):
                excel_data = create_excel_file(df1, date1, df2, date2)
            
            st.subheader("3. Download Result")
            st.download_button(
                label="Download Formatted Excel File",
                data=excel_data,
                file_name=f'Consolidated_Data_{date1}_vs_{date2}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
            st.markdown("Your Excel file is ready!")
            
        except Exception as e:
            st.error(f"An error occurred during Excel file creation. Please check your data format: {e}")

elif uploaded_file_msb1 or uploaded_file_msb2:
    st.warning("Please upload **both** MSB 1 and MSB 2 data files to proceed.")
