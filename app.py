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
    # Note: Renaming for consistency, but the original PSum (W) is used for kW calculation
    df = df.rename(columns={'PSum (W)': 'Active Power (W)'})

    # Convert 'Date & Time' to datetime objects (handling the DD/MM/YYYY HH:MM:SS format)
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
    # Using absolute value based on the previous implementation for 'Book1' format
    df['kW'] = (df['Active Power (W)'] / 1000).abs()

    # Select columns and set 'Local Time Stamp' as index for the merge
    # The columns are suffixed with a unique number (index) during merging in the main function
    df_processed = df[['Local Time Stamp', 'Active Power (W)', 'kW']]
    df_processed = df_processed.set_index('Local Time Stamp')

    return df_processed, date_str

def create_excel_file_multi(processed_sheets):
    """
    Merges all processed DataFrames on 'Local Time Stamp' and creates the final Excel file
    with a dynamic multi-level header structure.
    """
    if not processed_sheets:
        return None

    # 1. Merge all DataFrames based on the 'Local Time Stamp' index
    # Start with an empty DataFrame whose index will hold all unique timestamps
    merged_df = pd.DataFrame(index=pd.Index([], name='Local Time Stamp'))
    
    # Iterate through the dictionary of processed sheets and merge them sequentially
    for sheet_id, data in processed_sheets.items():
        df_to_merge = data['df']
        # The merge uses the unique sheet_id (integer 1, 2, 3...) as the suffix
        merged_df = merged_df.merge(df_to_merge, 
                                    left_index=True, 
                                    right_index=True, 
                                    how='outer', 
                                    suffixes=('_old', f'_{sheet_id}'))

    # Sort the index (Local Time Stamp) to ensure chronological order
    merged_df.sort_index(inplace=True)

    # 2. Prepare Data for Manual Excel Write
    
    # Define the headers and column data mapping dynamically
    headers_top = ['UTC Offset (minutes)']
    headers_bottom = ['']
    
    # Get the raw data from the merged DataFrame, including the timestamp index
    data_rows = merged_df.reset_index()
    data_rows = data_rows.rename(columns={'Local Time Stamp': 'Timestamp_Index'})
    
    # Column mapping for data retrieval
    column_map = ['Timestamp_Index'] # Start with the index column
    
    # Dynamically generate headers for each sheet
    for sheet_id, data in processed_sheets.items():
        sheet_name = data['name']
        date_str = data['date']
        
        # Add a blank separator column before the next sheet, unless it's the first sheet
        if sheet_id > 1:
            headers_top.append('') # Top separator
            headers_bottom.append('') # Bottom separator
            column_map.append(None) # No data column for separator

        # Add the three columns for the current sheet
        top_level = f'{sheet_name} ({date_str})'
        headers_top.extend([top_level, top_level, top_level])
        headers_bottom.extend(['Local Time Stamp', 'Active Power (W)', 'kW'])
        
        # Add column names from merged_df
        column_map.extend([
            'Timestamp_Index', # This column will be populated by the index (time)
            f'Active Power (W)_{sheet_id}', 
            f'kW_{sheet_id}'
        ])

    # 3. Create the Excel file in memory with Manual Header Write (Fix for MultiIndex error)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        workbook = writer.book
        sheet_name = 'Consolidated'
        worksheet = workbook.create_sheet(sheet_name)
        
        # Manually write the top row of the MultiIndex header
        worksheet.append(headers_top)
        
        # Manually write the bottom row of the MultiIndex header
        worksheet.append(headers_bottom)
        
        # Write the data rows
        # Loop through the rows of the data_rows DataFrame
        for index, row in data_rows.iterrows():
            new_row = []
            
            # Construct the row based on the defined column_map
            for col_name in column_map:
                if col_name is None:
                    # Separator column
                    new_row.append('')
                elif col_name == 'Timestamp_Index':
                    # Local Time Stamp column (repeated)
                    new_row.append(row[col_name])
                else:
                    # Active Power or kW column (use .get() to handle NaN/missing values gracefully)
                    value = row.get(col_name, '')
                    new_row.append(value if pd.notna(value) else '')

            worksheet.append(new_row)
        
        # Remove the default empty sheet created by openpyxl
        if 'Sheet' in workbook.sheetnames:
             workbook.remove(workbook['Sheet'])
    
    return output.getvalue()


# --- Streamlit UI ---
st.title("💡 Data Converter")
st.markdown("Upload a single Excel file. The application will process **all** sheets found inside and consolidate the data into one Excel file, with each sheet's data presented in side-by-side columns.")

# File Uploader
st.subheader("1. Upload Excel File")
uploaded_excel_file = st.file_uploader(
    "Upload Consolidated Data Excel File (.xlsx)",
    type=['xlsx'],
    key='excel_upload'
)

processed_sheets_data = {} # Dictionary to hold processed DataFrames and metadata

if uploaded_excel_file:
    # Read the sheet names from the uploaded file
    try:
        xls = pd.ExcelFile(uploaded_excel_file)
        sheet_names = xls.sheet_names
        st.info(f"Found {len(sheet_names)} sheet(s): {', '.join(sheet_names)}")
    except Exception as e:
        st.error(f"Error reading Excel file sheets: {e}")
        sheet_names = []
        xls = None

    if sheet_names and xls:
        st.subheader("2. Processing Sheets")
        
        # Use a list to store sheet data dictionaries {id, name, df, date}
        processed_sheets = {}
        successful_sheets = 0
        
        # Iterate over all sheets found
        for i, name in enumerate(sheet_names, 1):
            with st.spinner(f'Processing sheet "{name}" ({i}/{len(sheet_names)})...'):
                try:
                    df_sheet = xls.parse(name)
                    df_processed, date_str = process_data_sheet(df_sheet, name)
                    
                    if df_processed is not None:
                        # Store the processed DataFrame and metadata using 'i' as a unique ID/suffix
                        processed_sheets[i] = {'name': name, 'date': date_str, 'df': df_processed}
                        successful_sheets += 1
                    
                except Exception as e:
                    st.error(f"Skipping sheet '{name}' due to error: {e}")

        
        if successful_sheets >= 2:
            st.subheader("3. Consolidate and Generate")
            st.success(f"Successfully processed {successful_sheets} sheets. Ready to consolidate.")
            
            if st.button("Generate Consolidated Excel"):
                try:
                    with st.spinner('Generating final Excel file with side-by-side format...'):
                        # Pass the dictionary of all successfully processed sheets
                        excel_data = create_excel_file_multi(processed_sheets)
                    
                    st.subheader("4. Download Result")
                    st.download_button(
                        label="Download Formatted Excel File",
                        data=excel_data,
                        file_name=f'Consolidated_Data_{len(processed_sheets)}_Sheets.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )
                    st.markdown("Your Excel file is ready!")
                    
                except Exception as e:
                    # This fallback should ideally not be hit now, but kept for general safety
                    st.error(f"An unexpected error occurred during final Excel file creation: {e}")
        elif successful_sheets == 1:
            st.warning(f"Only 1 sheet ('{list(processed_sheets.values())[0]['name']}') was successfully processed. You need at least two sheets for the side-by-side comparison format.")
        else:
            st.error("No valid sheets were processed. Check if your sheets contain the required columns ('Date & Time' and 'PSum (W)').")

    elif uploaded_excel_file:
        st.warning("The uploaded file does not contain any readable sheets or is not a valid Excel file.")

elif not uploaded_excel_file:
    st.info("Please upload an Excel file to begin the processing.")
