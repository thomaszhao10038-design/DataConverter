import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side, numbers
from openpyxl.chart import LineChart, Reference

# --- Configuration ---
POWER_COL_OUT = 'PSumW'

# -----------------------------
# PROCESS SINGLE SHEET
# -----------------------------
def process_sheet(df, timestamp_col, psum_col):
    """
    Processes a single DataFrame sheet: cleans data, rounds timestamps to 10-minute intervals,
    calculates the sum of power, pads missing intervals, and filters out leading/trailing zeros.
    """
    df.columns = df.columns.astype(str).str.strip()
    df[timestamp_col] = pd.to_datetime(df[timestamp_col], errors="coerce", dayfirst=True)
    
    # Handle commas as decimal separators
    power_series = df[psum_col].astype(str).str.strip()
    power_series = power_series.str.replace(',', '.', regex=False)
    df[psum_col] = pd.to_numeric(power_series, errors='coerce')
    
    df = df.dropna(subset=[timestamp_col, psum_col])
    if df.empty:
        return pd.DataFrame()
    
    # 1. Filter out leading and trailing zero power data
    non_zero_indices = df[df[psum_col].abs() != 0].index
    if non_zero_indices.empty:
        return pd.DataFrame()
        
    first_valid_idx = non_zero_indices.min()
    last_valid_idx = non_zero_indices.max()
    df = df.loc[first_valid_idx:last_valid_idx].copy()

    # 2. Resample to 10-minute intervals and sum power
    df_indexed = df.set_index(timestamp_col)
    df_indexed.index = df_indexed.index.floor('10min')
    resampled_data = df_indexed[psum_col].groupby(level=0).sum()
    
    df_out = resampled_data.reset_index()
    df_out.columns = ['Rounded', POWER_COL_OUT] 
    if df_out.empty or df_out['Rounded'].isna().all():
        return pd.DataFrame()
        
    # 3. Create a full 10-minute time index for padding
    original_dates = set(df_out['Rounded'].dt.date)
    min_dt = df_out['Rounded'].min().floor('D')
    max_dt_exclusive = df_out['Rounded'].max().ceil('D') 
    if min_dt >= max_dt_exclusive:
        return pd.DataFrame()
        
    full_time_index = pd.date_range(
        start=min_dt.to_pydatetime(),
        end=max_dt_exclusive.to_pydatetime(),
        freq='10min',
        inclusive='left'
    )
    
    # 4. Reindex to the full index (introducing NaNs for missing intervals)
    df_indexed_for_reindex = df_out.set_index('Rounded')
    df_padded_series = df_indexed_for_reindex[POWER_COL_OUT].reindex(full_time_index)
    
    # 5. Final formatting and kW calculation
    grouped = df_padded_series.reset_index().rename(columns={'index': 'Rounded'})
    grouped.columns = ['Rounded', POWER_COL_OUT]
    grouped[POWER_COL_OUT] = grouped[POWER_COL_OUT].astype(float) 
    grouped["Date"] = grouped["Rounded"].dt.date
    grouped["Time"] = grouped["Rounded"].dt.strftime("%H:%M") 
    
    # Filter back to only include dates that had non-zero activity
    grouped = grouped[grouped["Date"].isin(original_dates)]
    
    # Calculate kW (absolute value)
    grouped['kW'] = grouped[POWER_COL_OUT].abs() / 1000
    
    return grouped

# -----------------------------
# BUILD EXCEL
# -----------------------------
def build_output_excel(sheets_dict):
    """
    Creates the output Excel workbook with processed data sheets and a summary "Total" sheet.
    """
    wb = Workbook()
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])

    # Define styles (optional but good practice)
    # header_fill = PatternFill(start_color='ADD8E6', end_color='ADD8E6', fill_type='solid')
    title_font = Font(bold=True, size=12)
    # thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
    #                      top=Side(style='thin'), bottom=Side(style='thin'))

    # Dictionary to hold daily max power per sheet for Total sheet
    total_data = {}

    for sheet_name, df in sheets_dict.items():
        ws = wb.create_sheet(sheet_name)
        dates = sorted(df["Date"].unique())
        col_start = 1
        max_row_used = 0
        daily_max_summary = []
        day_intervals = []

        for date in dates:
            day_data_full = df[df["Date"] == date].sort_values("Time")
            day_data_active = day_data_full.dropna(subset=[POWER_COL_OUT])
            
            n_rows = len(day_data_full)
            day_intervals.append(n_rows)
            merge_start = 3
            merge_end = merge_start + n_rows - 1

            date_str_full = date.strftime('%Y-%m-%d')
            date_str_short = date.strftime('%d-%b')

            # Merge date header
            ws.merge_cells(start_row=1, start_column=col_start, end_row=1, end_column=col_start+3)
            ws.cell(row=1, column=col_start, value=date_str_full).alignment = Alignment(horizontal="center")

            # Sub-headers
            ws.cell(row=2, column=col_start, value="UTC Offset (minutes)")
            ws.cell(row=2, column=col_start+1, value="Local Time Stamp")
            ws.cell(row=2, column=col_start+2, value="Active Power (W)")
            ws.cell(row=2, column=col_start+3, value="kW")

            # Merge UTC column
            ws.merge_cells(start_row=merge_start, start_column=col_start, end_row=merge_end, end_column=col_start)
            ws.cell(row=merge_start, column=col_start, value=date_str_full).alignment = Alignment(horizontal="center", vertical="center")

            # Fill data
            for idx, r in enumerate(day_data_full.itertuples(), start=merge_start):
                ws.cell(row=idx, column=col_start+1, value=r.Time)
                ws.cell(row=idx, column=col_start+2, value=getattr(r, POWER_COL_OUT))
                ws.cell(row=idx, column=col_start+3, value=r.kW)

            # Summary stats
            stats_row_start = merge_end + 1
            sum_w = day_data_active[POWER_COL_OUT].sum()
            mean_w = day_data_active[POWER_COL_OUT].mean()
            max_w = day_data_active[POWER_COL_OUT].max()
            sum_kw = day_data_active['kW'].sum()
            mean_kw = day_data_active['kW'].mean()
            max_kw = day_data_active['kW'].max()

            ws.cell(row=stats_row_start, column=col_start+1, value="Total")
            ws.cell(row=stats_row_start, column=col_start+2, value=sum_w)
            ws.cell(row=stats_row_start, column=col_start+3, value=sum_kw)
            ws.cell(row=stats_row_start+1, column=col_start+1, value="Average")
            ws.cell(row=stats_row_start+1, column=col_start+2, value=mean_w)
            ws.cell(row=stats_row_start+1, column=col_start+3, value=mean_kw)
            ws.cell(row=stats_row_start+2, column=col_start+1, value="Max")
            ws.cell(row=stats_row_start+2, column=col_start+2, value=max_w)
            ws.cell(row=stats_row_start+2, column=col_start+3, value=max_kw)

            max_row_used = max(max_row_used, stats_row_start+2)
            
            # Save max kw for Total sheet
            daily_max_summary.append((date_str_short, max_kw))
            total_data.setdefault(date_str_short, {})[sheet_name] = max_kw

            col_start += 4

        # Line chart for individual sheet
        if dates:
            chart = LineChart()
            chart.title = f"{sheet_name} - Power Consumption"
            chart.y_axis.title = "Power (kW)"
            chart.x_axis.title = "Time"
            max_rows = max(day_intervals) if day_intervals else 0
            if max_rows > 0:
                categories_ref = Reference(ws, min_col=2, min_row=3, max_row=2+max_rows)
                col_start = 1
                for n_rows in day_intervals:
                    data_ref = Reference(ws, min_col=col_start+3, min_row=3, max_row=2+n_rows)
                    series = chart.add_data(data_ref, titles_from_data=False)
                    # Add titles based on date
                    # series.title = ws.cell(row=1, column=col_start).value # Optional: Use date as title, but looks crowded
                    col_start += 4
                chart.set_categories(categories_ref)
                ws.add_chart(chart, f'G{max_row_used+2}')

    # -----------------------------
    # Add Total Sheet
    # -----------------------------
    ws_total = wb.create_sheet("Total")
    all_dates = sorted(total_data.keys())
    sheet_names = list(sheets_dict.keys())
    num_sheets = len(sheet_names)

    # Headers
    ws_total.cell(row=1, column=1, value="Date").font = title_font
    for i, sheet_name in enumerate(sheet_names):
        ws_total.cell(row=1, column=2+i, value=sheet_name).font = title_font
    ws_total.cell(row=1, column=2+num_sheets, value="Total Load").font = title_font

    # Fill data
    for r_idx, date_val in enumerate(all_dates, start=2):
        ws_total.cell(row=r_idx, column=1, value=date_val)
        total_load = 0
        for c_idx, sheet_name in enumerate(sheet_names, start=2):
            value = total_data[date_val].get(sheet_name, 0)
            ws_total.cell(row=r_idx, column=c_idx, value=value)
            total_load += value
        
        # Total Load column
        ws_total.cell(row=r_idx, column=2+num_sheets, value=total_load)

    # Format columns width
    for col in range(1, 3+num_sheets):
        ws_total.column_dimensions[chr(64+col)].width = 15

    # --- ADD LINE CHART TO TOTAL SHEET (New Addition) ---
    if all_dates:
        total_chart = LineChart()
        total_chart.title = "Daily Max Total Load Summary"
        total_chart.y_axis.title = "Total Max Load (kW)"
        total_chart.x_axis.title = "Date"
        
        # Calculate max row index (1-based)
        max_row_total = len(all_dates) + 1
        
        # Data series: Total Load (kW) - Located in the last data column
        data_col = 2 + num_sheets
        data = Reference(ws_total,
                         min_col=data_col,
                         min_row=2,
                         max_row=max_row_total)
        
        # Categories: Dates - Located in column 1
        cats = Reference(ws_total,
                         min_col=1, 
                         min_row=2,
                         max_row=max_row_total)

        total_chart.add_data(data, titles_from_data=False)
        total_chart.set_categories(cats)
        
        # Place chart below the table
        ws_total.add_chart(total_chart, f'A{max_row_total + 3}') 

    # Save workbook to BytesIO
    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    return stream

# -----------------------------
# STREAMLIT APP
# -----------------------------
def app():
    st.set_page_config(layout="wide", page_title="Electricity Data Converter")
    st.title("📊 Excel 10-Minute Electricity Data Converter")
    st.markdown("""
        Upload an **Excel file (.xlsx)** with time-series data. Each sheet is processed to calculate total absolute power (W) in 10-minute intervals. 
        
        Leading and trailing zero values (representing missing readings) are filtered out and appear blank, but zero values *within* the active recording period are kept.
        
        The output Excel file includes a **line chart**, a **Max Power Summary table**, and a **Total sheet** with daily max power for all sheets and an **aggregated chart**.
    """)

    uploaded = st.file_uploader("Upload .xlsx file", type=["xlsx"])
    if uploaded:
        xls = pd.ExcelFile(uploaded)
        result_sheets = {}
        st.write("---")

        for sheet_name in xls.sheet_names:
            st.markdown(f"**Processing sheet: `{sheet_name}`**")
            try:
                # Use engine='openpyxl' for compatibility with modern Excel files if issues arise
                df = pd.read_excel(uploaded, sheet_name=sheet_name)
            except Exception as e:
                st.error(f"Error reading sheet '{sheet_name}': {e}")
                continue

            df.columns = df.columns.astype(str).str.strip()
            # Dynamic column detection
            timestamp_col = next((c for c in df.columns if c in ["Date & Time","Date&Time","Timestamp","DateTime","Local Time","TIME","ts","date"]), None)
            psum_col = next((c for c in df.columns if c in ["PSum (W)","Psum (W)","PSum","P (W)","Power","W"]), None)
            
            if not timestamp_col or not psum_col:
                st.error(f"Sheet '{sheet_name}' missing required columns. Looked for Timestamp in: `Date & Time, Timestamp, DateTime, Local Time, TIME, ts, date`. Looked for Power in: `PSum (W), Psum (W), PSum, P (W), Power, W`.")
                continue

            processed = process_sheet(df, timestamp_col, psum_col)
            if not processed.empty:
                result_sheets[sheet_name] = processed
                st.success(f"Sheet '{sheet_name}' processed successfully with {len(processed['Date'].unique())} day(s) of data.")
            else:
                st.warning(f"Sheet '{sheet_name}' had no usable data.")

        if result_sheets:
            output_stream = build_output_excel(result_sheets)
            st.download_button(
                label="📥 Download Converted Excel (Converted_Output.xlsx)",
                data=output_stream,
                file_name="Converted_Output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    app()
