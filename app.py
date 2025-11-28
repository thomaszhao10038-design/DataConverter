import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side, numbers
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.series import Series  # Crucial import

# --- Configuration ---
POWER_COL_OUT = 'PSumW'

# -----------------------------
# PROCESS SINGLE SHEET
# -----------------------------
def process_sheet(df, timestamp_col, psum_col):
    df.columns = df.columns.astype(str).str.strip()
    df[timestamp_col] = pd.to_datetime(df[timestamp_col], errors="coerce", dayfirst=True)
    power_series = df[psum_col].astype(str).str.strip().str.replace(',', '.', regex=False)
    df[psum_col] = pd.to_numeric(power_series, errors='coerce')
    df = df.dropna(subset=[timestamp_col, psum_col])

    if df.empty:
        return pd.DataFrame()

    df_indexed = df.set_index(timestamp_col)
    df_indexed.index = df_indexed.index.floor('10min')
    resampled_data = df_indexed[psum_col].groupby(level=0).sum()
    df_out = resampled_data.reset_index()
    df_out.columns = ['Rounded', POWER_COL_OUT]

    if df_out.empty or df_out['Rounded'].isna().all():
        return pd.DataFrame()

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

    df_padded_series = df_out.set_index('Rounded')[POWER_COL_OUT].reindex(full_time_index, fill_value=0)
    grouped = df_padded_series.reset_index().rename(columns={'index': 'Rounded'})
    grouped.columns = ['Rounded', POWER_COL_OUT]
    grouped["Date"] = grouped["Rounded"].dt.date
    grouped["Time"] = grouped["Rounded"].dt.strftime("%H:%M")
    grouped = grouped[grouped["Date"].isin(original_dates)]
    return grouped


# -----------------------------
# BUILD EXCEL WITH CHART (FIXED SERIES CREATION)
# -----------------------------
def build_output_excel(sheets_dict):
    wb = Workbook()
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])

    # Styles
    header_fill = PatternFill(start_color='ADD8E6', end_color='ADD8E6', fill_type='solid')
    title_font = Font(bold=True, size=12)
    header_font = Font(bold=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))
    data_fill_alt = PatternFill(start_color='F0F8FF', end_color='F0F8FF', fill_type='solid')

    for sheet_name, df in sheets_dict.items():
        ws = wb.create_sheet(sheet_name)
        dates = sorted(df["Date"].unique())
        col_start = 1
        daily_max_summary = []
        max_row_used = 0

        chart_categories_ref = None
        chart_series_list = []   # Will hold proper Series objects

        for date in dates:
            date_str_short = date.strftime('%d-%b')
            date_str_full = date.strftime('%Y-%m-%d')
            day_data = df[df["Date"] == date].sort_values("Time")
            data_rows_count = len(day_data)
            merge_start_row = 3
            merge_end_row = 2 + data_rows_count

            # Header (merged)
            ws.merge_cells(start_row=1, start_column=col_start, end_row=1, end_column=col_start + 3)
            ws.cell(row=1, column=col_start, value=date_str_full).alignment = Alignment(horizontal="center", vertical="center")

            # Sub-headers
            ws.cell(row=2, column=col_start, value="UTC Offset (minutes)")
            ws.cell(row=2, column=col_start + 1, value="Local Time Stamp")
            ws.cell(row=2, column=col_start + 2, value="Active Power (W)")
            ws.cell(row=2, column=col_start + 3, value="kW")

            # Merge UTC column
            if data_rows_count > 0:
                ws.merge_cells(start_row=merge_start_row, start_column=col_start,
                               end_row=merge_end_row, end_column=col_start)
                ws.cell(row=merge_start_row, column=col_start, value=date_str_full).alignment = Alignment(horizontal="center", vertical="center")

            # Write data rows
            for idx, r in enumerate(day_data.itertuples(), start=3):
                ws.cell(row=idx, column=col_start + 1, value=r.Time)
                power_w = getattr(r, POWER_COL_OUT)
                ws.cell(row=idx, column=col_start + 2, value=power_w)
                ws.cell(row=idx, column=col_start + 3, value=abs(power_w) / 1000)

            # Summary stats
            if data_rows_count > 0:
                sum_w = day_data[POWER_COL_OUT].sum()
                mean_w = day_data[POWER_COL_OUT].mean()
                max_w = day_data[POWER_COL_OUT].max()
                max_kw_abs = day_data[POWER_COL_OUT].abs().max() / 1000

                stats_row_start = merge_end_row + 1
                # Total
                ws.cell(row=stats_row_start, column=col_start + 1, value="Total")
                ws.cell(row=stats_row_start, column=col_start + 3, value=day_data[POWER_COL_OUT].abs().sum() / 1000)
                # Average
                ws.cell(row=stats_row_start + 1, column=col_start + 1, value="Average")
                ws.cell(row=stats_row_start + 1, column=col_start + 3, value=day_data[POWER_COL_OUT].abs().mean() / 1000)
                # Max
                ws.cell(row=stats_row_start + 2, column=col_start + 1, value="Max")
                ws.cell(row=stats_row_start + 2, column=col_start + 3, value=max_kw_abs)

                max_row_used = max(max_row_used, stats_row_start + 2)
                daily_max_summary.append((date_str_short, max_kw_abs))

            # === CHART DATA COLLECTION (FIXED) ===
            if chart_categories_ref is None:
                chart_categories_ref = Reference(ws,
                                                 min_col=col_start + 1,
                                                 min_row=merge_start_row,
                                                 max_row=merge_end_row)

            data_ref = Reference(ws,
                                 min_col=col_start + 3,
                                 min_row=merge_start_row,
                                 max_row=merge_end_row)

            title_ref = Reference(ws, min_col=col_start, min_row=1, max_col=col_start + 3)

            # Correct way to create a Series (no manual idx in constructor)
            series = Series(data_ref)              # ← Fixed: only pass values
            series.title_from_data = False
            series.title = title_ref
            series.idx = len(chart_series_list)    # ← Set index safely after creation
            chart_series_list.append(series)

            col_start += 4

        # === ADD LINE CHART ===
        if chart_series_list and chart_categories_ref:
            chart = LineChart()
            chart.title = f"Daily 10-Minute Absolute Power Profile - {sheet_name}"
            chart.style = 10
            chart.y_axis.title = "Power (kW)"
            chart.x_axis.title = "Time of Day"

            for ser in chart_series_list:
                chart.append(ser)                  # append Series objects

            chart.set_categories(chart_categories_ref)
            ws.add_chart(chart, f"G{max_row_used + 5}")
            max_row_used += 25  # leave space for chart

        # === FINAL SUMMARY TABLE ===
        if daily_max_summary:
            summary_row = max_row_used + 4
            title_cell = ws.cell(row=summary_row, column=1, value="Daily Max Power (kW) Summary")
            ws.merge_cells(start_row=summary_row, start_column=1, end_row=summary_row, end_column=2)
            title_cell.font = title_font
            title_cell.alignment = Alignment(horizontal="center", vertical="center")

            summary_row += 1
            ws.cell(row=summary_row, column=1, value="Day").fill = header_fill
            ws.cell(row=summary_row, column=2, value="Max (kW)").fill = header_fill

            for i, (day_str, max_kw) in enumerate(daily_max_summary):
                row = summary_row + 1 + i
                fill = data_fill_alt if i % 2 == 0 else PatternFill(fill_type=None)
                ws.cell(row=row, column=1, value=day_str).fill = fill
                cell = ws.cell(row=row, column=2, value=round(max_kw, 2))
                cell.number_format = "0.00"
                cell.fill = fill

            ws.column_dimensions['A'].width = 15
            ws.column_dimensions['B'].width = 15

    # Save to BytesIO
    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    return stream


# -----------------------------
# STREAMLIT APP
# -----------------------------
def app():
    st.title("Excel 10-Minute Electricity Data Converter")
    st.markdown("""
    Upload an **Excel file (.xlsx)** → get 10-minute summed power data with daily profiles, 
    a line chart per sheet, and a max kW summary table.
    """)

    uploaded = st.file_uploader("Upload .xlsx file", type=["xlsx"])

    if uploaded:
        xls = pd.ExcelFile(uploaded)
        result_sheets = {}

        for sheet_name in xls.sheet_names:
            st.info(f"Processing sheet: **{sheet_name}**")
            try:
                df = pd.read_excel(uploaded, sheet_name=sheet_name)
            except Exception as e:
                st.error(f"Cannot read sheet '{sheet_name}': {e}")
                continue

            df.columns = df.columns.astype(str).str.strip()

            time_cols = ["Date & Time", "Date&Time", "Timestamp", "DateTime", "Local Time", "TIME", "ts"]
            timestamp_col = next((c for c in df.columns if c in time_cols), None)
            if not timestamp_col:
                st.error(f"No timestamp column found in '{sheet_name}'")
                continue

            power_cols = ["PSum (W)", "Psum (W)", "PSum", "P (W)", "Power"]
            psum_col = next((c for c in df.columns if c in power_cols), None)
            if not psum_col:
                st.error(f"No power column found in '{sheet_name}'")
                continue

            processed = process_sheet(df, timestamp_col, psum_col)
            if not processed.empty:
                result_sheets[sheet_name] = processed
                st.success(f"Sheet **{sheet_name}** → {len(processed['Date'].unique())} days")
            else:
                st.warning(f"Sheet **{sheet_name}** had no valid data")

        if result_sheets:
            output = build_output_excel(result_sheets)
            st.success("All sheets processed!")
            st.download_button(
                label="Download Converted Excel",
                data=output,
                file_name="10min_Power_Converted.xlsx",
                mime="application/vnd.openpyxl"
            )
        else:
            st.error("No sheets could be processed.")

if __name__ == '__main__':
    app()
