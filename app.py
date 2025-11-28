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
    df.columns = df.columns.astype(str).str.strip()
    df[timestamp_col] = pd.to_datetime(df[timestamp_col], errors="coerce", dayfirst=True)

    power_series = df[psum_col].astype(str).str.strip()
    power_series = power_series.str.replace(',', '.', regex=False)
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

    df_indexed_for_reindex = df_out.set_index('Rounded')
    df_padded_series = df_indexed_for_reindex[POWER_COL_OUT].reindex(full_time_index, fill_value=0)
    grouped = df_padded_series.reset_index().rename(columns={'index': 'Rounded'})
    grouped.columns = ['Rounded', POWER_COL_OUT]

    grouped["Date"] = grouped["Rounded"].dt.date
    grouped["Time"] = grouped["Rounded"].dt.strftime("%H:%M")
    grouped = grouped[grouped["Date"].isin(original_dates)]
    return grouped

# -----------------------------
# BUILD EXCEL FORMAT
# -----------------------------
def build_output_excel(sheets_dict):
    wb = Workbook()
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])

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
        chart_series_data = []

        # Track max intervals per day separately
        day_intervals = {}

        for date in dates:
            date_str_short = date.strftime('%d-%b')
            date_str_full = date.strftime('%Y-%m-%d')
            day_data = df[df["Date"] == date].sort_values("Time")
            data_rows_count = len(day_data)
            merge_start_row = 3
            merge_end_row = 2 + data_rows_count
            day_intervals[date_str_short] = data_rows_count

            # Merge date header
            ws.merge_cells(start_row=1, start_column=col_start, end_row=1, end_column=col_start+3)
            ws.cell(row=1, column=col_start, value=date_str_full).alignment = Alignment(horizontal="center", vertical="center")

            # Sub-headers
            ws.cell(row=2, column=col_start, value="UTC Offset (minutes)")
            ws.cell(row=2, column=col_start+1, value="Local Time Stamp")
            ws.cell(row=2, column=col_start+2, value="Active Power (W)")
            ws.cell(row=2, column=col_start+3, value="kW")

            if data_rows_count > 0:
                ws.merge_cells(start_row=merge_start_row, start_column=col_start, end_row=merge_end_row, end_column=col_start)
                ws.cell(row=merge_start_row, column=col_start, value=date_str_full).alignment = Alignment(horizontal="center", vertical="center")

            # Fill rows
            for idx, r in enumerate(day_data.itertuples(), start=3):
                ws.cell(row=idx, column=col_start+1, value=r.Time)
                power_w = getattr(r, POWER_COL_OUT)
                ws.cell(row=idx, column=col_start+2, value=power_w)
                ws.cell(row=idx, column=col_start+3, value=abs(power_w)/1000)

            # Summary stats
            if data_rows_count > 0:
                sum_w = day_data[POWER_COL_OUT].sum()
                mean_w = day_data[POWER_COL_OUT].mean()
                max_w = day_data[POWER_COL_OUT].max()
                sum_kw_abs = day_data[POWER_COL_OUT].abs().sum() / 1000
                mean_kw_abs = day_data[POWER_COL_OUT].abs().mean() / 1000
                max_kw_abs = day_data[POWER_COL_OUT].abs().max() / 1000

                stats_row_start = merge_end_row + 1
                ws.cell(row=stats_row_start, column=col_start+1, value="Total")
                ws.cell(row=stats_row_start, column=col_start+2, value=sum_w)
                ws.cell(row=stats_row_start, column=col_start+3, value=sum_kw_abs)

                ws.cell(row=stats_row_start+1, column=col_start+1, value="Average")
                ws.cell(row=stats_row_start+1, column=col_start+2, value=mean_w)
                ws.cell(row=stats_row_start+1, column=col_start+3, value=mean_kw_abs)

                ws.cell(row=stats_row_start+2, column=col_start+1, value="Max")
                ws.cell(row=stats_row_start+2, column=col_start+2, value=max_w)
                ws.cell(row=stats_row_start+2, column=col_start+3, value=max_kw_abs)

                max_row_used = max(max_row_used, stats_row_start+2)
                daily_max_summary.append((date_str_short, max_kw_abs))
                chart_series_data.append((col_start+3, date_str_short))

            col_start += 4

        final_summary_row = max_row_used + 2

        # Summary Table
        if daily_max_summary:
            title_cell = ws.cell(row=final_summary_row, column=1, value="Daily Max Power (kW) Summary")
            ws.merge_cells(start_row=final_summary_row, start_column=1, end_row=final_summary_row, end_column=2)
            title_cell.alignment = Alignment(horizontal="center", vertical="center")
            title_cell.font = title_font

            header_row = final_summary_row + 1
            ws.cell(row=header_row, column=1, value="Day").fill = header_fill
            ws.cell(row=header_row, column=1).font = header_font
            ws.cell(row=header_row, column=1).border = thin_border
            ws.cell(row=header_row, column=1).alignment = Alignment(horizontal="center")

            ws.cell(row=header_row, column=2, value="Max (kW)").fill = header_fill
            ws.cell(row=header_row, column=2).font = header_font
            ws.cell(row=header_row, column=2).border = thin_border
            ws.cell(row=header_row, column=2).alignment = Alignment(horizontal="center")

            current_row = header_row
            for date_str, max_kw in daily_max_summary:
                current_row += 1
                fill_style = data_fill_alt if (current_row % 2)==0 else PatternFill(fill_type=None)
                day_cell = ws.cell(row=current_row, column=1, value=date_str)
                day_cell.border = thin_border
                day_cell.fill = fill_style
                day_cell.alignment = Alignment(horizontal="center")

                max_cell = ws.cell(row=current_row, column=2, value=max_kw)
                max_cell.number_format = numbers.FORMAT_NUMBER_00
                max_cell.border = thin_border
                max_cell.fill = fill_style
                max_cell.alignment = Alignment(horizontal="right")

        # Chart
        if chart_series_data:
            chart = LineChart()
            chart.title = "Daily Absolute Power Profile (kW)"
            chart.style = 10

            for kw_col_idx, day_str in chart_series_data:
                num_rows = day_intervals[day_str]
                values = Reference(ws, min_col=kw_col_idx, min_row=3, max_col=kw_col_idx, max_row=2+num_rows)
                chart.add_data(values, titles_from_data=False)
                chart.series[-1].title = day_str  # Assign series title safely

            # X-axis categories (first day times)
            first_day_rows = day_intervals[chart_series_data[0][1]]
            categories = Reference(ws, min_col=2, min_row=3, max_col=2, max_row=2+first_day_rows)
            chart.set_categories(categories)

            chart.x_axis.title = "10-Minute Interval"
            chart.y_axis.title = "Absolute Power (kW)"
            ws.add_chart(chart, f'D{final_summary_row}')
            chart.width = 18
            chart.height = 12

    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    return stream

# -----------------------------
# STREAMLIT UI
# -----------------------------
def app():
    st.title("📊 Excel 10-Minute Electricity Data Converter")
    st.markdown("""
        Upload an **Excel file (.xlsx)** with time-series data. Each sheet is processed 
        separately to calculate the total absolute power (W) consumed/generated 
        in fixed **10-minute intervals**.
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
                st.error(f"Error reading sheet '{sheet_name}': {e}")
                continue

            df.columns = df.columns.astype(str).str.strip()
            possible_time_cols = ["Date & Time", "Date&Time", "Timestamp", "DateTime", "Local Time", "TIME", "ts"]
            timestamp_col = next((col for col in df.columns if col in possible_time_cols), None)
            if timestamp_col is None:
                st.error(f"No valid **Timestamp** column found in sheet **{sheet_name}**.")
                continue

            possible_psum_cols = ["PSum (W)", "Psum (W)", "PSum", "P (W)", "Power"]
            psum_col = next((col for col in df.columns if col in possible_psum_cols), None)
            if psum_col is None:
                st.error(f"No valid **PSum** column found in sheet **{sheet_name}**.")
                continue

            processed = process_sheet(df, timestamp_col, psum_col)
            if not processed.empty:
                result_sheets[sheet_name] = processed
                st.success(f"Sheet **{sheet_name}** processed successfully ({len(processed['Date'].unique())} days).")
            else:
                st.warning(f"Sheet **{sheet_name}** contained no usable data.")

        if result_sheets:
            output_stream = build_output_excel(result_sheets)
            st.download_button(
                label="📥 Download Converted Excel",
                data=output_stream,
                file_name="Converted_Output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("No sheets were successfully processed.")

if __name__ == '__main__':
    app()
