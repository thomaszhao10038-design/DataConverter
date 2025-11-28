import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import io

st.set_page_config(page_title="Power Data Converter", layout="wide")
st.title("📊 Power Data to Daily 4-Column Excel Converter")
st.markdown("""
Upload one or more `.xlsx` files.  
Each sheet will be processed independently and returned in a single output file with:
- 4 columns per day (starting from column A, E, I, ...)
- Merged date headers for same days
- 10-minute intervals (00:00 → 23:50)
- Active Power (W) = sum of instantaneous PSum (W) in that 10-min bin
- kW column = Active Power (W) / 1000
""")

uploaded_files = st.file_uploader(
    "Upload Excel files (.xlsx)", type=["xlsx"], accept_multiple_files=True
)

if not uploaded_files:
    st.info("Please upload at least one Excel file.")
    st.stop()

@st.cache_data
def process_file(file):
    xls = pd.ExcelFile(file)
    output_sheets = {}

    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)

        # === Expected columns (adjust if your column names differ slightly) ===
        if "PSum (W)" not in df.columns:
            st.error(f"Sheet '{sheet_name}' in {file.name} does not have 'PSum (W)' column.")
            continue

        # Try to find timestamp column (common names)
        time_col = None
        for col in ["Local Time Stamp", "Time", "Timestamp", "DateTime", "LocalTime"]:
            if col in df.columns:
                time_col = col
                break
        if time_col is None:
            st.error(f"Could not find timestamp column in sheet '{sheet_name}'")
            continue

        df[time_col] = pd.to_datetime(df[time_col], errors='coerce')
        df = df.dropna(subset=[time_col])  # remove invalid timestamps

        # Extract UTC offset if available (optional)
        utc_offset_col = None
        for col in ["UTC Offset", "UTC Offset (minutes)", "Offset"]:
            if col in df.columns:
                utc_offset_col = col
                break

        # Create 10-minute bins
        df["bin_start"] = df[time_col].dt.floor('10min')

        # Group by date and 10-min bin
        grouped = df.groupby([
            df["bin_start"].dt.date,
            df["bin_start"].dt.time
        ])["PSum (W)"].sum().reset_index()
        grouped.rename(columns={"bin_start": "time"}, inplace=True)

        # Create full 10-min grid for each day present
        dates = sorted(grouped["date"].unique())
        all_times = [datetime(2000,1,1, h, m) for h in range(24) for m in (0,10,20,30,40,50)]
        time_labels = [t.strftime("%H:%M") for t in all_times]

        final_data = []
        date_objects = []

        for date in dates:
            day_data = grouped[grouped["date"] == date]
            power_dict = dict(zip(zip(day_data["date"], day_data["time"]), day_data["PSum (W)"]))

            row_power = []
            for t in all_times:
                key = (date, t.time())
                row_power.append(power_dict.get(key, None))

            final_data.append(row_power)
            date_objects.append(date)

        # Build output DataFrame for this sheet
        columns_per_day = 4
        total_days = len(dates)
        total_cols = total_days * columns_per_day

        # Create multi-level columns
        arrays = []
        for i, date in enumerate(dates):
            dt = pd.to_datetime(date)
            date_str = dt.strftime("%Y-%m-%d")
            arrays.append([date_str, date_str, date_str, date_str])

        arrays.append(["UTC Offset (minutes)", "Local Time Stamp", "Active Power (W)", "kW"])

        tuples = list(zip(*arrays)) if total_days > 0 else []
        multi_cols = pd.MultiIndex.from_tuples(tuples)

        # Output DataFrame
        out_df = pd.DataFrame(index=range(len(time_labels)), columns=multi_cols)

        # Fill time labels and data
        for i, time_label in enumerate(time_labels):
            out_df.loc[i, (slice(None), "Local Time Stamp")] = time_label

        for day_idx, date in enumerate(dates):
            base_col = day_idx * 4
            power_col = out_df.columns[base_col + 2]   # Active Power (W)
            kw_col = out_df.columns[base_col + 3]      # kW

            for row_idx, power in enumerate(final_data[day_idx]):
                if power is not None and pd.notna(power):
                    out_df.iat[row_idx, power_col[0]*4 + power_col[1]] = int(power) if power == int(power) else power
                    out_df.iat[row_idx, kw_col[0]*4 + kw_col[1]] = round(power / 1000, 3)

            # Fill UTC Offset if we have it
            if utc_offset_col is not None:
                offset_val = df[utc_offset_col].iloc[0] if len(df[utc_offset_col].dropna()) > 0 else ""
                offset_col = out_df.columns[base_col]
                out_df[offset_col] = offset_val

        # Reorder columns properly (pandas sometimes messes order)
        ordered_cols = []
        for i in range(total_days):
            start = i * 4
            ordered_cols.extend(out_df.columns[start:start+4])
        out_df = out_df[ordered_cols]

        output_sheets[sheet_name] = out_df

    return output_sheets

if st.button("🚀 Process All Files"):
    with st.spinner("Processing files..."):
        all_sheets = {}
        for file in uploaded_files:
            sheets = process_file(file)
            for name, df in sheets.items():
                new_name = f"{file.name}_{name}"
                all_sheets[new_name] = df

        if not all_sheets:
            st.error("No valid data was processed.")
            st.stop()

        # Write to Excel with merged headers
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            for sheet_name, df in all_sheets.items():
                df.to_excel(writer, sheet_name=sheet_name[:31], index=False)

                workbook = writer.book
                worksheet = writer.sheets[sheet_name[:31]]

                # Merge cells for same date headers
                col_idx = 0
                while col_idx < len(df.columns):
                    date_group_start = col_idx
                    current_date = df.columns[col_idx][0]

                    while col_idx < len(df.columns) and df.columns[col_idx][0] == current_date:
                        col_idx += 1

                    if (col_idx - date_group_start) > 1:
                        worksheet.merge_range(
                            0, date_group_start, 0, col_idx - 1,
                            current_date,
                            workbook.add_format({'align': 'center', 'bold': True, 'bg_color': '#D9E1F2'})
                        )

                # Format headers row 2
                header_format = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#A6A6A6'})
                for col_num, value in enumerate(df.columns.get_level_values(1)):
                    worksheet.write(1, col_num, value, header_format)

                # Auto-fit columns
                for i, col in enumerate(df.columns):
                    max_len = max(
                        df[col].astype(str).map(len).max(),
                        len(str(col[1]))
                    ) + 2
                    worksheet.set_column(i, i, min(max_len, 30))

        output.seek(0)
        st.success("Processing complete!")
        st.download_button(
            label="📥 Download Converted Excel File",
            data=output,
            file_name=f"Converted_Power_Data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
