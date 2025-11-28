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
    
    # Convert to kW
    grouped['kW'] = grouped[POWER_COL_OUT].abs() / 1000

    # -----------------------------
    # Remove leading/trailing zeros for each day
    # -----------------------------
    def trim_zeros(day_df):
        vals = day_df[POWER_COL_OUT].values
        non_zero_idx = (vals != 0).nonzero()[0]
        if len(non_zero_idx) == 0:
            return day_df  # all zeros, keep as-is
        first, last = non_zero_idx[0], non_zero_idx[-1]
        day_df.loc[:first-1, POWER_COL_OUT] = None
        day_df.loc[:first-1, 'kW'] = None
        day_df.loc[last+1:, POWER_COL_OUT] = None
        day_df.loc[last+1:, 'kW'] = None
        return day_df

    grouped = grouped.groupby('Date', group_keys=False).apply(trim_zeros)

    return grouped
