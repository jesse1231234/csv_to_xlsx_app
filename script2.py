# -------------------- script2.py --------------------
import pandas as pd
import numpy as np


def process_gradebook(df: pd.DataFrame) -> pd.DataFrame:
    """
    Process gradebook DataFrame and return cleaned DataFrame with summary rows.
    """
    df2 = df.copy()

    # 1. Drop rows where first column contains "Student, Test"
    mask = df2.iloc[:, 0].astype(str).str.contains("Student, Test", na=False)
    df2 = df2[~mask].reset_index(drop=True)

    # 2. Drop unwanted columns except 'Final Grade'
    to_drop = ["Student","ID","SIS User ID","SIS Login ID",
               "Current Grade","Unposted Current Grade","Unposted Final Grade"]
    df2.drop(columns=[c for c in to_drop if c in df2.columns], inplace=True)

    # 3. Drop cols where rows 3+ (idx>=2) are all zeros/empty except 'Final Grade'
    drop_cols = []
    for col in df2.columns:
        if col == 'Final Grade': continue
        s = pd.to_numeric(df2[col].iloc[2:], errors='coerce')
        if s.fillna(0).eq(0).all():
            drop_cols.append(col)
    df2.drop(columns=drop_cols, inplace=True)

    # 4. Fill empty cells & convert to numeric (except 'Final Grade')
    for col in df2.columns:
        if col == 'Final Grade': continue
        df2[col] = pd.to_numeric(df2[col], errors='coerce').fillna(0)

    # 5. Replace '(read only)' in second row with max of col
    for col in df2.columns:
        if col == 'Final Grade': continue
        val = df2.at[1, col]
        if isinstance(val, str) and '(read only)' in val:
            data_vals = pd.to_numeric(df2[col].iloc[2:], errors='coerce').dropna()
            if not data_vals.empty:
                df2.at[1, col] = data_vals.max()

    # 6. Insert 'Row Titles' column and label rows
    df2.insert(0, 'Row Titles', '')
    if len(df2) > 1:
        df2.at[1, 'Row Titles'] = 'Points Possible'

    # 7. Append summary rows: 'Average' & 'Average Excluding Zeros'
    numeric_cols = [c for c in df2.columns if c not in ['Row Titles','Final Grade']]
    data_rows = df2.iloc[2:] if len(df2) > 2 else pd.DataFrame()
    avg  = data_rows[numeric_cols].mean(numeric_only=True)
    avg0 = data_rows[numeric_cols][data_rows[numeric_cols] > 0].mean(numeric_only=True)

    avg_row  = {col: (avg[col] / df2.at[1, col] if col in avg.index else None)  for col in numeric_cols}
    avg0_row = {col: (avg0[col]/ df2.at[1, col] if col in avg0.index else None) for col in numeric_cols}
    avg_row.update({'Row Titles':'Average'});  avg0_row.update({'Row Titles':'Average Excluding Zeros'})

    df_summary = pd.DataFrame([avg_row, avg0_row], columns=df2.columns)
    df2 = pd.concat([df2, df_summary], ignore_index=True)

    return df2
