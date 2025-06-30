# -------------------- script1.py --------------------
import re
import pandas as pd
import numpy as np
from datetime import timedelta


def time_to_seconds(ts: str) -> int:
    if pd.isna(ts) or ts == "":
        return 0
    parts = list(map(int, ts.split(':')))
    while len(parts) < 3:
        parts.insert(0, 0)
    h, m, s = parts
    return h*3600 + m*60 + s


def seconds_to_hms(sec: float) -> str:
    return "" if pd.isna(sec) else str(timedelta(seconds=int(sec)))


def natural_key(s: str):
    return [int(chunk) if chunk.isdigit() else chunk.lower()
            for chunk in re.split(r'(\d+)', s)]


def process_echo360(df: pd.DataFrame) -> pd.DataFrame:
    """
    Process Echo360 DataFrame and return a summary DataFrame.
    """
    df2 = df.copy()
    df2['Duration_sec']      = df2['Duration'].apply(time_to_seconds)
    df2['TotalViewTime_sec'] = df2['Total View Time'].apply(time_to_seconds)
    df2['AvgViewTime_sec']   = df2['Average View Time'].apply(time_to_seconds)
    df2['Row View %']        = df2['TotalViewTime_sec'] / df2['Duration_sec'].replace(0, np.nan)

    grp = df2.groupby('Media Name', sort=False)
    titles = list(grp.groups.keys())

    summary_core = pd.DataFrame({
        'Media Title':              titles,
        'Video Duration':           [grp.get_group(t)['Duration_sec'].iloc[0] for t in titles],
        'Number of Unique Viewers': grp['User Name'].nunique().values,
        'Average View %':           grp['Row View %'].mean().fillna(0).values,
        'Total View %':             (grp['TotalViewTime_sec'].sum() / grp['Duration_sec'].sum()).values,
        'Total View Time':          grp['TotalViewTime_sec'].sum().values,
        'Average View Time':        grp['AvgViewTime_sec'].mean().values,
        'Average Total View Time':  grp['TotalViewTime_sec'].mean().values,
    })

    # Natural sort
    summary_core['sort_key'] = summary_core['Media Title'].apply(natural_key)
    summary_core = (
        summary_core
        .sort_values('sort_key')
        .drop(columns='sort_key')
        .reset_index(drop=True)
    )

    # Grand Total (averages)
    metrics = ['Video Duration','Total View Time','Average View Time','Average Total View Time']
    means = summary_core[metrics].mean()
    viewers_mean = summary_core['Number of Unique Viewers'].mean()
    summary_core.loc[len(summary_core)] = {
        'Media Title':               'Grand Total',
        'Video Duration':            means['Video Duration'],
        'Number of Unique Viewers':  viewers_mean,
        'Average View %':            summary_core['Average View %'].mean(),
        'Total View %':              summary_core['Total View %'].mean(),
        'Total View Time':           means['Total View Time'],
        'Average View Time':         means['Average View Time'],
        'Average Total View Time':   means['Average Total View Time'],
    }

    # Average Video Length and Watch Time
    n = len(summary_core) - 1
    means2 = summary_core.loc[:n-1, metrics].mean()
    summary_core.loc[len(summary_core)] = {
        'Media Title':               'Average Video Length and Watch Time',
        'Video Duration':            means2['Video Duration'],
        'Number of Unique Viewers':  np.nan,
        'Average View %':            summary_core.loc[:n-1, 'Average View %'].mean(),
        'Total View %':              summary_core.loc[:n-1, 'Total View %'].mean(),
        'Total View Time':           means2['Total View Time'],
        'Average View Time':         means2['Average View Time'],
        'Average Total View Time':   means2['Average Total View Time'],
    }

    # Convert time metrics to hh:mm:ss strings
    for col in ['Video Duration','Total View Time','Average View Time','Average Total View Time']:
        summary_core[col] = summary_core[col].apply(seconds_to_hms)

    # Select and order columns
    summary = summary_core[[
        'Media Title','Video Duration','Number of Unique Viewers',
        'Average View %','Total View %','Total View Time',
        'Average View Time','Average Total View Time'
    ]]
    return summary