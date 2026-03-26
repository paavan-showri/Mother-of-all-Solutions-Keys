import io
import re

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

from session_utils import require_workbook

st.set_page_config(page_title='Waste Analysis', layout='wide')
st.title('6 Waste Analysis')
ctx = require_workbook()


def normalize_text(x):
    if pd.isna(x):
        return ''
    return str(x).strip()


def duration_to_seconds(x):
    if pd.isna(x):
        return 0.0
    if isinstance(x, pd.Timedelta):
        return x.total_seconds()
    if isinstance(x, (int, float, np.integer, np.floating)):
        val = float(x)
        if 0 < val < 1:
            return val * 24 * 3600
        return val
    s = str(x).strip()
    if ':' in s:
        parts = s.split(':')
        parts = [float(p) for p in parts]
        if len(parts) == 3:
            h, m, sec = parts
            return h * 3600 + m * 60 + sec
        if len(parts) == 2:
            m, sec = parts
            return m * 60 + sec
    try:
        return float(s)
    except Exception:
        return 0.0


def find_header_row(raw_df):
    required = {'Step', 'Description', 'Activity', 'Duration (Sec)'}
    for i, row in raw_df.iterrows():
        vals = {normalize_text(v) for v in row.tolist() if pd.notna(v)}
        if required.issubset(vals):
            return i
    return None


def classify_waste(row):
    desc = normalize_text(row.get('Description', '')).lower()
    activity = normalize_text(row.get('Activity', '')).upper()
    if activity == 'W' or 'wait' in desc:
        return 'Waiting'
    if activity == 'I' or 'inspect' in desc or 'check' in desc:
        return 'Inspection'
    if any(k in desc for k in ['walk', 'move', 'carry', 'turn back', 'turn around', 'search', 'look for', 'go to']):
        return 'Motion'
    if activity == 'S':
        return 'Inventory'
    if activity == 'D':
        return 'Extra Processing'
    return 'Other'


bio = io.BytesIO(ctx['excel_file'].getvalue())
raw = pd.read_excel(bio, sheet_name=ctx['sheet_name'], header=None)
header_row = find_header_row(raw)
if header_row is None:
    st.error('Could not find the Flow Process Chart table.')
    st.stop()

bio = io.BytesIO(ctx['excel_file'].getvalue())
df = pd.read_excel(bio, sheet_name=ctx['sheet_name'], header=header_row)
df.columns = [normalize_text(c) for c in df.columns]
df = df.dropna(subset=['Description']).copy()
df['Duration_sec'] = df['Duration (Sec)'].apply(duration_to_seconds)
df['Waste_Type'] = df.apply(classify_waste, axis=1)

summary = df.groupby('Waste_Type', as_index=False)['Duration_sec'].sum().sort_values('Duration_sec', ascending=False)
st.dataframe(summary, use_container_width=True)
fig = px.bar(summary, x='Waste_Type', y='Duration_sec', title='Waste time by category')
st.plotly_chart(fig, use_container_width=True)
st.dataframe(df[['Step', 'Description', 'Activity', 'Duration_sec', 'Waste_Type']], use_container_width=True)
