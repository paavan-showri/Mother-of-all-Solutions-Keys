import pandas as pd
import streamlit as st

from session_utils import require_workbook
from modules.m01_workbook_loader import load_current_state_steps
from modules.m02_nlp_normalization import normalize_steps
from modules.m05_macro_tasks import generate_macro_tasks

st.set_page_config(page_title='Macro Tasks', layout='wide')
st.title('10 Macro Tasks')
ctx = require_workbook()
use_zero_shot = st.checkbox('Use zero-shot waste classification', value=False)
use_manufactubert = st.checkbox('Use ManufactuBERT semantic hints', value=False)
use_planner_durations = st.checkbox('Use planner-style future-state durations', value=True)
steps = load_current_state_steps(ctx['excel_file'], sheet_name=ctx['sheet_name'])
normalized = normalize_steps(steps, use_zero_shot=use_zero_shot, use_manufactubert=use_manufactubert)
tasks = generate_macro_tasks(normalized, use_planner_durations=use_planner_durations)
df = pd.DataFrame([{
    'Task ID': t.task_id, 'Task Name': t.name, 'Duration (sec)': t.duration_sec,
    'Immediate Predecessors': ', '.join(str(p) for p in t.predecessors) if t.predecessors else '—',
    'Resources': ', '.join(t.resources) if t.resources else '—', 'Type': t.internal_external,
} for t in tasks])
st.dataframe(df, use_container_width=True)
