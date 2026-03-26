import streamlit as st

from session_utils import require_workbook
from modules.m01_workbook_loader import load_current_state_steps, load_resource_capacities
from modules.m02_nlp_normalization import normalize_steps
from modules.m05_macro_tasks import generate_macro_tasks
from modules.m07_rcpsp_solver import solve_rcpsp

st.set_page_config(page_title='RCPSP Schedule', layout='wide')
st.title('12 RCPSP Schedule')
ctx = require_workbook()
use_zero_shot = st.checkbox('Use zero-shot waste classification', value=False)
use_manufactubert = st.checkbox('Use ManufactuBERT semantic hints', value=False)
use_planner_durations = st.checkbox('Use planner-style future-state durations', value=True)
steps = load_current_state_steps(ctx['excel_file'], sheet_name=ctx['sheet_name'])
normalized = normalize_steps(steps, use_zero_shot=use_zero_shot, use_manufactubert=use_manufactubert)
tasks = generate_macro_tasks(normalized, use_planner_durations=use_planner_durations)
try:
    capacities = load_resource_capacities(ctx['excel_file'], sheet_name=ctx['resource_sheet'])
except Exception:
    capacities = {'Man': 1, 'Plate': 1, 'Toaster': 1, 'Knife': 1, 'Butter': 1}
schedule = solve_rcpsp(tasks, capacities)
st.write('Makespan (sec):', schedule.attrs.get('makespan'))
st.dataframe(schedule, use_container_width=True)
