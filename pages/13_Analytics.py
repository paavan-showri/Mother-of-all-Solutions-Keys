import streamlit as st

from session_utils import require_workbook
from modules.m01_workbook_loader import load_current_state_steps, load_resource_capacities
from modules.m02_nlp_normalization import normalize_steps
from modules.m03_ie_ontology import map_steps_to_ie_ontology
from modules.m05_macro_tasks import generate_macro_tasks
from modules.m07_rcpsp_solver import solve_rcpsp
from modules.m08_analytics import build_analytics

st.set_page_config(page_title='Analytics', layout='wide')
st.title('13 Analytics')
ctx = require_workbook()
use_zero_shot = st.checkbox('Use zero-shot waste classification', value=False)
use_manufactubert = st.checkbox('Use ManufactuBERT semantic hints', value=False)
use_planner_durations = st.checkbox('Use planner-style future-state durations', value=True)
steps = load_current_state_steps(ctx['excel_file'], sheet_name=ctx['sheet_name'])
normalized = normalize_steps(steps, use_zero_shot=use_zero_shot, use_manufactubert=use_manufactubert)
ontology = map_steps_to_ie_ontology(normalized)
tasks = generate_macro_tasks(normalized, use_planner_durations=use_planner_durations)
try:
    capacities = load_resource_capacities(ctx['excel_file'], sheet_name=ctx['resource_sheet'])
except Exception:
    capacities = {'Man': 1, 'Plate': 1, 'Toaster': 1, 'Knife': 1, 'Butter': 1}
schedule = solve_rcpsp(tasks, capacities)
analytics = build_analytics(steps, ontology, tasks, capacities, schedule)
st.subheader('Comparison')
st.dataframe(analytics['comparison'], use_container_width=True)
st.subheader('Bottleneck')
st.json(analytics['bottleneck'])
st.subheader('Resource utilization')
st.dataframe(analytics['resource_utilization'], use_container_width=True)
st.subheader('Waste summary')
st.dataframe(analytics['waste_summary'], use_container_width=True)
