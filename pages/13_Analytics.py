import streamlit as st

from session_utils import require_workbook
from modules.m01_workbook_loader import load_current_state_steps, load_precedence_tasks, load_resource_capacities
from modules.m02_nlp_normalization import normalize_steps
from modules.m03_ie_ontology import map_steps_to_ie_ontology
from modules.m07_rcpsp_solver import solve_rcpsp
from modules.m08_analytics import build_analytics

st.set_page_config(page_title="Analytics", layout="wide")
st.title("13 Analytics")

ctx = require_workbook()
steps = load_current_state_steps(ctx["excel_file"], sheet_name=ctx["sheet_name"])

ctx = require_workbook()
ontology = map_steps_to_ie_ontology(normalize_steps(steps))
tasks = load_precedence_tasks(ctx["excel_file"], sheet_name=ctx["precedence_sheet"])

ctx = require_workbook()
capacities = load_resource_capacities(ctx["excel_file"], sheet_name=ctx["resource_sheet"])
schedule = solve_rcpsp(tasks, capacities)
analytics = build_analytics(steps, ontology, tasks, capacities, schedule)

st.subheader("Comparison")
st.dataframe(analytics["comparison"], use_container_width=True)

st.subheader("Bottleneck")
st.json(analytics["bottleneck"])

st.subheader("Resource utilization")
st.dataframe(analytics["resource_utilization"], use_container_width=True)

st.subheader("Waste summary")
st.dataframe(analytics["waste_summary"], use_container_width=True)
