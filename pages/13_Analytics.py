import streamlit as st
from ie_modules_v2.m01_workbook_loader import load_current_state_steps, load_precedence_tasks, load_resource_capacities
from ie_modules_v2.m02_nlp_normalization import normalize_steps
from ie_modules_v2.m03_ie_ontology import map_steps_to_ie_ontology
from ie_modules_v2.m07_rcpsp_solver import solve_rcpsp
from ie_modules_v2.m08_analytics import build_analytics

st.title("13 Analytics")
if "uploaded_excel" not in st.session_state:
    st.warning("Upload the workbook in Home first.")
    st.stop()

file = st.session_state["uploaded_excel"]
steps = load_current_state_steps(file)
ontology = map_steps_to_ie_ontology(normalize_steps(steps))
tasks = load_precedence_tasks(file)
capacities = load_resource_capacities(file)
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
