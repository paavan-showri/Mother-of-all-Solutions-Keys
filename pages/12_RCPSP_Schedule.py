import streamlit as st
from ie_modules_v2.m01_workbook_loader import load_precedence_tasks, load_resource_capacities
from ie_modules_v2.m07_rcpsp_solver import solve_rcpsp

st.title("12 RCPSP Schedule")
if "uploaded_excel" not in st.session_state:
    st.warning("Upload the workbook in Home first.")
    st.stop()

file = st.session_state["uploaded_excel"]
tasks = load_precedence_tasks(file)
capacities = load_resource_capacities(file)
schedule = solve_rcpsp(tasks, capacities)

st.write("Makespan (sec):", schedule.attrs["makespan"])
st.dataframe(schedule, use_container_width=True)
