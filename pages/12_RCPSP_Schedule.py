import streamlit as st

from session_utils import require_workbook
from modules.m01_workbook_loader import load_precedence_tasks, load_resource_capacities
from modules.m07_rcpsp_solver import solve_rcpsp

st.set_page_config(page_title="RCPSP Schedule", layout="wide")
st.title("12 RCPSP Schedule")

ctx = require_workbook()

try:
    tasks = load_precedence_tasks(ctx["excel_file"], sheet_name=ctx["precedence_sheet"])
    capacities = load_resource_capacities(ctx["excel_file"], sheet_name=ctx["resource_sheet"])
    schedule = solve_rcpsp(tasks, capacities)
except Exception as e:  # noqa: BLE001
    st.error(f"Could not build the RCPSP schedule: {e}")
    st.stop()

st.write("Makespan (sec):", schedule.attrs.get("makespan"))
st.dataframe(schedule, use_container_width=True)
