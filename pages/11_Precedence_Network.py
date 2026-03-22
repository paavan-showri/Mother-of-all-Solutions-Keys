import streamlit as st

from session_utils import require_workbook
from modules.m01_workbook_loader import load_precedence_tasks
from modules.m06_precedence_network import build_precedence_outputs

st.set_page_config(page_title="Precedence Network", layout="wide")
st.title("11 Precedence Network")

ctx = require_workbook()
tasks = load_precedence_tasks(ctx["excel_file"], sheet_name=ctx["precedence_sheet"])
outputs = build_precedence_outputs(tasks)

st.subheader("Technological precedence")
st.dataframe(outputs["technological_precedence"], use_container_width=True)

st.subheader("Resource-linked conflicts")
st.dataframe(outputs["resource_linked_conflicts"], use_container_width=True)

st.subheader("Parallelizable task pairs")
st.dataframe(outputs["parallelizable_tasks"], use_container_width=True)

st.write("Critical path approximation:", outputs["critical_path_task_ids"])
