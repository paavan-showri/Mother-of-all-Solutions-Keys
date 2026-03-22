import streamlit as st
from ie_modules_v2.m01_workbook_loader import load_precedence_tasks
from ie_modules_v2.m06_precedence_network import build_precedence_outputs

st.title("11 Precedence Network")
if "uploaded_excel" not in st.session_state:
    st.warning("Upload the workbook in Home first.")
    st.stop()

tasks = load_precedence_tasks(st.session_state["uploaded_excel"])
outputs = build_precedence_outputs(tasks)

st.subheader("Technological precedence")
st.dataframe(outputs["technological_precedence"], use_container_width=True)

st.subheader("Resource-linked conflicts")
st.dataframe(outputs["resource_linked_conflicts"], use_container_width=True)

st.subheader("Parallelizable task pairs")
st.dataframe(outputs["parallelizable_tasks"], use_container_width=True)

st.write("Critical path approximation:", outputs["critical_path_task_ids"])
