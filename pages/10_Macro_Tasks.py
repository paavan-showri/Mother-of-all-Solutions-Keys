import pandas as pd
import streamlit as st

from session_utils import require_workbook
from modules.m01_workbook_loader import load_precedence_tasks

st.set_page_config(page_title="Macro Tasks", layout="wide")
st.title("10 Macro Tasks")

ctx = require_workbook()

try:
    manual_tasks = load_precedence_tasks(ctx["excel_file"], sheet_name=ctx["precedence_sheet"])
except Exception as e:  # noqa: BLE001
    st.error(f"Could not load macro tasks from worksheet '{ctx['precedence_sheet']}': {e}")
    st.stop()

macro_df = pd.DataFrame([
    {
        "Task ID": t.task_id,
        "Task Name": t.name,
        "Duration (sec)": t.duration_sec,
        "Immediate Predecessors": ", ".join(str(p) for p in t.predecessors) if t.predecessors else "—",
        "Resources": ", ".join(t.resources) if t.resources else "—",
        "Type": t.internal_external,
    }
    for t in manual_tasks
])

st.subheader("Macro tasks from Precedence Network sheet")
st.dataframe(macro_df, use_container_width=True)
st.caption("This page now depends only on the Precedence Network worksheet.")
