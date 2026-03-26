import streamlit as st

from session_utils import require_workbook
from modules.m01_workbook_loader import load_current_state_steps
from modules.m02_nlp_normalization import normalize_steps
from modules.m05_macro_tasks import generate_macro_tasks, tasks_to_df

st.set_page_config(page_title="Macro Tasks", layout="wide")
st.title("Macro Tasks")
st.caption("Discovered automatically from the current-state NLP output and Lean rules.")

ctx = require_workbook()
steps = load_current_state_steps(ctx["excel_file"], sheet_name=ctx["sheet_name"])
normalized = normalize_steps(steps)
tasks = generate_macro_tasks(normalized)
st.dataframe(tasks_to_df(tasks), use_container_width=True)
