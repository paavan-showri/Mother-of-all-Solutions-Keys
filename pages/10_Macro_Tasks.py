import pandas as pd
import streamlit as st

from session_utils import require_workbook
from modules.m01_workbook_loader import load_precedence_tasks, load_current_state_steps
from modules.m02_nlp_normalization import normalize_steps
from modules.m05_macro_tasks import generate_macro_tasks

st.set_page_config(page_title="Macro Tasks", layout="wide")
st.title("10 Macro Tasks")

ctx = require_workbook()

manual_tasks = load_precedence_tasks(ctx["excel_file"], sheet_name=ctx["precedence_sheet"])
st.subheader("Macro tasks from Precedence Network sheet")
st.dataframe(pd.DataFrame([vars(t) for t in manual_tasks]), use_container_width=True)

ctx = require_workbook()
steps = load_current_state_steps(ctx["excel_file"], sheet_name=ctx["sheet_name"])
normalized = normalize_steps(steps)
generated = generate_macro_tasks(normalized)

st.subheader("Macro tasks generated from FPC rows")
st.dataframe(pd.DataFrame([vars(t) for t in generated]), use_container_width=True)
