import pandas as pd
import streamlit as st

from session_utils import require_workbook
from modules.m01_workbook_loader import load_current_state_steps
from modules.m02_nlp_normalization import normalize_steps

st.set_page_config(page_title="NLP Normalization", layout="wide")
st.title("7 NLP Normalization")

ctx = require_workbook()
steps = load_current_state_steps(ctx["excel_file"], sheet_name=ctx["sheet_name"])
enable_transformers = st.toggle("Enable Hugging Face zero-shot classifier", value=False)
normalized = normalize_steps(steps, enable_transformers=enable_transformers)
df = pd.DataFrame([vars(x) for x in normalized])

st.dataframe(df, use_container_width=True)
