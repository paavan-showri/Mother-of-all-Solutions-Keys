import streamlit as st

from session_utils import require_workbook
from modules.m01_workbook_loader import load_current_state_steps
from modules.m02_nlp_normalization import normalize_steps, normalized_to_df

st.set_page_config(page_title="NLP Normalization", layout="wide")
st.title("NLP Normalization")
st.caption("Shows the cleaned NLP normalization table and flags rows that may need review.")

ctx = require_workbook()
show_review_only = st.checkbox("Show only flagged rows", value=False)

steps = load_current_state_steps(ctx["excel_file"], sheet_name=ctx["sheet_name"])
normalized = normalize_steps(steps)
df = normalized_to_df(normalized)

if not df.empty and "review_flag" in df.columns:
    flagged_count = int(df["review_flag"].sum())
    c1, c2, c3 = st.columns(3)
    c1.metric("Rows", len(df))
    c2.metric("Flagged", flagged_count)
    c3.metric("Flag Rate", f"{(100 * flagged_count / max(len(df), 1)):.1f}%")
    if show_review_only:
        df = df[df["review_flag"]].copy()

st.dataframe(df, use_container_width=True)
