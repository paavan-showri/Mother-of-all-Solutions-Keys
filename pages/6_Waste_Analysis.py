import pandas as pd
import plotly.express as px
import streamlit as st

from session_utils import require_workbook
from modules.m01_workbook_loader import load_current_state_steps
from modules.m02_nlp_normalization import normalize_steps

st.set_page_config(page_title="Waste Analysis", layout="wide")
st.title("Waste Analysis")
ctx = require_workbook()

steps = load_current_state_steps(ctx["excel_file"], sheet_name=ctx["sheet_name"])
normalized = normalize_steps(steps)
df = pd.DataFrame(
    [
        {
            "Step": s.step,
            "Description": s.raw_description,
            "Activity": s.activity,
            "Duration (sec)": s.duration_sec,
            "Waste": s.waste_pred,
            "Review": s.review_flag,
        }
        for s in normalized
    ]
)
summary = df.groupby("Waste", as_index=False)["Duration (sec)"].sum().sort_values("Duration (sec)", ascending=False)

left, right = st.columns([1, 2])
with left:
    st.dataframe(summary, use_container_width=True)
with right:
    fig = px.bar(summary, x="Waste", y="Duration (sec)", title="Waste Time by Category")
    st.plotly_chart(fig, use_container_width=True)

st.dataframe(df, use_container_width=True)
