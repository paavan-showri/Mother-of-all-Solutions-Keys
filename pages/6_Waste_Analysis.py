import io
import re
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px

st.set_page_config(page_title="Waste Analysis", layout="wide")
st.title("Waste Analysis")
st.caption("Cost and effort reduction through waste elimination from the Flow Process Chart")

# =========================================================
# SESSION STATE CHECK
# =========================================================
if "excel_file_bytes" not in st.session_state:
    st.error("No Excel file found. Please upload the Excel file on the Home page first.")
    st.stop()

excel_bytes = st.session_state["excel_file_bytes"]
sheet_name = st.session_state.get("sheet_name", "FPC_Current State")
excel_name = st.session_state.get("excel_file_name", "uploaded_file.xlsx")

st.success(f"Using file from Home page: {excel_name}")
st.info(f"Using sheet: {sheet_name}")

# =========================================================
# HELPERS
# =========================================================
def normalize_text(x):
    if pd.isna(x):
        return ""
    return str(x).strip()


def duration_to_seconds(x):
    """
    Converts duration values into seconds.
    Handles:
    - Excel time fraction
    - pandas Timedelta
    - strings like 00:00:05, 00:05, 5
    - numeric seconds
    """
    if pd.isna(x):
        return 0.0

    if isinstance(x, pd.Timedelta):
        return x.total_seconds()

    if isinstance(x, (int, float, np.integer, np.floating)):
        val = float(x)
        if 0 < val < 1:
            return val * 24 * 3600
        return val

    s = str(x).strip()
    if s == "":
        return 0.0

    if ":" in s:
        parts = s.split(":")
        try:
            parts = [float(p) for p in parts]
            if len(parts) == 3:
                h, m, sec = parts
                return h * 3600 + m * 60 + sec
            elif len(parts) == 2:
                m, sec = parts
                return m * 60 + sec
            elif len(parts) == 1:
                return parts[0]
        except Exception:
            return 0.0

    try:
        return float(s)
    except Exception:
        return 0.0


def contains_keyword(text, keywords):
    text = normalize_text(text).lower()
    return any(k in text for k in keywords)


def find_header_row(raw_df):
    required_headers = {"Step", "Description", "Activity", "Duration"}
    for i, row in raw_df.iterrows():
        row_values = {normalize_text(v) for v in row.tolist() if pd.notna(v)}
        if required_headers.issubset(row_values):
            return i
    return None


def extract_item_name(description):
    """
    Extract generic item name from action text.
    Keeps the logic universal rather than hardcoding bread/butter/knife/wife.
    """
    text = normalize_text(description).lower()

    text = re.sub(r"\b\d+(st|nd|rd|th)?\b", " ", text)
    text = re.sub(r"\b\d+\b", " ", text)

    patterns = [
        r"search for (.+)",
        r"walk to (.+)",
        r"walk with (.+?) to",
        r"grasp (.+)",
        r"get (.+)",
        r"pick up (.+)",
        r"retrieve (.+)",
        r"bring (.+?) to",
        r"move (.+?) to",
        r"carry (.+?) to",
        r"go to (.+)",
        r"reach for (.+)",
        r"fetch (.+)",
        r"open (.+)",
        r"take (.+)",
    ]

    candidate = None
    for pat in patterns:
        m = re.search(pat, text)
        if m:
            candidate = m.group(1).strip()
            break

    if candidate is None:
        candidate = text

    candidate = re.split(r"\bto\b|\bfrom\b|\bon\b|\bin\b|\bnear\b|\bat\b", candidate)[0].strip()
    candidate = re.sub(r"[^a-zA-Z\s]", " ", candidate)

    stop_words = {
        "the", "a", "an", "to", "from", "with", "for", "of", "on", "in", "into",
        "near", "counter", "top", "slot", "side", "left", "right", "and", "by",
        "at", "up", "down"
    }

    words = [w for w in candidate.split() if w not in stop_words]

    if not words:
        return "Unspecified Item"

    return " ".join(words[:3]).title()


def classify_waste(row):
    """
    Universal waste classification logic.

    Walking waste:
    - transport-like actions
    - retrieval/search/movement text

    Pure waste:
    - informing/deciding/waiting type actions
    """
    desc = normalize_text(row.get("Description", ""))
    activity = normalize_text(row.get("Activity", "")).upper()
    va_flag = normalize_text(row.get("VA / NVA / NNVA", "")).upper()

    walking_keywords = [
        "walk", "search", "go to", "bring", "carry", "move", "transport",
        "retrieve", "fetch", "get", "pick", "grasp", "reach", "travel"
    ]

    pure_waste_keywords = [
        "inform", "tell", "say", "decide", "discuss", "wait",
        "idle", "think", "confirm", "check again", "look around"
    ]

    if va_flag in {"NVA", "NNVA"}:
        if activity == "T" or contains_keyword(desc, walking_keywords):
            return "Walking Waste"
        if activity in {"W", "D"} or contains_keyword(desc, pure_waste_keywords):
            return "Pure Waste"

    if activity == "T" or contains_keyword(desc, walking_keywords):
        return "Walking Waste"

    if activity in {"W"} or contains_keyword(desc, pure_waste_keywords):
        return "Pure Waste"

    return "Other"


def load_fpc_table(excel_bytes_data, selected_sheet):
    bio = io.BytesIO(excel_bytes_data)

    raw = pd.read_excel(bio, sheet_name=selected_sheet, header=None)
    header_row = find_header_row(raw)

    if header_row is None:
        raise ValueError(
            "Could not find the Flow Process Chart table. Required headers: "
            "Step, Description, Activity, Duration."
        )

    bio = io.BytesIO(excel_bytes_data)
    df = pd.read_excel(bio, sheet_name=selected_sheet, header=header_row)
    df.columns = [normalize_text(c) for c in df.columns]

    df = df.loc[:, ~df.columns.str.contains("^Unnamed", na=False)].copy()

    required_cols = ["Step", "Description", "Activity", "Duration"]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    df = df.dropna(subset=["Description"]).copy()
    df["Description"] = df["Description"].astype(str).str.strip()
    df["Activity"] = df["Activity"].astype(str).str.strip().str.upper()
    df["Duration_sec"] = df["Duration"].apply(duration_to_seconds)

    if "VA / NVA / NNVA" in df.columns:
        df["VA / NVA / NNVA"] = df["VA / NVA / NNVA"].astype(str).str.strip().str.upper()
    else:
        df["VA / NVA / NNVA"] = ""

    df = df[df["Description"].str.strip() != ""].copy()
    df = df[df["Duration_sec"] >= 0].copy()

    return df


def build_analysis(df):
    df = df.copy()
    df["Waste_Type"] = df.apply(classify_waste, axis=1)

    walking_df = df[df["Waste_Type"] == "Walking Waste"].copy()
    walking_df["Waste_Item"] = walking_df["Description"].apply(extract_item_name)

    walking_breakdown = (
        walking_df.groupby("Waste_Item", as_index=False)["Duration_sec"]
        .sum()
        .sort_values("Duration_sec", ascending=False)
    )
    walking_breakdown["Category_Label"] = "Getting " + walking_breakdown["Waste_Item"]

    pure_df = df[df["Waste_Type"] == "Pure Waste"].copy()
    pure_breakdown = (
        pure_df.groupby("Description", as_index=False)["Duration_sec"]
        .sum()
        .sort_values("Duration_sec", ascending=False)
    )

    total_time = df["Duration_sec"].sum()
    walking_total = walking_breakdown["Duration_sec"].sum() if not walking_breakdown.empty else 0.0
    pure_total = pure_breakdown["Duration_sec"].sum() if not pure_breakdown.empty else 0.0
    total_waste = walking_total + pure_total

    walking_pct = (walking_total / total_time * 100) if total_time > 0 else 0
    pure_pct = (pure_total / total_time * 100) if total_time > 0 else 0
    total_waste_pct = (total_waste / total_time * 100) if total_time > 0 else 0

    summary_df = pd.DataFrame({
        "Category": ["Total Time", "Walking Waste", "Pure Waste", "Total Waste"],
        "Time (seconds)": [total_time, walking_total, pure_total, total_waste]
    })

    return {
        "classified_df": df,
        "walking_breakdown": walking_breakdown,
        "pure_breakdown": pure_breakdown,
        "summary_df": summary_df,
        "total_time": total_time,
        "walking_total": walking_total,
        "pure_total": pure_total,
        "total_waste": total_waste,
        "walking_pct": walking_pct,
        "pure_pct": pure_pct,
        "total_waste_pct": total_waste_pct,
    }


def make_chart(summary_df, total_waste_pct):
    fig = px.bar(
        summary_df,
        x="Category",
        y="Time (seconds)",
        title=f"Waste Analysis (Total Waste: {total_waste_pct:.1f}%)",
        text="Time (seconds)"
    )

    fig.update_traces(texttemplate="%{text:.1f}", textposition="outside")
    fig.update_layout(
        xaxis_title="Category",
        yaxis_title="Time (seconds)",
        template="plotly_white",
        height=600,
        margin=dict(l=40, r=40, t=70, b=40)
    )
    return fig


def to_excel_bytes(analysis):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        analysis["classified_df"].to_excel(writer, sheet_name="Classified Data", index=False)
        analysis["summary_df"].to_excel(writer, sheet_name="Summary", index=False)
        analysis["walking_breakdown"].to_excel(writer, sheet_name="Walking Waste", index=False)
        analysis["pure_breakdown"].to_excel(writer, sheet_name="Pure Waste", index=False)
    output.seek(0)
    return output.getvalue()


# =========================================================
# PROCESS
# =========================================================
try:
    df = load_fpc_table(excel_bytes, sheet_name)
    analysis = build_analysis(df)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Time", f"{analysis['total_time']:.1f} sec")
    c2.metric("Walking Waste", f"{analysis['walking_total']:.1f} sec", f"{analysis['walking_pct']:.1f}%")
    c3.metric("Pure Waste", f"{analysis['pure_total']:.1f} sec", f"{analysis['pure_pct']:.1f}%")
    c4.metric("Total Waste", f"{analysis['total_waste']:.1f} sec", f"{analysis['total_waste_pct']:.1f}%")

    fig = make_chart(analysis["summary_df"], analysis["total_waste_pct"])
    st.plotly_chart(fig, use_container_width=True)

    left, right = st.columns(2)

    with left:
        st.subheader("Walking Waste Breakdown")
        if analysis["walking_breakdown"].empty:
            st.info("No walking waste detected.")
        else:
            wb = analysis["walking_breakdown"].copy()
            wb["Percent of Total Time"] = np.where(
                analysis["total_time"] > 0,
                wb["Duration_sec"] / analysis["total_time"] * 100,
                0
            )
            wb = wb.rename(columns={
                "Waste_Item": "Item",
                "Category_Label": "Category",
                "Duration_sec": "Time (sec)"
            })
            wb["Percent of Total Time"] = wb["Percent of Total Time"].round(1)
            wb["Time (sec)"] = wb["Time (sec)"].round(1)
            st.dataframe(wb[["Category", "Item", "Time (sec)", "Percent of Total Time"]], use_container_width=True)

    with right:
        st.subheader("Pure Waste Breakdown")
        if analysis["pure_breakdown"].empty:
            st.info("No pure waste detected.")
        else:
            pb = analysis["pure_breakdown"].copy()
            pb["Percent of Total Time"] = np.where(
                analysis["total_time"] > 0,
                pb["Duration_sec"] / analysis["total_time"] * 100,
                0
            )
            pb = pb.rename(columns={
                "Description": "Description",
                "Duration_sec": "Time (sec)"
            })
            pb["Percent of Total Time"] = pb["Percent of Total Time"].round(1)
            pb["Time (sec)"] = pb["Time (sec)"].round(1)
            st.dataframe(pb[["Description", "Time (sec)", "Percent of Total Time"]], use_container_width=True)

    with st.expander("Show classified Flow Process Chart data"):
        st.dataframe(analysis["classified_df"], use_container_width=True)

    with st.expander("Show summary table"):
        st.dataframe(analysis["summary_df"], use_container_width=True)

    excel_output = to_excel_bytes(analysis)
    st.download_button(
        label="Download Waste Analysis Excel",
        data=excel_output,
        file_name="waste_analysis_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

except Exception as e:
    st.error(f"Error: {e}")
