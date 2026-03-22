import io
from datetime import datetime, time, timedelta
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st

st.set_page_config(page_title="Pareto Frequency", layout="wide")
st.title("Pareto Analysis by Frequency")


def require_upload():
    if "excel_file_bytes" not in st.session_state or "sheet_name" not in st.session_state:
        st.warning("Please upload the Excel file first on the Home page.")
        st.stop()
    return st.session_state["excel_file_bytes"], st.session_state["sheet_name"]


def find_header_row(file_bytes, sheet_name):
    raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None)
    for i, row in raw.iterrows():
        vals = [str(x).strip() for x in row.tolist()]
        if "Step" in vals and "Description" in vals and "Activity" in vals:
            return i
    raise ValueError("Could not find the Flow Process Chart table in the sheet.")


def excel_time_to_seconds(val):
    if val is None or pd.isna(val):
        return 0
    if isinstance(val, timedelta):
        return int(val.total_seconds())
    if isinstance(val, pd.Timestamp):
        return val.hour * 3600 + val.minute * 60 + val.second
    if isinstance(val, datetime):
        return val.hour * 3600 + val.minute * 60 + val.second
    if isinstance(val, time):
        return val.hour * 3600 + val.minute * 60 + val.second
    if isinstance(val, (int, float)):
        if 0 <= val < 1:
            return int(round(val * 24 * 3600))
        return int(round(val))

    s = str(val).strip()
    for fmt in ("%M:%S", "%H:%M:%S", "%I:%M:%S %p", "%I:%M %p"):
        try:
            t = datetime.strptime(s, fmt)
            return t.hour * 3600 + t.minute * 60 + t.second
        except ValueError:
            pass

    parts = s.split(":")
    try:
        parts = [float(p) for p in parts]
        if len(parts) == 3:
            h, m, sec = parts
            return int(h * 3600 + m * 60 + sec)
        if len(parts) == 2:
            m, sec = parts
            return int(m * 60 + sec)
        if len(parts) == 1:
            return int(parts[0])
    except Exception:
        return 0
    return 0


def load_full_data(file_bytes, sheet_name):
    header_row = find_header_row(file_bytes, sheet_name)
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=header_row)
    df.columns = df.columns.astype(str).str.strip()
    df = df.dropna(how="all")
    df = df[df["Step"].notna()].copy()
    df["Step"] = pd.to_numeric(df["Step"], errors="coerce")
    df = df[df["Step"].notna()].copy()
    df["Step"] = df["Step"].astype(int)
    df["Activity"] = df["Activity"].astype(str).str.strip().str.upper()
    return df.reset_index(drop=True)


def build_activity_summary(df):
    duration_col = None
    for col in df.columns:
        if "Duration" in str(col):
            duration_col = col
            break
    if duration_col is None:
        raise ValueError("No Duration column found.")

    duration_seconds = []
    for _, row in df.iterrows():
        dur_sec = excel_time_to_seconds(row.get(duration_col))
        if dur_sec == 0:
            start_sec = excel_time_to_seconds(row.get("Start time"))
            end_sec = excel_time_to_seconds(row.get("End time"))
            dur_sec = max(0, end_sec - start_sec)
        duration_seconds.append(dur_sec)

    df = df.copy()
    df["Duration_seconds"] = duration_seconds

    summary = df.groupby("Activity", as_index=False).agg(
        Frequency_of_Occurrence=("Activity", "count"),
        Total_Time_of_Occurrence=("Duration_seconds", "sum")
    )

    activity_order = ["O", "T", "M", "I", "W", "S", "D"]
    all_symbols_df = pd.DataFrame({"Activity": activity_order})
    summary = all_symbols_df.merge(summary, on="Activity", how="left").fillna(0)

    summary["Frequency_of_Occurrence"] = summary["Frequency_of_Occurrence"].astype(int)
    summary["Total_Time_of_Occurrence"] = summary["Total_Time_of_Occurrence"].astype(int)

    names = {"O": "Operation", "T": "Transport", "M": "Handling", "I": "Inspection", "W": "Delay", "S": "Storage", "D": "Decision"}
    summary["Activity Name"] = summary["Activity"].map(names)
    return summary


def make_pareto_frequency_chart(summary):
    pareto = summary.sort_values(by="Frequency_of_Occurrence", ascending=False).reset_index(drop=True)
    total = pareto["Frequency_of_Occurrence"].sum()
    pareto["Percent"] = (pareto["Frequency_of_Occurrence"] / total * 100) if total > 0 else 0
    pareto["Cumulative Percent"] = pareto["Percent"].cumsum()

    labels = pareto["Activity"] + " - " + pareto["Activity Name"]

    fig, ax1 = plt.subplots(figsize=(11, 5))
    bars = ax1.bar(labels, pareto["Frequency_of_Occurrence"])
    ax1.set_xlabel("Activity Symbol")
    ax1.set_ylabel("Frequency of Occurrence")
    ax1.set_title("Pareto Analysis by Frequency of Occurrence")
    ax1.tick_params(axis="x", rotation=45)

    for bar in bars:
        h = bar.get_height()
        ax1.text(bar.get_x() + bar.get_width() / 2, h + 0.1, f"{int(h)}", ha="center", va="bottom")

    ax2 = ax1.twinx()
    ax2.plot(labels, pareto["Cumulative Percent"], marker="o")
    ax2.set_ylabel("Cumulative Percentage (%)")
    ax2.set_ylim(0, 110)

    for i, val in enumerate(pareto["Cumulative Percent"]):
        ax2.text(i, val + 2, f"{val:.1f}%", ha="center", va="bottom", fontsize=9)

    plt.tight_layout()
    return fig, pareto


try:
    file_bytes, sheet_name = require_upload()
    df = load_full_data(file_bytes, sheet_name)
    summary = build_activity_summary(df)
    fig, pareto = make_pareto_frequency_chart(summary)
    st.pyplot(fig, use_container_width=True)
    with st.expander("Show table"):
        st.dataframe(pareto, use_container_width=True)
except Exception as e:
    st.error(f"Error: {e}")
