import io
from datetime import datetime, time, timedelta
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.path import Path
from matplotlib.lines import Line2D
import streamlit as st

st.set_page_config(page_title="Scatter Plot", layout="wide")
st.title("Scatter Plot: Total Time vs Frequency by Activity")


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


def handling_marker():
    theta = np.linspace(np.pi / 2, 3 * np.pi / 2, 40)
    r = 0.55
    cx, cy = -0.15, 0.0

    arc_x = cx + r * np.cos(theta)
    arc_y = cy + r * np.sin(theta)

    verts = []
    codes = []

    verts.append((arc_x[0], arc_y[0]))
    codes.append(Path.MOVETO)

    for x, y in zip(arc_x[1:], arc_y[1:]):
        verts.append((x, y))
        codes.append(Path.LINETO)

    verts.append((0.15, -0.25))
    codes.append(Path.LINETO)
    verts.append((0.65, 0.0))
    codes.append(Path.LINETO)
    verts.append((0.15, 0.25))
    codes.append(Path.LINETO)
    verts.append((arc_x[0], arc_y[0]))
    codes.append(Path.CLOSEPOLY)

    return Path(verts, codes)


def make_scatter_chart(summary_df):
    scatter_df = summary_df.copy()

    activity_names = {
        "O": "Operation",
        "T": "Transport",
        "M": "Handling",
        "I": "Inspection",
        "W": "Delay",
        "S": "Storage",
        "D": "Decision"
    }

    activity_markers = {
        "O": "o",
        "T": r"$\rightarrow$",
        "M": handling_marker(),
        "I": "s",
        "W": r"$D$",
        "S": "v",
        "D": "D"
    }

    fig, ax = plt.subplots(figsize=(11, 4))
    plot_marker_size = 100

    for _, row in scatter_df.iterrows():
        x = row["Frequency_of_Occurrence"]
        y = row["Total_Time_of_Occurrence"]
        act = row["Activity"]

        ax.scatter(
            x,
            y,
            s=plot_marker_size,
            marker=activity_markers.get(act, "o"),
            facecolors="white",
            edgecolors="black",
            linewidths=1.5
        )

    ax.set_title("Scatter Plot: Total Time vs Frequency by Activity", fontsize=16)
    ax.set_xlabel("Frequency of Occurrence", fontsize=13)
    ax.set_ylabel("Total Time of Occurrence (seconds)", fontsize=13)
    ax.grid(True, linestyle="--", alpha=0.4)

    if not scatter_df.empty:
        x_min = scatter_df["Frequency_of_Occurrence"].min()
        x_max = scatter_df["Frequency_of_Occurrence"].max()
        y_min = scatter_df["Total_Time_of_Occurrence"].min()
        y_max = scatter_df["Total_Time_of_Occurrence"].max()

        ax.set_xlim(max(0, x_min - 2), x_max + 3)
        ax.set_ylim(max(0, y_min - 3), y_max + 5)

    legend_order = ["D", "I", "M", "O", "S", "T", "W"]
    legend_handles = []

    for code in legend_order:
        if code in scatter_df["Activity"].values:
            legend_handles.append(
                Line2D(
                    [0], [0],
                    marker=activity_markers[code],
                    linestyle="None",
                    markerfacecolor="white",
                    markeredgecolor="black",
                    markeredgewidth=1.3,
                    markersize=9,
                    label=f"{code} - {activity_names[code]}"
                )
            )

    ax.legend(
        handles=legend_handles,
        title="Activity",
        loc="center left",
        bbox_to_anchor=(1.02, 0.5),
        frameon=True,
        fontsize=11,
        title_fontsize=12
    )

    plt.tight_layout(rect=[0, 0, 0.82, 1])
    return fig, scatter_df


try:
    file_bytes, sheet_name = require_upload()
    df = load_full_data(file_bytes, sheet_name)
    summary = build_activity_summary(df)
    fig, scatter_df = make_scatter_chart(summary)
    st.pyplot(fig, use_container_width=True)
    with st.expander("Show table"):
        st.dataframe(scatter_df, use_container_width=True)
except Exception as e:
    st.error(f"Error: {e}")
