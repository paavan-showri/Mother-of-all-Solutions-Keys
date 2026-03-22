import io
from datetime import datetime, time
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Resource Utilization", layout="wide")
st.title("Resource Utilization Chart")


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


def load_resource_data(file_bytes, sheet_name):
    header_row = find_header_row(file_bytes, sheet_name)
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=header_row)
    df.columns = df.columns.astype(str).str.strip()

    needed_cols = [
        "Step",
        "Description",
        "Activity",
        "Start time",
        "End time",
        "Duration (Sec)",
        "Resources",
    ]
    missing = [c for c in needed_cols if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    return df[needed_cols].dropna(how="all").reset_index(drop=True)


def to_seconds(val):
    if pd.isna(val):
        return None

    if isinstance(val, (int, float)):
        return int(round(float(val) * 86400))

    if isinstance(val, pd.Timestamp):
        return val.hour * 3600 + val.minute * 60 + val.second

    if isinstance(val, datetime):
        return val.hour * 3600 + val.minute * 60 + val.second

    if isinstance(val, time):
        return val.hour * 3600 + val.minute * 60 + val.second

    s = str(val).strip()
    for fmt in ("%M:%S", "%H:%M:%S", "%I:%M:%S %p", "%I:%M %p"):
        try:
            t = datetime.strptime(s, fmt)
            return t.hour * 3600 + t.minute * 60 + t.second
        except ValueError:
            pass

    parts = s.split(":")
    try:
        parts = [int(p) for p in parts]
        if len(parts) == 2:
            return parts[0] * 60 + parts[1]
        if len(parts) == 3:
            return parts[0] * 3600 + parts[1] * 60 + parts[2]
    except Exception:
        pass

    return None


def format_mmss(val):
    if pd.isna(val):
        return ""
    sec = int(val)
    m = sec // 60
    s = sec % 60
    return f"{m}:{s:02d}"


def normalize_resource_name(name):
    name = str(name).strip()
    if name == "Breader":
        return "Toaster"
    return name


def extract_resources(resource_text):
    resources = set()
    text = str(resource_text).replace(",", "&")
    for part in text.split("&"):
        p = normalize_resource_name(part)
        if p:
            resources.add(p)
    return resources


def make_resource_chart(df):
    df = df.copy()

    df["StartSec"] = df["Start time"].apply(to_seconds)
    df["EndSec"] = df["End time"].apply(to_seconds)
    df["DurationSec"] = df["Duration (Sec)"].apply(to_seconds)

    mask_bad = df["DurationSec"].isna()
    if mask_bad.any():
        numeric_duration = pd.to_numeric(df.loc[mask_bad, "Duration (Sec)"], errors="coerce")
        df.loc[mask_bad, "DurationSec"] = numeric_duration

    df.loc[df["DurationSec"].isna(), "DurationSec"] = df["EndSec"] - df["StartSec"]

    df = df.dropna(subset=["StartSec", "DurationSec", "Resources"]).copy()
    df = df[df["DurationSec"] > 0].copy()

    df["StartSec"] = df["StartSec"].astype(int)
    df["DurationSec"] = df["DurationSec"].astype(int)
    df["EndSecCalc"] = df["StartSec"] + df["DurationSec"]
    df["StartLabel"] = df["StartSec"].apply(format_mmss)
    df["EndLabel"] = df["EndSecCalc"].apply(format_mmss)
    df["TimeRange"] = df["StartLabel"] + " - " + df["EndLabel"]

    resources = ["Man", "Bread", "Plate", "Toaster", "Butter", "Knife", "Toast"]

    active_color = "#2ca02c"
    inactive_color = "#f40909"

    fig = go.Figure()

    for res in resources:
        active_rows = []
        inactive_rows = []

        for _, row in df.iterrows():
            mentioned = extract_resources(row["Resources"])

            row_payload = {
                "Step": row["Step"],
                "Description": row["Description"],
                "Activity": row["Activity"],
                "TimeRange": row["TimeRange"],
                "StartSec": row["StartSec"],
                "DurationSec": row["DurationSec"],
            }

            if res in mentioned:
                active_rows.append(row_payload)
            else:
                inactive_rows.append(row_payload)

        if inactive_rows:
            x_inactive = [r["DurationSec"] for r in inactive_rows]
            base_inactive = [r["StartSec"] for r in inactive_rows]
            custom_inactive = [
                [r["Step"], r["Description"], r["Activity"], r["TimeRange"]]
                for r in inactive_rows
            ]

            fig.add_trace(
                go.Bar(
                    x=x_inactive,
                    y=[res] * len(inactive_rows),
                    base=base_inactive,
                    orientation="h",
                    name="Inactive",
                    legendgroup="Inactive",
                    showlegend=(res == resources[0]),
                    marker=dict(color=inactive_color),
                    customdata=custom_inactive,
                    hovertemplate=(
                        "<b>Resource:</b> %{y}<br>"
                        "<b>Step:</b> %{customdata[0]}<br>"
                        "<b>Description:</b> %{customdata[1]}<br>"
                        "<b>Activity:</b> %{customdata[2]}<br>"
                        "<b>Time:</b> %{customdata[3]}"
                        "<extra></extra>"
                    ),
                )
            )

        if active_rows:
            x_active = [r["DurationSec"] for r in active_rows]
            base_active = [r["StartSec"] for r in active_rows]
            custom_active = [
                [r["Step"], r["Description"], r["Activity"], r["TimeRange"]]
                for r in active_rows
            ]

            fig.add_trace(
                go.Bar(
                    x=x_active,
                    y=[res] * len(active_rows),
                    base=base_active,
                    orientation="h",
                    name="Active",
                    legendgroup="Active",
                    showlegend=(res == resources[0]),
                    marker=dict(color=active_color),
                    customdata=custom_active,
                    hovertemplate=(
                        "<b>Resource:</b> %{y}<br>"
                        "<b>Step:</b> %{customdata[0]}<br>"
                        "<b>Description:</b> %{customdata[1]}<br>"
                        "<b>Activity:</b> %{customdata[2]}<br>"
                        "<b>Time:</b> %{customdata[3]}"
                        "<extra></extra>"
                    ),
                )
            )

    max_time = int(df["EndSecCalc"].max())
    upper_limit = ((max_time // 30) + 1) * 30
    tick_vals = list(range(0, upper_limit + 1, 30))
    tick_text = [format_mmss(v) for v in tick_vals]

    fig.update_layout(
        barmode="overlay",
        height=650,
        xaxis=dict(
            title="Time (m:ss)",
            range=[0, upper_limit],
            tickmode="array",
            tickvals=tick_vals,
            ticktext=tick_text,
            showgrid=True,
            gridcolor="rgba(0,0,0,0.15)",
        ),
        yaxis=dict(
            title="Resources",
            categoryorder="array",
            categoryarray=resources[::-1],
        ),
        title="Resource Utilization Chart",
        hoverlabel=dict(bgcolor="white", font_size=16, font_color="black"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )

    return fig


try:
    file_bytes, sheet_name = require_upload()
    df = load_resource_data(file_bytes, sheet_name)
    fig = make_resource_chart(df)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": True})
except Exception as e:
    st.error(f"Error: {e}")
