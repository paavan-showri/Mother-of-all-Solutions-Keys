
import io
import math
import re
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from pandas.io.formats.style import Styler

st.set_page_config(page_title="Impact vs Effort", layout="wide")
st.title("Impact vs Effort Matrix")

BOTTOM_NOTE = (
    "The Impact–Effort Matrix was developed by evaluating each activity using therblig-based motion "
    "analysis and Lean/Six Sigma principles. Impact was determined by whether the motion directly "
    "advanced the work or contributed to process performance, while effort was determined from time "
    "consumption, motion intensity, repetition, resource usage, and process complexity. As a result, "
    "effective motions could fall into either High Impact/Low Effort or High Impact/High Effort, "
    "while ineffective motions could fall into either Low Impact/Low Effort or Low Impact/High Effort "
    "depending on their actual resource demand."
)

def require_upload():
    if "excel_file_bytes" not in st.session_state or "sheet_name" not in st.session_state:
        st.warning("Please upload the Excel file first on the Home page.")
        st.stop()
    return st.session_state["excel_file_bytes"], st.session_state["sheet_name"]

def normalize_text(value):
    if pd.isna(value):
        return ""
    text = str(value).strip().lower()
    return re.sub(r"\s+", " ", text)

def matches_any(text, patterns):
    return any(re.search(pattern, text) for pattern in patterns)

def find_header_row(file_bytes, sheet_name):
    raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None)
    for i, row in raw.iterrows():
        vals = [str(x).strip() for x in row.tolist()]
        if "Step" in vals and "Description" in vals and "Activity" in vals:
            return i
    raise ValueError("Could not find the Flow Process Chart table in the sheet.")

def to_duration_seconds(value):
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        value = float(value)
        if 0 <= value < 1:
            return value * 24 * 60 * 60
        return value
    text = str(value).strip()
    if not text:
        return 0.0
    try:
        return pd.to_timedelta(text).total_seconds()
    except Exception:
        pass
    if re.fullmatch(r"\d{1,2}:\d{2}", text):
        mins, secs = text.split(":")
        return int(mins) * 60 + int(secs)
    if re.fullmatch(r"\d{1,2}:\d{2}:\d{2}", text):
        hrs, mins, secs = text.split(":")
        return int(hrs) * 3600 + int(mins) * 60 + int(secs)
    parsed = pd.to_numeric(text, errors="coerce")
    return 0.0 if pd.isna(parsed) else float(parsed)

def resource_count(value):
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return 0
    parts = re.split(r"\s*&\s*|\s*,\s*|\s*/\s*|\s+and\s+", text, flags=re.IGNORECASE)
    return len([p.strip() for p in parts if p.strip()])

def load_full_data(file_bytes, sheet_name):
    header_row = find_header_row(file_bytes, sheet_name)
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=header_row)
    df.columns = df.columns.astype(str).str.strip()

    df = df.dropna(how="all")
    df = df[df["Step"].notna()].copy()
    df["Step"] = pd.to_numeric(df["Step"], errors="coerce")
    df = df[df["Step"].notna()].copy()
    df["Step"] = df["Step"].astype(int)

    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].astype(str).str.strip()

    if "Description" not in df.columns:
        df["Description"] = ""
    if "Activity Type" not in df.columns and "VA / NVA / NNVA" in df.columns:
        df["Activity Type"] = df["VA / NVA / NNVA"]
    if "Activity Type" not in df.columns:
        df["Activity Type"] = ""
    if "Resources" not in df.columns:
        df["Resources"] = ""
    if "Duration (Sec)" not in df.columns:
        df["Duration (Sec)"] = 0

    df["Activity"] = df["Activity"].astype(str).str.strip().str.upper()
    df["Description"] = df["Description"].astype(str).str.strip()
    df["Activity Type"] = df["Activity Type"].astype(str).str.strip().str.upper()
    df["Resources"] = df["Resources"].astype(str).str.strip()
    df["Duration Seconds"] = df["Duration (Sec)"].apply(to_duration_seconds).fillna(0.0)
    df["Resource Count"] = df["Resources"].apply(resource_count)
    return df.reset_index(drop=True)

NON_VALUE_ACTIVITY_TEXT = {
    "T": "Transport is non-value-added unless it directly supports transformation.",
    "M": "Handling/motion is non-value-added when it is walking, repositioning, searching, or other extra motion.",
    "I": "Inspection is generally non-value-added from the customer perspective.",
    "W": "Waiting/delay is a classic Lean waste.",
    "S": "Storage represents inventory waste.",
    "D": "Decision/planning does not directly transform the product.",
}

def classify_therblig(desc, activity, activity_type):
    text = normalize_text(desc)

    if activity == "W":
        if matches_any(text, [r"\bsearch\b", r"\blook for\b", r"\bfind\b", r"\blocate\b"]):
            return "Search", "Ineffective"
        return "Delay", "Ineffective"

    if activity == "I":
        return "Inspect/Check", "Ineffective"

    if activity == "D":
        if matches_any(text, [r"\bdecide\b", r"\bplan\b", r"\breact\b", r"\bcomplain\b", r"\bdetermine\b"]):
            return "Plan/Decision", "Ineffective"
        return "Non-Value Motion", "Ineffective"

    if activity == "S":
        return "Position", "Ineffective"

    if matches_any(text, [r"\bsearch\b", r"\blook for\b", r"\blocate\b", r"\bfind\b"]):
        return "Search", "Ineffective"
    if matches_any(text, [r"\bwait\b", r"\bdelay\b", r"\bidle\b"]):
        return "Delay", "Ineffective"
    if matches_any(text, [r"\bwalk\b", r"\bgo to\b", r"\breturn\b", r"\bwalk back\b"]):
        return "Walk/Transport Empty", "Ineffective"
    if matches_any(text, [r"\bturn around\b", r"\bturn back\b", r"\breorient\b"]):
        return "Reorient/Extra Motion", "Ineffective"
    if matches_any(text, [r"\breposition\b", r"\bmove .* back\b", r"\bmove knife to new location\b", r"\bmove plate\b", r"\bmove butter plate\b", r"\bmove fruit bowl\b", r"\bmove plate with toast\b"]):
        return "Reposition/Relocate", "Ineffective"
    if matches_any(text, [r"\bposition\b", r"\badjust\b", r"\balign\b", r"\borient\b"]):
        return "Position", "Ineffective"
    if matches_any(text, [r"\bhold\b", r"\bsteady\b"]):
        return "Hold", "Ineffective"

    if matches_any(text, [r"\bgrasp\b", r"\bpick\b", r"\bgrab\b", r"\btake\b"]):
        return "Grasp", "Effective"
    if matches_any(text, [r"\bdrop\b", r"\bplace\b", r"\bput\b", r"\bset down\b", r"\brelease\b"]):
        return "Release Load / Pre-Position", "Effective"
    if matches_any(text, [r"\bopen\b", r"\bclose\b", r"\bturn on\b", r"\bserve\b", r"\bflip\b"]):
        return "Use", "Effective"
    if matches_any(text, [r"\bcut\b", r"\bslice\b"]):
        if matches_any(text, [r"\bstack\b"]):
            return "Assemble", "Effective"
        return "Use / Disassemble", "Effective"
    if matches_any(text, [r"\bstack\b", r"\bassemble\b", r"\bcombine\b", r"\bjoin\b"]):
        return "Assemble", "Effective"
    if matches_any(text, [r"\bmove\b", r"\bcarry\b", r"\bbring\b", r"\btransfer\b"]):
        return "Move", "Effective"

    if activity == "O":
        return "Use", "Effective"
    if activity in {"T", "M"}:
        if activity_type == "NVA":
            return "Reposition/Relocate", "Ineffective"
        return "Move", "Effective"
    return "Unclassified", "Ineffective"

def classify_impact(row):
    activity = row["Activity"]
    activity_type = row["Activity Type"]
    therblig = row["Therblig / Motion Type"]
    motion_class = row["Motion Class"]

    if activity in {"I", "W", "S", "D"}:
        return "Low"
    if therblig in {"Search", "Delay", "Walk/Transport Empty", "Reorient/Extra Motion", "Reposition/Relocate", "Position", "Hold", "Inspect/Check", "Plan/Decision", "Non-Value Motion"}:
        return "Low"
    if motion_class == "Effective":
        return "High"
    if activity_type == "VA":
        return "High"
    return "Low"

def classify_effort(row):
    activity = row["Activity"]
    therblig = row["Therblig / Motion Type"]
    duration = row["Duration Seconds"]
    res_count = row["Resource Count"]

    if therblig in {"Search", "Delay", "Walk/Transport Empty"}:
        return "High"
    if therblig in {"Inspect/Check", "Reorient/Extra Motion", "Non-Value Motion"} and duration >= 3:
        return "High"
    if therblig in {"Use", "Use / Disassemble", "Assemble", "Grasp", "Move", "Release Load / Pre-Position"}:
        if duration >= 3:
            return "High"
        if duration >= 2 and res_count >= 4:
            return "High"
        return "Low"
    if duration >= 4:
        return "High"
    if duration >= 3 and activity in {"I", "D", "M", "T"}:
        return "High"
    return "Low"

def get_quadrant(row):
    if row["Impact"] == "High" and row["Effort"] == "Low":
        return "Quick Wins"
    if row["Impact"] == "High" and row["Effort"] == "High":
        return "Major Projects"
    if row["Impact"] == "Low" and row["Effort"] == "Low":
        return "Fill-Ins"
    return "Time Sinks"

def build_logic_text(row):
    reasons = []
    activity = row["Activity"]
    therblig = row["Therblig / Motion Type"]
    duration = int(round(row["Duration Seconds"]))

    if row["Motion Class"] == "Effective":
        reasons.append(f"{therblig} directly advances the work.")
    else:
        reasons.append(f"{therblig} is treated as non-value-added motion.")
    if activity in NON_VALUE_ACTIVITY_TEXT:
        reasons.append(NON_VALUE_ACTIVITY_TEXT[activity])
    if duration >= 3:
        reasons.append(f"Duration={duration}s increases effort.")
    else:
        reasons.append(f"Duration={duration}s keeps effort lower.")
    if row["Resource Count"] >= 4:
        reasons.append("Multiple resources increase complexity.")
    elif row["Resource Count"] >= 3:
        reasons.append("Several resources are involved.")
    return " ".join(reasons)

def assign_grid_positions(group, center_x, center_y, width, height, ncols):
    group = group.sort_values("Step").copy().reset_index(drop=True)
    n = len(group)
    if n == 0:
        return group
    nrows = math.ceil(n / ncols)
    xs, ys = [], []
    for i in range(n):
        r = i // ncols
        c = i % ncols
        dx = 0 if ncols == 1 else width / (ncols - 1)
        x_offset = 0 if ncols == 1 else -width / 2 + c * dx
        dy = 0 if nrows == 1 else height / (nrows - 1)
        y_offset = 0 if nrows == 1 else height / 2 - r * dy
        xs.append(center_x + x_offset)
        ys.append(center_y + y_offset)
    group["x"] = xs
    group["y"] = ys
    return group

def make_impact_effort_matrix(df, selected_step=None):
    df = df.copy()
    df[["Therblig / Motion Type", "Motion Class"]] = df.apply(
        lambda row: pd.Series(classify_therblig(row["Description"], row["Activity"], row["Activity Type"])),
        axis=1,
    )
    df["Impact"] = df.apply(classify_impact, axis=1)
    df["Effort"] = df.apply(classify_effort, axis=1)
    df["Quadrant"] = df.apply(get_quadrant, axis=1)
    df["Lean / Six Sigma Logic"] = df.apply(build_logic_text, axis=1)
    df = df.sort_values("Step").reset_index(drop=True)

    quadrant_centers = {
        "Quick Wins": (1.0, 2.0),
        "Major Projects": (2.0, 2.0),
        "Fill-Ins": (1.0, 1.0),
        "Time Sinks": (2.0, 1.0),
    }
    quadrant_colors = {
        "Quick Wins": "#6FA8B6",
        "Major Projects": "#8E7CC3",
        "Fill-Ins": "#6D9EEB",
        "Time Sinks": "#EA9999",
    }

    frames = []
    for quad, (cx, cy) in quadrant_centers.items():
        temp = df[df["Quadrant"] == quad].copy()
        temp = assign_grid_positions(temp, cx, cy, 0.72, 0.72, 10)
        frames.append(temp)
    plot_df = pd.concat(frames, ignore_index=True) if frames else df.copy()

    fig = go.Figure()
    for quadrant, color in quadrant_colors.items():
        temp = plot_df[plot_df["Quadrant"] == quadrant].copy().reset_index(drop=True)
        if temp.empty:
            continue

        selected_points = None
        if selected_step is not None and selected_step in temp["Step"].tolist():
            selected_points = [temp.index[temp["Step"] == selected_step][0]]

        fig.add_trace(
            go.Scatter(
                x=temp["x"],
                y=temp["y"],
                mode="markers+text",
                text=temp["Step"].astype(str),
                textposition="middle center",
                textfont=dict(size=11, color="black"),
                marker=dict(size=28, color="rgba(0,0,0,0)", line=dict(color=color, width=2)),
                customdata=temp[[
                    "Step", "Activity", "Therblig / Motion Type", "Motion Class"
                ]].values,
                hovertemplate=(
                    "Activity: %{customdata[1]}"
                    "<br>Therblig: %{customdata[2]}"
                    "<br>Motion Class: %{customdata[3]}<extra></extra>"
                ),
                hoverlabel=dict(bgcolor="white", bordercolor="black", font=dict(color="black", size=13)),
                selectedpoints=selected_points,
                selected=dict(marker=dict(size=34, color="black", opacity=1.0), textfont=dict(color="black")),
                unselected=dict(marker=dict(opacity=0.55), textfont=dict(color="black")),
                showlegend=False,
            )
        )

    fig.add_shape(type="line", x0=1.5, x1=1.5, y0=0.28, y1=2.72, line=dict(color="gray", dash="dash", width=1.5))
    fig.add_shape(type="line", x0=0.28, x1=2.72, y0=1.5, y1=1.5, line=dict(color="gray", dash="dash", width=1.5))
    fig.update_layout(
        title=dict(text="Impact vs Effort Matrix", x=0.5, font=dict(color="black", size=24)),
        xaxis=dict(range=[0.28, 2.72], visible=False),
        yaxis=dict(range=[0.28, 2.72], visible=False),
        plot_bgcolor="#dcdcdc",
        paper_bgcolor="#dcdcdc",
        height=900,
        margin=dict(l=90, r=40, t=80, b=80),
        clickmode="event+select",
        dragmode="select",
        annotations=[
            dict(x=0.52, y=0.01, xref="paper", yref="paper", text="Effort", showarrow=False, font=dict(size=22, color="black")),
            dict(x=0.01, y=0.52, xref="paper", yref="paper", text="Impact", showarrow=False, textangle=-90, font=dict(size=22, color="black")),
            dict(x=0.28, y=0.05, xref="paper", yref="paper", text="Low", showarrow=False, font=dict(size=16, color="black")),
            dict(x=0.78, y=0.05, xref="paper", yref="paper", text="High", showarrow=False, font=dict(size=16, color="black")),
            dict(x=0.04, y=0.28, xref="paper", yref="paper", text="Low", showarrow=False, font=dict(size=16, color="black")),
            dict(x=0.04, y=0.78, xref="paper", yref="paper", text="High", showarrow=False, font=dict(size=16, color="black")),
        ],
    )
    return fig, plot_df

def style_logic_table(df, selected_step=None):
    def highlight_row(row):
        if selected_step is not None and int(row["Step"]) == int(selected_step):
            return ["background-color: #fff3b0; font-weight: bold;"] * len(row)
        return [""] * len(row)

    return df.style.apply(highlight_row, axis=1)

def get_selected_step_from_event(event):
    if not event:
        return None
    selection = event.get("selection", {}) if isinstance(event, dict) else getattr(event, "selection", {})
    points = selection.get("points", []) if hasattr(selection, "get") else []
    if not points:
        return None
    customdata = points[0].get("customdata", [])
    if not customdata:
        return None
    try:
        return int(customdata[0])
    except Exception:
        return None

st.markdown(
    """
**Classification logic used in this chart**

- It uses **Description + Activity + Activity Type + Duration (Sec) + Resources** together.
- **Impact** is driven mainly by whether the motion is an **effective therblig** or an **ineffective/non-value-added therblig**.
- The **Activity** column is explicitly considered. Inspection, delay/waiting, storage, and decision rows are treated as low-impact by default because they are Lean wastes or non-value-added support work.
- For handling / motion rows, the code uses the description to decide whether the step is a useful motion like **grasp, move, release, use** or a waste motion like **walk, search, reposition, turn back, inspect, wait**.
- **Effort** is scored separately from impact using duration, motion waste, repeated walking/searching/inspection, and resource involvement.
- This allows all four quadrants to appear: **Quick Wins, Major Projects, Fill-Ins, and Time Sinks**.
- Lean wastes considered in the logic include **Waiting, Transportation, Motion, Inventory/Storage, Extra Processing, and Defect-like recheck/inspection behavior**.
"""
)

try:
    file_bytes, sheet_name = require_upload()
    df = load_full_data(file_bytes, sheet_name)

    if "impact_effort_selected_step" not in st.session_state:
        st.session_state["impact_effort_selected_step"] = None

    initial_fig, full_table_df = make_impact_effort_matrix(
        df, selected_step=st.session_state["impact_effort_selected_step"]
    )

    chart_event = st.plotly_chart(
        initial_fig,
        use_container_width=True,
        theme=None,
        key="impact_effort_chart",
        on_select="rerun",
        selection_mode="points",
        config={"scrollZoom": False},
    )

    clicked_step = get_selected_step_from_event(chart_event)
    if clicked_step is not None:
        st.session_state["impact_effort_selected_step"] = clicked_step

    selected_step = st.session_state.get("impact_effort_selected_step")
    fig, table_df = make_impact_effort_matrix(df, selected_step=selected_step)

    if clicked_step is not None:
        st.rerun()

    st.markdown(
        f"""
<div style="margin-top: 10px; font-size: 14px; color: black; text-align: justify;">
{BOTTOM_NOTE}
</div>
""",
        unsafe_allow_html=True,
    )

    st.markdown("### Activity classification table")

    filter_col1, filter_col2, filter_col3 = st.columns([1, 1, 1.2])
    with filter_col1:
        impact_filter = st.multiselect(
            "Filter Impact",
            options=["Low", "High"],
            default=["Low", "High"],
        )
    with filter_col2:
        effort_filter = st.multiselect(
            "Filter Effort",
            options=["Low", "High"],
            default=["Low", "High"],
        )
    with filter_col3:
        if st.button("Clear marker selection", use_container_width=True):
            st.session_state["impact_effort_selected_step"] = None
            st.rerun()

    display_df = table_df.copy()
    display_df = display_df[
        display_df["Impact"].isin(impact_filter) &
        display_df["Effort"].isin(effort_filter)
    ].copy()

    display_df = display_df.sort_values(["Step", "Impact", "Effort"]).reset_index(drop=True)
    display_df["Duration"] = display_df["Duration Seconds"].round(0).astype(int).astype(str) + " s"

    table_cols = [
        "Step",
        "Description",
        "Activity",
        "Therblig / Motion Type",
        "Motion Class",
        "Impact",
        "Effort",
        "Quadrant",
        "Duration",
        "Lean / Six Sigma Logic",
    ]
    display_df = display_df[table_cols].rename(columns={
        "Therblig / Motion Type": "Therblig",
        "Lean / Six Sigma Logic": "Logic",
    })

    styled_table = style_logic_table(display_df, selected_step=selected_step)
    st.dataframe(styled_table, use_container_width=True, hide_index=True)

    if selected_step is not None:
        st.caption(f"Selected marker: Step {selected_step}. The matching row is highlighted below.")
    else:
        st.caption("Click a marker once to select it and highlight the matching row in the table.")
except Exception as e:
    st.error(f"Error: {e}")
