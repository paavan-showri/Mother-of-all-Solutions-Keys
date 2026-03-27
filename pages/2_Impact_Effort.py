import io
import math
import re
from typing import List, Tuple

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Impact vs Effort", layout="wide")
st.title("Impact vs Effort Matrix")

BOTTOM_NOTE = (
    "This page uses a recursive Impact–Effort decomposition instead of a flat 2×2 matrix. "
    "Each region is repeatedly split into four sub-regions like a quadtree, so dense clusters of activities "
    "become visible at finer levels. Impact is estimated from therblig effectiveness, value-added behavior, "
    "and activity meaning, while effort is estimated from duration, motion burden, and resource complexity."
)


# -------------------------
# Workbook helpers
# -------------------------
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


# -------------------------
# Rule engine
# -------------------------
NON_VALUE_ACTIVITY_TEXT = {
    "T": "Transport is non-value-added unless it directly supports transformation.",
    "M": "Handling or motion is non-value-added when it is walking, repositioning, searching, or extra motion.",
    "I": "Inspection is generally non-value-added from the customer perspective.",
    "W": "Waiting or delay is a classic Lean waste.",
    "S": "Storage represents inventory waste.",
    "D": "Decision or planning does not directly transform the product.",
}


THERBLIG_IMPACT_BASE = {
    "Search": 0.12,
    "Delay": 0.05,
    "Walk/Transport Empty": 0.15,
    "Reorient/Extra Motion": 0.18,
    "Reposition/Relocate": 0.22,
    "Position": 0.25,
    "Hold": 0.20,
    "Inspect/Check": 0.18,
    "Plan/Decision": 0.16,
    "Non-Value Motion": 0.14,
    "Grasp": 0.70,
    "Release Load / Pre-Position": 0.68,
    "Use": 0.88,
    "Use / Disassemble": 0.82,
    "Assemble": 0.92,
    "Move": 0.72,
    "Unclassified": 0.30,
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
    if therblig in {
        "Search", "Delay", "Walk/Transport Empty", "Reorient/Extra Motion", "Reposition/Relocate",
        "Position", "Hold", "Inspect/Check", "Plan/Decision", "Non-Value Motion"
    }:
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


def compute_scores(df):
    df = df.copy()

    max_duration = max(float(df["Duration Seconds"].max()), 1.0)
    max_resources = max(float(df["Resource Count"].max()), 1.0)

    duration_norm = (df["Duration Seconds"] / max_duration).clip(0, 1)
    resource_norm = (df["Resource Count"] / max_resources).clip(0, 1)

    waste_bonus = df["Therblig / Motion Type"].map({
        "Search": 0.22,
        "Delay": 0.30,
        "Walk/Transport Empty": 0.22,
        "Inspect/Check": 0.14,
        "Reorient/Extra Motion": 0.16,
        "Reposition/Relocate": 0.14,
        "Plan/Decision": 0.12,
        "Non-Value Motion": 0.14,
    }).fillna(0.0)

    motion_penalty = df["Motion Class"].map({"Effective": 0.10, "Ineffective": 0.0}).fillna(0.0)
    va_bonus = df["Activity Type"].map({"VA": 0.08, "NNVA": 0.03, "NVA": 0.0}).fillna(0.0)

    impact_base = df["Therblig / Motion Type"].map(THERBLIG_IMPACT_BASE).fillna(0.30)
    df["Impact Score"] = (impact_base + va_bonus - 0.05 * duration_norm).clip(0.02, 0.98)

    df["Effort Score"] = (
        0.62 * duration_norm +
        0.20 * resource_norm +
        waste_bonus -
        motion_penalty
    ).clip(0.02, 0.98)

    return df


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
    reasons.append(
        f"Scores → Impact={row['Impact Score']:.2f}, Effort={row['Effort Score']:.2f}."
    )
    return " ".join(reasons)


# -------------------------
# Quadtree helpers
# -------------------------
Cell = Tuple[pd.DataFrame, float, float, float, float, int]


def quadtree_partition(
    df: pd.DataFrame,
    x_min: float,
    x_max: float,
    y_min: float,
    y_max: float,
    depth: int,
    max_depth: int,
    min_points: int,
) -> List[Cell]:
    if df.empty:
        return []
    if len(df) <= min_points or depth >= max_depth:
        return [(df.copy(), x_min, x_max, y_min, y_max, depth)]

    x_mid = (x_min + x_max) / 2
    y_mid = (y_min + y_max) / 2

    upper_left = df[(df["Effort Score"] <= x_mid) & (df["Impact Score"] > y_mid)]
    upper_right = df[(df["Effort Score"] > x_mid) & (df["Impact Score"] > y_mid)]
    lower_left = df[(df["Effort Score"] <= x_mid) & (df["Impact Score"] <= y_mid)]
    lower_right = df[(df["Effort Score"] > x_mid) & (df["Impact Score"] <= y_mid)]

    quadrants = [
        (upper_left, x_min, x_mid, y_mid, y_max),
        (upper_right, x_mid, x_max, y_mid, y_max),
        (lower_left, x_min, x_mid, y_min, y_mid),
        (lower_right, x_mid, x_max, y_min, y_mid),
    ]

    result: List[Cell] = []
    for sub_df, sx0, sx1, sy0, sy1 in quadrants:
        if sub_df.empty:
            continue
        result.extend(quadtree_partition(sub_df, sx0, sx1, sy0, sy1, depth + 1, max_depth, min_points))
    return result


def assign_positions_in_cell(df, xmin, xmax, ymin, ymax):
    df = df.sort_values(["Step", "Effort Score", "Impact Score"]).copy().reset_index(drop=True)
    n = len(df)
    if n == 0:
        return df

    cols = max(1, math.ceil(math.sqrt(n)))
    rows = max(1, math.ceil(n / cols))

    x_pad = min(0.012, (xmax - xmin) * 0.12)
    y_pad = min(0.012, (ymax - ymin) * 0.12)
    usable_x0, usable_x1 = xmin + x_pad, xmax - x_pad
    usable_y0, usable_y1 = ymin + y_pad, ymax - y_pad

    xs, ys = [], []
    for i in range(n):
        r = i // cols
        c = i % cols
        x = usable_x0 + (c + 0.5) * max((usable_x1 - usable_x0) / cols, 1e-6)
        y = usable_y1 - (r + 0.5) * max((usable_y1 - usable_y0) / rows, 1e-6)
        xs.append(x)
        ys.append(y)

    df["x"] = xs
    df["y"] = ys
    return df


def density_fill(count, max_count):
    intensity = count / max(max_count, 1)
    alpha = 0.08 + 0.28 * intensity
    return f"rgba(70, 130, 180, {alpha:.3f})"


# -------------------------
# Chart builder
# -------------------------
def make_quadtree_matrix(df, selected_step=None, max_depth=3, min_points=6):
    df = df.copy()
    df[["Therblig / Motion Type", "Motion Class"]] = df.apply(
        lambda row: pd.Series(classify_therblig(row["Description"], row["Activity"], row["Activity Type"])),
        axis=1,
    )
    df["Impact"] = df.apply(classify_impact, axis=1)
    df["Effort"] = df.apply(classify_effort, axis=1)
    df["Quadrant"] = df.apply(get_quadrant, axis=1)
    df = compute_scores(df)
    df["Lean / Six Sigma Logic"] = df.apply(build_logic_text, axis=1)
    df = df.sort_values("Step").reset_index(drop=True)

    cells = quadtree_partition(df, 0.0, 1.0, 0.0, 1.0, depth=0, max_depth=max_depth, min_points=min_points)

    positioned_frames = []
    leaf_meta = []
    max_count = max((len(cell_df) for cell_df, *_ in cells), default=1)

    for cell_df, x0, x1, y0, y1, depth in cells:
        temp = assign_positions_in_cell(cell_df, x0, x1, y0, y1)
        temp["Cell X0"] = x0
        temp["Cell X1"] = x1
        temp["Cell Y0"] = y0
        temp["Cell Y1"] = y1
        temp["Tree Depth"] = depth
        positioned_frames.append(temp)
        leaf_meta.append(
            {
                "x0": x0,
                "x1": x1,
                "y0": y0,
                "y1": y1,
                "depth": depth,
                "count": len(cell_df),
                "fillcolor": density_fill(len(cell_df), max_count),
            }
        )

    plot_df = pd.concat(positioned_frames, ignore_index=True) if positioned_frames else df.copy()

    quadrant_outline_colors = {
        "Quick Wins": "#2a9d8f",
        "Major Projects": "#7b61ff",
        "Fill-Ins": "#3a86ff",
        "Time Sinks": "#ef476f",
    }

    fig = go.Figure()

    for meta in leaf_meta:
        fig.add_shape(
            type="rect",
            x0=meta["x0"], x1=meta["x1"], y0=meta["y0"], y1=meta["y1"],
            line=dict(color="rgba(70,70,70,0.45)", width=1),
            fillcolor=meta["fillcolor"],
            layer="below",
        )

    for quadrant, color in quadrant_outline_colors.items():
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
                textfont=dict(size=10, color="black"),
                marker=dict(size=28, color="rgba(255,255,255,0.65)", line=dict(color=color, width=2)),
                customdata=temp[[
                    "Step", "Description", "Activity", "Impact", "Effort", "Impact Score", "Effort Score",
                    "Quadrant", "Tree Depth", "Cell X0", "Cell X1", "Cell Y0", "Cell Y1"
                ]].values,
                hovertemplate=(
                    "Step: %{customdata[0]}"
                    "<br>Description: %{customdata[1]}"
                    "<br>Activity: %{customdata[2]}"
                    "<br>Impact: %{customdata[3]} (%{customdata[5]:.2f})"
                    "<br>Effort: %{customdata[4]} (%{customdata[6]:.2f})"
                    "<br>Quadrant: %{customdata[7]}"
                    "<br>Tree depth: %{customdata[8]}"
                    "<extra></extra>"
                ),
                hoverlabel=dict(bgcolor="white", bordercolor="black", font=dict(color="black", size=12)),
                selectedpoints=selected_points,
                selected=dict(marker=dict(size=34, color="rgba(255,255,0,0.95)", opacity=1.0)),
                unselected=dict(marker=dict(opacity=0.65)),
                showlegend=False,
            )
        )

    # Main 2x2 split lines
    fig.add_shape(type="line", x0=0.5, x1=0.5, y0=0.0, y1=1.0, line=dict(color="gray", dash="dash", width=1.5))
    fig.add_shape(type="line", x0=0.0, x1=1.0, y0=0.5, y1=0.5, line=dict(color="gray", dash="dash", width=1.5))

    fig.update_layout(
        title=dict(text="Recursive Impact vs Effort Matrix", x=0.5, font=dict(color="black", size=24)),
        xaxis=dict(range=[0, 1], visible=False, fixedrange=True),
        yaxis=dict(range=[0, 1], visible=False, fixedrange=True),
        plot_bgcolor="#e6e6e6",
        paper_bgcolor="#e6e6e6",
        height=920,
        margin=dict(l=90, r=40, t=80, b=90),
        clickmode="event+select",
        dragmode="select",
        annotations=[
            dict(x=0.52, y=0.01, xref="paper", yref="paper", text="Effort", showarrow=False, font=dict(size=22, color="black")),
            dict(x=0.01, y=0.52, xref="paper", yref="paper", text="Impact", showarrow=False, textangle=-90, font=dict(size=22, color="black")),
            dict(x=0.25, y=0.05, xref="paper", yref="paper", text="Low", showarrow=False, font=dict(size=16, color="black")),
            dict(x=0.75, y=0.05, xref="paper", yref="paper", text="High", showarrow=False, font=dict(size=16, color="black")),
            dict(x=0.04, y=0.25, xref="paper", yref="paper", text="Low", showarrow=False, font=dict(size=16, color="black")),
            dict(x=0.04, y=0.75, xref="paper", yref="paper", text="High", showarrow=False, font=dict(size=16, color="black")),
            dict(x=0.25, y=0.97, xref="paper", yref="paper", text="Quick Wins", showarrow=False, font=dict(size=15, color="#2a9d8f")),
            dict(x=0.75, y=0.97, xref="paper", yref="paper", text="Major Projects", showarrow=False, font=dict(size=15, color="#7b61ff")),
            dict(x=0.25, y=0.12, xref="paper", yref="paper", text="Fill-Ins", showarrow=False, font=dict(size=15, color="#3a86ff")),
            dict(x=0.75, y=0.12, xref="paper", yref="paper", text="Time Sinks", showarrow=False, font=dict(size=15, color="#ef476f")),
        ],
    )
    return fig, plot_df, leaf_meta


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
**How this version works**

- The page still classifies each row using **Description + Activity + Activity Type + Duration (Sec) + Resources**.
- Then it converts each activity into a continuous **Impact Score** and **Effort Score** between 0 and 1.
- The matrix is recursively split into **4 sub-quadrants**, and each occupied region is split again up to the selected tree depth.
- Darker cells indicate **higher local density**, so the chart behaves like a **quadtree + heatmap**.
- This helps separate crowded activities inside the same main quadrant instead of stacking them in one flat block.
"""
)

try:
    file_bytes, sheet_name = require_upload()
    df = load_full_data(file_bytes, sheet_name)

    if "impact_effort_selected_step" not in st.session_state:
        st.session_state["impact_effort_selected_step"] = None

    control_col1, control_col2, control_col3 = st.columns([1, 1, 2])
    with control_col1:
        max_depth = st.slider("Quadtree depth", min_value=1, max_value=6, value=3, step=1)
    with control_col2:
        min_points = st.slider("Min points per cell", min_value=1, max_value=15, value=6, step=1)
    with control_col3:
        st.caption(
            "Higher depth gives finer subdivision. Lower min points forces more splitting. "
            "Use both controls to tune cluster visibility."
        )

    initial_fig, _, _ = make_quadtree_matrix(
        df,
        selected_step=st.session_state["impact_effort_selected_step"],
        max_depth=max_depth,
        min_points=min_points,
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
    fig, table_df, leaf_meta = make_quadtree_matrix(
        df,
        selected_step=selected_step,
        max_depth=max_depth,
        min_points=min_points,
    )

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

    summary_col1, summary_col2, summary_col3 = st.columns(3)
    with summary_col1:
        st.metric("Activities", len(table_df))
    with summary_col2:
        st.metric("Leaf cells", len(leaf_meta))
    with summary_col3:
        st.metric("Max tree depth used", max((meta["depth"] for meta in leaf_meta), default=0))

    st.markdown("### Activity classification table")

    filter_col1, filter_col2, filter_col3 = st.columns([1, 1, 1.2])
    with filter_col1:
        impact_filter = st.multiselect("Filter Impact", options=["Low", "High"], default=["Low", "High"])
    with filter_col2:
        effort_filter = st.multiselect("Filter Effort", options=["Low", "High"], default=["Low", "High"])
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
    display_df["Impact Score"] = display_df["Impact Score"].round(2)
    display_df["Effort Score"] = display_df["Effort Score"].round(2)

    table_cols = [
        "Step",
        "Description",
        "Activity",
        "Therblig / Motion Type",
        "Motion Class",
        "Impact",
        "Effort",
        "Impact Score",
        "Effort Score",
        "Quadrant",
        "Tree Depth",
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
