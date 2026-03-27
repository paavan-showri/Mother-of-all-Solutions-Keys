import io
import math
import re
from collections import defaultdict

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Impact vs Effort", layout="wide")
st.title("Recursive Impact vs Effort Matrix")

BOTTOM_NOTE = (
    "This view keeps the original Lean impact–effort logic, but it lays the activities inside a "
    "recursive 2×2 structure. The full matrix is first divided into Quick Wins, Major Projects, "
    "Fill-Ins, and Time Sinks. Then each of those quadrants is divided again using the same "
    "impact-versus-effort split, and so on. Activities are therefore positioned inside nested "
    "sub-quadrants instead of being dumped into only one large box."
)

QUADRANT_COLORS = {
    "Quick Wins": "#0f9d9a",
    "Major Projects": "#635bff",
    "Fill-Ins": "#2f78ff",
    "Time Sinks": "#ff4d6d",
}

CELL_LINE = "rgba(95, 120, 150, 0.72)"
CELL_FILL = "rgba(68, 95, 127, 0.08)"
PLOT_BG = "#dcdcdc"


# -------------------------
# Workbook loading helpers
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
# Lean / therblig logic
# -------------------------
NON_VALUE_ACTIVITY_TEXT = {
    "T": "Transport is non-value-added unless it directly supports transformation.",
    "M": "Handling or motion is non-value-added when it is walking, repositioning, searching, or extra motion.",
    "I": "Inspection is generally non-value-added from the customer perspective.",
    "W": "Waiting or delay is a classic Lean waste.",
    "S": "Storage represents inventory waste.",
    "D": "Decision or planning does not directly transform the product.",
}

THERBLIG_IMPACT_ADJUST = {
    "Search": -0.18,
    "Delay": -0.20,
    "Walk/Transport Empty": -0.18,
    "Reorient/Extra Motion": -0.14,
    "Reposition/Relocate": -0.12,
    "Position": -0.10,
    "Hold": -0.10,
    "Inspect/Check": -0.14,
    "Plan/Decision": -0.15,
    "Non-Value Motion": -0.15,
    "Grasp": 0.05,
    "Release Load / Pre-Position": 0.05,
    "Use": 0.12,
    "Use / Disassemble": 0.10,
    "Assemble": 0.12,
    "Move": 0.04,
    "Unclassified": -0.05,
}

THERBLIG_EFFORT_ADJUST = {
    "Search": 0.20,
    "Delay": 0.20,
    "Walk/Transport Empty": 0.18,
    "Reorient/Extra Motion": 0.12,
    "Reposition/Relocate": 0.10,
    "Position": 0.06,
    "Hold": 0.07,
    "Inspect/Check": 0.10,
    "Plan/Decision": 0.10,
    "Non-Value Motion": 0.12,
    "Grasp": 0.03,
    "Release Load / Pre-Position": 0.03,
    "Use": 0.08,
    "Use / Disassemble": 0.09,
    "Assemble": 0.10,
    "Move": 0.06,
    "Unclassified": 0.05,
}

LOW_IMPACT_THERBLIGS = {
    "Search", "Delay", "Walk/Transport Empty", "Reorient/Extra Motion", "Reposition/Relocate",
    "Position", "Hold", "Inspect/Check", "Plan/Decision", "Non-Value Motion"
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
        if matches_any(text, [r"\bdecide\b", r"\bplan\b", r"\breact\b", r"\bdetermine\b"]):
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
    if row["Activity"] in {"I", "W", "S", "D"}:
        return "Low"
    if row["Therblig / Motion Type"] in LOW_IMPACT_THERBLIGS:
        return "Low"
    if row["Motion Class"] == "Effective" or row["Activity Type"] == "VA":
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
        if duration >= 3 or (duration >= 2 and res_count >= 4):
            return "High"
        return "Low"
    if duration >= 4 or (duration >= 3 and activity in {"I", "D", "M", "T"}):
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
    therblig = row["Therblig / Motion Type"]
    duration = int(round(row["Duration Seconds"]))

    if row["Motion Class"] == "Effective":
        reasons.append(f"{therblig} directly advances the work.")
    else:
        reasons.append(f"{therblig} is treated as non-value-added motion.")
    if row["Activity"] in NON_VALUE_ACTIVITY_TEXT:
        reasons.append(NON_VALUE_ACTIVITY_TEXT[row["Activity"]])
    if duration >= 3:
        reasons.append(f"Duration={duration}s increases effort.")
    else:
        reasons.append(f"Duration={duration}s keeps effort lower.")
    if row["Resource Count"] >= 4:
        reasons.append("Multiple resources increase complexity.")
    elif row["Resource Count"] >= 3:
        reasons.append("Several resources are involved.")
    return " ".join(reasons)


def compute_continuous_scores(df):
    df = df.copy()
    max_duration = max(float(df["Duration Seconds"].max()), 1.0)
    max_resources = max(int(df["Resource Count"].max()), 1)

    duration_norm = (df["Duration Seconds"] / max_duration).clip(0, 1)
    resource_norm = (df["Resource Count"] / max_resources).clip(0, 1)

    impact_base = df["Impact"].map({"High": 0.74, "Low": 0.26}).fillna(0.26)
    effort_base = df["Effort"].map({"High": 0.74, "Low": 0.26}).fillna(0.26)
    impact_adjust = df["Therblig / Motion Type"].map(THERBLIG_IMPACT_ADJUST).fillna(0.0)
    effort_adjust = df["Therblig / Motion Type"].map(THERBLIG_EFFORT_ADJUST).fillna(0.0)
    activity_type_adjust = df["Activity Type"].map({"VA": 0.08, "NNVA": -0.02, "NVA": -0.08}).fillna(0.0)
    motion_adjust = df["Motion Class"].map({"Effective": 0.05, "Ineffective": -0.05}).fillna(0.0)

    df["Impact Score"] = (
        impact_base + impact_adjust + activity_type_adjust + motion_adjust - 0.04 * duration_norm + 0.02 * resource_norm
    ).clip(0.02, 0.98)
    df["Effort Score"] = (
        effort_base + effort_adjust + 0.18 * duration_norm + 0.10 * resource_norm
    ).clip(0.02, 0.98)
    return df


# -------------------------
# Recursive layout
# -------------------------
def auto_levels(n_points):
    if n_points <= 20:
        return 2
    if n_points <= 80:
        return 3
    return 4


def build_full_recursive_cells(levels, xmin=0.0, xmax=1.0, ymin=0.0, ymax=1.0, depth=0):
    cells = [{"xmin": xmin, "xmax": xmax, "ymin": ymin, "ymax": ymax, "depth": depth}]
    if depth >= levels:
        return cells

    xmid = (xmin + xmax) / 2.0
    ymid = (ymin + ymax) / 2.0
    children = [
        (xmin, xmid, ymid, ymax),  # Quick Wins
        (xmid, xmax, ymid, ymax),  # Major Projects
        (xmin, xmid, ymin, ymid),  # Fill-Ins
        (xmid, xmax, ymin, ymid),  # Time Sinks
    ]
    for a, b, c, d in children:
        cells.extend(build_full_recursive_cells(levels, a, b, c, d, depth + 1))
    return cells


def get_leaf_bounds(x_score, y_score, levels):
    xmin, xmax, ymin, ymax = 0.0, 1.0, 0.0, 1.0
    path = []
    for _ in range(levels):
        xmid = (xmin + xmax) / 2.0
        ymid = (ymin + ymax) / 2.0
        if x_score <= xmid and y_score > ymid:
            xmax = xmid
            ymin = ymid
            path.append("Quick Wins")
        elif x_score > xmid and y_score > ymid:
            xmin = xmid
            ymin = ymid
            path.append("Major Projects")
        elif x_score <= xmid and y_score <= ymid:
            xmax = xmid
            ymax = ymid
            path.append("Fill-Ins")
        else:
            xmin = xmid
            ymax = ymid
            path.append("Time Sinks")
    return xmin, xmax, ymin, ymax, " > ".join(path)


def assign_positions_by_leaf(df, levels):
    temp = df.copy()
    leaf_info = temp.apply(
        lambda row: get_leaf_bounds(float(row["Effort Score"]), float(row["Impact Score"]), levels),
        axis=1,
        result_type="expand",
    )
    leaf_info.columns = ["Leaf XMin", "Leaf XMax", "Leaf YMin", "Leaf YMax", "Recursive Path"]
    temp = pd.concat([temp, leaf_info], axis=1)

    grouped = []
    for _, group in temp.groupby(["Leaf XMin", "Leaf XMax", "Leaf YMin", "Leaf YMax"], sort=False):
        group = group.sort_values(["Impact Score", "Effort Score", "Step"], ascending=[False, True, True]).reset_index(drop=True)
        n = len(group)
        cols = math.ceil(math.sqrt(n))
        rows = math.ceil(n / cols)

        xmin = float(group["Leaf XMin"].iloc[0])
        xmax = float(group["Leaf XMax"].iloc[0])
        ymin = float(group["Leaf YMin"].iloc[0])
        ymax = float(group["Leaf YMax"].iloc[0])

        pad_x = (xmax - xmin) * 0.14
        pad_y = (ymax - ymin) * 0.14
        ux0, ux1 = xmin + pad_x, xmax - pad_x
        uy0, uy1 = ymin + pad_y, ymax - pad_y

        xs = []
        ys = []
        for i in range(n):
            r = i // cols
            c = i % cols
            x = ux0 + (c + 0.5) * (ux1 - ux0) / cols
            y = uy1 - (r + 0.5) * (uy1 - uy0) / rows
            xs.append(x)
            ys.append(y)

        group["x"] = xs
        group["y"] = ys
        group["Cell Width"] = xmax - xmin
        group["Cell Height"] = ymax - ymin
        group["Recursive Level"] = levels
        grouped.append(group)

    return pd.concat(grouped, ignore_index=True) if grouped else temp


def marker_size_from_levels(levels):
    if levels >= 4:
        return 16
    if levels == 3:
        return 20
    return 24


def draw_recursive_grid(fig, levels):
    cells = build_full_recursive_cells(levels)
    for cell in cells:
        if cell["depth"] == 0:
            continue
        fill_alpha = min(0.025 + 0.012 * cell["depth"], 0.08)
        fig.add_shape(
            type="rect",
            x0=cell["xmin"],
            x1=cell["xmax"],
            y0=cell["ymin"],
            y1=cell["ymax"],
            line=dict(color=CELL_LINE, width=1),
            fillcolor=f"rgba(68, 95, 127, {fill_alpha:.3f})",
            layer="below",
        )


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
    df = compute_continuous_scores(df.sort_values("Step").reset_index(drop=True))

    levels = auto_levels(len(df))
    plot_df = assign_positions_by_leaf(df, levels)
    marker_size = marker_size_from_levels(levels)

    fig = go.Figure()
    draw_recursive_grid(fig, levels)

    for quadrant, color in QUADRANT_COLORS.items():
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
                textfont=dict(size=max(8, marker_size - 8), color="black"),
                marker=dict(size=marker_size, color="rgba(255,255,255,0)", line=dict(color=color, width=2)),
                customdata=temp[[
                    "Step", "Description", "Activity", "Therblig / Motion Type", "Motion Class",
                    "Impact Score", "Effort Score", "Recursive Path"
                ]].values,
                hovertemplate=(
                    "Step: %{customdata[0]}"
                    "<br>Description: %{customdata[1]}"
                    "<br>Activity: %{customdata[2]}"
                    "<br>Therblig: %{customdata[3]}"
                    "<br>Motion Class: %{customdata[4]}"
                    "<br>Impact Score: %{customdata[5]:.3f}"
                    "<br>Effort Score: %{customdata[6]:.3f}"
                    "<br>Recursive Path: %{customdata[7]}"
                    "<extra></extra>"
                ),
                hoverlabel=dict(bgcolor="white", bordercolor="black", font=dict(color="black", size=12)),
                selectedpoints=selected_points,
                selected=dict(marker=dict(size=marker_size + 6, color="black", opacity=1.0)),
                unselected=dict(marker=dict(opacity=0.70)),
                showlegend=False,
            )
        )

    fig.add_shape(type="line", x0=0.5, x1=0.5, y0=0.0, y1=1.0, line=dict(color="gray", dash="dash", width=1.6))
    fig.add_shape(type="line", x0=0.0, x1=1.0, y0=0.5, y1=0.5, line=dict(color="gray", dash="dash", width=1.6))

    fig.update_layout(
        xaxis=dict(range=[0, 1], visible=False, fixedrange=True),
        yaxis=dict(range=[0, 1], visible=False, fixedrange=True, scaleanchor="x", scaleratio=1),
        plot_bgcolor=PLOT_BG,
        paper_bgcolor=PLOT_BG,
        height=940,
        margin=dict(l=70, r=30, t=70, b=60),
        clickmode="event+select",
        dragmode="select",
        annotations=[
            dict(x=0.50, y=1.06, xref="paper", yref="paper", text="Recursive Impact vs Effort Matrix", showarrow=False, font=dict(size=22, color="black")),
            dict(x=0.50, y=0.01, xref="paper", yref="paper", text="Effort", showarrow=False, font=dict(size=18, color="black")),
            dict(x=0.01, y=0.50, xref="paper", yref="paper", text="Impact", showarrow=False, textangle=-90, font=dict(size=18, color="black")),
            dict(x=0.25, y=0.96, xref="paper", yref="paper", text="Quick Wins", showarrow=False, font=dict(size=14, color=QUADRANT_COLORS["Quick Wins"])),
            dict(x=0.75, y=0.96, xref="paper", yref="paper", text="Major Projects", showarrow=False, font=dict(size=14, color=QUADRANT_COLORS["Major Projects"])),
            dict(x=0.25, y=0.09, xref="paper", yref="paper", text="Fill-Ins", showarrow=False, font=dict(size=14, color=QUADRANT_COLORS["Fill-Ins"])),
            dict(x=0.75, y=0.09, xref="paper", yref="paper", text="Time Sinks", showarrow=False, font=dict(size=14, color=QUADRANT_COLORS["Time Sinks"])),
            dict(x=0.05, y=0.24, xref="paper", yref="paper", text="Low", showarrow=False, font=dict(size=13, color="black")),
            dict(x=0.05, y=0.76, xref="paper", yref="paper", text="High", showarrow=False, font=dict(size=13, color="black")),
            dict(x=0.24, y=0.05, xref="paper", yref="paper", text="Low", showarrow=False, font=dict(size=13, color="black")),
            dict(x=0.76, y=0.05, xref="paper", yref="paper", text="High", showarrow=False, font=dict(size=13, color="black")),
        ],
    )
    return fig, plot_df


# -------------------------
# Table helpers
# -------------------------
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
- **Effort** is scored separately using duration, motion waste, repeated walking/searching/inspection, and resource involvement.
- Each activity is converted into a continuous **Impact Score** and **Effort Score** between 0 and 1.
- The whole matrix is then recursively divided into **Quick Wins, Major Projects, Fill-Ins, and Time Sinks**, so the nested grid is visible across the full chart.
"""
)

try:
    file_bytes, sheet_name = require_upload()
    df = load_full_data(file_bytes, sheet_name)

    if "impact_effort_selected_step" not in st.session_state:
        st.session_state["impact_effort_selected_step"] = None

    fig, table_df = make_impact_effort_matrix(df, selected_step=st.session_state["impact_effort_selected_step"])

    chart_event = st.plotly_chart(
        fig,
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
        st.rerun()

    selected_step = st.session_state.get("impact_effort_selected_step")

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
        impact_filter = st.multiselect("Filter Impact", options=["Low", "High"], default=["Low", "High"])
    with filter_col2:
        effort_filter = st.multiselect("Filter Effort", options=["Low", "High"], default=["Low", "High"])
    with filter_col3:
        if st.button("Clear marker selection", use_container_width=True):
            st.session_state["impact_effort_selected_step"] = None
            st.rerun()

    display_df = table_df.copy()
    display_df = display_df[
        display_df["Impact"].isin(impact_filter) & display_df["Effort"].isin(effort_filter)
    ].copy()
    display_df = display_df.sort_values(["Step", "Impact", "Effort"]).reset_index(drop=True)
    display_df["Duration"] = display_df["Duration Seconds"].round(0).astype(int).astype(str) + " s"

    table_cols = [
        "Step", "Description", "Activity", "Therblig / Motion Type", "Motion Class",
        "Impact", "Effort", "Quadrant", "Impact Score", "Effort Score", "Recursive Path",
        "Duration", "Lean / Six Sigma Logic",
    ]
    display_df = display_df[table_cols].rename(columns={
        "Therblig / Motion Type": "Therblig",
        "Lean / Six Sigma Logic": "Logic",
    })
    display_df["Impact Score"] = display_df["Impact Score"].round(3)
    display_df["Effort Score"] = display_df["Effort Score"].round(3)

    styled_table = style_logic_table(display_df, selected_step=selected_step)
    st.dataframe(styled_table, use_container_width=True, hide_index=True)

    if selected_step is not None:
        st.caption(f"Selected marker: Step {selected_step}. The matching row is highlighted below.")
    else:
        st.caption("Click a marker once to select it and highlight the matching row in the table.")
except Exception as e:
    st.error(f"Error: {e}")
