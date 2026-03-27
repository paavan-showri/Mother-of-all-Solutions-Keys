import io
import math
import re
from typing import Dict, List, Tuple

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Impact vs Effort", layout="wide")
st.title("Recursive Impact vs Effort Matrix")

BOTTOM_NOTE = (
    "This view recursively subdivides each occupied quadrant using the activities inside that quadrant only. "
    "So after an activity first lands in Quick Wins, Major Projects, Fill-Ins, or Time Sinks, it is evaluated "
    "again relative to the other activities in that same parent quadrant and placed into a second-level 2×2 split, "
    "then again for deeper levels until each final box is small enough to show clearly."
)

QUADRANT_COLORS = {
    "Quick Wins": "#0f9d9a",
    "Major Projects": "#635bff",
    "Fill-Ins": "#2f78ff",
    "Time Sinks": "#ff4d6d",
}
QUADRANT_ORDER = ["Quick Wins", "Major Projects", "Fill-Ins", "Time Sinks"]
CHILD_FILL = {
    "Quick Wins": "rgba(15, 157, 154, 0.03)",
    "Major Projects": "rgba(99, 91, 255, 0.03)",
    "Fill-Ins": "rgba(47, 120, 255, 0.03)",
    "Time Sinks": "rgba(255, 77, 109, 0.03)",
}
PLOT_BG = "#dcdcdc"
CELL_LINE = "rgba(65, 105, 155, 0.85)"
MAX_DEPTH = 5
LEAF_CAPACITY = 6
POINT_SIZE = 22


# ------------------------------
# Workbook helpers
# ------------------------------
def require_upload():
    if "excel_file_bytes" not in st.session_state or "sheet_name" not in st.session_state:
        st.warning("Please upload the Excel file first on the Home page.")
        st.stop()
    return st.session_state["excel_file_bytes"], st.session_state["sheet_name"]


def normalize_text(value):
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).strip().lower())


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


# ------------------------------
# Lean classification logic
# ------------------------------
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
        return "Assemble", "Effective" if matches_any(text, [r"\bstack\b"]) else ("Use / Disassemble", "Effective")
    if matches_any(text, [r"\bstack\b", r"\bassemble\b", r"\bcombine\b", r"\bjoin\b"]):
        return "Assemble", "Effective"
    if matches_any(text, [r"\bmove\b", r"\bcarry\b", r"\bbring\b", r"\btransfer\b"]):
        return "Move", "Effective"
    if activity == "O":
        return "Use", "Effective"
    if activity in {"T", "M"}:
        return ("Reposition/Relocate", "Ineffective") if activity_type == "NVA" else ("Move", "Effective")
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
    reasons.append(f"Duration={duration}s {'increases' if duration >= 3 else 'keeps'} effort {'higher' if duration >= 3 else 'lower' }.")
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
    df["Impact Score"] = (impact_base + impact_adjust + activity_type_adjust + motion_adjust - 0.04 * duration_norm + 0.02 * resource_norm).clip(0.02, 0.98)
    df["Effort Score"] = (effort_base + effort_adjust + 0.18 * duration_norm + 0.10 * resource_norm).clip(0.02, 0.98)
    return df


# ------------------------------
# Recursive local quadrant logic
# ------------------------------
def child_bounds(bounds: Tuple[float, float, float, float], quadrant: str) -> Tuple[float, float, float, float]:
    xmin, xmax, ymin, ymax = bounds
    xmid = (xmin + xmax) / 2.0
    ymid = (ymin + ymax) / 2.0
    if quadrant == "Quick Wins":
        return xmin, xmid, ymid, ymax
    if quadrant == "Major Projects":
        return xmid, xmax, ymid, ymax
    if quadrant == "Fill-Ins":
        return xmin, xmid, ymin, ymid
    return xmid, xmax, ymin, ymid


def assign_local_quadrants(node_df: pd.DataFrame) -> Tuple[Dict[str, pd.DataFrame], float, float]:
    if node_df.empty:
        return {q: node_df.copy() for q in QUADRANT_ORDER}, 0.5, 0.5

    x_split = float(node_df["Effort Score"].median())
    y_split = float(node_df["Impact Score"].median())

    # Small jitter by Step to avoid too many equal-score ties collapsing into one child.
    temp = node_df.copy()
    temp["_jit"] = (temp["Step"].rank(method="dense") % 7) * 1e-6
    x = temp["Effort Score"] + temp["_jit"]
    y = temp["Impact Score"] + temp["_jit"]

    groups = {
        "Quick Wins": temp[(x <= x_split) & (y > y_split)].drop(columns=["_jit"]),
        "Major Projects": temp[(x > x_split) & (y > y_split)].drop(columns=["_jit"]),
        "Fill-Ins": temp[(x <= x_split) & (y <= y_split)].drop(columns=["_jit"]),
        "Time Sinks": temp[(x > x_split) & (y <= y_split)].drop(columns=["_jit"]),
    }
    return groups, x_split, y_split


def recursive_partition(node_df: pd.DataFrame, bounds: Tuple[float, float, float, float], path: List[str], depth: int,
                        split_rects: List[Dict], leaves: List[Dict]) -> None:
    if node_df.empty:
        return

    if len(node_df) <= LEAF_CAPACITY or depth >= MAX_DEPTH:
        leaves.append({"df": node_df.copy(), "bounds": bounds, "path": path[:]})
        return

    child_groups, x_split, y_split = assign_local_quadrants(node_df)

    # If all activities still collapse into one child, stop recursion to avoid a visually useless chain.
    non_empty_counts = sum(1 for q in QUADRANT_ORDER if not child_groups[q].empty)
    if non_empty_counts <= 1:
        leaves.append({"df": node_df.copy(), "bounds": bounds, "path": path[:]})
        return

    split_rects.append({
        "bounds": bounds,
        "path": path[:],
        "counts": {q: len(child_groups[q]) for q in QUADRANT_ORDER},
        "x_split": x_split,
        "y_split": y_split,
    })

    for quad in QUADRANT_ORDER:
        quad_df = child_groups[quad]
        if quad_df.empty:
            continue
        recursive_partition(
            quad_df,
            child_bounds(bounds, quad),
            path + [quad],
            depth + 1,
            split_rects,
            leaves,
        )


def path_label(path: List[str]) -> str:
    return "Root" if not path else " > ".join(path)


def assign_points_in_leaf(df: pd.DataFrame, bounds: Tuple[float, float, float, float]) -> pd.DataFrame:
    temp = df.sort_values(["Impact Score", "Effort Score", "Step"], ascending=[False, True, True]).reset_index(drop=True).copy()
    n = len(temp)
    if n == 0:
        return temp

    xmin, xmax, ymin, ymax = bounds
    cols = math.ceil(math.sqrt(n))
    rows = math.ceil(n / cols)

    pad_x = max((xmax - xmin) * 0.10, 0.004)
    pad_y = max((ymax - ymin) * 0.10, 0.004)
    ux0, ux1 = xmin + pad_x, xmax - pad_x
    uy0, uy1 = ymin + pad_y, ymax - pad_y

    xs, ys = [], []
    for i in range(n):
        r = i // cols
        c = i % cols
        x = ux0 + (c + 0.5) * (ux1 - ux0) / cols
        y = uy1 - (r + 0.5) * (uy1 - uy0) / rows
        xs.append(x)
        ys.append(y)

    temp["x"] = xs
    temp["y"] = ys
    return temp


def build_recursive_structure(df: pd.DataFrame) -> Tuple[List[Dict], List[Dict], pd.DataFrame]:
    root_groups, _, _ = assign_local_quadrants(df)
    split_rects: List[Dict] = []
    leaves: List[Dict] = []
    for quad in QUADRANT_ORDER:
        quad_df = root_groups[quad]
        if quad_df.empty:
            continue
        recursive_partition(quad_df, child_bounds((0.0, 1.0, 0.0, 1.0), quad), [quad], 1, split_rects, leaves)

    leaf_frames = []
    for leaf in leaves:
        placed = assign_points_in_leaf(leaf["df"], leaf["bounds"])
        placed["Recursive Path"] = path_label(leaf["path"])
        placed["Top Quadrant"] = leaf["path"][0] if leaf["path"] else "Root"
        leaf_frames.append(placed)
    plot_df = pd.concat(leaf_frames, ignore_index=True) if leaf_frames else pd.DataFrame()
    return split_rects, leaves, plot_df


# ------------------------------
# Plotting helpers
# ------------------------------
def add_top_level_rectangles(fig: go.Figure):
    root_bounds = (0.0, 1.0, 0.0, 1.0)
    for quad in QUADRANT_ORDER:
        xmin, xmax, ymin, ymax = child_bounds(root_bounds, quad)
        fig.add_shape(
            type="rect",
            x0=xmin,
            x1=xmax,
            y0=ymin,
            y1=ymax,
            line=dict(color=CELL_LINE, width=1.6),
            fillcolor=CHILD_FILL[quad],
            layer="below",
        )


def add_recursive_rectangles(fig: go.Figure, split_rects: List[Dict]):
    for node in split_rects:
        xmin, xmax, ymin, ymax = node["bounds"]
        xmid = (xmin + xmax) / 2.0
        ymid = (ymin + ymax) / 2.0
        fig.add_shape(type="line", x0=xmid, x1=xmid, y0=ymin, y1=ymax, line=dict(color=CELL_LINE, width=1.0))
        fig.add_shape(type="line", x0=xmin, x1=xmax, y0=ymid, y1=ymid, line=dict(color=CELL_LINE, width=1.0))


def add_global_crosshair(fig: go.Figure):
    fig.add_shape(type="line", x0=0.5, x1=0.5, y0=0.0, y1=1.0, line=dict(color="gray", dash="dash", width=1.4))
    fig.add_shape(type="line", x0=0.0, x1=1.0, y0=0.5, y1=0.5, line=dict(color="gray", dash="dash", width=1.4))


def add_activity_traces(fig: go.Figure, plot_df: pd.DataFrame, selected_step=None):
    if plot_df.empty:
        return
    for quad in QUADRANT_ORDER:
        temp = plot_df[plot_df["Top Quadrant"] == quad].copy().reset_index(drop=True)
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
                textfont=dict(size=9, color="black"),
                marker=dict(size=POINT_SIZE, color="rgba(255,255,255,0)", line=dict(color=QUADRANT_COLORS[quad], width=2)),
                customdata=temp[[
                    "Step", "Description", "Activity", "Therblig / Motion Type", "Motion Class",
                    "Impact Score", "Effort Score", "Recursive Path", "Quadrant"
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
                    "<br>Top-Level Quadrant: %{customdata[8]}"
                    "<extra></extra>"
                ),
                selectedpoints=selected_points,
                selected=dict(marker=dict(size=28, color="black", opacity=1.0)),
                unselected=dict(marker=dict(opacity=0.88)),
                showlegend=False,
            )
        )


def add_axis_and_titles(fig: go.Figure):
    annotations = [
        dict(x=0.50, y=1.05, xref="paper", yref="paper", text="Recursive Impact vs Effort Matrix", showarrow=False, font=dict(size=22, color="black")),
        dict(x=0.50, y=0.01, xref="paper", yref="paper", text="Effort", showarrow=False, font=dict(size=18, color="black")),
        dict(x=0.01, y=0.50, xref="paper", yref="paper", text="Impact", showarrow=False, textangle=-90, font=dict(size=18, color="black")),
        dict(x=0.05, y=0.24, xref="paper", yref="paper", text="Low", showarrow=False, font=dict(size=13, color="black")),
        dict(x=0.05, y=0.76, xref="paper", yref="paper", text="High", showarrow=False, font=dict(size=13, color="black")),
        dict(x=0.24, y=0.05, xref="paper", yref="paper", text="Low", showarrow=False, font=dict(size=13, color="black")),
        dict(x=0.76, y=0.05, xref="paper", yref="paper", text="High", showarrow=False, font=dict(size=13, color="black")),
        dict(x=0.26, y=0.96, xref="paper", yref="paper", text="Quick Wins", showarrow=False, font=dict(size=14, color=QUADRANT_COLORS["Quick Wins"])),
        dict(x=0.74, y=0.96, xref="paper", yref="paper", text="Major Projects", showarrow=False, font=dict(size=14, color=QUADRANT_COLORS["Major Projects"])),
        dict(x=0.26, y=0.08, xref="paper", yref="paper", text="Fill-Ins", showarrow=False, font=dict(size=14, color=QUADRANT_COLORS["Fill-Ins"])),
        dict(x=0.74, y=0.08, xref="paper", yref="paper", text="Time Sinks", showarrow=False, font=dict(size=14, color=QUADRANT_COLORS["Time Sinks"])),
    ]
    fig.update_layout(
        xaxis=dict(range=[0, 1], visible=False, fixedrange=True),
        yaxis=dict(range=[0, 1], visible=False, fixedrange=True, scaleanchor="x", scaleratio=1),
        plot_bgcolor=PLOT_BG,
        paper_bgcolor=PLOT_BG,
        height=920,
        margin=dict(l=70, r=30, t=90, b=60),
        clickmode="event+select",
        dragmode=False,
        annotations=annotations,
    )


# ------------------------------
# Tables and selection
# ------------------------------
def build_rcpsp_priority_table(df: pd.DataFrame) -> pd.DataFrame:
    temp = df.copy()
    temp["Quadrant Weight"] = temp["Quadrant"].map({"Quick Wins": 4, "Major Projects": 3, "Fill-Ins": 2, "Time Sinks": 1}).fillna(1)
    temp["RCPSP Priority Score"] = 0.45 * temp["Quadrant Weight"] + 0.35 * temp["Impact Score"] - 0.20 * temp["Effort Score"]
    temp = temp.sort_values(["RCPSP Priority Score", "Impact Score", "Effort Score", "Step"], ascending=[False, False, True, True]).reset_index(drop=True)
    temp["RCPSP Priority Rank"] = range(1, len(temp) + 1)
    return temp[["RCPSP Priority Rank", "Step", "Description", "Quadrant", "Impact", "Effort", "Impact Score", "Effort Score", "RCPSP Priority Score", "Recursive Path"]]


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


# ------------------------------
# Build full page data and figure
# ------------------------------
def make_impact_effort_matrix(df: pd.DataFrame, selected_step=None):
    temp = df.copy()
    temp[["Therblig / Motion Type", "Motion Class"]] = temp.apply(
        lambda row: pd.Series(classify_therblig(row["Description"], row["Activity"], row["Activity Type"])),
        axis=1,
    )
    temp["Impact"] = temp.apply(classify_impact, axis=1)
    temp["Effort"] = temp.apply(classify_effort, axis=1)
    temp["Quadrant"] = temp.apply(get_quadrant, axis=1)
    temp["Lean / Six Sigma Logic"] = temp.apply(build_logic_text, axis=1)
    temp = compute_continuous_scores(temp.sort_values("Step").reset_index(drop=True))

    split_rects, leaves, plot_df = build_recursive_structure(temp)

    # Attach recursive path back to full table.
    if not plot_df.empty:
        path_map = plot_df[["Step", "Recursive Path"]].drop_duplicates()
        temp = temp.merge(path_map, on="Step", how="left")
    else:
        temp["Recursive Path"] = ""

    fig = go.Figure()
    add_top_level_rectangles(fig)
    add_recursive_rectangles(fig, split_rects)
    add_global_crosshair(fig)
    add_activity_traces(fig, plot_df, selected_step=selected_step)
    add_axis_and_titles(fig)

    return fig, temp, leaves


st.markdown(
    """
**How this recursive view works**

- The first split classifies activities into **Quick Wins, Major Projects, Fill-Ins, and Time Sinks**.
- Inside each occupied quadrant, the activities in that quadrant are split **again** into a local 2×2 matrix using only that quadrant's own activities.
- The same logic repeats until the final boxes are small enough to display clearly.
- This lets you see, for example, which activities inside **Major Projects** are the more urgent local quick wins versus the lower-value local time sinks.
"""
)

try:
    file_bytes, sheet_name = require_upload()
    df = load_full_data(file_bytes, sheet_name)

    if "impact_effort_selected_step" not in st.session_state:
        st.session_state["impact_effort_selected_step"] = None

    fig, full_df, leaves = make_impact_effort_matrix(df, selected_step=st.session_state["impact_effort_selected_step"])

    chart_event = st.plotly_chart(
        fig,
        use_container_width=True,
        theme=None,
        key="impact_effort_chart",
        on_select="rerun",
        selection_mode="points",
        config={"scrollZoom": False, "displayModeBar": False},
    )

    clicked_step = get_selected_step_from_event(chart_event)
    if clicked_step is not None:
        st.session_state["impact_effort_selected_step"] = clicked_step
        st.rerun()

    st.markdown(
        f"<div style='margin-top: 10px; font-size: 14px; color: black; text-align: justify;'>{BOTTOM_NOTE}</div>",
        unsafe_allow_html=True,
    )

    st.markdown("### Recursive leaf summary")
    leaf_rows = []
    for leaf in leaves:
        leaf_rows.append({
            "Recursive Path": path_label(leaf["path"]),
            "Activity Count": len(leaf["df"]),
            "Min Step": int(leaf["df"]["Step"].min()),
            "Max Step": int(leaf["df"]["Step"].max()),
        })
    leaf_df = pd.DataFrame(leaf_rows).sort_values(["Recursive Path", "Activity Count"], ascending=[True, False]) if leaf_rows else pd.DataFrame()
    st.dataframe(leaf_df, use_container_width=True, hide_index=True)

    st.markdown("### RCPSP auto-priority list")
    rcpsp_df = build_rcpsp_priority_table(full_df)
    rcpsp_df[["Impact Score", "Effort Score", "RCPSP Priority Score"]] = rcpsp_df[["Impact Score", "Effort Score", "RCPSP Priority Score"]].round(3)
    st.dataframe(rcpsp_df, use_container_width=True, hide_index=True)

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

    display_df = full_df.copy()
    display_df = display_df[display_df["Impact"].isin(impact_filter) & display_df["Effort"].isin(effort_filter)].copy()
    display_df = display_df.sort_values(["Step", "Impact", "Effort"]).reset_index(drop=True)
    display_df["Duration"] = display_df["Duration Seconds"].round(0).astype(int).astype(str) + " s"
    table_cols = [
        "Step", "Description", "Activity", "Therblig / Motion Type", "Motion Class", "Impact", "Effort",
        "Quadrant", "Impact Score", "Effort Score", "Recursive Path", "Duration", "Lean / Six Sigma Logic",
    ]
    display_df = display_df[table_cols].rename(columns={"Therblig / Motion Type": "Therblig", "Lean / Six Sigma Logic": "Logic"})
    display_df[["Impact Score", "Effort Score"]] = display_df[["Impact Score", "Effort Score"]].round(3)
    styled_table = style_logic_table(display_df, selected_step=st.session_state.get("impact_effort_selected_step"))
    st.dataframe(styled_table, use_container_width=True, hide_index=True)

    if st.session_state.get("impact_effort_selected_step") is not None:
        st.caption(f"Selected marker: Step {st.session_state['impact_effort_selected_step']}. The matching row is highlighted below.")
    else:
        st.caption("Click an activity marker to highlight its matching row in the table.")
except Exception as e:
    st.error(f"Error: {e}")
