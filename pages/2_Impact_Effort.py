import io
import math
import re
from typing import Dict, List, Tuple

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Impact vs Effort", layout="wide")
st.title("Zoomable Impact vs Effort Quadtree")

BOTTOM_NOTE = (
    "This view keeps the original Lean impact-effort logic, but instead of showing every recursive "
    "level at once, it shows one quadtree level at a time. Click a quadrant label to expand that "
    "subdivision, or click an activity marker to highlight the matching row. This avoids the heavy "
    "clustering that happens when all nested levels are drawn together."
)

QUADRANT_COLORS = {
    "Quick Wins": "#0f9d9a",
    "Major Projects": "#635bff",
    "Fill-Ins": "#2f78ff",
    "Time Sinks": "#ff4d6d",
}

QUADRANT_ORDER = ["Quick Wins", "Major Projects", "Fill-Ins", "Time Sinks"]
QUADRANT_SHORT = {
    "Quick Wins": "QW",
    "Major Projects": "MP",
    "Fill-Ins": "FI",
    "Time Sinks": "TS",
}

PLOT_BG = "#dcdcdc"
CELL_LINE = "rgba(65, 105, 155, 0.85)"
CELL_FILL = "rgba(68, 95, 127, 0.06)"
CHILD_FILL = {
    "Quick Wins": "rgba(15, 157, 154, 0.06)",
    "Major Projects": "rgba(99, 91, 255, 0.06)",
    "Fill-Ins": "rgba(47, 120, 255, 0.06)",
    "Time Sinks": "rgba(255, 77, 109, 0.06)",
}
MAX_DEPTH = 5
SHOW_POINTS_THRESHOLD = 16


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
    if matches_any(text, [
        r"\breposition\b", r"\bmove .* back\b", r"\bmove knife to new location\b", r"\bmove plate\b",
        r"\bmove butter plate\b", r"\bmove fruit bowl\b", r"\bmove plate with toast\b"
    ]):
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
# Quadtree helpers
# -------------------------
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


def path_to_bounds(path: List[str]) -> Tuple[float, float, float, float]:
    bounds = (0.0, 1.0, 0.0, 1.0)
    for quadrant in path:
        bounds = child_bounds(bounds, quadrant)
    return bounds


def path_label(path: List[str]) -> str:
    if not path:
        return "Root"
    return " > ".join(path)


def score_to_path(effort_score: float, impact_score: float, levels: int) -> List[str]:
    path = []
    bounds = (0.0, 1.0, 0.0, 1.0)
    for _ in range(levels):
        xmin, xmax, ymin, ymax = bounds
        xmid = (xmin + xmax) / 2.0
        ymid = (ymin + ymax) / 2.0
        if effort_score <= xmid and impact_score > ymid:
            quad = "Quick Wins"
        elif effort_score > xmid and impact_score > ymid:
            quad = "Major Projects"
        elif effort_score <= xmid and impact_score <= ymid:
            quad = "Fill-Ins"
        else:
            quad = "Time Sinks"
        path.append(quad)
        bounds = child_bounds(bounds, quad)
    return path


def assign_recursive_paths(df: pd.DataFrame, levels: int) -> pd.DataFrame:
    temp = df.copy()
    temp["Recursive Path List"] = temp.apply(
        lambda row: score_to_path(float(row["Effort Score"]), float(row["Impact Score"]), levels),
        axis=1,
    )
    temp["Recursive Path"] = temp["Recursive Path List"].apply(lambda x: " > ".join(x))
    return temp


def path_starts_with(full_path: List[str], prefix: List[str]) -> bool:
    return full_path[: len(prefix)] == prefix


def get_node_df(df: pd.DataFrame, focus_path: List[str]) -> pd.DataFrame:
    return df[df["Recursive Path List"].apply(lambda p: path_starts_with(p, focus_path))].copy()


def get_child_df(df: pd.DataFrame, focus_path: List[str], child_name: str) -> pd.DataFrame:
    child_path = focus_path + [child_name]
    return df[df["Recursive Path List"].apply(lambda p: path_starts_with(p, child_path))].copy()


def assign_points_in_bounds(df: pd.DataFrame, bounds: Tuple[float, float, float, float]) -> pd.DataFrame:
    temp = df.sort_values(["Impact Score", "Effort Score", "Step"], ascending=[False, True, True]).reset_index(drop=True).copy()
    n = len(temp)
    if n == 0:
        return temp

    xmin, xmax, ymin, ymax = bounds
    cols = math.ceil(math.sqrt(n))
    rows = math.ceil(n / cols)
    pad_x = (xmax - xmin) * 0.14
    pad_y = (ymax - ymin) * 0.14
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


def build_focus_plot_df(df: pd.DataFrame, focus_path: List[str]) -> pd.DataFrame:
    pieces = []
    for quad in QUADRANT_ORDER:
        child_df = get_child_df(df, focus_path, quad)
        if child_df.empty:
            continue
        bounds = child_bounds(path_to_bounds(focus_path), quad)
        placed = assign_points_in_bounds(child_df, bounds)
        placed["Visible Child"] = quad
        pieces.append(placed)
    if not pieces:
        return pd.DataFrame()
    return pd.concat(pieces, ignore_index=True)


def get_child_summary(df: pd.DataFrame, focus_path: List[str]) -> pd.DataFrame:
    rows = []
    for quad in QUADRANT_ORDER:
        child_df = get_child_df(df, focus_path, quad)
        bounds = child_bounds(path_to_bounds(focus_path), quad)
        count = len(child_df)
        if count == 0:
            avg_impact = None
            avg_effort = None
        else:
            avg_impact = float(child_df["Impact Score"].mean())
            avg_effort = float(child_df["Effort Score"].mean())
        rows.append({
            "Quadrant": quad,
            "Count": count,
            "Avg Impact Score": avg_impact,
            "Avg Effort Score": avg_effort,
            "bounds": bounds,
            "path": focus_path + [quad],
        })
    return pd.DataFrame(rows)


def marker_size_from_count(count: int) -> int:
    if count <= 6:
        return 28
    if count <= 12:
        return 24
    return 20


def add_child_rectangles(fig: go.Figure, child_summary: pd.DataFrame):
    for _, row in child_summary.iterrows():
        xmin, xmax, ymin, ymax = row["bounds"]
        quad = row["Quadrant"]
        fig.add_shape(
            type="rect",
            x0=xmin,
            x1=xmax,
            y0=ymin,
            y1=ymax,
            line=dict(color=CELL_LINE, width=1.4),
            fillcolor=CHILD_FILL[quad],
            layer="below",
        )


def add_cell_click_targets(fig: go.Figure, child_summary: pd.DataFrame):
    centers_x, centers_y, labels, custom = [], [], [], []
    for _, row in child_summary.iterrows():
        xmin, xmax, ymin, ymax = row["bounds"]
        quad = row["Quadrant"]
        count = int(row["Count"])
        centers_x.append((xmin + xmax) / 2.0)
        centers_y.append((ymin + ymax) / 2.0)
        labels.append(f"{quad}<br>{count}")
        custom.append(["cell", quad, " > ".join(row["path"]), count])

    fig.add_trace(
        go.Scatter(
            x=centers_x,
            y=centers_y,
            mode="markers+text",
            text=labels,
            textposition="middle center",
            textfont=dict(size=16, color="black"),
            marker=dict(size=56, opacity=0.01, color="rgba(0,0,0,0)"),
            customdata=custom,
            hovertemplate="Click to expand %{customdata[1]}<br>Count: %{customdata[3]}<extra></extra>",
            showlegend=False,
        )
    )


def add_activity_traces(fig: go.Figure, plot_df: pd.DataFrame, selected_step=None):
    if plot_df.empty:
        return
    for quad in QUADRANT_ORDER:
        temp = plot_df[plot_df["Visible Child"] == quad].copy().reset_index(drop=True)
        if temp.empty:
            continue
        marker_size = marker_size_from_count(len(temp))
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
                textfont=dict(size=max(8, marker_size - 10), color="black"),
                marker=dict(size=marker_size, color="rgba(255,255,255,0)", line=dict(color=QUADRANT_COLORS[quad], width=2)),
                customdata=temp[[
                    "Step", "Description", "Activity", "Therblig / Motion Type", "Motion Class",
                    "Impact Score", "Effort Score", "Recursive Path", "Visible Child"
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
                selectedpoints=selected_points,
                selected=dict(marker=dict(size=marker_size + 6, color="black", opacity=1.0)),
                unselected=dict(marker=dict(opacity=0.75)),
                showlegend=False,
            )
        )


def add_global_crosshair(fig: go.Figure):
    fig.add_shape(type="line", x0=0.5, x1=0.5, y0=0.0, y1=1.0, line=dict(color="gray", dash="dash", width=1.5))
    fig.add_shape(type="line", x0=0.0, x1=1.0, y0=0.5, y1=0.5, line=dict(color="gray", dash="dash", width=1.5))


def make_impact_effort_matrix(df, focus_path: List[str], selected_step=None):
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
    df = assign_recursive_paths(df, MAX_DEPTH)

    node_df = get_node_df(df, focus_path)
    child_summary = get_child_summary(df, focus_path)
    plot_df = build_focus_plot_df(df, focus_path)

    fig = go.Figure()
    add_child_rectangles(fig, child_summary)
    add_global_crosshair(fig)
    add_cell_click_targets(fig, child_summary)
    add_activity_traces(fig, plot_df, selected_step=selected_step)

    focus_title = f"Focus: {path_label(focus_path)}"
    if focus_path:
        subtitle = f"Showing immediate subdivisions inside {path_label(focus_path)}"
    else:
        subtitle = "Showing the first recursive split of the full matrix"

    annotations = [
        dict(x=0.50, y=1.06, xref="paper", yref="paper", text="Zoomable Impact vs Effort Quadtree", showarrow=False, font=dict(size=22, color="black")),
        dict(x=0.50, y=1.02, xref="paper", yref="paper", text=focus_title, showarrow=False, font=dict(size=15, color="black")),
        dict(x=0.50, y=0.99, xref="paper", yref="paper", text=subtitle, showarrow=False, font=dict(size=12, color="#333333")),
        dict(x=0.50, y=0.01, xref="paper", yref="paper", text="Effort", showarrow=False, font=dict(size=18, color="black")),
        dict(x=0.01, y=0.50, xref="paper", yref="paper", text="Impact", showarrow=False, textangle=-90, font=dict(size=18, color="black")),
        dict(x=0.05, y=0.24, xref="paper", yref="paper", text="Low", showarrow=False, font=dict(size=13, color="black")),
        dict(x=0.05, y=0.76, xref="paper", yref="paper", text="High", showarrow=False, font=dict(size=13, color="black")),
        dict(x=0.24, y=0.05, xref="paper", yref="paper", text="Low", showarrow=False, font=dict(size=13, color="black")),
        dict(x=0.76, y=0.05, xref="paper", yref="paper", text="High", showarrow=False, font=dict(size=13, color="black")),
    ]

    fig.update_layout(
        xaxis=dict(range=[0, 1], visible=False, fixedrange=True),
        yaxis=dict(range=[0, 1], visible=False, fixedrange=True, scaleanchor="x", scaleratio=1),
        plot_bgcolor=PLOT_BG,
        paper_bgcolor=PLOT_BG,
        height=900,
        margin=dict(l=70, r=30, t=90, b=60),
        clickmode="event+select",
        dragmode="select",
        annotations=annotations,
    )
    return fig, df, node_df, child_summary


def build_rcpsp_priority_table(df: pd.DataFrame, focus_path: List[str]) -> pd.DataFrame:
    temp = get_node_df(df, focus_path).copy()
    temp["Quadrant Weight"] = temp["Quadrant"].map({
        "Quick Wins": 4,
        "Major Projects": 3,
        "Fill-Ins": 2,
        "Time Sinks": 1,
    }).fillna(1)
    temp["RCPSP Priority Score"] = (
        0.45 * temp["Quadrant Weight"] +
        0.35 * temp["Impact Score"] -
        0.20 * temp["Effort Score"]
    )
    temp = temp.sort_values(["RCPSP Priority Score", "Impact Score", "Effort Score", "Step"], ascending=[False, False, True, True]).reset_index(drop=True)
    temp["RCPSP Priority Rank"] = range(1, len(temp) + 1)
    return temp[[
        "RCPSP Priority Rank", "Step", "Description", "Quadrant", "Impact", "Effort",
        "Impact Score", "Effort Score", "RCPSP Priority Score", "Recursive Path"
    ]]


def style_logic_table(df, selected_step=None):
    def highlight_row(row):
        if selected_step is not None and int(row["Step"]) == int(selected_step):
            return ["background-color: #fff3b0; font-weight: bold;"] * len(row)
        return [""] * len(row)

    return df.style.apply(highlight_row, axis=1)


def get_selected_payload_from_event(event):
    if not event:
        return None
    selection = event.get("selection", {}) if isinstance(event, dict) else getattr(event, "selection", {})
    points = selection.get("points", []) if hasattr(selection, "get") else []
    if not points:
        return None
    return points[0].get("customdata", None)


st.markdown(
    """
**How this zoomable view works**

- The original chart logic still classifies each activity into **Quick Wins, Major Projects, Fill-Ins, and Time Sinks**.
- Each activity is also converted to a continuous **Impact Score** and **Effort Score** between 0 and 1.
- The matrix is recursively divided into the same four quadrants again and again.
- To avoid heavy clustering, this page shows **one recursive level at a time**. Click a quadrant label to expand it.
- The **RCPSP priority table** below turns the currently focused node into an execution priority list that can feed a scheduling model.
"""
)

try:
    file_bytes, sheet_name = require_upload()
    df = load_full_data(file_bytes, sheet_name)

    if "impact_effort_selected_step" not in st.session_state:
        st.session_state["impact_effort_selected_step"] = None
    if "impact_effort_focus_path" not in st.session_state:
        st.session_state["impact_effort_focus_path"] = []

    focus_path = st.session_state["impact_effort_focus_path"]
    fig, full_df, node_df, child_summary = make_impact_effort_matrix(
        df,
        focus_path=focus_path,
        selected_step=st.session_state["impact_effort_selected_step"],
    )

    nav1, nav2, nav3 = st.columns([1, 1, 2])
    with nav1:
        if st.button("Back one level", use_container_width=True, disabled=(len(focus_path) == 0)):
            st.session_state["impact_effort_focus_path"] = focus_path[:-1]
            st.rerun()
    with nav2:
        if st.button("Reset to root", use_container_width=True):
            st.session_state["impact_effort_focus_path"] = []
            st.rerun()
    with nav3:
        st.caption(f"Current focus path: {path_label(focus_path)}")

    chart_event = st.plotly_chart(
        fig,
        use_container_width=True,
        theme=None,
        key="impact_effort_chart",
        on_select="rerun",
        selection_mode="points",
        config={"scrollZoom": False},
    )

    payload = get_selected_payload_from_event(chart_event)
    if payload is not None:
        if payload[0] == "cell":
            clicked_quad = payload[1]
            clicked_count = int(payload[3])
            if clicked_count > 0 and len(focus_path) < MAX_DEPTH:
                st.session_state["impact_effort_focus_path"] = focus_path + [clicked_quad]
                st.rerun()
        else:
            st.session_state["impact_effort_selected_step"] = int(payload[0])
            st.rerun()

    st.markdown(
        f"""
<div style="margin-top: 10px; font-size: 14px; color: black; text-align: justify;">
{BOTTOM_NOTE}
</div>
""",
        unsafe_allow_html=True,
    )

    st.markdown("### Visible subdivision summary")
    summary_df = child_summary[["Quadrant", "Count", "Avg Impact Score", "Avg Effort Score"]].copy()
    summary_df["Avg Impact Score"] = summary_df["Avg Impact Score"].round(3)
    summary_df["Avg Effort Score"] = summary_df["Avg Effort Score"].round(3)
    st.dataframe(summary_df, use_container_width=True, hide_index=True)

    st.markdown("### RCPSP auto-priority list for the current focus")
    rcpsp_df = build_rcpsp_priority_table(full_df, focus_path)
    rcpsp_df["Impact Score"] = rcpsp_df["Impact Score"].round(3)
    rcpsp_df["Effort Score"] = rcpsp_df["Effort Score"].round(3)
    rcpsp_df["RCPSP Priority Score"] = rcpsp_df["RCPSP Priority Score"].round(3)
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
    display_df = display_df[
        display_df["Impact"].isin(impact_filter) &
        display_df["Effort"].isin(effort_filter)
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

    styled_table = style_logic_table(display_df, selected_step=st.session_state.get("impact_effort_selected_step"))
    st.dataframe(styled_table, use_container_width=True, hide_index=True)

    if st.session_state.get("impact_effort_selected_step") is not None:
        st.caption(f"Selected marker: Step {st.session_state['impact_effort_selected_step']}. The matching row is highlighted below.")
    else:
        st.caption("Click an activity marker to highlight it, or click a quadrant label/count to zoom into that node.")
except Exception as e:
    st.error(f"Error: {e}")
