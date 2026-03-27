import io
import math
import re
from collections import defaultdict

import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from sklearn.cluster import DBSCAN

st.set_page_config(page_title="Impact vs Effort", layout="wide")
st.title("Impact vs Effort Matrix")

BOTTOM_NOTE = (
    "This page keeps the original Lean impact-effort logic, but now it supports zoomable recursive "
    "quadrants, optional DBSCAN cluster detection, and a priority handoff table for future RCPSP "
    "scheduling. Activities stay anchored to their continuous impact and effort scores instead of "
    "being packed into one corner of a cell."
)

QUADRANT_COLORS = {
    "Quick Wins": "#0f9d9a",
    "Major Projects": "#635bff",
    "Fill-Ins": "#2f78ff",
    "Time Sinks": "#ff4d6d",
}

PLOT_BG = "#dcdcdc"
GRID_LINE = "rgba(55, 95, 150, 0.75)"
FILL_TEMPLATE = "rgba(80, 110, 150, {alpha})"
ROOT_BOUNDS = (0.0, 1.0, 0.0, 1.0)
LEAF_CAPACITY = 10
MAX_DEPTH = 5


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
# Recursive nodes and clustering
# -------------------------
def split_bounds(bounds):
    xmin, xmax, ymin, ymax = bounds
    xmid = (xmin + xmax) / 2.0
    ymid = (ymin + ymax) / 2.0
    return {
        "Quick Wins": (xmin, xmid, ymid, ymax),
        "Major Projects": (xmid, xmax, ymid, ymax),
        "Fill-Ins": (xmin, xmid, ymin, ymid),
        "Time Sinks": (xmid, xmax, ymin, ymid),
    }


def quadrant_for_point(effort_score, impact_score, bounds):
    child_boxes = split_bounds(bounds)
    xmid = (bounds[0] + bounds[1]) / 2.0
    ymid = (bounds[2] + bounds[3]) / 2.0
    if effort_score <= xmid and impact_score > ymid:
        return "Quick Wins", child_boxes["Quick Wins"]
    if effort_score > xmid and impact_score > ymid:
        return "Major Projects", child_boxes["Major Projects"]
    if effort_score <= xmid and impact_score <= ymid:
        return "Fill-Ins", child_boxes["Fill-Ins"]
    return "Time Sinks", child_boxes["Time Sinks"]


def build_quadtree_nodes(df, max_depth=MAX_DEPTH, leaf_capacity=LEAF_CAPACITY):
    nodes = []

    def recurse(node_df, bounds, path, depth):
        node_id = "Root" if not path else " > ".join(path)
        node = {
            "id": node_id,
            "path": list(path),
            "bounds": bounds,
            "depth": depth,
            "count": len(node_df),
            "is_leaf": len(node_df) <= leaf_capacity or depth >= max_depth,
        }
        nodes.append(node)
        if node["is_leaf"] or len(node_df) == 0:
            return
        buckets = defaultdict(list)
        for _, row in node_df.iterrows():
            qname, child_bounds = quadrant_for_point(float(row["Effort Score"]), float(row["Impact Score"]), bounds)
            key = (qname, child_bounds)
            buckets[key].append(row)
        for (qname, child_bounds), rows in buckets.items():
            child_df = pd.DataFrame(rows)
            recurse(child_df, child_bounds, path + [qname], depth + 1)

    recurse(df.copy(), ROOT_BOUNDS, [], 0)
    return nodes


def assign_recursive_path(df, max_depth=MAX_DEPTH):
    rows = []
    for _, row in df.iterrows():
        bounds = ROOT_BOUNDS
        path = []
        for _depth in range(max_depth):
            qname, child_bounds = quadrant_for_point(float(row["Effort Score"]), float(row["Impact Score"]), bounds)
            path.append(qname)
            bounds = child_bounds
        row_copy = row.copy()
        row_copy["Recursive Path"] = " > ".join(path)
        row_copy["Leaf XMin"] = bounds[0]
        row_copy["Leaf XMax"] = bounds[1]
        row_copy["Leaf YMin"] = bounds[2]
        row_copy["Leaf YMax"] = bounds[3]
        rows.append(row_copy)
    return pd.DataFrame(rows)


def add_deterministic_jitter(df, x_col="Effort Score", y_col="Impact Score"):
    df = df.copy().sort_values("Step").reset_index(drop=True)
    groups = defaultdict(list)
    for idx, row in df.iterrows():
        key = (round(float(row[x_col]), 2), round(float(row[y_col]), 2), row["Quadrant"])
        groups[key].append(idx)

    jitter_x = [0.0] * len(df)
    jitter_y = [0.0] * len(df)
    for indices in groups.values():
        n = len(indices)
        if n == 1:
            continue
        radius = min(0.012 + 0.002 * math.sqrt(n), 0.028)
        for pos, idx in enumerate(indices):
            angle = 2 * math.pi * pos / n
            jitter_x[idx] = radius * math.cos(angle)
            jitter_y[idx] = radius * math.sin(angle)

    df["plot_x"] = (df[x_col] + pd.Series(jitter_x)).clip(0.01, 0.99)
    df["plot_y"] = (df[y_col] + pd.Series(jitter_y)).clip(0.01, 0.99)
    return df


def apply_dbscan(df, eps=0.065, min_samples=3):
    coords = df[["Effort Score", "Impact Score"]].to_numpy()
    model = DBSCAN(eps=eps, min_samples=min_samples)
    labels = model.fit_predict(coords)
    out = df.copy()
    out["Cluster"] = labels
    return out


def cluster_boxes(df):
    shapes = []
    label_ann = []
    clustered = df[df["Cluster"] >= 0].copy()
    for cluster_id, grp in clustered.groupby("Cluster"):
        pad = 0.018
        x0 = max(float(grp["plot_x"].min()) - pad, 0.0)
        x1 = min(float(grp["plot_x"].max()) + pad, 1.0)
        y0 = max(float(grp["plot_y"].min()) - pad, 0.0)
        y1 = min(float(grp["plot_y"].max()) + pad, 1.0)
        shapes.append(
            dict(
                type="rect",
                x0=x0,
                x1=x1,
                y0=y0,
                y1=y1,
                line=dict(color="rgba(0,0,0,0.45)", width=1.5, dash="dot"),
                fillcolor="rgba(255,255,255,0)",
                layer="below",
            )
        )
        label_ann.append(
            dict(
                x=(x0 + x1) / 2,
                y=y1,
                xref="x",
                yref="y",
                text=f"C{int(cluster_id)}",
                showarrow=False,
                yshift=10,
                font=dict(size=11, color="black"),
            )
        )
    return shapes, label_ann


def node_options(nodes):
    options = ["Root"]
    options.extend([n["id"] for n in nodes if n["id"] != "Root" and n["count"] > 0])
    return options


def find_node(nodes, node_id):
    for node in nodes:
        if node["id"] == node_id:
            return node
    return {"id": "Root", "bounds": ROOT_BOUNDS, "depth": 0, "count": 0}


def parent_chain(node_id):
    if node_id == "Root":
        return ["Root"]
    parts = node_id.split(" > ")
    chain = ["Root"]
    for i in range(1, len(parts) + 1):
        chain.append(" > ".join(parts[:i]))
    return chain


def bounds_with_padding(bounds, pad=0.01):
    xmin, xmax, ymin, ymax = bounds
    return [max(0.0, xmin - pad), min(1.0, xmax + pad)], [max(0.0, ymin - pad), min(1.0, ymax + pad)]


def visible_nodes(nodes, focus_bounds, max_extra_depth=1):
    fx0, fx1, fy0, fy1 = focus_bounds
    out = []
    for node in nodes:
        x0, x1, y0, y1 = node["bounds"]
        inside = x0 >= fx0 - 1e-9 and x1 <= fx1 + 1e-9 and y0 >= fy0 - 1e-9 and y1 <= fy1 + 1e-9
        if inside:
            out.append(node)
    return out


def marker_size_for_range(xrng, yrng):
    width = xrng[1] - xrng[0]
    span = min(width, yrng[1] - yrng[0])
    if span <= 0.08:
        return 24
    if span <= 0.18:
        return 18
    return 13


def base_annotations():
    return [
        dict(x=0.25, y=0.96, xref="paper", yref="paper", text="Quick Wins", showarrow=False,
             font=dict(size=14, color=QUADRANT_COLORS["Quick Wins"])),
        dict(x=0.75, y=0.96, xref="paper", yref="paper", text="Major Projects", showarrow=False,
             font=dict(size=14, color=QUADRANT_COLORS["Major Projects"])),
        dict(x=0.25, y=0.09, xref="paper", yref="paper", text="Fill-Ins", showarrow=False,
             font=dict(size=14, color=QUADRANT_COLORS["Fill-Ins"])),
        dict(x=0.75, y=0.09, xref="paper", yref="paper", text="Time Sinks", showarrow=False,
             font=dict(size=14, color=QUADRANT_COLORS["Time Sinks"])),
        dict(x=0.05, y=0.24, xref="paper", yref="paper", text="Low", showarrow=False, font=dict(size=13, color="black")),
        dict(x=0.05, y=0.76, xref="paper", yref="paper", text="High", showarrow=False, font=dict(size=13, color="black")),
        dict(x=0.24, y=0.05, xref="paper", yref="paper", text="Low", showarrow=False, font=dict(size=13, color="black")),
        dict(x=0.76, y=0.05, xref="paper", yref="paper", text="High", showarrow=False, font=dict(size=13, color="black")),
    ]


def build_shapes(nodes, focus_bounds, use_dbscan=False, cluster_shapes=None):
    shapes = []
    for node in visible_nodes(nodes, focus_bounds):
        if node["depth"] == 0:
            continue
        x0, x1, y0, y1 = node["bounds"]
        lw = 2.0 if node["depth"] == 1 else 1.1
        alpha = min(0.02 + 0.01 * node["depth"], 0.08)
        shapes.append(
            dict(
                type="rect",
                x0=x0,
                x1=x1,
                y0=y0,
                y1=y1,
                line=dict(color=GRID_LINE, width=lw),
                fillcolor=FILL_TEMPLATE.format(alpha=f"{alpha:.3f}"),
                layer="below",
            )
        )
    shapes.append(dict(type="line", x0=0.5, x1=0.5, y0=0.0, y1=1.0, line=dict(color="gray", dash="dash", width=1.8)))
    shapes.append(dict(type="line", x0=0.0, x1=1.0, y0=0.5, y1=0.5, line=dict(color="gray", dash="dash", width=1.8)))
    if use_dbscan and cluster_shapes:
        shapes.extend(cluster_shapes)
    return shapes


def add_traces(fig, plot_df, marker_size, selected_step=None):
    for quadrant, color in QUADRANT_COLORS.items():
        temp = plot_df[plot_df["Quadrant"] == quadrant].copy().reset_index(drop=True)
        if temp.empty:
            continue
        selected_points = None
        if selected_step is not None and selected_step in temp["Step"].tolist():
            selected_points = [temp.index[temp["Step"] == selected_step][0]]
        fig.add_trace(
            go.Scatter(
                x=temp["plot_x"],
                y=temp["plot_y"],
                mode="markers+text",
                text=temp["Step"].astype(str),
                textposition="middle center",
                textfont=dict(size=max(8, marker_size - 6), color="black"),
                marker=dict(size=marker_size, color="rgba(255,255,255,0)", line=dict(color=color, width=2)),
                customdata=temp[[
                    "Step", "Description", "Activity", "Therblig / Motion Type", "Motion Class",
                    "Impact Score", "Effort Score", "Recursive Path", "Cluster"
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
                    "<br>Cluster: %{customdata[8]}"
                    "<extra></extra>"
                ),
                hoverlabel=dict(bgcolor="white", bordercolor="black", font=dict(color="black", size=12)),
                selectedpoints=selected_points,
                selected=dict(marker=dict(size=marker_size + 5, color="black", opacity=1.0)),
                unselected=dict(marker=dict(opacity=0.82)),
                showlegend=False,
            )
        )


def make_main_figure(plot_df, nodes, focus_node_id="Root", selected_step=None, use_dbscan=False):
    focus_node = find_node(nodes, focus_node_id)
    xrng, yrng = bounds_with_padding(focus_node["bounds"], pad=0.01)
    visible_df = plot_df[
        plot_df["Effort Score"].between(xrng[0], xrng[1]) &
        plot_df["Impact Score"].between(yrng[0], yrng[1])
    ].copy()
    cluster_shapes, cluster_ann = ([], [])
    if use_dbscan:
        cluster_shapes, cluster_ann = cluster_boxes(visible_df)
    marker_size = marker_size_for_range(xrng, yrng)

    fig = go.Figure()
    add_traces(fig, visible_df, marker_size=marker_size, selected_step=selected_step)
    fig.update_layout(
        xaxis=dict(range=xrng, visible=False, fixedrange=True),
        yaxis=dict(range=yrng, visible=False, fixedrange=True, scaleanchor="x", scaleratio=1),
        plot_bgcolor=PLOT_BG,
        paper_bgcolor=PLOT_BG,
        height=940,
        margin=dict(l=70, r=30, t=70, b=60),
        clickmode="event+select",
        dragmode="select",
        shapes=build_shapes(nodes, focus_node["bounds"], use_dbscan=use_dbscan, cluster_shapes=cluster_shapes),
        annotations=[
            dict(x=0.50, y=1.06, xref="paper", yref="paper", text="Impact vs Effort Matrix", showarrow=False,
                 font=dict(size=22, color="black")),
            dict(x=0.50, y=0.01, xref="paper", yref="paper", text="Effort", showarrow=False,
                 font=dict(size=18, color="black")),
            dict(x=0.01, y=0.50, xref="paper", yref="paper", text="Impact", showarrow=False, textangle=-90,
                 font=dict(size=18, color="black")),
            dict(x=0.84, y=1.03, xref="paper", yref="paper", text=f"View: {focus_node_id}", showarrow=False,
                 font=dict(size=12, color="black")),
        ] + base_annotations() + cluster_ann,
    )
    return fig, visible_df


def make_animation_figure(plot_df, nodes, target_node_id, selected_step=None):
    chain = parent_chain(target_node_id)
    first_node = find_node(nodes, chain[0])
    xrng, yrng = bounds_with_padding(first_node["bounds"], pad=0.01)
    marker_size = marker_size_for_range(xrng, yrng)
    fig = go.Figure()
    add_traces(fig, plot_df, marker_size=marker_size, selected_step=selected_step)

    frames = []
    for node_id in chain:
        node = find_node(nodes, node_id)
        fx, fy = bounds_with_padding(node["bounds"], pad=0.01)
        frames.append(
            go.Frame(
                name=node_id,
                layout=go.Layout(
                    xaxis=dict(range=fx, visible=False, fixedrange=True),
                    yaxis=dict(range=fy, visible=False, fixedrange=True, scaleanchor="x", scaleratio=1),
                    shapes=build_shapes(nodes, node["bounds"]),
                    annotations=[
                        dict(x=0.50, y=1.06, xref="paper", yref="paper", text="Animated drill-down", showarrow=False,
                             font=dict(size=22, color="black")),
                        dict(x=0.84, y=1.03, xref="paper", yref="paper", text=f"Frame: {node_id}", showarrow=False,
                             font=dict(size=12, color="black")),
                    ] + base_annotations(),
                ),
            )
        )

    fig.frames = frames
    fig.update_layout(
        xaxis=dict(range=xrng, visible=False, fixedrange=True),
        yaxis=dict(range=yrng, visible=False, fixedrange=True, scaleanchor="x", scaleratio=1),
        plot_bgcolor=PLOT_BG,
        paper_bgcolor=PLOT_BG,
        height=700,
        margin=dict(l=70, r=30, t=70, b=60),
        updatemenus=[{
            "type": "buttons",
            "buttons": [
                {
                    "label": "Play",
                    "method": "animate",
                    "args": [None, {"frame": {"duration": 700, "redraw": True}, "transition": {"duration": 450}, "fromcurrent": True}],
                },
                {
                    "label": "Pause",
                    "method": "animate",
                    "args": [[None], {"frame": {"duration": 0, "redraw": False}, "mode": "immediate", "transition": {"duration": 0}}],
                },
            ],
            "x": 0.01,
            "y": 1.12,
            "showactive": False,
        }],
        sliders=[{
            "currentvalue": {"prefix": "Level: "},
            "steps": [
                {"label": name, "method": "animate", "args": [[name], {"mode": "immediate", "frame": {"duration": 0, "redraw": True}, "transition": {"duration": 0}}]}
                for name in chain
            ]
        }],
        shapes=build_shapes(nodes, first_node["bounds"]),
        annotations=[
            dict(x=0.50, y=1.06, xref="paper", yref="paper", text="Animated drill-down", showarrow=False,
                 font=dict(size=22, color="black")),
        ] + base_annotations(),
    )
    return fig


def build_priority_table(df):
    df = df.copy()
    quadrant_rank = {"Quick Wins": 1, "Major Projects": 2, "Fill-Ins": 3, "Time Sinks": 4}
    df["Priority Bucket"] = df["Quadrant"].map({
        "Quick Wins": "Do first",
        "Major Projects": "Plan and schedule",
        "Fill-Ins": "Bundle when capacity exists",
        "Time Sinks": "Reduce, redesign, or eliminate",
    })
    df["Priority Score"] = (
        (1 - df["Effort Score"]) * 0.45 + df["Impact Score"] * 0.55
    ).round(4)
    df["Quadrant Rank"] = df["Quadrant"].map(quadrant_rank)
    df = df.sort_values(["Quadrant Rank", "Priority Score", "Duration Seconds", "Step"], ascending=[True, False, False, True])
    df["RCPSP Priority"] = range(1, len(df) + 1)
    return df


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
- **Impact** is driven mainly by whether the motion is an **effective therblig** or an **ineffective / non-value-added therblig**.
- **Effort** is scored separately using duration, motion waste, repeated walking / searching / inspection, and resource involvement.
- Each activity is converted into a continuous **Impact Score** and **Effort Score** between 0 and 1.
- The page now supports a **zoomable recursive quadrant view**, **animated drill-down**, **optional DBSCAN cluster boxes**, and a **priority handoff table** for RCPSP scheduling.
"""
)

try:
    file_bytes, sheet_name = require_upload()
    df = load_full_data(file_bytes, sheet_name)

    if "impact_effort_selected_step" not in st.session_state:
        st.session_state["impact_effort_selected_step"] = None
    if "impact_effort_focus_node" not in st.session_state:
        st.session_state["impact_effort_focus_node"] = "Root"

    df[["Therblig / Motion Type", "Motion Class"]] = df.apply(
        lambda row: pd.Series(classify_therblig(row["Description"], row["Activity"], row["Activity Type"])),
        axis=1,
    )
    df["Impact"] = df.apply(classify_impact, axis=1)
    df["Effort"] = df.apply(classify_effort, axis=1)
    df["Quadrant"] = df.apply(get_quadrant, axis=1)
    df["Lean / Six Sigma Logic"] = df.apply(build_logic_text, axis=1)
    df = compute_continuous_scores(df.sort_values("Step").reset_index(drop=True))
    nodes = build_quadtree_nodes(df)
    df = assign_recursive_path(df)

    controls = st.columns([1.2, 1.3, 1.1, 1.1, 1.2])
    with controls[0]:
        focus_node = st.selectbox("Zoom region", options=node_options(nodes), index=node_options(nodes).index(st.session_state.get("impact_effort_focus_node", "Root")))
    with controls[1]:
        view_mode = st.selectbox("View mode", options=["Zoomable quadtree", "Animated drill-down"])
    with controls[2]:
        clustering_mode = st.selectbox("Clustering", options=["None", "DBSCAN"])
    with controls[3]:
        dbscan_eps = st.slider("DBSCAN eps", min_value=0.03, max_value=0.15, value=0.065, step=0.005)
    with controls[4]:
        dbscan_min = st.slider("DBSCAN min samples", min_value=2, max_value=8, value=3, step=1)

    st.session_state["impact_effort_focus_node"] = focus_node
    use_dbscan = clustering_mode == "DBSCAN"
    if use_dbscan:
        df = apply_dbscan(df, eps=dbscan_eps, min_samples=dbscan_min)
    else:
        df["Cluster"] = -1
    df = add_deterministic_jitter(df)

    if view_mode == "Zoomable quadtree":
        fig, visible_df = make_main_figure(
            df,
            nodes,
            focus_node_id=focus_node,
            selected_step=st.session_state["impact_effort_selected_step"],
            use_dbscan=use_dbscan,
        )
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
            selected_row = df[df["Step"] == clicked_step]
            if not selected_row.empty:
                st.session_state["impact_effort_focus_node"] = focus_node
            st.rerun()
    else:
        anim_fig = make_animation_figure(df, nodes, target_node_id=focus_node, selected_step=st.session_state["impact_effort_selected_step"])
        st.plotly_chart(anim_fig, use_container_width=True, theme=None, key="impact_effort_anim", config={"scrollZoom": False})

    action_cols = st.columns([1, 1, 1])
    with action_cols[0]:
        if st.button("Reset zoom", use_container_width=True):
            st.session_state["impact_effort_focus_node"] = "Root"
            st.rerun()
    with action_cols[1]:
        if st.button("Zoom to selected step path", use_container_width=True):
            step = st.session_state.get("impact_effort_selected_step")
            if step is not None:
                row = df[df["Step"] == step]
                if not row.empty:
                    parts = row.iloc[0]["Recursive Path"].split(" > ")[:3]
                    st.session_state["impact_effort_focus_node"] = " > ".join(parts) if parts else "Root"
                    st.rerun()
    with action_cols[2]:
        if st.button("Clear marker selection", use_container_width=True):
            st.session_state["impact_effort_selected_step"] = None
            st.rerun()

    st.markdown(
        f"""
<div style="margin-top: 10px; font-size: 14px; color: black; text-align: justify;">
{BOTTOM_NOTE}
</div>
""",
        unsafe_allow_html=True,
    )

    st.markdown("### RCPSP priority handoff")
    priority_df = build_priority_table(df)
    priority_view = priority_df[[
        "RCPSP Priority", "Step", "Description", "Quadrant", "Priority Bucket", "Impact Score", "Effort Score",
        "Duration Seconds", "Recursive Path", "Cluster"
    ]].copy()
    priority_view["Impact Score"] = priority_view["Impact Score"].round(3)
    priority_view["Effort Score"] = priority_view["Effort Score"].round(3)
    priority_view.rename(columns={"Duration Seconds": "Duration (s)"}, inplace=True)
    st.dataframe(priority_view, use_container_width=True, hide_index=True)

    csv_bytes = priority_view.to_csv(index=False).encode("utf-8")
    st.download_button("Download RCPSP priority CSV", data=csv_bytes, file_name="impact_effort_rcpsp_priority.csv", mime="text/csv")

    st.markdown("### Activity classification table")
    filter_col1, filter_col2 = st.columns([1, 1])
    with filter_col1:
        impact_filter = st.multiselect("Filter Impact", options=["Low", "High"], default=["Low", "High"])
    with filter_col2:
        effort_filter = st.multiselect("Filter Effort", options=["Low", "High"], default=["Low", "High"])

    display_df = df.copy()
    display_df = display_df[
        display_df["Impact"].isin(impact_filter) & display_df["Effort"].isin(effort_filter)
    ].copy()
    display_df = display_df.sort_values(["Step", "Impact", "Effort"]).reset_index(drop=True)
    display_df["Duration"] = display_df["Duration Seconds"].round(0).astype(int).astype(str) + " s"

    table_cols = [
        "Step", "Description", "Activity", "Therblig / Motion Type", "Motion Class", "Impact", "Effort", "Quadrant",
        "Impact Score", "Effort Score", "Recursive Path", "Cluster", "Duration", "Lean / Six Sigma Logic",
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
        st.caption("Click a marker once to select it and highlight the matching row in the table.")
except Exception as e:
    st.error(f"Error: {e}")
