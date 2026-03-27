import textwrap

import networkx as nx
import plotly.graph_objects as go
import streamlit as st

from session_utils import require_workbook
from modules.m01_workbook_loader import load_current_state_steps
from modules.m02_nlp_normalization import normalize_steps
from modules.m05_macro_tasks import generate_macro_tasks
from modules.m06_precedence_network import build_precedence_outputs


st.set_page_config(page_title="Precedence Network", layout="wide")
st.title("Precedence Network")
ctx = require_workbook()

st.markdown(
    """
    <style>
    div[data-testid="stPlotlyChart"] .modebar {
        transform: scale(1.45);
        transform-origin: top right;
        opacity: 1 !important;
        right: 18px !important;
        top: 12px !important;
    }
    div[data-testid="stPlotlyChart"] .modebar-group {
        background: rgba(255,255,255,0.96) !important;
        border: 1px solid #cfcfcf !important;
        border-radius: 10px !important;
        padding: 3px 5px !important;
        margin-left: 6px !important;
        box-shadow: 0 1px 6px rgba(0,0,0,0.12) !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

top1, top2, top3 = st.columns([1, 1, 1])
with top1:
    show_simplified = st.checkbox("Show simplified graph", value=True)
with top2:
    show_task_ids = st.checkbox("Show task IDs in labels", value=False)
with top3:
    curved_edges = st.checkbox("Use curved edges", value=False)

controls = st.columns(8)
with controls[0]:
    chart_height = st.slider("Chart height", min_value=650, max_value=1800, value=1000, step=50)
with controls[1]:
    chart_width = st.slider("Chart width", min_value=1200, max_value=4200, value=2200, step=100)
with controls[2]:
    node_size = st.slider("Node size", min_value=34, max_value=95, value=58, step=1)
with controls[3]:
    text_size = st.slider("Text size", min_value=10, max_value=24, value=14, step=1)
with controls[4]:
    x_gap = st.slider("Horizontal spacing", min_value=1.8, max_value=7.0, value=3.0, step=0.1)
with controls[5]:
    y_gap = st.slider("Base vertical spacing", min_value=12.0, max_value=16.0, value=3.2, step=0.1)
with controls[6]:
    lane_spread = st.slider("Lane spread multiplier", min_value=1.0, max_value=3.0, value=1.6, step=0.1)
with controls[7]:
    label_width = st.slider("Label wrap width", min_value=10, max_value=28, value=16, step=1)

download_cols = st.columns(3)
with download_cols[0]:
    download_scale = st.slider("Download scale", min_value=1, max_value=5, value=2, step=1)
with download_cols[1]:
    download_name = st.text_input("Download filename", value="precedence_network")
with download_cols[2]:
    st.caption("Use the toolbar at top-right to pan, zoom, reset, and download.")

steps = load_current_state_steps(ctx["excel_file"], sheet_name=ctx["sheet_name"])
normalized = normalize_steps(steps)
tasks = generate_macro_tasks(normalized)
outputs = build_precedence_outputs(tasks)

full_graph = outputs["graph"]


def _copy_attrs(source: nx.DiGraph, target: nx.DiGraph) -> nx.DiGraph:
    for n, attrs in source.nodes(data=True):
        if n in target.nodes:
            target.nodes[n].update(attrs)
    return target


if show_simplified:
    try:
        graph = nx.transitive_reduction(full_graph)
        graph = _copy_attrs(full_graph, graph)
    except Exception:
        graph = full_graph.copy()
else:
    graph = full_graph.copy()


def build_left_to_right_positions(g: nx.DiGraph, x_step: float, y_step: float, spread_mult: float):
    level = {n: 0 for n in g.nodes()}
    for n in nx.topological_sort(g):
        preds = list(g.predecessors(n))
        if preds:
            level[n] = max(level[p] for p in preds) + 1

    grouped = {}
    for node, lv in level.items():
        grouped.setdefault(lv, []).append(node)

    max_group = max((len(nodes) for nodes in grouped.values()), default=1)
    auto_y = y_step * spread_mult

    pos = {}
    for lv in sorted(grouped):
        nodes = sorted(grouped[lv])
        center = (len(nodes) - 1) / 2.0
        local_y = auto_y * (1.0 + 0.18 * max(0, len(nodes) - 2))
        for idx, node in enumerate(nodes):
            x = lv * x_step
            y = (center - idx) * local_y
            pos[node] = (x, y)
    return pos, level, max_group


def short_label(g: nx.DiGraph, node_id: int, width: int) -> str:
    name = str(g.nodes[node_id].get("name", node_id))
    lines = textwrap.wrap(name, width=width)[:2]
    base = "<br>".join(lines) if lines else name
    if show_task_ids:
        return f"<b>{node_id}</b><br>{base}"
    return base


def hover_label(g: nx.DiGraph, node_id: int) -> str:
    attrs = g.nodes[node_id]
    name = attrs.get("name", str(node_id))
    duration = attrs.get("duration_sec", "—")
    preds = attrs.get("predecessors", [])
    resources = attrs.get("resources", [])
    preds_text = ", ".join(map(str, preds)) if preds else "—"
    if isinstance(resources, (list, tuple, set)):
        resources_text = ", ".join(map(str, resources)) if resources else "—"
    else:
        resources_text = str(resources) if resources else "—"
    return (
        f"<b>{name}</b><br>"
        f"Task ID: {node_id}<br>"
        f"Duration (sec): {duration}<br>"
        f"Predecessors: {preds_text}<br>"
        f"Resources: {resources_text}"
    )


pos, level, max_group = build_left_to_right_positions(graph, x_gap, y_gap, lane_spread)
critical_ids = outputs.get("critical_path_task_ids", [])
critical_nodes = set(critical_ids)
critical_edges = set(zip(critical_ids[:-1], critical_ids[1:]))

edge_x_normal, edge_y_normal = [], []
edge_x_critical, edge_y_critical = [], []

for u, v in graph.edges():
    x0, y0 = pos[u]
    x1, y1 = pos[v]

    if curved_edges and level[u] != level[v]:
        mid_x = (x0 + x1) / 2.0
        xs = [x0, mid_x, x1, None]
        ys = [y0, y0, y1, None]
    else:
        xs = [x0, x1, None]
        ys = [y0, y1, None]

    if (u, v) in critical_edges:
        edge_x_critical.extend(xs)
        edge_y_critical.extend(ys)
    else:
        edge_x_normal.extend(xs)
        edge_y_normal.extend(ys)

line_shape = "spline" if curved_edges else "linear"

edge_trace_normal = go.Scatter(
    x=edge_x_normal,
    y=edge_y_normal,
    mode="lines",
    line=dict(width=1.5, color="#A0A7AE", shape=line_shape),
    hoverinfo="skip",
    showlegend=False,
)

edge_trace_critical = go.Scatter(
    x=edge_x_critical,
    y=edge_y_critical,
    mode="lines",
    line=dict(width=3.0, color="#C0392B", shape=line_shape),
    hoverinfo="skip",
    showlegend=False,
)

x_normal, y_normal, text_normal, hover_normal = [], [], [], []
x_critical, y_critical, text_critical, hover_critical = [], [], [], []

for node in graph.nodes():
    x, y = pos[node]
    label = short_label(graph, node, label_width)
    hover = hover_label(graph, node)
    if node in critical_nodes:
        x_critical.append(x)
        y_critical.append(y)
        text_critical.append(label)
        hover_critical.append(hover)
    else:
        x_normal.append(x)
        y_normal.append(y)
        text_normal.append(label)
        hover_normal.append(hover)

node_trace_normal = go.Scatter(
    x=x_normal,
    y=y_normal,
    mode="markers+text",
    text=text_normal,
    textposition="middle center",
    textfont=dict(size=text_size, color="#111111"),
    hovertext=hover_normal,
    hovertemplate="%{hovertext}<extra></extra>",
    marker=dict(size=node_size, color="#D6EAF8", line=dict(width=2, color="#1F4E79")),
    showlegend=False,
)

node_trace_critical = go.Scatter(
    x=x_critical,
    y=y_critical,
    mode="markers+text",
    text=text_critical,
    textposition="middle center",
    textfont=dict(size=text_size, color="#111111"),
    hovertext=hover_critical,
    hovertemplate="%{hovertext}<extra></extra>",
    marker=dict(size=node_size + 6, color="#F9D6D5", line=dict(width=2.2, color="#922B21")),
    showlegend=False,
)

xs = [p[0] for p in pos.values()] if pos else [0]
ys = [p[1] for p in pos.values()] if pos else [0]
x_pad = max(0.7, (max(xs) - min(xs)) * 0.04) if len(xs) > 1 else 1
y_pad = max(1.8, (max(ys) - min(ys)) * 0.18) if len(ys) > 1 else 2.5

fig = go.Figure(data=[edge_trace_normal, edge_trace_critical, node_trace_normal, node_trace_critical])

fig.update_layout(
    height=chart_height,
    margin=dict(l=10, r=10, t=10, b=10),
    paper_bgcolor="white",
    plot_bgcolor="white",
    dragmode="pan",
    xaxis=dict(
        visible=False,
        range=[min(xs) - x_pad, max(xs) + x_pad],
        fixedrange=False,
    ),
    yaxis=dict(
        visible=False,
        range=[min(ys) - y_pad, max(ys) + y_pad],
        fixedrange=False,
        scaleanchor="x",
        scaleratio=1,
    ),
)

st.plotly_chart(
    fig,
    width="stretch",
    config={
        "scrollZoom": True,
        "displayModeBar": True,
        "displaylogo": False,
        "modeBarButtonsToAdd": ["zoom2d", "pan2d", "resetScale2d", "toImage"],
        "toImageButtonOptions": {
            "format": "png",
            "filename": download_name,
            "width": chart_width,
            "height": chart_height,
            "scale": download_scale,
        },
    },
)

summary_cols = st.columns(4)
summary_cols[0].metric("Tasks shown", len(graph.nodes()))
summary_cols[1].metric("Precedence links shown", len(graph.edges()))
summary_cols[2].metric("Critical path tasks", len(critical_nodes))
summary_cols[3].metric("Max tasks in a layer", max_group)

with st.expander("Technological precedence table", expanded=False):
    st.dataframe(outputs["technological_precedence"], width="stretch")

with st.expander("Resource-linked conflicts", expanded=False):
    st.dataframe(outputs["resource_linked_conflicts"], width="stretch")

if "parallelizable_tasks" in outputs:
    with st.expander("Parallelizable task pairs", expanded=False):
        st.dataframe(outputs["parallelizable_tasks"], width="stretch")
