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

show_simplified = st.checkbox("Show simplified graph", value=True)

controls = st.columns(6)
with controls[0]:
    chart_height = st.slider("Chart height", min_value=600, max_value=1800, value=950, step=50)
with controls[1]:
    x_gap = st.slider("Horizontal spacing", min_value=2.0, max_value=8.0, value=4.2, step=0.2)
with controls[2]:
    y_gap = st.slider("Vertical spacing", min_value=1.0, max_value=5.0, value=2.1, step=0.1)
with controls[3]:
    node_size = st.slider("Node size", min_value=35, max_value=120, value=68, step=1)
with controls[4]:
    text_size = st.slider("Text size", min_value=10, max_value=26, value=16, step=1)
with controls[5]:
    wrap_width = st.slider("Wrap width", min_value=10, max_value=28, value=18, step=1)

steps = load_current_state_steps(ctx["excel_file"], sheet_name=ctx["sheet_name"])
normalized = normalize_steps(steps)
tasks = generate_macro_tasks(normalized)
outputs = build_precedence_outputs(tasks)

full_graph = outputs["graph"]


def copy_attrs(source: nx.DiGraph, target: nx.DiGraph) -> nx.DiGraph:
    for n, attrs in source.nodes(data=True):
        if n in target.nodes:
            target.nodes[n].update(attrs)
    return target


if show_simplified:
    try:
        graph = nx.transitive_reduction(full_graph)
        graph = copy_attrs(full_graph, graph)
    except Exception:
        graph = full_graph.copy()
else:
    graph = full_graph.copy()


def build_layered_positions(g: nx.DiGraph, x_gap_value: float, y_gap_value: float):
    level = {n: 0 for n in g.nodes()}
    for n in nx.topological_sort(g):
        preds = list(g.predecessors(n))
        if preds:
            level[n] = max(level[p] for p in preds) + 1

    grouped = {}
    for node, lv in level.items():
        grouped.setdefault(lv, []).append(node)

    pos = {}
    for lv, nodes in grouped.items():
        nodes = sorted(nodes)
        center = (len(nodes) - 1) / 2.0
        for i, node in enumerate(nodes):
            pos[node] = (lv * x_gap_value, (center - i) * y_gap_value)
    return pos


def node_label(g: nx.DiGraph, node_id: int, width: int) -> str:
    name = g.nodes[node_id].get("name", str(node_id))
    wrapped = "<br>".join(textwrap.wrap(str(name), width=width)[:3])
    return f"<b>{node_id}</b><br>{wrapped}" if wrapped else f"<b>{node_id}</b>"


def node_hover(g: nx.DiGraph, node_id: int) -> str:
    attrs = g.nodes[node_id]
    name = attrs.get("name", str(node_id))
    duration = attrs.get("duration_sec", "—")
    preds = attrs.get("predecessors", [])
    resources = attrs.get("resources", [])
    if isinstance(resources, (list, tuple, set)):
        resources_text = ", ".join(map(str, resources)) if resources else "—"
    else:
        resources_text = str(resources) if resources else "—"
    preds_text = ", ".join(map(str, preds)) if preds else "—"
    return (
        f"<b>{name}</b><br>"
        f"Task ID: {node_id}<br>"
        f"Duration (sec): {duration}<br>"
        f"Predecessors: {preds_text}<br>"
        f"Resources: {resources_text}"
    )


pos = build_layered_positions(graph, x_gap, y_gap)
critical_ids = outputs.get("critical_path_task_ids", [])
critical_nodes = set(critical_ids)
critical_edges = set(zip(critical_ids[:-1], critical_ids[1:]))

edge_x_normal, edge_y_normal = [], []
edge_x_critical, edge_y_critical = [], []

for u, v in graph.edges():
    x0, y0 = pos[u]
    x1, y1 = pos[v]
    bucket_x = edge_x_critical if (u, v) in critical_edges else edge_x_normal
    bucket_y = edge_y_critical if (u, v) in critical_edges else edge_y_normal
    bucket_x.extend([x0, x1, None])
    bucket_y.extend([y0, y1, None])

edge_trace_normal = go.Scatter(
    x=edge_x_normal,
    y=edge_y_normal,
    mode="lines",
    line=dict(width=1.8, color="#98a3a3"),
    hoverinfo="skip",
    showlegend=False,
)

edge_trace_critical = go.Scatter(
    x=edge_x_critical,
    y=edge_y_critical,
    mode="lines",
    line=dict(width=3.2, color="#c0392b"),
    hoverinfo="skip",
    showlegend=False,
)

x_normal, y_normal, text_normal, hover_normal = [], [], [], []
x_critical, y_critical, text_critical, hover_critical = [], [], [], []

for n in graph.nodes():
    x, y = pos[n]
    label = node_label(graph, n, wrap_width)
    hover = node_hover(graph, n)
    if n in critical_nodes:
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
    marker=dict(
        size=node_size,
        color="#d6eaf8",
        line=dict(width=2, color="#1f4e79"),
    ),
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
    marker=dict(
        size=max(node_size + 6, int(node_size * 1.08)),
        color="#f9d6d5",
        line=dict(width=2.2, color="#922b21"),
    ),
    showlegend=False,
)

xs = [p[0] for p in pos.values()] if pos else [0]
ys = [p[1] for p in pos.values()] if pos else [0]
x_pad = max(0.6, (max(xs) - min(xs)) * 0.03) if len(xs) > 1 else 1
y_pad = max(0.8, (max(ys) - min(ys)) * 0.15) if len(ys) > 1 else 1

fig = go.Figure(
    data=[edge_trace_normal, edge_trace_critical, node_trace_normal, node_trace_critical]
)

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
        "displaylogo": False,
        "modeBarButtonsToAdd": ["zoom2d", "pan2d", "resetScale2d"],
    },
)

with st.expander("Technological precedence table", expanded=False):
    st.dataframe(outputs["technological_precedence"], width="stretch")

with st.expander("Resource-linked conflicts", expanded=False):
    st.dataframe(outputs["resource_linked_conflicts"], width="stretch")

if "parallelizable_tasks" in outputs:
    with st.expander("Parallelizable task pairs", expanded=False):
        st.dataframe(outputs["parallelizable_tasks"], width="stretch")
