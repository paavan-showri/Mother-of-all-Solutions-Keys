import textwrap

import matplotlib.pyplot as plt
import networkx as nx
import streamlit as st

from session_utils import require_workbook
from modules.m01_workbook_loader import load_current_state_steps
from modules.m02_nlp_normalization import normalize_steps
from modules.m05_macro_tasks import generate_macro_tasks
from modules.m06_precedence_network import build_precedence_outputs


st.set_page_config(page_title="Precedence Network", layout="wide")
st.title("Precedence Network")
ctx = require_workbook()

# Keep UI simple and avoid deprecated container-width behavior.
show_simplified = st.checkbox("Show simplified graph (recommended)", value=True)

controls = st.columns(6)
with controls[0]:
    chart_width = st.slider("Chart width", min_value=16, max_value=40, value=24, step=1)
with controls[1]:
    chart_height = st.slider("Chart height", min_value=8, max_value=24, value=14, step=1)
with controls[2]:
    node_size = st.slider("Node size", min_value=3000, max_value=22000, value=11000, step=500)
with controls[3]:
    font_size = st.slider("Text size", min_value=8, max_value=22, value=14, step=1)
with controls[4]:
    h_gap = st.slider("Horizontal spacing", min_value=2.0, max_value=8.0, value=4.2, step=0.2)
with controls[5]:
    v_gap = st.slider("Vertical spacing", min_value=1.0, max_value=5.0, value=2.0, step=0.1)

label_width = st.slider("Wrap width", min_value=10, max_value=28, value=18, step=1)

steps = load_current_state_steps(ctx["excel_file"], sheet_name=ctx["sheet_name"])
normalized = normalize_steps(steps)
tasks = generate_macro_tasks(normalized)
outputs = build_precedence_outputs(tasks)

full_graph = outputs.get("graph")
graph = outputs.get("display_graph", full_graph) if show_simplified else full_graph


def _copy_attrs_from(source: nx.DiGraph, target: nx.DiGraph) -> nx.DiGraph:
    for n, attrs in source.nodes(data=True):
        if n in target.nodes:
            target.nodes[n].update(attrs)
    return target


if show_simplified and graph is full_graph:
    # Build a display graph even if the module does not provide one.
    try:
        graph = nx.transitive_reduction(full_graph)
        graph = _copy_attrs_from(full_graph, graph)
    except Exception:
        graph = full_graph.copy()


def build_layered_positions(g: nx.DiGraph, x_gap: float, y_gap: float):
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
        n_nodes = len(nodes)
        center = (n_nodes - 1) / 2.0
        for i, node in enumerate(nodes):
            pos[node] = (lv * x_gap, (center - i) * y_gap)
    return pos


def wrapped_label(g: nx.DiGraph, n: int, width: int) -> str:
    name = g.nodes[n].get("name", str(n))
    wrapped = "\n".join(textwrap.wrap(str(name), width=width)[:3])
    return f"{n}\n{wrapped}" if wrapped else str(n)


pos = build_layered_positions(graph, h_gap, v_gap)
labels = {n: wrapped_label(graph, n, label_width) for n in graph.nodes()}

cp_nodes = set(outputs.get("critical_path_task_ids", []))
cp_edges = set(zip(outputs.get("critical_path_task_ids", [])[:-1], outputs.get("critical_path_task_ids", [])[1:]))

fig, ax = plt.subplots(figsize=(chart_width, chart_height), dpi=120)
fig.patch.set_facecolor("white")
ax.set_facecolor("white")

normal_edges = [e for e in graph.edges() if e not in cp_edges]
normal_nodes = [n for n in graph.nodes() if n not in cp_nodes]
highlight_nodes = [n for n in graph.nodes() if n in cp_nodes]
highlight_edges = [e for e in graph.edges() if e in cp_edges]

nx.draw_networkx_edges(
    graph,
    pos,
    edgelist=normal_edges,
    ax=ax,
    arrows=True,
    edge_color="#98a3a3",
    width=1.8,
    arrowsize=18,
    connectionstyle="arc3,rad=0.03",
)
if highlight_edges:
    nx.draw_networkx_edges(
        graph,
        pos,
        edgelist=highlight_edges,
        ax=ax,
        arrows=True,
        edge_color="#c0392b",
        width=3.2,
        arrowsize=20,
        connectionstyle="arc3,rad=0.03",
    )

if normal_nodes:
    nx.draw_networkx_nodes(
        graph,
        pos,
        nodelist=normal_nodes,
        ax=ax,
        node_size=node_size,
        node_color="#d6eaf8",
        edgecolors="#1f4e79",
        linewidths=1.3,
    )
if highlight_nodes:
    nx.draw_networkx_nodes(
        graph,
        pos,
        nodelist=highlight_nodes,
        ax=ax,
        node_size=int(node_size * 1.08),
        node_color="#f9d6d5",
        edgecolors="#922b21",
        linewidths=1.5,
    )

nx.draw_networkx_labels(
    graph,
    pos,
    labels=labels,
    ax=ax,
    font_size=font_size,
    font_weight="bold",
    verticalalignment="center",
)

# Tighten bounds aggressively to remove excess white space.
if pos:
    xs = [p[0] for p in pos.values()]
    ys = [p[1] for p in pos.values()]
    x_pad = max(0.6, (max(xs) - min(xs)) * 0.02)
    y_pad = max(0.6, (max(ys) - min(ys)) * 0.10)
    ax.set_xlim(min(xs) - x_pad, max(xs) + x_pad)
    ax.set_ylim(min(ys) - y_pad, max(ys) + y_pad)

ax.margins(x=0.01, y=0.02)
ax.axis("off")
fig.subplots_adjust(left=0.01, right=0.99, top=0.99, bottom=0.01)

st.pyplot(fig)

with st.expander("Technological precedence table", expanded=False):
    st.dataframe(outputs["technological_precedence"], width="stretch")

with st.expander("Resource-linked conflicts", expanded=False):
    st.dataframe(outputs["resource_linked_conflicts"], width="stretch")

if "parallelizable_tasks" in outputs:
    with st.expander("Parallelizable task pairs", expanded=False):
        st.dataframe(outputs["parallelizable_tasks"], width="stretch")
