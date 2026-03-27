import textwrap

import matplotlib.pyplot as plt
import networkx as nx
import streamlit as st

from session_utils import require_workbook
from modules.m01_workbook_loader import load_current_state_steps
from modules.m02_nlp_normalization import normalize_steps
from modules.m05_macro_tasks import generate_macro_tasks, tasks_to_df
from modules.m06_precedence_network import build_precedence_outputs

st.set_page_config(page_title="Precedence Network", layout="wide")
st.title("Precedence Network")
ctx = require_workbook()

steps = load_current_state_steps(ctx["excel_file"], sheet_name=ctx["sheet_name"])
normalized = normalize_steps(steps)
tasks = generate_macro_tasks(normalized)
outputs = build_precedence_outputs(tasks)

control_a, control_b, control_c = st.columns([1, 1, 1])
with control_a:
    show_reduced = st.checkbox(
        "Show simplified network",
        value=True,
        help="Shows only essential precedence links for readability.",
    )
with control_b:
    zoom = st.slider("Graph zoom", min_value=100, max_value=220, value=150, step=10)
with control_c:
    node_scale = st.slider("Node size", min_value=80, max_value=180, value=125, step=5)

graph = outputs["display_graph"] if show_reduced else outputs["graph"]


def build_layered_positions(graph: nx.DiGraph, x_gap: float, y_gap: float):
    if graph.number_of_nodes() == 0:
        return {}
    level = {n: 0 for n in graph.nodes()}
    for n in nx.topological_sort(graph):
        preds = list(graph.predecessors(n))
        if preds:
            level[n] = max(level[p] for p in preds) + 1
    grouped = {}
    for node, lv in level.items():
        grouped.setdefault(lv, []).append(node)
    pos = {}
    for lv in sorted(grouped):
        nodes = sorted(grouped[lv])
        center = (len(nodes) - 1) / 2.0
        for i, node in enumerate(nodes):
            pos[node] = (lv * x_gap, (center - i) * y_gap)
    return pos


zoom_factor = zoom / 100.0
pos = build_layered_positions(graph, x_gap=9.0 * zoom_factor, y_gap=3.1 * zoom_factor)
label_width = 16 if graph.number_of_nodes() <= 18 else 13 if graph.number_of_nodes() <= 28 else 11
labels = {
    n: "\n".join(textwrap.wrap(str(graph.nodes[n].get("name", n)), width=label_width)[:4])
    for n in graph.nodes()
}
cp_nodes = set(outputs["critical_path_task_ids"])
cp_edges = set(
    outputs["critical_path_display_edges"]
    if show_reduced
    else zip(outputs["critical_path_task_ids"][:-1], outputs["critical_path_task_ids"][1:])
)

levels = [x for x, _ in pos.values()] or [0]
ys = [y for _, y in pos.values()] or [0]
span_x = max(levels) - min(levels) if levels else 0
span_y = max(ys) - min(ys) if ys else 0
fig_w = max(22, 8 + span_x * 0.34)
fig_h = max(12, 6 + span_y * 10)
fig, ax = plt.subplots(figsize=(fig_w, fig_h), dpi=180)

normal_edges = [e for e in graph.edges() if e not in cp_edges]
normal_nodes = [n for n in graph.nodes() if n not in cp_nodes]
base_node_size = 7800 if graph.number_of_nodes() <= 16 else 6400 if graph.number_of_nodes() <= 28 else 5200
node_size = int(base_node_size * (node_scale / 100.0) ** 2)
font_size = 12 if graph.number_of_nodes() <= 16 else 10 if graph.number_of_nodes() <= 28 else 9
font_size = max(8, int(font_size * (zoom_factor ** 0.35)))

nx.draw_networkx_edges(
    graph,
    pos,
    edgelist=normal_edges,
    ax=ax,
    arrows=True,
    edge_color="#9aa5a8",
    width=2.4,
    arrowsize=24,
    connectionstyle="arc3,rad=0.03",
    min_source_margin=18,
    min_target_margin=18,
)
if cp_edges:
    nx.draw_networkx_edges(
        graph,
        pos,
        edgelist=list(cp_edges),
        ax=ax,
        arrows=True,
        edge_color="#c0392b",
        width=3.8,
        arrowsize=26,
        connectionstyle="arc3,rad=0.03",
        min_source_margin=18,
        min_target_margin=18,
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
        linewidths=1.6,
    )
critical_visible = list(cp_nodes & set(graph.nodes()))
if critical_visible:
    nx.draw_networkx_nodes(
        graph,
        pos,
        nodelist=critical_visible,
        ax=ax,
        node_size=int(node_size * 1.06),
        node_color="#f9d6d5",
        edgecolors="#922b21",
        linewidths=1.7,
    )

nx.draw_networkx_labels(graph, pos, labels=labels, ax=ax, font_size=font_size, font_weight="bold")

x_values = [x for x, _ in pos.values()] or [0]
y_values = [y for _, y in pos.values()] or [0]
ax.set_xlim(min(x_values) - 6 * zoom_factor, max(x_values) + 6 * zoom_factor)
ax.set_ylim(min(y_values) - 4 * zoom_factor, max(y_values) + 4 * zoom_factor)
ax.set_axis_off()
plt.tight_layout(pad=1.2)

st.pyplot(fig, use_container_width=False)
st.caption("Red edges show the critical path. Increase Graph zoom for a larger network view, or turn off simplified view to inspect every original precedence link.")

st.subheader("Discovered tasks")
st.dataframe(tasks_to_df(tasks), use_container_width=True)

st.subheader("Displayed precedence")
st.dataframe(outputs["display_precedence"] if show_reduced else outputs["technological_precedence"], use_container_width=True)

with st.expander("All precedence links"):
    st.dataframe(outputs["technological_precedence"], use_container_width=True)

with st.expander("Resource-linked conflicts"):
    st.dataframe(outputs["resource_linked_conflicts"], use_container_width=True)
