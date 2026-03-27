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

show_reduced = st.checkbox(
    "Show simplified network",
    value=True,
    help="Shows only essential precedence links for readability.",
)
graph = outputs["display_graph"] if show_reduced else outputs["graph"]


def build_layered_positions(graph: nx.DiGraph):
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
    x_gap = 6.0
    y_gap = 2.2
    for lv in sorted(grouped):
        nodes = sorted(grouped[lv])
        center = (len(nodes) - 1) / 2.0
        for i, node in enumerate(nodes):
            pos[node] = (lv * x_gap, (center - i) * y_gap)
    return pos


pos = build_layered_positions(graph)
label_width = 14 if graph.number_of_nodes() <= 18 else 11
labels = {
    n: "\n".join(textwrap.wrap(str(graph.nodes[n].get("name", n)), width=label_width)[:3])
    for n in graph.nodes()
}
cp_nodes = set(outputs["critical_path_task_ids"])
cp_edges = set(outputs["critical_path_display_edges"] if show_reduced else zip(outputs["critical_path_task_ids"][:-1], outputs["critical_path_task_ids"][1:]))

levels = [x for x, _ in pos.values()] or [0]
ys = [y for _, y in pos.values()] or [0]
fig_w = max(14, 4 + (max(levels) - min(levels)) * 0.9)
fig_h = max(8, 3 + (max(ys) - min(ys)) * 0.55)
fig, ax = plt.subplots(figsize=(fig_w, fig_h))

normal_edges = [e for e in graph.edges() if e not in cp_edges]
normal_nodes = [n for n in graph.nodes() if n not in cp_nodes]
node_size = 7200 if graph.number_of_nodes() <= 16 else 5600 if graph.number_of_nodes() <= 28 else 4200
font_size = 10 if graph.number_of_nodes() <= 16 else 9 if graph.number_of_nodes() <= 28 else 8

nx.draw_networkx_edges(graph, pos, edgelist=normal_edges, ax=ax, arrows=True, edge_color="#9aa5a8", width=1.8, arrowsize=18, connectionstyle="arc3,rad=0.02")
if cp_edges:
    nx.draw_networkx_edges(graph, pos, edgelist=list(cp_edges), ax=ax, arrows=True, edge_color="#c0392b", width=3.0, arrowsize=20, connectionstyle="arc3,rad=0.02")
if normal_nodes:
    nx.draw_networkx_nodes(graph, pos, nodelist=normal_nodes, ax=ax, node_size=node_size, node_color="#d6eaf8", edgecolors="#1f4e79", linewidths=1.3)
critical_visible = list(cp_nodes & set(graph.nodes()))
if critical_visible:
    nx.draw_networkx_nodes(graph, pos, nodelist=critical_visible, ax=ax, node_size=node_size * 1.04, node_color="#f9d6d5", edgecolors="#922b21", linewidths=1.4)

nx.draw_networkx_labels(graph, pos, labels=labels, ax=ax, font_size=font_size, font_weight="bold")
ax.set_axis_off()
plt.tight_layout()

st.pyplot(fig, use_container_width=True)
st.caption("Red edges show the critical path. Turn off simplified view to inspect every original precedence link.")

st.subheader("Discovered tasks")
st.dataframe(tasks_to_df(tasks), use_container_width=True)

st.subheader("Displayed precedence")
st.dataframe(outputs["display_precedence"] if show_reduced else outputs["technological_precedence"], use_container_width=True)

with st.expander("All precedence links"):
    st.dataframe(outputs["technological_precedence"], use_container_width=True)

with st.expander("Resource-linked conflicts"):
    st.dataframe(outputs["resource_linked_conflicts"], use_container_width=True)
