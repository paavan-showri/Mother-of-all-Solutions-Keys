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


def build_layered_positions(graph: nx.DiGraph):
    level = {n: 0 for n in graph.nodes()}
    for n in nx.topological_sort(graph):
        preds = list(graph.predecessors(n))
        if preds:
            level[n] = max(level[p] for p in preds) + 1
    grouped = {}
    for node, lv in level.items():
        grouped.setdefault(lv, []).append(node)
    pos = {}
    for lv, nodes in grouped.items():
        nodes = sorted(nodes)
        n = len(nodes)
        for i, node in enumerate(nodes):
            pos[node] = (lv * 4.4, (n - 1) * 2.4 / 2 - i * 2.4)
    return pos


pos = build_layered_positions(outputs["graph"])
labels = {
    n: "\n".join(textwrap.wrap(outputs["graph"].nodes[n]["name"], width=14)[:3])
    for n in outputs["graph"].nodes()
}
cp_nodes = set(outputs["critical_path_task_ids"])
cp_edges = set(zip(outputs["critical_path_task_ids"][:-1], outputs["critical_path_task_ids"][1:]))

fig, ax = plt.subplots(figsize=(16, 10))
normal_edges = [e for e in outputs["graph"].edges() if e not in cp_edges]
normal_nodes = [n for n in outputs["graph"].nodes() if n not in cp_nodes]
nx.draw_networkx_edges(outputs["graph"], pos, edgelist=normal_edges, ax=ax, arrows=True, edge_color="#7f8c8d", width=2)
nx.draw_networkx_edges(outputs["graph"], pos, edgelist=list(cp_edges), ax=ax, arrows=True, edge_color="#c0392b", width=3.5)
nx.draw_networkx_nodes(outputs["graph"], pos, nodelist=normal_nodes, ax=ax, node_size=8500, node_color="#d6eaf8", edgecolors="#1f4e79")
nx.draw_networkx_nodes(outputs["graph"], pos, nodelist=list(cp_nodes), ax=ax, node_size=9500, node_color="#f9d6d5", edgecolors="#922b21")
nx.draw_networkx_labels(outputs["graph"], pos, labels=labels, ax=ax, font_size=10, font_weight="bold")
ax.axis("off")

st.pyplot(fig, use_container_width=True)
st.dataframe(tasks_to_df(tasks), use_container_width=True)
st.dataframe(outputs["technological_precedence"], use_container_width=True)
st.dataframe(outputs["resource_linked_conflicts"], use_container_width=True)
