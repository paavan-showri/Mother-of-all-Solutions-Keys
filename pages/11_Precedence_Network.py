import streamlit as st
import pandas as pd
import networkx as nx
import matplotlib.pyplot as plt

from session_utils import require_workbook
from modules.m01_workbook_loader import load_precedence_tasks
from modules.m06_precedence_network import build_precedence_outputs


st.set_page_config(page_title="Precedence Network", layout="wide")
st.title("Precedence Network")


def build_layered_positions(g: nx.DiGraph):
    """
    Create a clean left-to-right layered layout based on topological levels.
    """
    topo = list(nx.topological_sort(g))

    level = {}
    for node in topo:
        preds = list(g.predecessors(node))
        if not preds:
            level[node] = 0
        else:
            level[node] = max(level[p] for p in preds) + 1

    levels = {}
    for node, lv in level.items():
        levels.setdefault(lv, []).append(node)

    pos = {}
    for lv, nodes in levels.items():
        nodes_sorted = sorted(nodes)
        n = len(nodes_sorted)
        for i, node in enumerate(nodes_sorted):
            # x = level, y spaced vertically
            pos[node] = (lv, -i + (n - 1) / 2)

    return pos


def draw_precedence_graph(g: nx.DiGraph, critical_path_task_ids, tasks):
    """
    Draw precedence network with critical path highlighted.
    """
    fig, ax = plt.subplots(figsize=(16, 8))

    pos = build_layered_positions(g)

    critical_nodes = set(critical_path_task_ids)
    critical_edges = set(zip(critical_path_task_ids[:-1], critical_path_task_ids[1:]))

    # Build labels: TaskID + short task name
    task_name_map = {t.task_id: t.name for t in tasks}
    labels = {}
    for node in g.nodes():
        task_name = str(task_name_map.get(node, ""))
        short_name = task_name if len(task_name) <= 22 else task_name[:22] + "..."
        labels[node] = f"{node}\n{short_name}"

    normal_nodes = [n for n in g.nodes() if n not in critical_nodes]
    normal_edges = [e for e in g.edges() if e not in critical_edges]
    crit_nodes = [n for n in g.nodes() if n in critical_nodes]
    crit_edges = [e for e in g.edges() if e in critical_edges]

    # Draw normal edges
    nx.draw_networkx_edges(
        g,
        pos,
        edgelist=normal_edges,
        ax=ax,
        arrows=True,
        arrowstyle="->",
        arrowsize=20,
        width=1.8,
        edge_color="gray",
        alpha=0.7,
        min_source_margin=12,
        min_target_margin=12,
    )

    # Draw critical edges
    nx.draw_networkx_edges(
        g,
        pos,
        edgelist=crit_edges,
        ax=ax,
        arrows=True,
        arrowstyle="->",
        arrowsize=24,
        width=3.0,
        edge_color="red",
        alpha=0.95,
        min_source_margin=12,
        min_target_margin=12,
    )

    # Draw normal nodes
    nx.draw_networkx_nodes(
        g,
        pos,
        nodelist=normal_nodes,
        ax=ax,
        node_size=3000,
        node_color="lightblue",
        edgecolors="black",
        linewidths=1.2,
    )

    # Draw critical nodes
    nx.draw_networkx_nodes(
        g,
        pos,
        nodelist=crit_nodes,
        ax=ax,
        node_size=3200,
        node_color="#ff9999",
        edgecolors="darkred",
        linewidths=2.0,
    )

    # Draw labels
    nx.draw_networkx_labels(
        g,
        pos,
        labels=labels,
        ax=ax,
        font_size=9,
        font_weight="bold",
    )

    ax.set_title("Precedence Network Diagram", fontsize=16, fontweight="bold")
    ax.axis("off")
    plt.tight_layout()

    return fig


ctx = require_workbook()
tasks = load_precedence_tasks(
    ctx["excel_file"],
    sheet_name=ctx["precedence_sheet"]
)
outputs = build_precedence_outputs(tasks)

g = outputs["graph"]
critical_path = outputs["critical_path_task_ids"]

# -----------------------------
# NETWORK DIAGRAM
# -----------------------------
st.subheader("Precedence Network Diagram")
fig = draw_precedence_graph(g, critical_path, tasks)
st.pyplot(fig, use_container_width=True)

# -----------------------------
# CRITICAL PATH
# -----------------------------
st.subheader("Critical Path Approximation")
if critical_path:
    st.write(" → ".join(str(x) for x in critical_path))
else:
    st.write("No critical path found.")

# -----------------------------
# TABLES
# -----------------------------
st.subheader("Technological Precedence")
st.dataframe(outputs["technological_precedence"], use_container_width=True)

st.subheader("Resource-Linked Conflicts")
resource_conflicts_df = outputs["resource_linked_conflicts"]
if resource_conflicts_df.empty:
    st.info("No resource-linked conflicts found.")
else:
    st.dataframe(resource_conflicts_df, use_container_width=True)

st.subheader("Parallelizable Task Pairs")
parallel_df = outputs["parallelizable_tasks"]
if parallel_df.empty:
    st.info("No parallelizable task pairs found.")
else:
    st.dataframe(parallel_df, use_container_width=True)
