import textwrap

import matplotlib.pyplot as plt
import networkx as nx
import streamlit as st

from session_utils import require_workbook
from modules.m01_workbook_loader import load_precedence_tasks
from modules.m06_precedence_network import build_precedence_outputs

st.set_page_config(page_title="Precedence Network", layout="wide")
st.title("11 Precedence Network")


def build_layered_positions(g: nx.DiGraph):
    topo = list(nx.topological_sort(g))
    level = {}
    for node in topo:
        preds = list(g.predecessors(node))
        level[node] = 0 if not preds else max(level[p] for p in preds) + 1

    grouped = {}
    for node, lv in level.items():
        grouped.setdefault(lv, []).append(node)

    pos = {}
    for lv, nodes in grouped.items():
        nodes_sorted = sorted(nodes)
        n = len(nodes_sorted)
        for i, node in enumerate(nodes_sorted):
            pos[node] = (lv * 4.4, (n - 1) * 2.4 / 2 - i * 2.4)
    return pos


def make_labels(g: nx.DiGraph):
    labels = {}
    for node, attrs in g.nodes(data=True):
        task_name = str(attrs.get("name", ""))
        wrapped = "\n".join(textwrap.wrap(task_name, width=16)[:3])
        labels[node] = f"{node}\n{wrapped}" if wrapped else str(node)
    return labels


def draw_precedence_network(g: nx.DiGraph, critical_path_task_ids):
    pos = build_layered_positions(g)
    labels = make_labels(g)

    critical_nodes = set(critical_path_task_ids)
    critical_edges = set(zip(critical_path_task_ids[:-1], critical_path_task_ids[1:]))
    normal_nodes = [n for n in g.nodes() if n not in critical_nodes]
    normal_edges = [e for e in g.edges() if e not in critical_edges]
    cp_nodes = [n for n in g.nodes() if n in critical_nodes]
    cp_edges = [e for e in g.edges() if e in critical_edges]

    fig_w = max(18, len({x for x, _ in pos.values()}) * 3.3)
    fig_h = max(10, len(g.nodes()) * 0.95)
    fig, ax = plt.subplots(figsize=(fig_w, fig_h))

    nx.draw_networkx_edges(
        g,
        pos,
        edgelist=normal_edges,
        ax=ax,
        arrows=True,
        arrowstyle="-|>",
        arrowsize=26,
        width=2.2,
        edge_color="#7f8c8d",
        min_source_margin=28,
        min_target_margin=28,
    )

    nx.draw_networkx_edges(
        g,
        pos,
        edgelist=cp_edges,
        ax=ax,
        arrows=True,
        arrowstyle="-|>",
        arrowsize=30,
        width=3.8,
        edge_color="#c0392b",
        min_source_margin=30,
        min_target_margin=30,
    )

    nx.draw_networkx_nodes(
        g,
        pos,
        nodelist=normal_nodes,
        ax=ax,
        node_shape="o",
        node_size=9500,
        node_color="#d6eaf8",
        edgecolors="#1f4e79",
        linewidths=2.2,
    )

    nx.draw_networkx_nodes(
        g,
        pos,
        nodelist=cp_nodes,
        ax=ax,
        node_shape="o",
        node_size=11000,
        node_color="#f9d6d5",
        edgecolors="#922b21",
        linewidths=2.8,
    )

    nx.draw_networkx_labels(
        g,
        pos,
        labels=labels,
        ax=ax,
        font_size=10,
        font_weight="bold",
        verticalalignment="center",
        horizontalalignment="center",
    )

    ax.set_title("Precedence Graph", fontsize=18, fontweight="bold", pad=16)
    ax.axis("off")
    plt.tight_layout()
    return fig


ctx = require_workbook()

try:
    tasks = load_precedence_tasks(ctx["excel_file"], sheet_name=ctx["precedence_sheet"])
    outputs = build_precedence_outputs(tasks)
except Exception as e:  # noqa: BLE001
    st.error(f"Could not build precedence outputs from worksheet '{ctx['precedence_sheet']}': {e}")
    st.stop()

st.subheader("Precedence graph")
st.caption("Red bubbles and arrows show the critical path approximation.")
fig = draw_precedence_network(outputs["graph"], outputs["critical_path_task_ids"])
st.pyplot(fig, use_container_width=True)

st.subheader("Critical path approximation")
if outputs["critical_path_task_ids"]:
    st.write(" → ".join(str(x) for x in outputs["critical_path_task_ids"]))
else:
    st.write("No critical path found.")

st.subheader("Technological precedence")
st.dataframe(outputs["technological_precedence"], use_container_width=True)

st.subheader("Resource-linked conflicts")
resource_df = outputs["resource_linked_conflicts"]
if resource_df.empty:
    st.info("No resource-linked conflicts found.")
else:
    st.dataframe(resource_df, use_container_width=True)

st.subheader("Parallelizable task pairs")
parallel_df = outputs["parallelizable_tasks"]
if parallel_df.empty:
    st.info("No parallelizable task pairs found.")
else:
    st.dataframe(parallel_df, use_container_width=True)
