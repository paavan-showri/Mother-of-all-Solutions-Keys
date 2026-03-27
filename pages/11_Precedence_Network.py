import textwrap

import matplotlib.pyplot as plt
import networkx as nx
import streamlit as st

from session_utils import require_workbook
from modules.m01_workbook_loader import load_current_state_steps
from modules.m02_nlp_normalization import normalize_steps
from modules.m05_macro_tasks import generate_macro_tasks
from modules.m06_precedence_network import build_precedence_outputs

st.set_page_config(page_title='Precedence Network', layout='wide')
st.title('Precedence Network')
ctx = require_workbook()

steps = load_current_state_steps(ctx['excel_file'], sheet_name=ctx['sheet_name'])
normalized = normalize_steps(steps)
tasks = generate_macro_tasks(normalized)
outputs = build_precedence_outputs(tasks)

simplified_view = st.checkbox('Simplified view', value=True)
graph = outputs.get('display_graph', outputs['graph']) if simplified_view else outputs['graph']
critical_path_ids = outputs.get('critical_path_task_ids', [])

zoom = st.slider('Graph zoom', min_value=1.0, max_value=4.0, value=2.2, step=0.1)
node_size = st.slider('Node size', min_value=1200, max_value=12000, value=3600, step=200)
layer_spacing = st.slider('Horizontal spacing', min_value=3.0, max_value=10.0, value=6.5, step=0.5)
row_spacing = st.slider('Vertical spacing', min_value=1.5, max_value=6.0, value=3.2, step=0.1)
label_width = st.slider('Label wrap width', min_value=8, max_value=24, value=14, step=1)
fig_height = st.slider('Chart height', min_value=10, max_value=36, value=22, step=1)
fig_width = st.slider('Chart width', min_value=18, max_value=50, value=34, step=1)


def build_layered_positions(g: nx.DiGraph):
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
        for i, node in enumerate(nodes):
            pos[node] = (lv * layer_spacing, (n_nodes - 1) * row_spacing / 2 - i * row_spacing)
    return pos


pos = build_layered_positions(graph)
labels = {
    n: f"{n}\n" + "\n".join(textwrap.wrap(graph.nodes[n].get('name', str(n)), width=label_width)[:3])
    for n in graph.nodes()
}
cp_nodes = {n for n in critical_path_ids if n in graph.nodes()}
cp_edges = {(a, b) for a, b in zip(critical_path_ids[:-1], critical_path_ids[1:]) if graph.has_edge(a, b)}
normal_edges = [e for e in graph.edges() if e not in cp_edges]
normal_nodes = [n for n in graph.nodes() if n not in cp_nodes]

fig, ax = plt.subplots(figsize=(fig_width * zoom, fig_height * zoom), dpi=160)
nx.draw_networkx_edges(
    graph,
    pos,
    edgelist=normal_edges,
    ax=ax,
    arrows=True,
    edge_color='#7f8c8d',
    width=1.8,
    arrowsize=18,
    connectionstyle='arc3,rad=0.02',
)
nx.draw_networkx_edges(
    graph,
    pos,
    edgelist=list(cp_edges),
    ax=ax,
    arrows=True,
    edge_color='#c0392b',
    width=3.0,
    arrowsize=20,
    connectionstyle='arc3,rad=0.02',
)
nx.draw_networkx_nodes(
    graph,
    pos,
    nodelist=normal_nodes,
    ax=ax,
    node_size=node_size,
    node_color='#d6eaf8',
    edgecolors='#1f4e79',
    linewidths=1.4,
)
nx.draw_networkx_nodes(
    graph,
    pos,
    nodelist=list(cp_nodes),
    ax=ax,
    node_size=int(node_size * 1.15),
    node_color='#f9d6d5',
    edgecolors='#922b21',
    linewidths=1.6,
)
nx.draw_networkx_labels(graph, pos, labels=labels, ax=ax, font_size=10, font_weight='bold')
ax.margins(0.08, 0.15)
ax.axis('off')
plt.tight_layout(pad=2.5)
st.pyplot(fig, use_container_width=False)

with st.expander('Technological precedence table', expanded=False):
    st.dataframe(outputs['technological_precedence'], use_container_width=True)
with st.expander('Resource-linked conflicts', expanded=False):
    st.dataframe(outputs['resource_linked_conflicts'], use_container_width=True)
