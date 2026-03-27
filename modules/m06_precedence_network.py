from __future__ import annotations

from typing import Dict, List

import networkx as nx
import pandas as pd

from .m01_workbook_loader import Task


def _safe_transitive_reduction(graph: nx.DiGraph) -> nx.DiGraph:
    """Return a transitive reduction for display while preserving node attributes."""
    try:
        reduced = nx.transitive_reduction(graph)
        out = nx.DiGraph()
        out.add_nodes_from(graph.nodes(data=True))
        out.add_edges_from(reduced.edges())
        return out
    except Exception:
        return graph.copy()


def build_precedence_outputs(tasks: List[Task]) -> Dict[str, object]:
    graph = nx.DiGraph()
    for task in tasks:
        graph.add_node(
            task.task_id,
            name=task.name,
            duration_sec=task.duration_sec,
            resources=task.resources,
            stage_group=getattr(task, 'stage_group', 'other'),
            action_family=getattr(task, 'action_family', 'other'),
        )

    ids = {task.task_id for task in tasks}
    for task in tasks:
        for pred in task.predecessors:
            if pred not in ids:
                raise ValueError(f"Predecessor {pred} is not present in task list.")
            if pred != task.task_id:
                graph.add_edge(pred, task.task_id)

    if not nx.is_directed_acyclic_graph(graph):
        raise ValueError("Precedence network is not a DAG.")

    topo = list(nx.topological_sort(graph))
    longest = {n: 0 for n in topo}
    parent = {n: None for n in topo}
    durations = {task.task_id: max(int(task.duration_sec), 1) for task in tasks}

    for node in topo:
        for succ in graph.successors(node):
            candidate = longest[node] + durations[node]
            if candidate > longest[succ]:
                longest[succ] = candidate
                parent[succ] = node

    end_node = max(topo, key=lambda n: longest[n] + durations[n]) if topo else None
    critical_path: List[int] = []
    while end_node is not None:
        critical_path.append(end_node)
        end_node = parent[end_node]
    critical_path = list(reversed(critical_path))

    reduced_graph = _safe_transitive_reduction(graph)
    reduced_edges = set(reduced_graph.edges())
    cp_edge_pairs = [
        (u, v) for u, v in zip(critical_path[:-1], critical_path[1:]) if (u, v) in reduced_edges
    ]

    tech = pd.DataFrame([{"predecessor": u, "successor": v} for u, v in graph.edges()]).sort_values(["predecessor", "successor"]).reset_index(drop=True)
    reduced_tech = pd.DataFrame([{"predecessor": u, "successor": v} for u, v in reduced_graph.edges()]).sort_values(["predecessor", "successor"]).reset_index(drop=True)

    conflicts = []
    parallel = []
    for i, task_a in enumerate(tasks):
        for task_b in tasks[i + 1 :]:
            shared = sorted(set(task_a.resources) & set(task_b.resources))
            linked = nx.has_path(graph, task_a.task_id, task_b.task_id) or nx.has_path(graph, task_b.task_id, task_a.task_id)
            if shared:
                conflicts.append({"task_a": task_a.task_id, "task_b": task_b.task_id, "shared_resources": ", ".join(shared)})
            if not linked and not shared:
                parallel.append({"task_a": task_a.task_id, "task_b": task_b.task_id})

    return {
        "graph": graph,
        "display_graph": reduced_graph,
        "technological_precedence": tech,
        "display_precedence": reduced_tech,
        "resource_linked_conflicts": pd.DataFrame(conflicts),
        "parallelizable_tasks": pd.DataFrame(parallel),
        "critical_path_task_ids": critical_path,
        "critical_path_display_edges": cp_edge_pairs,
    }
