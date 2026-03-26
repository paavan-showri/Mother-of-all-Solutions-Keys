from __future__ import annotations

from typing import Dict, List

import networkx as nx
import pandas as pd

from .m01_workbook_loader import Task


def build_precedence_outputs(tasks: List[Task]) -> Dict[str, object]:
    g = nx.DiGraph()
    for task in tasks:
        g.add_node(task.task_id, name=task.name, duration_sec=task.duration_sec, resources=task.resources)

    ids = {task.task_id for task in tasks}
    for task in tasks:
        for pred in task.predecessors:
            if pred not in ids:
                raise ValueError(f"Predecessor {pred} is not present in task list.")
            g.add_edge(pred, task.task_id)

    if not nx.is_directed_acyclic_graph(g):
        raise ValueError("Precedence network is not a DAG.")

    topo = list(nx.topological_sort(g))
    longest = {n: 0 for n in topo}
    parent = {n: None for n in topo}
    durations = {t.task_id: t.duration_sec for t in tasks}

    for node in topo:
        for succ in g.successors(node):
            candidate = longest[node] + durations[node]
            if candidate > longest[succ]:
                longest[succ] = candidate
                parent[succ] = node

    end_node = max(topo, key=lambda n: longest[n] + durations[n]) if topo else None
    critical_path = []
    while end_node is not None:
        critical_path.append(end_node)
        end_node = parent[end_node]
    critical_path = list(reversed(critical_path))

    tech = pd.DataFrame([{"predecessor": u, "successor": v} for u, v in g.edges()])
    conflicts = []
    parallel = []
    for i, task_a in enumerate(tasks):
        for task_b in tasks[i + 1 :]:
            shared = sorted(set(task_a.resources) & set(task_b.resources))
            linked = nx.has_path(g, task_a.task_id, task_b.task_id) or nx.has_path(g, task_b.task_id, task_a.task_id)
            if shared:
                conflicts.append(
                    {"task_a": task_a.task_id, "task_b": task_b.task_id, "shared_resources": ", ".join(shared)}
                )
            if not linked and not shared:
                parallel.append({"task_a": task_a.task_id, "task_b": task_b.task_id})

    return {
        "graph": g,
        "technological_precedence": tech,
        "resource_linked_conflicts": pd.DataFrame(conflicts),
        "parallelizable_tasks": pd.DataFrame(parallel),
        "critical_path_task_ids": critical_path,
    }
