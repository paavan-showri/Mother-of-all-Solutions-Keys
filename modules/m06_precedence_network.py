from __future__ import annotations

from itertools import combinations
from typing import Dict, List, Tuple

import networkx as nx
import pandas as pd

from .m01_workbook_loader import Task


def build_precedence_outputs(tasks: List[Task]) -> Dict[str, object]:
    g = nx.DiGraph()
    for t in tasks:
        g.add_node(t.task_id, name=t.name, duration_sec=t.duration_sec, resources=t.resources)
    for t in tasks:
        for pred in t.predecessors:
            g.add_edge(pred, t.task_id)
    if not nx.is_directed_acyclic_graph(g):
        raise ValueError("Precedence network is not a DAG.")

    node_duration = {t.task_id: t.duration_sec for t in tasks}
    topo = list(nx.topological_sort(g))
    longest: Dict[int, int] = {n: 0 for n in topo}
    parent: Dict[int, int | None] = {n: None for n in topo}
    for n in topo:
        for succ in g.successors(n):
            cand = longest[n] + node_duration[n]
            if cand > longest[succ]:
                longest[succ] = cand
                parent[succ] = n
    end_node = max(topo, key=lambda n: longest[n] + node_duration[n]) if topo else None
    critical_path = []
    while end_node is not None:
        critical_path.append(end_node)
        end_node = parent[end_node]
    critical_path = list(reversed(critical_path))

    tech_edges = pd.DataFrame([
        {"predecessor": u, "successor": v} for u, v in g.edges()
    ])

    resource_conflicts = []
    parallelizable = []
    for a, b in combinations(tasks, 2):
        shared = sorted(set(a.resources) & set(b.resources))
        comparable = nx.has_path(g, a.task_id, b.task_id) or nx.has_path(g, b.task_id, a.task_id)
        if shared:
            resource_conflicts.append({
                "task_a": a.task_id, "task_b": b.task_id, "shared_resources": ", ".join(shared)
            })
        if not shared and not comparable:
            parallelizable.append({
                "task_a": a.task_id, "task_b": b.task_id
            })

    return {
        "graph": g,
        "technological_precedence": tech_edges,
        "resource_linked_conflicts": pd.DataFrame(resource_conflicts),
        "parallelizable_tasks": pd.DataFrame(parallelizable),
        "critical_path_task_ids": critical_path,
    }
