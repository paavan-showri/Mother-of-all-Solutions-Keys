from __future__ import annotations

from typing import Dict, List

import networkx as nx
import pandas as pd

from .m01_workbook_loader import Task


def build_precedence_outputs(tasks: List[Task]) -> Dict[str, object]:
    task_ids = {t.task_id for t in tasks}
    missing_preds = sorted({pred for t in tasks for pred in t.predecessors if pred not in task_ids})
    if missing_preds:
        raise ValueError(f"These predecessor IDs do not exist in Task ID: {missing_preds}.")

    g = nx.DiGraph()
    for t in tasks:
        g.add_node(t.task_id, name=t.name, duration_sec=t.duration_sec, resources=t.resources)
    for t in tasks:
        for pred in t.predecessors:
            g.add_edge(pred, t.task_id)
    if not nx.is_directed_acyclic_graph(g):
        raise ValueError("Precedence network is not a DAG.")

    topo = list(nx.topological_sort(g))
    duration = {t.task_id: t.duration_sec for t in tasks}
    longest = {n: 0 for n in topo}
    parent = {n: None for n in topo}
    for n in topo:
        for succ in g.successors(n):
            cand = longest[n] + duration[n]
            if cand > longest[succ]:
                longest[succ] = cand
                parent[succ] = n
    end_node = max(topo, key=lambda n: longest[n] + duration[n]) if topo else None
    critical_path = []
    while end_node is not None:
        critical_path.append(end_node)
        end_node = parent[end_node]
    critical_path.reverse()

    tech_df = pd.DataFrame([{"predecessor": u, "successor": v} for u, v in g.edges()])
    task_lookup = {t.task_id: t for t in tasks}

    parallel_rows = []
    conflict_rows = []
    for i, a in enumerate(tasks):
        for b in tasks[i + 1:]:
            shared = sorted(set(a.resources) & set(b.resources))
            comparable = nx.has_path(g, a.task_id, b.task_id) or nx.has_path(g, b.task_id, a.task_id)
            if shared:
                conflict_rows.append({
                    "task_a": a.task_id,
                    "task_a_name": a.name,
                    "task_b": b.task_id,
                    "task_b_name": b.name,
                    "shared_resources": ", ".join(shared),
                })
            if not shared and not comparable:
                parallel_rows.append({
                    "task_a": a.task_id,
                    "task_a_name": a.name,
                    "task_b": b.task_id,
                    "task_b_name": b.name,
                })

    cp_rows = [
        {"Task ID": tid, "Task Name": task_lookup[tid].name, "Duration (sec)": task_lookup[tid].duration_sec}
        for tid in critical_path
    ]

    return {
        "graph": g,
        "technological_precedence": tech_df,
        "resource_linked_conflicts": pd.DataFrame(conflict_rows),
        "parallelizable_tasks": pd.DataFrame(parallel_rows),
        "critical_path_task_ids": critical_path,
        "critical_path_table": pd.DataFrame(cp_rows),
    }
