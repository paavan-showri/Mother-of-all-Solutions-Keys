from __future__ import annotations

from typing import Dict, List

import networkx as nx
import pandas as pd

from .m01_workbook_loader import Task


def build_precedence_outputs(tasks: List[Task]) -> Dict[str, object]:
    g = nx.DiGraph()
    for t in tasks:
        g.add_node(t.task_id, name=t.name, duration_sec=t.duration_sec, resources=t.resources)
    ids = {t.task_id for t in tasks}
    for t in tasks:
        for pred in t.predecessors:
            if pred not in ids:
                raise ValueError(f'Predecessor {pred} is not present in task list.')
            g.add_edge(pred, t.task_id)
    if not nx.is_directed_acyclic_graph(g):
        raise ValueError('Precedence network is not a DAG.')

    topo = list(nx.topological_sort(g))
    longest = {n: 0 for n in topo}
    parent = {n: None for n in topo}
    dur = {t.task_id: t.duration_sec for t in tasks}
    for n in topo:
        for succ in g.successors(n):
            cand = longest[n] + dur[n]
            if cand > longest[succ]:
                longest[succ] = cand
                parent[succ] = n
    end_node = max(topo, key=lambda n: longest[n] + dur[n]) if topo else None
    cp = []
    while end_node is not None:
        cp.append(end_node)
        end_node = parent[end_node]
    cp = list(reversed(cp))

    tech = pd.DataFrame([{'predecessor': u, 'successor': v} for u, v in g.edges()])
    conflicts = []
    parallel = []
    for i, a in enumerate(tasks):
        for b in tasks[i + 1:]:
            shared = sorted(set(a.resources) & set(b.resources))
            linked = nx.has_path(g, a.task_id, b.task_id) or nx.has_path(g, b.task_id, a.task_id)
            if shared:
                conflicts.append({'task_a': a.task_id, 'task_b': b.task_id, 'shared_resources': ', '.join(shared)})
            if not linked and not shared:
                parallel.append({'task_a': a.task_id, 'task_b': b.task_id})
    return {
        'graph': g,
        'technological_precedence': tech,
        'resource_linked_conflicts': pd.DataFrame(conflicts),
        'parallelizable_tasks': pd.DataFrame(parallel),
        'critical_path_task_ids': cp,
    }
