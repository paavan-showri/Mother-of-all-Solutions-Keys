from __future__ import annotations

from typing import Dict, List

import pandas as pd
from ortools.sat.python import cp_model

from .m01_workbook_loader import Task


def solve_rcpsp(tasks: List[Task], capacities: Dict[str, int]) -> pd.DataFrame:
    model = cp_model.CpModel()
    horizon = max(1, sum(t.duration_sec for t in tasks))
    starts = {}
    ends = {}
    intervals = {}
    for t in tasks:
        starts[t.task_id] = model.NewIntVar(0, horizon, f'start_{t.task_id}')
        ends[t.task_id] = model.NewIntVar(0, horizon, f'end_{t.task_id}')
        intervals[t.task_id] = model.NewIntervalVar(starts[t.task_id], t.duration_sec, ends[t.task_id], f'ival_{t.task_id}')
    for t in tasks:
        for pred in t.predecessors:
            model.Add(starts[t.task_id] >= ends[pred])
    for resource, cap in capacities.items():
        using = [t for t in tasks if resource.title() in [r.title() for r in t.resources]]
        if not using:
            continue
        ivals = [intervals[t.task_id] for t in using]
        if cap <= 1:
            model.AddNoOverlap(ivals)
        else:
            model.AddCumulative(ivals, [1] * len(ivals), cap)
    makespan = model.NewIntVar(0, horizon, 'makespan')
    for t in tasks:
        model.Add(makespan >= ends[t.task_id])
    model.Minimize(makespan)
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 20
    status = solver.Solve(model)
    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        raise RuntimeError('No feasible RCPSP solution found.')
    rows = []
    for t in sorted(tasks, key=lambda x: x.task_id):
        rows.append({
            'Task ID': t.task_id, 'Task Name': t.name, 'Duration (sec)': t.duration_sec,
            'Start (sec)': solver.Value(starts[t.task_id]), 'End (sec)': solver.Value(ends[t.task_id]),
            'Immediate Predecessors': ', '.join(str(p) for p in t.predecessors) if t.predecessors else '—',
            'Resources': ', '.join(t.resources), 'Internal/External': t.internal_external,
        })
    out = pd.DataFrame(rows)
    out.attrs['makespan'] = solver.Value(makespan)
    return out
