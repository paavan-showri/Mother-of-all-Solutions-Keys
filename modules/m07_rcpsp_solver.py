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

    for task in tasks:
        starts[task.task_id] = model.NewIntVar(0, horizon, f"start_{task.task_id}")
        ends[task.task_id] = model.NewIntVar(0, horizon, f"end_{task.task_id}")
        intervals[task.task_id] = model.NewIntervalVar(
            starts[task.task_id], task.duration_sec, ends[task.task_id], f"interval_{task.task_id}"
        )

    for task in tasks:
        for pred in task.predecessors:
            model.Add(starts[task.task_id] >= ends[pred])

    for resource, cap in capacities.items():
        using = [t for t in tasks if resource in t.resources]
        if not using:
            continue
        ivals = [intervals[t.task_id] for t in using]
        if cap <= 1:
            model.AddNoOverlap(ivals)
        else:
            model.AddCumulative(ivals, [1] * len(ivals), int(cap))

    makespan = model.NewIntVar(0, horizon, "makespan")
    for task in tasks:
        model.Add(makespan >= ends[task.task_id])
    model.Minimize(makespan)

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 20.0
    status = solver.Solve(model)
    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        raise RuntimeError("No feasible RCPSP solution found.")

    rows = []
    for task in sorted(tasks, key=lambda t: t.task_id):
        rows.append({
            "Task ID": task.task_id,
            "Task Name": task.name,
            "Duration (sec)": task.duration_sec,
            "Start (sec)": solver.Value(starts[task.task_id]),
            "End (sec)": solver.Value(ends[task.task_id]),
            "Immediate Predecessors": task.predecessors,
            "Resources": ", ".join(task.resources),
            "Internal/External": task.internal_external,
        })
    df = pd.DataFrame(rows)
    df.attrs["makespan"] = solver.Value(makespan)
    return df
