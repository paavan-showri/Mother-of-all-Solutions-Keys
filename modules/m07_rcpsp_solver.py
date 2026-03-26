from __future__ import annotations

from typing import Dict, List

import pandas as pd
from ortools.sat.python import cp_model

from .m01_workbook_loader import Task


def solve_rcpsp(tasks: List[Task], capacities: Dict[str, int]) -> pd.DataFrame:
    if not tasks:
        out = pd.DataFrame(
            columns=[
                "Task ID",
                "Task Name",
                "Duration (sec)",
                "Start (sec)",
                "End (sec)",
                "Immediate Predecessors",
                "Resources",
                "Internal/External",
            ]
        )
        out.attrs["makespan"] = 0
        return out

    model = cp_model.CpModel()
    horizon = max(1, sum(max(int(t.duration_sec), 0) for t in tasks))
    starts = {}
    ends = {}
    intervals = {}

    for task in tasks:
        duration = max(int(task.duration_sec), 0)
        starts[task.task_id] = model.NewIntVar(0, horizon, f"start_{task.task_id}")
        ends[task.task_id] = model.NewIntVar(0, horizon, f"end_{task.task_id}")
        intervals[task.task_id] = model.NewIntervalVar(
            starts[task.task_id], duration, ends[task.task_id], f"ival_{task.task_id}"
        )

    for task in tasks:
        for pred in task.predecessors:
            if pred not in ends:
                raise ValueError(f"Predecessor {pred} is missing from the task list.")
            model.Add(starts[task.task_id] >= ends[pred])

    for resource, cap in capacities.items():
        using = [t for t in tasks if resource.title() in [r.title() for r in t.resources]]
        if not using:
            continue
        ivals = [intervals[t.task_id] for t in using]
        if cap <= 1:
            model.AddNoOverlap(ivals)
        else:
            model.AddCumulative(ivals, [1] * len(ivals), max(int(cap), 1))

    makespan = model.NewIntVar(0, horizon, "makespan")
    for task in tasks:
        model.Add(makespan >= ends[task.task_id])
    model.Minimize(makespan)

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 20
    status = solver.Solve(model)
    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        raise RuntimeError("No feasible RCPSP solution found.")

    rows = []
    for task in sorted(tasks, key=lambda x: x.task_id):
        rows.append(
            {
                "Task ID": task.task_id,
                "Task Name": task.name,
                "Duration (sec)": task.duration_sec,
                "Start (sec)": solver.Value(starts[task.task_id]),
                "End (sec)": solver.Value(ends[task.task_id]),
                "Immediate Predecessors": ", ".join(str(p) for p in task.predecessors) if task.predecessors else "—",
                "Resources": ", ".join(task.resources),
                "Internal/External": task.internal_external,
            }
        )
    out = pd.DataFrame(rows)
    out.attrs["makespan"] = solver.Value(makespan)
    return out
