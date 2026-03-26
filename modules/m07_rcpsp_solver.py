from __future__ import annotations

from typing import Dict, List

import pandas as pd

try:
    from ortools.sat.python import cp_model
except Exception:  # pragma: no cover
    cp_model = None

from .m01_workbook_loader import Task


def _task_row(task: Task, start: int, end: int) -> dict:
    return {
        "Task ID": task.task_id,
        "Task Name": task.name,
        "Duration (sec)": task.duration_sec,
        "Start (sec)": start,
        "End (sec)": end,
        "Immediate Predecessors": ", ".join(str(p) for p in task.predecessors) if task.predecessors else "—",
        "Resources": ", ".join(task.resources),
        "Internal/External": task.internal_external,
        "Stage": task.stage_group,
        "Family": task.action_family,
        "Source Steps": ", ".join(str(s) for s in task.source_steps),
    }


def _solve_greedy(tasks: List[Task], capacities: Dict[str, int]) -> pd.DataFrame:
    resource_next_free = {resource.title(): [0] * max(cap, 1) for resource, cap in capacities.items()}
    task_map = {task.task_id: task for task in tasks}
    finish_times: Dict[int, int] = {}
    rows = []
    for task in sorted(tasks, key=lambda x: x.task_id):
        start = 0
        if task.predecessors:
            start = max(finish_times[pred] for pred in task.predecessors if pred in finish_times)
        start = int(start)
        duration = max(1, int(task.duration_sec))
        for resource in task.resources:
            pools = resource_next_free.setdefault(resource.title(), [0])
            chosen_idx = min(range(len(pools)), key=lambda i: pools[i])
            start = max(start, pools[chosen_idx])
        end = start + duration
        for resource in task.resources:
            pools = resource_next_free.setdefault(resource.title(), [0])
            chosen_idx = min(range(len(pools)), key=lambda i: pools[i])
            pools[chosen_idx] = end
        finish_times[task.task_id] = end
        rows.append(_task_row(task, start, end))
    out = pd.DataFrame(rows)
    out.attrs["makespan"] = max(finish_times.values(), default=0)
    return out


def solve_rcpsp(tasks: List[Task], capacities: Dict[str, int]) -> pd.DataFrame:
    if not tasks:
        out = pd.DataFrame(columns=["Task ID", "Task Name", "Duration (sec)", "Start (sec)", "End (sec)"])
        out.attrs["makespan"] = 0
        return out

    if cp_model is None:
        return _solve_greedy(tasks, capacities)

    model = cp_model.CpModel()
    horizon = max(1, sum(max(1, t.duration_sec) for t in tasks))
    starts = {}
    ends = {}
    intervals = {}
    for task in tasks:
        duration = max(1, int(task.duration_sec))
        starts[task.task_id] = model.NewIntVar(0, horizon, f"start_{task.task_id}")
        ends[task.task_id] = model.NewIntVar(0, horizon, f"end_{task.task_id}")
        intervals[task.task_id] = model.NewIntervalVar(starts[task.task_id], duration, ends[task.task_id], f"ival_{task.task_id}")

    for task in tasks:
        for pred in task.predecessors:
            model.Add(starts[task.task_id] >= ends[pred])

    for resource, cap in capacities.items():
        using = [task for task in tasks if resource.title() in [r.title() for r in task.resources]]
        if not using:
            continue
        ivals = [intervals[task.task_id] for task in using]
        if cap <= 1:
            model.AddNoOverlap(ivals)
        else:
            model.AddCumulative(ivals, [1] * len(ivals), cap)

    makespan = model.NewIntVar(0, horizon, "makespan")
    for task in tasks:
        model.Add(makespan >= ends[task.task_id])
    model.Minimize(makespan)

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 20
    status = solver.Solve(model)
    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        return _solve_greedy(tasks, capacities)

    rows = []
    for task in sorted(tasks, key=lambda x: x.task_id):
        rows.append(_task_row(task, solver.Value(starts[task.task_id]), solver.Value(ends[task.task_id])))
    out = pd.DataFrame(rows)
    out.attrs["makespan"] = solver.Value(makespan)
    return out
