from __future__ import annotations

from collections import Counter
from dataclasses import asdict
from typing import Dict, List, Tuple

import pandas as pd

from .m01_workbook_loader import Task
from .m02_nlp_normalization import NormalizedStep
from .m03_ie_ontology import map_steps_to_ie_ontology
from .m04_rule_engine import apply_rule_engine

ACTION_NAME = {
    "search": "Find",
    "move": "Move",
    "retrieve": "Get",
    "handle": "Set Up",
    "wait": "Wait",
    "inspect": "Check",
    "start_machine": "Start",
    "process": "Run",
    "cut": "Cut",
    "apply": "Apply",
    "assemble": "Assemble",
    "serve": "Serve",
    "decision": "Decide",
    "other": "Do",
}


def _stage_rank(stage: str) -> int:
    return {
        "control": 0,
        "material_prep": 1,
        "machine_cycle": 2,
        "finish_prep": 3,
        "finishing": 4,
        "other": 5,
    }.get(stage, 5)


def _family(row: pd.Series) -> str:
    stage = row["stage_group"]
    action = row["action"]
    if stage == "machine_cycle":
        if action == "process":
            return "machine_run"
        if action == "start_machine":
            return "machine_start"
        return "machine_setup"
    if stage == "finish_prep":
        return "prep_finish"
    if stage == "material_prep":
        return "prep_material"
    if stage == "finishing":
        return "finish"
    if stage == "control":
        return "control"
    return "other"


def _window_group(row: pd.Series, seen_machine_start: bool, seen_machine_run: bool) -> str:
    if row["stage_group"] == "control":
        return "pre"
    if row["stage_group"] == "machine_cycle" and row["action"] != "process":
        return "pre"
    if row["stage_group"] == "machine_cycle" and row["action"] == "process":
        return "machine"
    if seen_machine_start and not seen_machine_run:
        return "during"
    if seen_machine_run:
        return "post"
    return "pre"


def _can_merge(prev: Dict[str, object], curr: Dict[str, object]) -> bool:
    if prev["window_group"] != curr["window_group"]:
        return False
    if prev["family"] == "machine_run" or curr["family"] == "machine_run":
        return False
    if prev["family"] == curr["family"] and (prev["obj"] == curr["obj"] or prev["stage_group"] == curr["stage_group"]):
        return True
    if prev["stage_group"] == curr["stage_group"] and set(prev["resources"]) == set(curr["resources"]):
        return True
    return False


def _name_task(rows: List[Dict[str, object]]) -> str:
    actions = Counter(r["action"] for r in rows if r["action"] != "decision")
    action = actions.most_common(1)[0][0] if actions else rows[0]["action"]
    objs = Counter(r["obj"] for r in rows if r["obj"])
    top_objs = [obj for obj, _ in objs.most_common(2)]
    if action == "machine_run":
        action = "process"
    verb = ACTION_NAME.get(action, "Do")
    if any(r["family"] == "machine_run" for r in rows):
        top_objs = top_objs or ["toaster"]
        if "toast" not in top_objs and "bread" in top_objs:
            top_objs = ["bread"]
        return f"Run {' and '.join(o.title() for o in top_objs)}"
    if not top_objs:
        stage = rows[0]["stage_group"]
        return {
            "material_prep": "Prepare Materials",
            "finish_prep": "Prepare Finish Materials",
            "finishing": "Finish Toast",
            "control": "Control Decision",
        }.get(stage, verb)
    return f"{verb} {' and '.join(o.title() for o in top_objs)}"


def _task_resources(rows: List[Dict[str, object]]) -> List[str]:
    resources = []
    for row in rows:
        resources.extend(row["resources"])
    unique = []
    for resource in resources:
        if resource not in unique:
            unique.append(resource)
    return unique or ["Man"]


def _task_duration(rows: List[Dict[str, object]]) -> int:
    total = 0
    for row in rows:
        future = int(row["future_duration_sec"])
        original = int(row["duration_sec"])
        if row["future_state_action"] == "remove_or_absorb":
            continue
        if row["family"] == "machine_run":
            total += max(original, future, 1)
        else:
            total += max(future, 1)
    return max(total, 1)


def _build_rows(normalized_steps: List[NormalizedStep]) -> List[Dict[str, object]]:
    ontology = map_steps_to_ie_ontology(normalized_steps)
    ruled = apply_rule_engine(ontology)
    ruled = ruled.sort_values("step").reset_index(drop=True)
    seen_machine_start = False
    seen_machine_run = False
    rows: List[Dict[str, object]] = []
    for _, row in ruled.iterrows():
        row_dict = row.to_dict()
        row_dict["family"] = _family(row)
        row_dict["window_group"] = _window_group(row, seen_machine_start, seen_machine_run)
        rows.append(row_dict)
        if row["action"] == "start_machine":
            seen_machine_start = True
        if row["action"] == "process":
            seen_machine_run = True
    return rows


def generate_macro_tasks(normalized_steps: List[NormalizedStep]) -> List[Task]:
    rows = _build_rows(normalized_steps)
    if not rows:
        return []

    groups: List[List[Dict[str, object]]] = []
    current: List[Dict[str, object]] = []
    for row in rows:
        if row["future_state_action"] == "remove_or_absorb" and row["stage_group"] == "control":
            continue
        if not current:
            current = [row]
            continue
        if _can_merge(current[-1], row):
            current.append(row)
        else:
            groups.append(current)
            current = [row]
    if current:
        groups.append(current)

    tasks: List[Task] = []
    for idx, group in enumerate(groups, start=1):
        stage_group = group[0]["stage_group"]
        family = group[0]["family"]
        internal_external = "external" if all(g["internal_external"] == "external" for g in group) else "internal"
        task = Task(
            task_id=idx * 10,
            name=_name_task(group),
            duration_sec=_task_duration(group),
            resources=_task_resources(group),
            internal_external=internal_external,
            source_steps=[int(g["step"]) for g in group],
            stage_group=stage_group,
            action_family=family,
            order_key=min(int(g["step"]) for g in group),
        )
        tasks.append(task)

    machine_start_task = next((t for t in tasks if t.action_family == "machine_start"), None)
    machine_run_task = next((t for t in tasks if t.action_family == "machine_run"), None)
    during_tasks = [t for t in tasks if t.internal_external == "external" and t.order_key > (machine_start_task.order_key if machine_start_task else -1)]
    post_tasks = [t for t in tasks if t.order_key > (machine_run_task.order_key if machine_run_task else 10**9)]

    prev_non_during: Task | None = None
    for task in tasks:
        if task is machine_run_task:
            if prev_non_during is not None:
                task.predecessors.append(prev_non_during.task_id)
            prev_non_during = task
            continue
        if task in during_tasks and machine_start_task is not None:
            task.predecessors.append(machine_start_task.task_id)
            continue
        if task.action_family == "machine_start":
            if prev_non_during is not None:
                task.predecessors.append(prev_non_during.task_id)
            prev_non_during = task
            continue
        if task in post_tasks and machine_run_task is not None and machine_run_task.task_id not in task.predecessors:
            task.predecessors.append(machine_run_task.task_id)
        if task in post_tasks:
            needed_preps = [t.task_id for t in during_tasks if t.order_key < task.order_key]
            task.predecessors.extend(needed_preps)
        if task not in during_tasks:
            if prev_non_during is not None and prev_non_during.task_id not in task.predecessors:
                task.predecessors.append(prev_non_during.task_id)
            prev_non_during = task

    for task in tasks:
        task.predecessors = sorted(set(p for p in task.predecessors if p != task.task_id))
    return tasks


def tasks_to_df(tasks: List[Task]) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "Task ID": t.task_id,
                "Task Name": t.name,
                "Duration (sec)": t.duration_sec,
                "Immediate Predecessors": ", ".join(str(p) for p in t.predecessors) if t.predecessors else "—",
                "Resources": ", ".join(t.resources) if t.resources else "—",
                "Type": t.internal_external,
                "Stage": t.stage_group,
                "Family": t.action_family,
                "Source Steps": ", ".join(str(s) for s in t.source_steps),
            }
            for t in tasks
        ]
    )
