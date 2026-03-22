from __future__ import annotations

from typing import Dict, List, Optional

from .m01_workbook_loader import Task
from .m02_nlp_normalization import NormalizedStep

ORDERED_TASK_NAMES = [
    "Get bread", "Get plate", "Place bread on plate", "Put bread in toaster",
    "Start toaster", "Toast bread", "Get butter", "Get knife", "Remove toast",
    "Cut butter", "Butter toast", "Stack toast", "Cut toast", "Serve",
]
DEFAULT_PRECEDENCE_BY_NAME: Dict[str, List[str]] = {
    "Place bread on plate": ["Get bread", "Get plate"],
    "Put bread in toaster": ["Place bread on plate"],
    "Start toaster": ["Put bread in toaster"],
    "Toast bread": ["Start toaster"],
    "Get butter": ["Start toaster"],
    "Get knife": ["Start toaster"],
    "Remove toast": ["Toast bread"],
    "Cut butter": ["Get butter", "Get knife"],
    "Butter toast": ["Remove toast", "Cut butter"],
    "Stack toast": ["Butter toast"],
    "Cut toast": ["Stack toast"],
    "Serve": ["Cut toast"],
}


def match_task_name(step: NormalizedStep) -> Optional[str]:
    text = step.raw_description.lower()
    if "tell wife" in text or step.verb == "serve":
        return "Serve"
    if "cuts the toast stack" in text or (step.verb == "cut" and step.obj == "toast"):
        return "Cut toast"
    if "stack" in text or step.verb == "assemble":
        return "Stack toast"
    if step.verb == "apply":
        return "Butter toast"
    if step.verb == "cut" and step.obj == "butter":
        return "Cut butter"
    if step.verb in {"wait", "inspect", "process"} and (step.obj in {"bread", "toast", "toaster"} or "Toaster" in step.resources):
        return "Toast bread"
    if step.verb == "start_machine":
        return "Start toaster"
    if step.obj == "knife" and step.verb in {"retrieve", "move", "handle"}:
        return "Get knife"
    if step.obj == "butter" and step.verb in {"retrieve", "move", "search", "handle"}:
        return "Get butter"
    if "toast" in text and "plate" in text and step.verb in {"retrieve", "handle"}:
        return "Remove toast"
    if "slot" in text or ("toaster" in text and step.obj == "bread"):
        return "Put bread in toaster"
    if "plate" in text and step.obj == "bread":
        return "Place bread on plate"
    if step.obj == "plate" and step.verb in {"retrieve", "move", "search", "handle"}:
        return "Get plate"
    if step.obj == "bread" and step.verb in {"retrieve", "move", "search", "handle"}:
        return "Get bread"
    return None


def generate_macro_tasks(normalized_steps: List[NormalizedStep]) -> List[Task]:
    grouped: Dict[str, List[NormalizedStep]] = {}
    for step in normalized_steps:
        task_name = match_task_name(step)
        if task_name:
            grouped.setdefault(task_name, []).append(step)

    id_map = {name: (i + 1) * 10 for i, name in enumerate(ORDERED_TASK_NAMES)}
    tasks: List[Task] = []
    for name in ORDERED_TASK_NAMES:
        group = grouped.get(name, [])
        if not group:
            continue
        resources = sorted({r for s in group for r in s.resources})
        tasks.append(Task(
            task_id=id_map[name],
            name=name,
            duration_sec=sum(s.duration_sec for s in group),
            resources=resources,
            internal_external="external" if name in {"Get butter", "Get knife", "Cut butter"} else "internal",
        ))
    name_to_id = {t.name: t.task_id for t in tasks}
    for task in tasks:
        task.predecessors = [name_to_id[p] for p in DEFAULT_PRECEDENCE_BY_NAME.get(task.name, []) if p in name_to_id]
    return tasks
