from __future__ import annotations

from collections import defaultdict
from typing import Dict, List, Optional

from .m01_workbook_loader import Task
from .m02_nlp_normalization import NormalizedStep

PLANNER_STYLE_DURATIONS = {
    "Get Knife": 2,
    "Get Butter": 15,
    "Get Bread": 2,
    "Get Plate": 4,
    "Place Bread On Plate": 5,
    "Return Bread Bag": 1,
    "Put Bread In Toaster": 7,
    "Start Toaster": 2,
    "Toast Bread": 133,
    "Remove Toast": 2,
    "Cut And Spread Butter": 46,
    "Stack Toast": 4,
    "Cut Toast": 4,
    "Serve": 4,
}

ORDER = [
    "Get Bread",
    "Get Plate",
    "Place Bread On Plate",
    "Return Bread Bag",
    "Put Bread In Toaster",
    "Start Toaster",
    "Toast Bread",
    "Get Butter",
    "Get Knife",
    "Remove Toast",
    "Cut And Spread Butter",
    "Stack Toast",
    "Cut Toast",
    "Serve",
]

PREDS = {
    "Place Bread On Plate": ["Get Bread", "Get Plate"],
    "Return Bread Bag": ["Place Bread On Plate"],
    "Put Bread In Toaster": ["Place Bread On Plate"],
    "Start Toaster": ["Put Bread In Toaster"],
    "Toast Bread": ["Start Toaster"],
    "Get Butter": ["Start Toaster"],
    "Get Knife": ["Start Toaster"],
    "Remove Toast": ["Toast Bread"],
    "Cut And Spread Butter": ["Get Knife", "Get Butter", "Remove Toast"],
    "Stack Toast": ["Cut And Spread Butter"],
    "Cut Toast": ["Stack Toast"],
    "Serve": ["Cut Toast"],
}

DEFAULT_RESOURCES = {
    "Get Bread": ["Man", "Bread"],
    "Get Plate": ["Man", "Plate"],
    "Place Bread On Plate": ["Man", "Bread", "Plate"],
    "Return Bread Bag": ["Man", "Bread"],
    "Put Bread In Toaster": ["Man", "Bread", "Toaster"],
    "Start Toaster": ["Man", "Toaster"],
    "Toast Bread": ["Toaster"],
    "Get Butter": ["Man", "Butter"],
    "Get Knife": ["Man", "Knife"],
    "Remove Toast": ["Man", "Toaster", "Plate"],
    "Cut And Spread Butter": ["Man", "Knife", "Butter", "Plate"],
    "Stack Toast": ["Man", "Plate"],
    "Cut Toast": ["Man", "Knife", "Plate"],
    "Serve": ["Man", "Plate"],
}

TASK_INTERNAL_EXTERNAL = {
    "Get Bread": "internal",
    "Get Plate": "internal",
    "Place Bread On Plate": "internal",
    "Return Bread Bag": "internal",
    "Put Bread In Toaster": "internal",
    "Start Toaster": "internal",
    "Toast Bread": "internal",
    "Get Butter": "external",
    "Get Knife": "external",
    "Remove Toast": "internal",
    "Cut And Spread Butter": "internal",
    "Stack Toast": "internal",
    "Cut Toast": "internal",
    "Serve": "internal",
}


def _contains(text: str, *parts: str) -> bool:
    text = text.lower()
    return all(p.lower() in text for p in parts)


def _step_bucket(step: NormalizedStep) -> Optional[str]:
    text = step.raw_description.lower()
    obj = (step.obj or "").lower()
    action = step.action

    if "wife" in text or _contains(text, "serve") or action == "communicate":
        return "Serve"
    if "stack" in text or "press" in text or "flip" in text:
        return "Stack Toast"
    if _contains(text, "cut", "toast") or ("toast stack" in text and action == "cut"):
        return "Cut Toast"
    if obj in {"butter", "knife"} and action in {"cut", "apply", "handle", "retrieve", "move"} and "toast" in text:
        return "Cut And Spread Butter"
    if action == "apply" and ("toast" in text or obj == "butter"):
        return "Cut And Spread Butter"
    if "butter" in text and "toast" in text:
        return "Cut And Spread Butter"
    if "knife" in text and action in {"cut", "apply", "handle"}:
        return "Cut And Spread Butter"
    if "remove" in text and "toast" in text:
        return "Remove Toast"
    if "toast" in text and "plate" in text and any(k in text for k in ["place", "drop", "grasp", "remove", "put"]):
        return "Remove Toast"
    if action in {"wait", "inspect", "process"} and ("toaster" in text or "toast" in text):
        return "Toast Bread"
    if action == "start_machine":
        return "Start Toaster"
    if ("toaster" in text and "bread" in text) or "slot" in text:
        return "Put Bread In Toaster"
    if ("return" in text and "bread" in text) or ("close" in text and "bread" in text) or ("bag" in text and "counter" in text and action == "handle"):
        return "Return Bread Bag"
    if "plate" in text and "bread" in text:
        return "Place Bread On Plate"
    if obj == "plate":
        return "Get Plate"
    if obj == "bread":
        return "Get Bread"
    if obj == "knife":
        return "Get Knife"
    if obj == "butter":
        return "Get Butter"
    return None


def _reduction_factor(step: NormalizedStep, task_name: str) -> float:
    action = step.action
    waste = (step.waste_pred or "").lower()
    text = step.raw_description.lower()

    if task_name == "Toast Bread":
        return 1.0 if action == "process" else 0.25
    if task_name in {"Start Toaster", "Remove Toast"}:
        return 1.0
    if action == "search":
        return 0.10
    if action == "wait":
        return 0.0
    if action == "inspect":
        return 0.20
    if action == "move":
        return 0.45 if "with" in text or "to" in text else 0.55
    if action == "retrieve":
        return 0.65
    if action == "handle":
        return 0.70 if step.setup_related else 0.75
    if action == "communicate":
        return 0.35
    if action == "cognitive":
        return 0.20
    if action in {"cut", "apply", "assemble"}:
        return 0.90
    if waste in {"handling", "setup"}:
        return 0.70
    if waste in {"motion", "transportation"}:
        return 0.45
    return 0.85


def _discover_grouped_steps(normalized_steps: List[NormalizedStep]) -> Dict[str, List[NormalizedStep]]:
    grouped: Dict[str, List[NormalizedStep]] = defaultdict(list)
    for step in normalized_steps:
        bucket = _step_bucket(step)
        if bucket:
            grouped[bucket].append(step)
    return grouped


def _automatic_duration(task_name: str, group: List[NormalizedStep]) -> int:
    if not group:
        return max(1, PLANNER_STYLE_DURATIONS.get(task_name, 1))

    if task_name == "Toast Bread":
        process_time = sum(s.duration_sec for s in group if s.action == "process")
        wait_or_check = sum(s.duration_sec for s in group if s.action in {"wait", "inspect"})
        duration = process_time + round(wait_or_check * 0.15)
        return max(duration, 1)

    reduced = 0.0
    for step in group:
        reduced += step.duration_sec * _reduction_factor(step, task_name)

    if task_name in {"Get Butter", "Get Knife"}:
        reduced *= 0.90
    if task_name == "Cut And Spread Butter":
        reduced = max(reduced, sum(s.duration_sec for s in group) * 0.75)

    duration = int(round(reduced))
    return max(duration, 1)


def _collect_resources(group: List[NormalizedStep], task_name: str) -> List[str]:
    # Keep task resources tight and planner-like. Pulling every raw row resource into a macro task
    # creates noisy conflicts such as Toaster appearing in Get Butter.
    defaults = DEFAULT_RESOURCES.get(task_name, ["Man"])
    return list(defaults)


def generate_macro_tasks(normalized_steps: List[NormalizedStep], use_planner_durations: bool = False) -> List[Task]:
    grouped = _discover_grouped_steps(normalized_steps)
    id_map = {name: (i + 1) * 10 for i, name in enumerate(ORDER)}

    tasks: List[Task] = []
    for name in ORDER:
        if name not in grouped and name not in PREDS:
            continue

        group = grouped.get(name, [])
        duration = PLANNER_STYLE_DURATIONS[name] if use_planner_durations else _automatic_duration(name, group)
        resources = _collect_resources(group, name)

        tasks.append(
            Task(
                task_id=id_map[name],
                name=name,
                duration_sec=duration,
                predecessors=[],
                resources=resources,
                internal_external=TASK_INTERNAL_EXTERNAL.get(name, "internal"),
            )
        )

    name_to_id = {t.name: t.task_id for t in tasks}
    for task in tasks:
        task.predecessors = [name_to_id[p] for p in PREDS.get(task.name, []) if p in name_to_id]

    return tasks
