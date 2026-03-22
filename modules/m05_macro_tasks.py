from __future__ import annotations

from typing import Dict, List

import pandas as pd

from .m01_workbook_loader import Task

ORDERED_TASK_NAMES = [
    "Get bread",
    "Get plate",
    "Place bread on plate",
    "Put bread in toaster",
    "Start toaster",
    "Toast bread",
    "Get butter",
    "Get knife",
    "Remove toast",
    "Cut butter",
    "Butter toast",
    "Stack toast",
    "Cut toast",
    "Serve",
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


def generate_macro_tasks(ruled_df: pd.DataFrame) -> List[Task]:
    grouped: Dict[str, pd.DataFrame] = {
        name: ruled_df[ruled_df["macro_task_hint"] == name].copy()
        for name in ORDERED_TASK_NAMES
    }

    tasks: List[Task] = []
    id_map = {name: (idx + 1) * 10 for idx, name in enumerate(ORDERED_TASK_NAMES)}

    for name in ORDERED_TASK_NAMES:
        df = grouped.get(name)
        if df is None or df.empty:
            continue
        resources = sorted({r for resources in df["resources"] for r in resources})
        internal_external = "external" if name in {"Get butter", "Get knife", "Cut butter"} else "internal"
        tasks.append(
            Task(
                task_id=id_map[name],
                name=name,
                duration_sec=int(df["duration_sec"].sum()),
                resources=resources,
                internal_external=internal_external,
            )
        )

    name_to_id = {t.name: t.task_id for t in tasks}
    for task in tasks:
        task.predecessors = [name_to_id[p] for p in DEFAULT_PRECEDENCE_BY_NAME.get(task.name, []) if p in name_to_id]
    return tasks
