from __future__ import annotations

from collections import defaultdict
from typing import Dict, List, Optional

from .m01_workbook_loader import Task
from .m02_nlp_normalization import NormalizedStep

PLANNER_STYLE_DURATIONS = {
    'Get Knife': 2,
    'Get Butter': 15,
    'Get Bread': 2,
    'Get Plate': 4,
    'Place Bread On Plate': 5,
    'Return Bread Bag': 1,
    'Put Bread In Toaster': 7,
    'Start Toaster': 2,
    'Toast Bread': 133,
    'Remove Toast': 2,
    'Cut And Spread Butter': 46,
    'Stack Toast': 4,
    'Cut Toast': 4,
    'Serve': 4,
}

ORDER = [
    'Get Knife', 'Get Butter', 'Get Bread', 'Get Plate', 'Place Bread On Plate', 'Return Bread Bag',
    'Put Bread In Toaster', 'Start Toaster', 'Toast Bread', 'Remove Toast', 'Cut And Spread Butter',
    'Stack Toast', 'Cut Toast', 'Serve'
]

PREDS = {
    'Place Bread On Plate': ['Get Bread', 'Get Plate'],
    'Return Bread Bag': ['Place Bread On Plate'],
    'Put Bread In Toaster': ['Place Bread On Plate'],
    'Start Toaster': ['Put Bread In Toaster'],
    'Toast Bread': ['Start Toaster'],
    'Get Butter': ['Start Toaster'],
    'Get Knife': ['Start Toaster'],
    'Remove Toast': ['Toast Bread'],
    'Cut And Spread Butter': ['Get Knife', 'Get Butter', 'Remove Toast'],
    'Stack Toast': ['Cut And Spread Butter'],
    'Cut Toast': ['Stack Toast'],
    'Serve': ['Cut Toast'],
}


def _match_task(step: NormalizedStep) -> Optional[str]:
    text = step.raw_description.lower()
    if 'wife' in text or step.action == 'serve':
        return 'Serve'
    if 'stack' in text or 'press' in text or 'flip' in text:
        return 'Stack Toast'
    if 'cuts the toast stack' in text or (step.action == 'cut' and (step.obj == 'toast' or 'toast stack' in text)):
        return 'Cut Toast'
    if any(k in text for k in ['butter', 'spread']) or (step.action in {'cut', 'apply'} and step.obj == 'butter'):
        return 'Cut And Spread Butter'
    if 'toast' in text and 'plate' in text and any(k in text for k in ['place', 'grasp', 'remove']):
        return 'Remove Toast'
    if step.action in {'wait', 'inspect', 'process'} and ('toaster' in text or 'toast' in text):
        return 'Toast Bread'
    if step.action == 'start_machine':
        return 'Start Toaster'
    if 'slot' in text or ('toaster' in text and 'bread' in text):
        return 'Put Bread In Toaster'
    if 'close bag of bread' in text or 'return bag' in text or 'position bag' in text:
        return 'Return Bread Bag'
    if 'plate' in text and 'bread' in text:
        return 'Place Bread On Plate'
    if step.obj == 'plate':
        return 'Get Plate'
    if step.obj == 'bread':
        return 'Get Bread'
    if step.obj == 'knife':
        return 'Get Knife'
    if step.obj == 'butter':
        return 'Get Butter'
    return None


def generate_macro_tasks(normalized_steps: List[NormalizedStep], use_planner_durations: bool = True) -> List[Task]:
    grouped: Dict[str, List[NormalizedStep]] = defaultdict(list)
    for s in normalized_steps:
        task_name = _match_task(s)
        if task_name:
            grouped[task_name].append(s)

    id_map = {name: (i + 1) * 10 for i, name in enumerate(ORDER)}
    tasks: List[Task] = []
    for name in ORDER:
        if name not in grouped and name not in PREDS and name not in PLANNER_STYLE_DURATIONS:
            continue
        group = grouped.get(name, [])
        if not group and name not in {'Get Knife', 'Get Butter', 'Get Bread', 'Get Plate', 'Place Bread On Plate', 'Put Bread In Toaster', 'Start Toaster', 'Toast Bread', 'Remove Toast', 'Cut And Spread Butter', 'Stack Toast', 'Cut Toast', 'Serve'}:
            continue
        resources = sorted({r for s in group for r in s.resources})
        if not resources:
            default_resources = {
                'Get Knife': ['Man'], 'Get Butter': ['Man'], 'Get Bread': ['Man'], 'Get Plate': ['Man'],
                'Place Bread On Plate': ['Man', 'Plate'], 'Return Bread Bag': ['Man'],
                'Put Bread In Toaster': ['Man', 'Toaster'], 'Start Toaster': ['Man', 'Toaster'], 'Toast Bread': ['Toaster'],
                'Remove Toast': ['Man', 'Toaster', 'Plate'], 'Cut And Spread Butter': ['Man', 'Knife', 'Butter', 'Plate'],
                'Stack Toast': ['Man', 'Plate'], 'Cut Toast': ['Man', 'Knife', 'Plate'], 'Serve': ['Man', 'Plate'],
            }
            resources = default_resources.get(name, ['Man'])
        duration = sum(s.duration_sec for s in group)
        if use_planner_durations and name in PLANNER_STYLE_DURATIONS:
            duration = PLANNER_STYLE_DURATIONS[name]
        if duration <= 0:
            duration = PLANNER_STYLE_DURATIONS.get(name, 1)
        tasks.append(Task(
            task_id=id_map[name], name=name, duration_sec=duration, resources=resources,
            internal_external='external' if name in {'Get Knife', 'Get Butter'} else 'internal'
        ))
    name_to_id = {t.name: t.task_id for t in tasks}
    for t in tasks:
        t.predecessors = [name_to_id[p] for p in PREDS.get(t.name, []) if p in name_to_id]
    return tasks
