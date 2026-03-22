from __future__ import annotations

import re
from dataclasses import dataclass
from typing import List, Optional

from .m01_workbook_loader import FPCStep


@dataclass
class NormalizedStep:
    step: int
    raw_description: str
    actor: Optional[str]
    verb: str
    obj: Optional[str]
    location: Optional[str]
    equipment: List[str]
    standardized_text: str
    duration_sec: int
    resources: List[str]
    activity_type: str
    activity: str


VERB_PATTERNS = [
    (r"\b(search|look for|locate|find)\b", "search"),
    (r"\b(walk|move|carry|transport|bring|go to|return|turn back|turn around)\b", "move"),
    (r"\b(grasp|pick|collect|retrieve|get|take)\b", "retrieve"),
    (r"\b(open|close|position|place|drop|adjust|hold)\b", "handle"),
    (r"\b(wait)\b", "wait"),
    (r"\b(inspect|check|verify)\b", "inspect"),
    (r"\b(start|turn on|activate)\b", "start_machine"),
    (r"\b(cut|slice)\b", "cut"),
    (r"\b(butter|spread)\b", "apply"),
    (r"\b(stack|assemble|press|flip)\b", "assemble"),
    (r"\b(serve|deliver|tell)\b", "serve"),
    (r"\b(toast)\b", "process"),
]

OBJECT_PATTERNS = {
    "bread": ["bread", "slice"],
    "plate": ["plate"],
    "toaster": ["toaster"],
    "butter": ["butter"],
    "knife": ["knife"],
    "toast": ["toast"],
}

LOCATION_PATTERNS = {
    "cabinet": ["cabinet"],
    "counter": ["counter"],
    "refrigerator": ["refrigerator", "fridge"],
    "living_room": ["living room"],
}

SYNONYM_MAP = {
    "collect": "retrieve",
    "pick": "retrieve",
    "fetch": "retrieve",
    "bring": "move",
    "carry": "move",
    "verify": "inspect",
}


def _detect(patterns, text: str) -> Optional[str]:
    t = text.lower()
    for label, variants in patterns.items():
        if any(v in t for v in variants):
            return label
    return None


def detect_verb(text: str) -> str:
    t = text.lower()
    for pattern, label in VERB_PATTERNS:
        if re.search(pattern, t):
            return SYNONYM_MAP.get(label, label)
    return "other"


def normalize_step(step: FPCStep) -> NormalizedStep:
    verb = detect_verb(step.description)
    obj = _detect(OBJECT_PATTERNS, step.description)
    location = _detect(LOCATION_PATTERNS, step.description)
    actor = "operator" if "Man" in step.resources else None
    equipment = [r for r in step.resources if r in {"Toaster", "Knife"}]
    standardized_text = " | ".join([x for x in [actor, verb, obj, location] if x])
    return NormalizedStep(
        step=step.step,
        raw_description=step.description,
        actor=actor,
        verb=verb,
        obj=obj,
        location=location,
        equipment=equipment,
        standardized_text=standardized_text,
        duration_sec=step.duration_sec,
        resources=step.resources,
        activity_type=step.activity_type,
        activity=step.activity,
    )


def normalize_steps(steps: List[FPCStep]) -> List[NormalizedStep]:
    return [normalize_step(s) for s in steps]
