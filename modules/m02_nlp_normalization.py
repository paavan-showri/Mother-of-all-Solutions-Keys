from __future__ import annotations

import re
from dataclasses import asdict, dataclass
from typing import Dict, Iterable, List, Optional

import pandas as pd

from .m01_workbook_loader import FPCStep


@dataclass
class NormalizedStep:
    step: int
    raw_description: str
    standardized_text: str
    actor: Optional[str]
    action: str
    obj: Optional[str]
    location: Optional[str]
    equipment: List[str]
    materials: List[str]
    tools: List[str]
    resource_entities: List[str]
    machine_related: bool
    setup_related: bool
    delay_signal: bool
    inspection_signal: bool
    motion_signal: bool
    search_signal: bool
    duration_sec: int
    resources: List[str]
    activity_type: str
    activity: str
    va_flag: str
    waste_pred: str = ""
    waste_score: float = 0.0
    review_flag: bool = False
    review_note: str = ""


ACTION_PATTERNS = [
    (r"\b(search|look for|locate|find)\b", "search"),
    (r"\b(walk|move|carry|transport|bring|go|return|turn around|turn back)\b", "move"),
    (r"\b(grasp|pick|collect|get|take|fetch|reach)\b", "retrieve"),
    (r"\b(open|close|position|place|drop|adjust|hold|set)\b", "handle"),
    (r"\b(wait|idle)\b", "wait"),
    (r"\b(inspect|check|verify|monitor)\b", "inspect"),
    (r"\b(turn on|start|activate)\b", "start_machine"),
    (r"\b(toast|heat|process|run)\b", "process"),
    (r"\b(cut|slice|trim)\b", "cut"),
    (r"\b(butter|spread|apply|coat)\b", "apply"),
    (r"\b(stack|assemble|press|flip)\b", "assemble"),
    (r"\b(serve|deliver)\b", "serve"),
    (r"\b(decide|react|complain|tell|inform|reject)\b", "decision"),
]

OBJECT_WORDS: Dict[str, List[str]] = {
    "bread": ["bag of bread", "bread bag", "bread", "slice"],
    "plate": ["plate"],
    "toaster": ["toaster", "slot"],
    "butter": ["butter"],
    "knife": ["butter knife", "knife"],
    "toast": ["toast", "toasts"],
    "drawer": ["drawer"],
    "refrigerator": ["refrigerator", "fridge"],
    "cabinet": ["cabinet"],
    "counter": ["counter", "counter top", "kitchen counter"],
    "living_room": ["living room"],
}
LOCATION_WORDS: Dict[str, List[str]] = {
    "cabinet": ["cabinet"],
    "counter": ["counter", "counter top", "kitchen counter"],
    "refrigerator": ["refrigerator", "fridge"],
    "living_room": ["living room"],
    "drawer": ["drawer"],
    "microwave": ["microwave"],
    "toaster": ["toaster"],
}
TOOL_WORDS = {"knife"}
MATERIAL_WORDS = {"bread", "butter", "toast", "plate"}
EQUIPMENT_WORDS = {"toaster", "microwave", "refrigerator", "cabinet", "drawer"}


def _contains_any(text: str, phrases: Iterable[str]) -> bool:
    return any(p in text for p in phrases)


def _detect_label(pattern_map: Dict[str, List[str]], text: str) -> Optional[str]:
    t = text.lower()
    ordered = sorted(pattern_map.items(), key=lambda kv: max(len(v) for v in kv[1]), reverse=True)
    for label, variants in ordered:
        if _contains_any(t, [v.lower() for v in variants]):
            return label
    return None


def _detect_action(text: str) -> str:
    t = text.lower()
    for pattern, label in ACTION_PATTERNS:
        if re.search(pattern, t):
            return label
    return "other"


def _extract_actor(text: str, resources: List[str]) -> Optional[str]:
    t = text.lower()
    if "wife" in t:
        return "wife"
    if "man " in f"{t} " or " husband" in f" {t}" or "operator" in t:
        return "operator"
    if any(r.lower() == "man" for r in resources):
        return "operator"
    return None


def _rule_waste(text: str, activity: str, va_flag: str) -> str:
    t = text.lower()
    va = (va_flag or "").upper()
    if activity == "W" or _contains_any(t, ["search", "wait", "idle"]):
        return "Waiting"
    if activity == "I" or _contains_any(t, ["inspect", "check", "verify", "monitor"]):
        return "Inspection"
    if _contains_any(t, ["walk", "turn around", "turn back", "go to", "walk back", "reach"]):
        return "Motion"
    if activity == "M" and _contains_any(t, ["move", "carry", "transport", "bring"]):
        return "Transportation"
    if activity == "S":
        return "Inventory"
    if activity == "D":
        return "Decision"
    if activity == "O" and va == "VA":
        return "Transformation"
    if _contains_any(t, ["open", "close", "place", "position", "drop", "adjust", "grasp", "hold", "set"]):
        return "Handling"
    return "Handling" if va != "NVA" else "Extra Processing"


def _resource_buckets(resources: List[str]) -> tuple[List[str], List[str], List[str]]:
    lowered = {r.lower(): r.title() for r in resources}
    equipment = sorted({title for low, title in lowered.items() if low in EQUIPMENT_WORDS or "toaster" in low})
    tools = sorted({title for low, title in lowered.items() if low in TOOL_WORDS})
    materials = sorted({title for low, title in lowered.items() if low in MATERIAL_WORDS})
    return equipment, materials, tools


def _review_note(text: str, action: str, waste_pred: str) -> tuple[bool, str]:
    notes: List[str] = []
    if action == "decision":
        notes.append("cognitive or communication step")
    if "wife" in text.lower() or "reject" in text.lower():
        notes.append("customer interaction step")
    if "butter knife" in text.lower():
        notes.append("tool phrase detected")
    if waste_pred == "Decision":
        notes.append("decision category")
    return (bool(notes), "; ".join(notes))


def normalize_step(step: FPCStep) -> NormalizedStep:
    text = step.description.strip()
    action = _detect_action(text)
    actor = _extract_actor(text, step.resources)
    obj = _detect_label(OBJECT_WORDS, text)
    location = _detect_label(LOCATION_WORDS, text)
    equipment, materials, tools = _resource_buckets(step.resources)
    if "butter knife" in text.lower():
        tools = sorted(set(tools + ["Knife"]))
        obj = "knife"
    machine_related = bool(equipment) or action in {"start_machine", "process"}
    setup_related = action == "handle" and _contains_any(text.lower(), ["open", "close", "position", "set"])
    delay_signal = action == "wait"
    inspection_signal = action == "inspect"
    motion_signal = action in {"move", "retrieve"}
    search_signal = action == "search"
    standardized_text = " | ".join(
        [x for x in [actor, action, obj, location] if x]
    )
    waste_pred = _rule_waste(text, step.activity, step.va_flag)
    review_flag, review_note = _review_note(text, action, waste_pred)
    return NormalizedStep(
        step=step.step,
        raw_description=text,
        standardized_text=standardized_text,
        actor=actor,
        action=action,
        obj=obj,
        location=location,
        equipment=equipment,
        materials=materials,
        tools=tools,
        resource_entities=list(step.resources),
        machine_related=machine_related,
        setup_related=setup_related,
        delay_signal=delay_signal,
        inspection_signal=inspection_signal,
        motion_signal=motion_signal,
        search_signal=search_signal,
        duration_sec=max(int(step.duration_sec or 0), 0),
        resources=list(step.resources),
        activity_type=step.activity_type,
        activity=step.activity,
        va_flag=step.va_flag,
        waste_pred=waste_pred,
        waste_score=1.0,
        review_flag=review_flag,
        review_note=review_note,
    )


def normalize_steps(steps: List[FPCStep]) -> List[NormalizedStep]:
    return [normalize_step(step) for step in steps if step.description]


def normalized_to_df(normalized_steps: List[NormalizedStep]) -> pd.DataFrame:
    df = pd.DataFrame(asdict(step) for step in normalized_steps)
    if df.empty:
        return df
    order = [
        "step", "raw_description", "standardized_text", "actor", "action", "obj", "location",
        "equipment", "materials", "tools", "resource_entities", "machine_related", "setup_related",
        "delay_signal", "inspection_signal", "motion_signal", "search_signal", "duration_sec",
        "resources", "activity_type", "activity", "va_flag", "waste_pred", "waste_score",
        "review_flag", "review_note",
    ]
    return df[[c for c in order if c in df.columns]]
