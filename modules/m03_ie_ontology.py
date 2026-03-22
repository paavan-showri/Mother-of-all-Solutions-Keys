from __future__ import annotations

from dataclasses import asdict
from typing import Dict, List

import pandas as pd

from .m02_nlp_normalization import NormalizedStep

ACTIVITY_TO_IE_CLASS = {
    "search": "retrieval_search",
    "wait": "delay",
    "inspect": "inspection",
    "move": "transport_motion",
    "retrieve": "retrieval",
    "handle": "handling",
    "start_machine": "machine_start",
    "process": "machine_run",
    "load_unload": "load_unload",
    "cut": "transformation",
    "apply": "transformation",
    "assemble": "assembly",
    "serve": "delivery",
    "rework": "rework",
    "other": "other",
}

LEAN_WASTE_BY_ACTIVITY = {
    "W": ["waiting"],
    "I": ["extra_processing"],
    "S": ["inventory"],
    "D": ["non_utilized_talent"],
    "M": ["motion"],
    "T": ["transportation"],
    "O": ["value_added_operation"],
}


def _combine_labels(*label_lists: List[str]) -> List[str]:
    seen = []
    for labels in label_lists:
        for label in labels:
            if label and label not in seen:
                seen.append(label)
    return seen


def _derive_resource_role(row: Dict[str, object]) -> str:
    if row.get("machine"):
        return "equipment"
    if row.get("tool"):
        return "tool"
    if row.get("material"):
        return "material"
    if row.get("product"):
        return "product"
    return "personnel_or_other"


def _derive_smed_bucket(row: Dict[str, object]) -> str:
    action = row.get("action")
    machine = row.get("machine")
    if action in {"wait", "inspect", "process", "start_machine", "load_unload"} or machine:
        return "internal"
    if action in {"retrieve", "search", "move", "handle"}:
        return "external"
    return "review"


def _derive_lean_bucket(row: Dict[str, object], waste_labels: List[str]) -> str:
    activity_type = str(row.get("activity_type", "")).upper()
    if activity_type == "VA":
        return "VA"
    if activity_type == "NNVA":
        return "NNVA"
    if activity_type == "NVA":
        return "NVA"
    if "value_added_operation" in waste_labels:
        return "VA"
    if any(x in waste_labels for x in ["waiting", "motion", "transportation", "inventory", "extra_processing", "defects", "overproduction", "non_utilized_talent"]):
        return "NVA"
    return "NNVA"


def map_steps_to_ie_ontology(normalized_steps: List[NormalizedStep]) -> pd.DataFrame:
    rows = []
    for s in normalized_steps:
        row = asdict(s)
        ie_class = ACTIVITY_TO_IE_CLASS.get(s.action, "other")
        activity_waste = LEAN_WASTE_BY_ACTIVITY.get(str(s.activity).upper(), [])
        waste_labels = _combine_labels(activity_waste, s.waste_hints, s.classifier_labels)
        row["ie_class"] = ie_class
        row["waste_labels"] = waste_labels
        row["primary_waste_label"] = waste_labels[0] if waste_labels else "unclassified"
        row["resource_role"] = _derive_resource_role(row)
        row["smed_bucket"] = _derive_smed_bucket(row)
        row["lean_bucket"] = _derive_lean_bucket(row, waste_labels)
        row["manufacturing_stage"] = (
            "machine_run" if ie_class in {"machine_run", "machine_start", "load_unload"}
            else "quality" if ie_class == "inspection"
            else "material_tool_prep" if ie_class in {"retrieval", "retrieval_search", "handling", "transport_motion"}
            else "transformation" if ie_class in {"transformation", "assembly", "delivery"}
            else "other"
        )
        row["machine_or_station"] = row.get("machine") or row.get("location")
        rows.append(row)
    return pd.DataFrame(rows)
