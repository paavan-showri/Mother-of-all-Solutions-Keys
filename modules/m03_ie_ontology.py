from __future__ import annotations

from dataclasses import asdict
from typing import List

import pandas as pd

from .m02_nlp_normalization import NormalizedStep

ACTION_TO_IE = {
    "search": "search",
    "move": "transport_motion",
    "retrieve": "retrieval",
    "handle": "handling",
    "wait": "delay",
    "inspect": "inspection",
    "start_machine": "machine_start",
    "process": "machine_run",
    "cut": "transformation",
    "apply": "transformation",
    "assemble": "assembly",
    "serve": "delivery",
    "decision": "decision",
    "other": "other",
}

WASTE_TO_LEAN = {
    "Waiting": "NVA",
    "Motion": "NVA",
    "Transportation": "NVA",
    "Inspection": "NNVA",
    "Extra Processing": "NVA",
    "Inventory": "NVA",
    "Handling": "NNVA",
    "Decision": "NNVA",
    "Transformation": "VA",
}


def _internal_external(row: dict) -> str:
    action = row.get("action")
    obj = row.get("obj")
    machine_related = bool(row.get("machine_related"))
    if action in {"process", "start_machine"}:
        return "internal"
    if obj in {"butter", "knife", "refrigerator", "drawer"}:
        return "external"
    if action in {"search", "move", "retrieve"} and not machine_related:
        return "external"
    return "internal"


def _stage_group(row: dict) -> str:
    text = str(row.get("raw_description", "")).lower()
    action = row.get("action")
    obj = row.get("obj")
    if action == "decision":
        return "control"
    if action in {"search", "move", "retrieve", "handle"} and any(k in text for k in ["bread", "plate", "cabinet"]):
        return "material_prep"
    if action in {"start_machine", "process"} or "toaster" in text:
        return "machine_cycle"
    if obj in {"butter", "knife", "refrigerator", "drawer"} or any(k in text for k in ["butter", "knife"]):
        return "finish_prep"
    if obj == "toast" or any(k in text for k in ["toast", "stack", "press", "flip", "serve", "wife", "living room"]):
        return "finishing"
    return "other"


def map_steps_to_ie_ontology(normalized_steps: List[NormalizedStep]) -> pd.DataFrame:
    rows = []
    for step in normalized_steps:
        row = asdict(step)
        row["ie_class"] = ACTION_TO_IE.get(step.action, "other")
        row["lean_bucket"] = WASTE_TO_LEAN.get(step.waste_pred, step.va_flag or "")
        row["internal_external"] = _internal_external(row)
        row["stage_group"] = _stage_group(row)
        row["resource_role"] = "machine" if step.machine_related else ("tool" if step.tools else ("material" if step.materials else "labor"))
        row["bottleneck_candidate"] = step.machine_related or step.delay_signal
        rows.append(row)
    return pd.DataFrame(rows)
