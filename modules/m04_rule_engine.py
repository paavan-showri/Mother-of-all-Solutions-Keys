from __future__ import annotations

from typing import List

import pandas as pd


def _to_list(value) -> List[str]:
    if isinstance(value, list):
        return value
    if pd.isna(value) or value in (None, ""):
        return []
    return [str(value)]


def apply_rule_engine(ontology_df: pd.DataFrame) -> pd.DataFrame:
    df = ontology_df.copy()

    ecrs_actions = []
    internal_external = []
    macro_task_hints = []
    precedence_hints = []

    for _, row in df.iterrows():
        wastes = _to_list(row.get("waste_labels"))
        action = str(row.get("action", ""))
        text = str(row.get("raw_description", "")).lower()
        material = str(row.get("material", "") or "")
        product = str(row.get("product", "") or "")
        machine = str(row.get("machine", "") or "")
        tool = str(row.get("tool", "") or "")

        suggestions: List[str] = []
        if any(w in wastes for w in ["waiting", "inventory"]):
            suggestions.append("E: eliminate_or_reduce_delay")
        if any(w in wastes for w in ["motion", "transportation", "extra_processing"]):
            suggestions.append("C/S: combine_and_simplify")
        if "search" in action or "search" in text:
            suggestions.append("E/5S: visual_management_and_point_of_use")
        if action in {"retrieve", "move", "handle"}:
            suggestions.append("R: resequence_for_point_of_use")
        if action in {"retrieve", "search", "move"} and any(x in text for x in ["knife", "butter", "tool", "material"]):
            suggestions.append("SPLIT: externalize_from_machine_cycle")
        if action == "inspect":
            suggestions.append("S: standardize_quality_check")
        if action == "rework" or "defects" in wastes:
            suggestions.append("E: defect_prevention_priority")
        if not suggestions:
            suggestions.append("KEEP: review_for_standard_work")

        ext_int = "internal"
        if action in {"retrieve", "search", "move", "handle"} and not machine:
            ext_int = "external"
        elif action in {"process", "start_machine", "load_unload", "inspect", "wait"} or machine:
            ext_int = "internal"

        if material == "bread" and action in {"retrieve", "move", "search", "handle"}:
            macro = "Get bread" if "plate" not in text else "Place bread on plate"
        elif product == "plate" and action in {"retrieve", "move", "search", "handle"}:
            macro = "Get plate"
        elif "slot" in text or (machine == "toaster" and material == "bread"):
            macro = "Put bread in toaster"
        elif action == "start_machine":
            macro = "Start toaster"
        elif action in {"wait", "inspect", "process"} and machine == "toaster":
            macro = "Toast bread"
        elif material == "butter" and action in {"retrieve", "move", "search", "handle"}:
            macro = "Get butter"
        elif tool == "knife" and action in {"retrieve", "move", "search", "handle"}:
            macro = "Get knife"
        elif "toast" in text and "plate" in text and action in {"retrieve", "move", "handle", "load_unload"}:
            macro = "Remove toast"
        elif material == "butter" and action == "cut":
            macro = "Cut butter"
        elif action == "apply":
            macro = "Butter toast"
        elif action == "assemble":
            macro = "Stack toast"
        elif action == "cut" and product == "toast":
            macro = "Cut toast"
        elif action == "serve":
            macro = "Serve"
        else:
            macro = "Review"

        preds: List[str] = []
        if macro == "Place bread on plate":
            preds = ["Get bread", "Get plate"]
        elif macro == "Put bread in toaster":
            preds = ["Place bread on plate"]
        elif macro == "Start toaster":
            preds = ["Put bread in toaster"]
        elif macro == "Toast bread":
            preds = ["Start toaster"]
        elif macro in {"Get butter", "Get knife"}:
            preds = ["Start toaster"]
        elif macro == "Remove toast":
            preds = ["Toast bread"]
        elif macro == "Cut butter":
            preds = ["Get butter", "Get knife"]
        elif macro == "Butter toast":
            preds = ["Remove toast", "Cut butter"]
        elif macro == "Stack toast":
            preds = ["Butter toast"]
        elif macro == "Cut toast":
            preds = ["Stack toast"]
        elif macro == "Serve":
            preds = ["Cut toast"]

        ecrs_actions.append(suggestions)
        internal_external.append(ext_int)
        macro_task_hints.append(macro)
        precedence_hints.append(preds)

    df["ecrs_actions"] = ecrs_actions
    df["internal_external_hint"] = internal_external
    df["macro_task_hint"] = macro_task_hints
    df["precedence_hint_names"] = precedence_hints
    return df
