from __future__ import annotations

from typing import List

import pandas as pd


def apply_rule_engine(ontology_df: pd.DataFrame) -> pd.DataFrame:
    df = ontology_df.copy()

    def suggestion(row):
        text = str(row.get("raw_description", "")).lower()
        ie_class = row.get("ie_class", "other")
        waste = row.get("waste_class", "other")

        suggestions = []
        if waste in {"waiting", "decision", "storage"}:
            suggestions.append("eliminate_candidate")
        if ie_class in {"handling", "retrieval", "transport_motion"}:
            suggestions.append("combine_or_simplify_candidate")
        if waste == "search":
            suggestions.append("standardize_candidate")
        if "toaster" in text and ie_class in {"retrieval", "handling"}:
            suggestions.append("rearrange_candidate")
        if any(k in text for k in ["knife", "butter"]) and row.get("ie_class") in {"retrieval", "handling", "transformation"}:
            suggestions.append("split_external_candidate")
        if "inspect" in text:
            suggestions.append("simplify_candidate")
        if "search" in text or "look for" in text:
            suggestions.append("5s_visual_management_candidate")
        if row.get("activity_type", "").upper() == "NVA":
            suggestions.append("high_waste_priority")
        if not suggestions:
            suggestions.append("keep_or_review")
        return ", ".join(dict.fromkeys(suggestions))

    df["ecrssa_suggestions"] = df.apply(suggestion, axis=1)
    df["internal_external_hint"] = df.apply(
        lambda r: "external" if any(k in str(r.get("raw_description", "")).lower() for k in ["knife", "butter", "drawer", "refrigerator"]) else "internal",
        axis=1,
    )
    return df
