from __future__ import annotations

from dataclasses import asdict
from typing import List

import pandas as pd

from .m02_nlp_normalization import NormalizedStep

IE_CLASS_RULES = {
    "wait": "delay",
    "inspect": "inspection",
    "move": "transport_motion",
    "retrieve": "retrieval",
    "handle": "handling",
    "start_machine": "machine_start",
    "process": "machine_run",
    "cut": "transformation",
    "apply": "transformation",
    "assemble": "assembly",
    "serve": "delivery",
    "search": "search",
    "other": "other",
}

WASTE_MAP_BY_ACTIVITY = {
    "W": "waiting",
    "I": "inspection",
    "S": "storage",
    "D": "decision",
    "M": "motion_handling",
    "T": "transport",
    "O": "operation",
}


def map_steps_to_ie_ontology(normalized_steps: List[NormalizedStep]) -> pd.DataFrame:
    rows = []
    for s in normalized_steps:
        row = asdict(s)
        row["ie_class"] = IE_CLASS_RULES.get(s.verb, "other")
        row["waste_class"] = WASTE_MAP_BY_ACTIVITY.get(s.activity.upper(), "other")
        rows.append(row)
    return pd.DataFrame(rows)
