from __future__ import annotations

from typing import Dict, List

import pandas as pd

from .m01_workbook_loader import FPCStep, Task


def build_analytics(
    steps: List[FPCStep],
    ontology_df: pd.DataFrame,
    tasks: List[Task],
    capacities: Dict[str, int],
    optimized_schedule: pd.DataFrame,
) -> Dict[str, object]:
    current_makespan = max((s.end_sec or 0) for s in steps) if steps else 0
    optimized_makespan = int(optimized_schedule.attrs.get("makespan", 0))
    improvement_sec = current_makespan - optimized_makespan
    improvement_pct = (improvement_sec / current_makespan * 100.0) if current_makespan else 0.0

    exploded = ontology_df.copy()
    if "waste_labels" in exploded.columns:
        exploded = exploded.explode("waste_labels")
    waste_summary = (
        exploded.groupby(["lean_bucket", "ie_class", "waste_labels"], dropna=False)["duration_sec"]
        .sum()
        .reset_index()
        .sort_values("duration_sec", ascending=False)
    )

    total_makespan = max(optimized_makespan, 1)
    util_rows = []
    for resource, cap in capacities.items():
        busy = sum(t.duration_sec for t in tasks if resource in t.resources)
        util_rows.append({
            "resource": resource,
            "capacity": cap,
            "busy_sec": busy,
            "utilization_pct": round(100.0 * busy / total_makespan / max(cap, 1), 2),
        })
    resource_util = pd.DataFrame(util_rows).sort_values("utilization_pct", ascending=False)
    bottleneck = resource_util.iloc[0].to_dict() if not resource_util.empty else {}

    comparison = pd.DataFrame([
        {"metric": "Current makespan (sec)", "value": current_makespan},
        {"metric": "Optimized makespan (sec)", "value": optimized_makespan},
        {"metric": "Improvement (sec)", "value": improvement_sec},
        {"metric": "Improvement (%)", "value": round(improvement_pct, 2)},
    ])

    return {
        "waste_summary": waste_summary,
        "resource_utilization": resource_util,
        "bottleneck": bottleneck,
        "comparison": comparison,
    }
