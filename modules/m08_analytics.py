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
    if steps:
        current_makespan = max((step.end_sec or 0) for step in steps)
        if current_makespan == 0:
            current_makespan = sum(step.duration_sec for step in steps)
    else:
        current_makespan = 0

    optimized_makespan = int(optimized_schedule.attrs.get("makespan", 0))
    improvement_sec = current_makespan - optimized_makespan
    improvement_pct = round((improvement_sec / current_makespan * 100), 2) if current_makespan else 0.0

    waste_summary = (
        ontology_df.groupby(["waste_pred", "lean_bucket"], dropna=False)["duration_sec"]
        .sum()
        .reset_index()
        .sort_values("duration_sec", ascending=False)
    )

    total = max(optimized_makespan, 1)
    util_rows = []
    for resource, cap in capacities.items():
        busy = sum(task.duration_sec for task in tasks if resource.title() in [r.title() for r in task.resources])
        util_rows.append(
            {
                "resource": resource,
                "capacity": cap,
                "busy_sec": busy,
                "utilization_pct": round(100 * busy / total / max(cap, 1), 2),
            }
        )
    util_df = pd.DataFrame(util_rows).sort_values("utilization_pct", ascending=False)
    bottleneck = util_df.iloc[0].to_dict() if not util_df.empty else {}

    comparison = pd.DataFrame(
        [
            {"metric": "Current makespan (sec)", "value": current_makespan},
            {"metric": "Optimized makespan (sec)", "value": optimized_makespan},
            {"metric": "Improvement (sec)", "value": improvement_sec},
            {"metric": "Improvement (%)", "value": improvement_pct},
        ]
    )
    return {
        "comparison": comparison,
        "bottleneck": bottleneck,
        "resource_utilization": util_df,
        "waste_summary": waste_summary,
    }
