from __future__ import annotations

from typing import Dict, List

import pandas as pd

from .m01_workbook_loader import FPCStep, Task


def _current_makespan(steps: List[FPCStep]) -> int:
    if not steps:
        return 0
    end_based = max((s.end_sec or 0) for s in steps)
    if end_based > 0:
        return int(end_based)
    return int(sum(max(int(s.duration_sec or 0), 0) for s in steps))


def build_analytics(
    steps: List[FPCStep],
    ontology_df: pd.DataFrame,
    tasks: List[Task],
    capacities: Dict[str, int],
    optimized_schedule: pd.DataFrame,
) -> Dict[str, object]:
    current_makespan = _current_makespan(steps)
    optimized_makespan = int(optimized_schedule.attrs.get("makespan", 0))
    improvement_sec = current_makespan - optimized_makespan
    improvement_pct = round((improvement_sec / current_makespan * 100), 2) if current_makespan else 0.0

    if ontology_df.empty:
        waste_summary = pd.DataFrame(columns=["waste_pred", "lean_bucket", "duration_sec"])
    else:
        waste_summary = (
            ontology_df.groupby(["waste_pred", "lean_bucket"], dropna=False)["duration_sec"]
            .sum()
            .reset_index()
            .sort_values("duration_sec", ascending=False)
        )

    util_rows = []
    total = max(optimized_makespan, 1)
    for resource, cap in capacities.items():
        normalized_resource = resource.title()
        busy = sum(t.duration_sec for t in tasks if normalized_resource in [r.title() for r in t.resources])
        util_rows.append(
            {
                "resource": normalized_resource,
                "capacity": cap,
                "busy_sec": busy,
                "utilization_pct": round(100 * busy / total / max(cap, 1), 2),
            }
        )
    util_df = pd.DataFrame(util_rows).sort_values("utilization_pct", ascending=False) if util_rows else pd.DataFrame()
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
