from __future__ import annotations

from typing import Dict, List

import pandas as pd

from .m01_workbook_loader import FPCStep, Task


def build_analytics(steps: List[FPCStep], ontology_df: pd.DataFrame, tasks: List[Task], capacities: Dict[str, int], optimized_schedule: pd.DataFrame) -> Dict[str, object]:
    current_makespan = max((s.end_sec or 0) for s in steps) if steps else 0
    optimized_makespan = int(optimized_schedule.attrs.get('makespan', 0))
    improvement_sec = current_makespan - optimized_makespan
    improvement_pct = round((improvement_sec / current_makespan * 100), 2) if current_makespan else 0.0

    waste_summary = ontology_df.groupby(['waste_pred', 'lean_bucket'], dropna=False)['duration_sec'].sum().reset_index().sort_values('duration_sec', ascending=False)

    util_rows = []
    total = max(optimized_makespan, 1)
    for resource, cap in capacities.items():
        busy = sum(t.duration_sec for t in tasks if resource.title() in [r.title() for r in t.resources])
        util_rows.append({'resource': resource, 'capacity': cap, 'busy_sec': busy, 'utilization_pct': round(100 * busy / total / cap, 2)})
    util_df = pd.DataFrame(util_rows).sort_values('utilization_pct', ascending=False)
    bottleneck = util_df.iloc[0].to_dict() if not util_df.empty else {}

    comparison = pd.DataFrame([
        {'metric': 'Current makespan (sec)', 'value': current_makespan},
        {'metric': 'Optimized makespan (sec)', 'value': optimized_makespan},
        {'metric': 'Improvement (sec)', 'value': improvement_sec},
        {'metric': 'Improvement (%)', 'value': improvement_pct},
    ])
    return {'comparison': comparison, 'bottleneck': bottleneck, 'resource_utilization': util_df, 'waste_summary': waste_summary}
