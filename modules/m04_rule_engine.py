from __future__ import annotations

import pandas as pd

ECRS_TEMPLATES = {
    "Waiting": ["eliminate_wait", "shift_to_parallel_window"],
    "Motion": ["reduce_walking", "point_of_use"],
    "Transportation": ["shorten_distance", "combine_moves"],
    "Inspection": ["absorb_or_reduce_checks"],
    "Extra Processing": ["eliminate_or_merge"],
    "Handling": ["combine_or_simplify"],
    "Decision": ["standardize_or_remove"],
    "Inventory": ["reduce_holding"],
    "Transformation": ["keep_value_add"],
}

FUTURE_REDUCTION = {
    "Waiting": 0.20,
    "Motion": 0.45,
    "Transportation": 0.50,
    "Inspection": 0.50,
    "Extra Processing": 0.40,
    "Handling": 0.75,
    "Decision": 0.10,
    "Inventory": 0.60,
    "Transformation": 0.95,
}


def _future_state_action(row) -> str:
    stage = row.get("stage_group", "other")
    waste = row.get("waste_pred", "")
    text = str(row.get("raw_description", "")).lower()
    if stage == "control":
        return "remove_or_absorb"
    if waste == "Waiting":
        return "eliminate_or_parallelize"
    if stage == "finish_prep" and row.get("internal_external") == "external":
        return "parallelize_during_machine_run"
    if stage == "machine_cycle" and "inspect" in text:
        return "absorb_into_machine_cycle"
    if waste in {"Motion", "Transportation", "Handling"}:
        return "combine_and_simplify"
    if waste == "Transformation":
        return "keep_value_add"
    return "review"


def _fivew1h(row) -> str:
    return (
        f"What={row.get('action')} | Why={row.get('waste_pred')} | Where={row.get('location') or 'workstation'} | "
        f"When={row.get('stage_group')} | Who={row.get('actor') or 'operator'} | How={row.get('resource_role')}"
    )


def apply_rule_engine(ontology_df: pd.DataFrame) -> pd.DataFrame:
    df = ontology_df.copy()
    df["ecrs_suggestions"] = df["waste_pred"].map(lambda x: ", ".join(ECRS_TEMPLATES.get(x, ["review"])))
    df["future_state_action"] = df.apply(_future_state_action, axis=1)
    df["fivew1h_summary"] = df.apply(_fivew1h, axis=1)
    df["reduction_factor"] = df["waste_pred"].map(lambda x: FUTURE_REDUCTION.get(x, 0.8)).astype(float)
    df["future_duration_sec"] = (
        (df["duration_sec"].fillna(0).astype(float) * df["reduction_factor"]).round().clip(lower=0)
    ).astype(int)
    df["toc_note"] = df.apply(
        lambda r: "exploit_machine_cycle"
        if bool(r.get("bottleneck_candidate"))
        else ("prepare_off_bottleneck" if r.get("internal_external") == "external" else "review"),
        axis=1,
    )
    return df
