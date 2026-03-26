from __future__ import annotations

from typing import List

import pandas as pd


ECRS_TEMPLATES = {
    'Waiting': ['eliminate_wait', 'subordinate_to_bottleneck'],
    'Motion': ['rearrange_point_of_use', 'combine_handling'],
    'Transportation': ['reduce_distance', 'combine_moves'],
    'Inspection': ['simplify_or_absorb_check'],
    'Extra Processing': ['eliminate_or_merge'],
    'Handling': ['combine_or_simplify'],
    'Setup': ['externalize_if_possible'],
}


def _future_state_action(row) -> str:
    stage = row.get('stage_group', 'other')
    text = str(row.get('raw_description', '')).lower()
    waste = row.get('waste_pred', '')
    if 'inspect toast' in text or ('monitor' in text and 'toaster' in text):
        return 'absorb_into_machine_run'
    if waste == 'Waiting':
        return 'eliminate_or_shift_parallel'
    if stage == 'butter_prep' and row.get('internal_external') == 'external':
        return 'shift_into_machine_run_window'
    if waste in {'Motion', 'Transportation', 'Handling'}:
        return 'combine_with_adjacent_task'
    return 'keep'


def _fivew1h(row) -> str:
    return f"What={row.get('action')} | Why={row.get('waste_pred')} | Where={row.get('location') or 'workstation'} | When={row.get('stage_group')} | Who={row.get('actor') or 'operator'} | How={row.get('resource_role')}"


def apply_rule_engine(ontology_df: pd.DataFrame) -> pd.DataFrame:
    df = ontology_df.copy()
    df['ecrs_suggestions'] = df['waste_pred'].map(lambda x: ', '.join(ECRS_TEMPLATES.get(x, ['review'])))
    df['future_state_action'] = df.apply(_future_state_action, axis=1)
    df['fivew1h_summary'] = df.apply(_fivew1h, axis=1)
    df['toc_note'] = df.apply(
        lambda r: 'exploit_machine_cycle' if bool(r.get('bottleneck_candidate')) else ('prepare_off_bottleneck' if r.get('internal_external') == 'external' else 'review'),
        axis=1,
    )
    df['macro_task_hint'] = df['stage_group'].map({
        'material_prep': 'prep_materials',
        'machine_cycle': 'run_machine',
        'butter_prep': 'prepare_finish_materials',
        'finishing': 'finish_and_serve',
    }).fillna('review')
    return df
