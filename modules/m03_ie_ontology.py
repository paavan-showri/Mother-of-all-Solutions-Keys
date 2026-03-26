from __future__ import annotations

from dataclasses import asdict
from typing import List

import pandas as pd

from .m02_nlp_normalization import NormalizedStep

ACTION_TO_IE = {
    'search': 'search', 'move': 'transport_motion', 'retrieve': 'retrieval', 'handle': 'handling',
    'wait': 'delay', 'inspect': 'inspection', 'start_machine': 'machine_start', 'process': 'machine_run',
    'cut': 'transformation', 'apply': 'transformation', 'assemble': 'assembly', 'serve': 'delivery', 'other': 'other'
}

WASTE_TO_LEAN = {
    'Waiting': 'NVA', 'Motion': 'NVA', 'Transportation': 'NVA', 'Inspection': 'NNVA',
    'Extra Processing': 'NVA', 'Defects': 'NVA', 'Inventory': 'NVA', 'Overproduction': 'NVA',
    'Non-Utilized Talent': 'NVA', 'Setup': 'NNVA', 'Handling': 'NNVA'
}


def _internal_external(row) -> str:
    if row['machine_related'] and row['action'] in {'start_machine', 'process'}:
        return 'internal'
    if row['obj'] in {'butter', 'knife'} or row['location'] in {'drawer', 'refrigerator'}:
        return 'external'
    if row['action'] in {'retrieve', 'search', 'move'} and not row['machine_related']:
        return 'external'
    return 'internal'


def _stage_group(row) -> str:
    text = row['raw_description'].lower()
    if row['action'] in {'retrieve', 'search'} and any(k in text for k in ['bread', 'plate']):
        return 'material_prep'
    if row['action'] in {'handle'} and 'plate' in text and 'bread' in text:
        return 'material_prep'
    if row['action'] in {'start_machine', 'process'} or 'toaster' in text:
        return 'machine_cycle'
    if any(k in text for k in ['butter', 'knife']):
        return 'butter_prep'
    if 'toast' in text and any(k in text for k in ['stack', 'cut', 'serve', 'wife']):
        return 'finishing'
    return 'other'


def map_steps_to_ie_ontology(normalized_steps: List[NormalizedStep]) -> pd.DataFrame:
    rows = []
    for s in normalized_steps:
        row = asdict(s)
        row['ie_class'] = ACTION_TO_IE.get(s.action, 'other')
        row['lean_bucket'] = WASTE_TO_LEAN.get(s.waste_pred, s.va_flag or '')
        row['internal_external'] = _internal_external(row)
        row['stage_group'] = _stage_group(row)
        row['resource_role'] = 'machine' if s.machine_related else ('tool' if s.tools else ('material' if s.materials else 'labor'))
        row['bottleneck_candidate'] = s.machine_related or s.delay_signal
        rows.append(row)
    return pd.DataFrame(rows)
