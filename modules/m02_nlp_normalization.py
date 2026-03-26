from __future__ import annotations

import re
from dataclasses import dataclass, asdict
from functools import lru_cache
from typing import Dict, List, Optional

import pandas as pd

from .m01_workbook_loader import FPCStep

try:
    import spacy
except Exception:  # pragma: no cover
    spacy = None


@dataclass
class NormalizedStep:
    step: int
    raw_description: str
    standardized_text: str
    actor: Optional[str]
    action: str
    obj: Optional[str]
    location: Optional[str]
    equipment: List[str]
    materials: List[str]
    tools: List[str]
    resource_entities: List[str]
    machine_related: bool
    setup_related: bool
    delay_signal: bool
    inspection_signal: bool
    motion_signal: bool
    search_signal: bool
    duration_sec: int
    resources: List[str]
    activity_type: str
    activity: str
    va_flag: str
    waste_pred: str = ""
    waste_score: float = 0.0
    manuf_semantic_hint: str = ""


ACTION_PATTERNS = [
    (r"(search|look for|locate|find)", "search"),
    (r"(walk|move|carry|transport|bring|go to|return|turn back|turn around)", "move"),
    (r"(grasp|pick|collect|retrieve|get|take|fetch|reach)", "retrieve"),
    (r"(open|close|position|place|drop|adjust|hold|set)", "handle"),
    (r"(wait|idle)", "wait"),
    (r"(inspect|check|verify|monitor)", "inspect"),
    (r"(start|turn on|activate)", "start_machine"),
    (r"(run|process|toast|heat|machine)", "process"),
    (r"(cut|slice|trim)", "cut"),
    (r"(butter|spread|apply|coat)", "apply"),
    (r"(stack|assemble|press|flip)", "assemble"),
    (r"(serve|deliver|tell|inform)", "serve"),
]

OBJECT_WORDS = {
    'bread': ['bread', 'slice'],
    'plate': ['plate'],
    'toaster': ['toaster'],
    'butter': ['butter'],
    'knife': ['knife'],
    'toast': ['toast'],
    'drawer': ['drawer'],
    'refrigerator': ['refrigerator', 'fridge'],
    'cabinet': ['cabinet'],
}
LOCATION_WORDS = {
    'cabinet': ['cabinet'], 'counter': ['counter'], 'refrigerator': ['refrigerator', 'fridge'],
    'living_room': ['living room'], 'drawer': ['drawer'], 'microwave': ['microwave'],
}
TOOL_WORDS = {'knife'}
MATERIAL_WORDS = {'bread', 'butter', 'toast'}
EQUIPMENT_WORDS = {'toaster', 'microwave', 'refrigerator', 'cabinet', 'drawer'}
ZERO_SHOT_LABELS = [
    'Waiting', 'Motion', 'Transportation', 'Inspection', 'Extra Processing',
    'Defects', 'Inventory', 'Overproduction', 'Non-Utilized Talent', 'Setup', 'Handling'
]


@lru_cache(maxsize=1)
def _get_nlp():
    if spacy is None:
        return None
    try:
        return spacy.load('en_core_web_sm')
    except Exception:
        return spacy.blank('en')


@lru_cache(maxsize=2)
def _get_zero_shot_pipeline(model_name: str = 'facebook/bart-large-mnli'):
    try:
        from transformers import pipeline
        return pipeline('zero-shot-classification', model=model_name)
    except Exception:
        return None


@lru_cache(maxsize=2)
def _get_feature_pipeline(model_name: str = 'rarmingaud/ManufactuBERT'):
    try:
        from transformers import pipeline
        return pipeline('feature-extraction', model=model_name, tokenizer=model_name)
    except Exception:
        return None


def _detect_label(pattern_map: Dict[str, List[str]], text: str) -> Optional[str]:
    t = text.lower()
    for label, variants in pattern_map.items():
        if any(v in t for v in variants):
            return label
    return None


def _detect_action(text: str) -> str:
    t = text.lower()
    for pattern, label in ACTION_PATTERNS:
        if re.search(pattern, t):
            return label
    return 'other'


def _extract_spacy_parts(text: str):
    nlp = _get_nlp()
    if nlp is None:
        return None, None, None
    doc = nlp(text)
    actor = None
    obj = None
    location = None
    for token in doc:
        if token.dep_ in {'nsubj', 'nsubjpass'} and actor is None:
            actor = token.text.lower()
        if token.dep_ in {'dobj', 'pobj', 'attr'} and obj is None:
            obj = token.text.lower()
    for ent in getattr(doc, 'ents', []):
        if ent.label_.lower() in {'gpe', 'loc', 'facility'} and location is None:
            location = ent.text.lower()
    return actor, obj, location


def _rule_waste(text: str, activity: str, va_flag: str) -> str:
    t = text.lower()
    if activity == 'W' or 'wait' in t:
        return 'Waiting'
    if activity == 'I' or any(k in t for k in ['inspect', 'check', 'verify', 'monitor']):
        return 'Inspection'
    if any(k in t for k in ['walk', 'move', 'carry', 'turn back', 'turn around']):
        return 'Motion'
    if any(k in t for k in ['search', 'look for', 'locate', 'find']):
        return 'Motion'
    if any(k in t for k in ['transport', 'bring to', 'move to']) or activity == 'T':
        return 'Transportation'
    if any(k in t for k in ['open', 'close', 'position', 'adjust', 'handle', 'reposition']):
        return 'Handling'
    if activity == 'S':
        return 'Inventory'
    if activity == 'D':
        return 'Extra Processing' if va_flag.upper() == 'NVA' else 'Handling'
    return 'Extra Processing' if va_flag.upper() == 'NVA' else 'Handling'


def normalize_step(step: FPCStep, use_zero_shot: bool = False, use_manufactubert: bool = False) -> NormalizedStep:
    text = step.description.strip()
    actor, obj_spacy, loc_spacy = _extract_spacy_parts(text)
    action = _detect_action(text)
    obj = _detect_label(OBJECT_WORDS, text) or obj_spacy
    location = _detect_label(LOCATION_WORDS, text) or loc_spacy
    resource_entities = [r for r in step.resources]
    equipment = sorted({r for r in step.resources if r.lower() in EQUIPMENT_WORDS or 'toaster' in r.lower()})
    materials = sorted({r for r in step.resources if r.lower() in MATERIAL_WORDS})
    tools = sorted({r for r in step.resources if r.lower() in TOOL_WORDS})
    machine_related = action in {'start_machine', 'process'} or bool(equipment)
    setup_related = any(k in text.lower() for k in ['open', 'close', 'setup', 'prepare', 'position'])
    delay_signal = action == 'wait'
    inspection_signal = action == 'inspect'
    motion_signal = action in {'move', 'retrieve'}
    search_signal = action == 'search'
    standardized_text = ' | '.join([x for x in [actor or ('operator' if 'Man' in step.resources else None), action, obj, location] if x])
    waste_pred = _rule_waste(text, step.activity, step.va_flag)
    waste_score = 0.51
    if use_zero_shot:
        clf = _get_zero_shot_pipeline()
        if clf is not None:
            try:
                out = clf(text, candidate_labels=ZERO_SHOT_LABELS, multi_label=False)
                waste_pred = out['labels'][0]
                waste_score = float(out['scores'][0])
            except Exception:
                pass
    manuf_hint = ''
    if use_manufactubert:
        feat = _get_feature_pipeline()
        if feat is not None:
            try:
                feat(text[:256])
                manuf_hint = 'manufacturing_embedding_ready'
            except Exception:
                manuf_hint = ''
    return NormalizedStep(
        step=step.step, raw_description=text, standardized_text=standardized_text,
        actor=actor or ('operator' if 'Man' in step.resources else None), action=action, obj=obj,
        location=location, equipment=equipment, materials=materials, tools=tools,
        resource_entities=resource_entities, machine_related=machine_related,
        setup_related=setup_related, delay_signal=delay_signal, inspection_signal=inspection_signal,
        motion_signal=motion_signal, search_signal=search_signal, duration_sec=step.duration_sec,
        resources=step.resources, activity_type=step.activity_type, activity=step.activity,
        va_flag=step.va_flag, waste_pred=waste_pred, waste_score=waste_score, manuf_semantic_hint=manuf_hint,
    )


def normalize_steps(steps: List[FPCStep], use_zero_shot: bool = False, use_manufactubert: bool = False) -> List[NormalizedStep]:
    return [normalize_step(s, use_zero_shot=use_zero_shot, use_manufactubert=use_manufactubert) for s in steps]


def normalized_to_df(steps: List[NormalizedStep]) -> pd.DataFrame:
    return pd.DataFrame([asdict(s) for s in steps])
