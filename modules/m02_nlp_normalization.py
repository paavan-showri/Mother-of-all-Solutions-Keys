from __future__ import annotations

import re
from dataclasses import dataclass, asdict
from functools import lru_cache
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

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
    review_flag: bool = False
    review_note: str = ""


ACTION_PATTERNS: List[Tuple[str, str]] = [
    (r"\b(search|look\s+for|locate|find)\b", "search"),
    (r"\b(wait|idle)\b", "wait"),
    (r"\b(inspect|check|verify|monitor)\b", "inspect"),
    (r"\b(decide|react|complain|reject)\b", "cognitive"),
    (r"\b(tell|inform|serve|deliver)\b", "communicate"),
    (r"\b(turn\s+on|start|activate)\b", "start_machine"),
    (r"\b(run|process|toast|heat)\b", "process"),
    (r"\b(cut|slice|trim)\b", "cut"),
    (r"\b(butter|spread|apply|coat)\b", "apply"),
    (r"\b(stack|assemble|press|flip)\b", "assemble"),
    (r"\b(walk|move|carry|transport|bring|go\s+to|return|turn\s+back|turn\s+around|reach)\b", "move"),
    (r"\b(grasp|pick|collect|retrieve|get|take|fetch|hold)\b", "retrieve"),
    (r"\b(open|close|position|place|drop|adjust|set)\b", "handle"),
]
]

PHRASE_OBJECTS: List[Tuple[str, str]] = [
    ("butter knife", "knife"),
    ("fruit bowl", "fruit_bowl"),
    ("bag of bread", "bread"),
    ("slice of butter", "butter"),
    ("slices of toast", "toast"),
    ("slice of toast", "toast"),
    ("slices of bread", "bread"),
    ("slice of bread", "bread"),
]

OBJECT_WORDS: Dict[str, List[str]] = {
    "knife": ["knife"],
    "butter": ["butter"],
    "toast": ["toast"],
    "bread": ["bread", "slice"],
    "plate": ["plate"],
    "toaster": ["toaster"],
    "drawer": ["drawer"],
    "refrigerator": ["refrigerator", "fridge"],
    "cabinet": ["cabinet"],
    "microwave": ["microwave"],
    "counter": ["counter", "counter top", "countertop"],
    "wife": ["wife"],
}
LOCATION_WORDS: Dict[str, List[str]] = {
    "cabinet": ["cabinet"],
    "counter": ["counter", "counter top", "countertop", "kitchen counter"],
    "refrigerator": ["refrigerator", "fridge"],
    "living_room": ["living room"],
    "drawer": ["drawer"],
    "microwave": ["microwave"],
    "toaster": ["toaster"],
}
TOOL_WORDS = {"knife"}
MATERIAL_WORDS = {"bread", "butter", "toast"}
EQUIPMENT_WORDS = {"toaster", "microwave", "refrigerator", "cabinet", "drawer"}
PERSON_WORDS = {"man", "wife", "operator"}
ZERO_SHOT_LABELS = [
    "Waiting",
    "Motion",
    "Transportation",
    "Inspection",
    "Extra Processing",
    "Defects",
    "Inventory",
    "Overproduction",
    "Non-Utilized Talent",
    "Setup",
    "Handling",
    "Value Added",
]


def _clean_text(text: str) -> str:
    text = (text or "").replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _slug(label: str) -> str:
    return label.lower().replace(" ", "_")


@lru_cache(maxsize=1)
def _get_nlp():
    if spacy is None:
        return None
    try:
        return spacy.load("en_core_web_sm")
    except Exception:
        try:
            return spacy.blank("en")
        except Exception:
            return None


@lru_cache(maxsize=2)
def _get_zero_shot_pipeline(model_name: str = "facebook/bart-large-mnli"):
    try:
        from transformers import pipeline

        return pipeline("zero-shot-classification", model=model_name)
    except Exception:
        return None


@lru_cache(maxsize=2)
def _get_feature_pipeline(model_name: str = "rarmingaud/ManufactuBERT"):
    try:
        from transformers import pipeline

        return pipeline("feature-extraction", model=model_name, tokenizer=model_name)
    except Exception:
        return None


def _contains_any(text: str, phrases: Sequence[str]) -> bool:
    return any(re.search(rf"\b{re.escape(p)}\b", text) for p in phrases)


def _resource_list(step: FPCStep) -> List[str]:
    values: List[str] = []
    seen = set()
    for raw in list(getattr(step, "resources", []) or []):
        item = _clean_text(str(raw))
        if not item:
            continue
        key = item.lower()
        if key not in seen:
            seen.add(key)
            values.append(item)
    return values


def _detect_action(text: str) -> str:
    for pattern, label in ACTION_PATTERNS:
        if re.search(pattern, text):
            return label
    return "other"


def _detect_label(pattern_map: Dict[str, List[str]], text: str) -> Optional[str]:
    for label, variants in pattern_map.items():
        if _contains_any(text, variants):
            return label
    return None


def _extract_spacy_parts(text: str) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    nlp = _get_nlp()
    if nlp is None:
        return None, None, None
    doc = nlp(text)
    actor = None
    obj = None
    location = None
    for token in doc:
        dep = getattr(token, "dep_", "")
        if dep in {"nsubj", "nsubjpass"} and actor is None:
            actor = token.text.lower()
        if dep in {"dobj", "pobj", "attr", "obj"} and obj is None:
            obj = token.text.lower()
    for ent in getattr(doc, "ents", []):
        label = getattr(ent, "label_", "").lower()
        if label in {"gpe", "loc", "facility"} and location is None:
            location = ent.text.lower()
    return actor, obj, location


def _infer_actor(text: str, resources: Sequence[str], actor_spacy: Optional[str]) -> Optional[str]:
    lowered_resources = {r.lower() for r in resources}
    if "wife" in text:
        return "wife" if text.startswith("wife ") else "operator"
    if "man" in lowered_resources:
        return "operator"
    if actor_spacy in PERSON_WORDS:
        return "operator" if actor_spacy == "man" else actor_spacy
    if any(p in text for p in ("tell wife", "walk", "grasp", "place", "open", "close", "cut", "butter")):
        return "operator"
    return actor_spacy


def _infer_object(text: str, resources: Sequence[str], obj_spacy: Optional[str]) -> Optional[str]:
    for phrase, label in PHRASE_OBJECTS:
        if phrase in text:
            return label
    label = _detect_label(OBJECT_WORDS, text)
    if label:
        return label
    lowered_resources = [r.lower() for r in resources]
    if "knife" in text:
        return "knife"
    for candidate in ["toast", "butter", "bread", "plate", "toaster", "refrigerator", "drawer", "cabinet"]:
        if candidate in lowered_resources:
            return candidate
    return obj_spacy


def _infer_location(text: str, resources: Sequence[str], loc_spacy: Optional[str]) -> Optional[str]:
    label = _detect_label(LOCATION_WORDS, text)
    if label:
        return label
    lowered_resources = [r.lower() for r in resources]
    for candidate in ["toaster", "refrigerator", "drawer", "cabinet"]:
        if candidate in text and candidate in lowered_resources:
            return candidate
    return loc_spacy


def _extract_entities(text: str, resources: Sequence[str]) -> Tuple[List[str], List[str], List[str]]:
    equipment = set()
    materials = set()
    tools = set()
    tokens = set(re.findall(r"[a-z_]+", text))
    for item in list(resources) + list(tokens):
        token = str(item).lower().strip()
        if token in EQUIPMENT_WORDS:
            equipment.add(token.title() if token != "toaster" else "Toaster")
        if token in MATERIAL_WORDS:
            materials.add(token.title())
        if token in TOOL_WORDS:
            tools.add(token.title())
    return sorted(equipment), sorted(materials), sorted(tools)


def _semantic_hint_from_action(action: str, text: str, location: Optional[str], machine_related: bool) -> str:
    if action in {"search", "wait"}:
        return "delay_or_search"
    if action == "inspect":
        return "inspection_check"
    if action == "start_machine":
        return "machine_start"
    if action == "process":
        return "machine_processing"
    if action == "cut":
        return "material_transformation"
    if action == "apply":
        return "finishing_application"
    if action == "assemble":
        return "assembly_or_stack"
    if action == "communicate":
        return "communication_step"
    if action == "cognitive":
        return "decision_or_reaction"
    if action == "move":
        if location or _contains_any(text, ["to", "from", "back"]):
            return "transport_or_motion"
        return "operator_motion"
    if action == "retrieve":
        return "retrieval_handling"
    if action == "handle":
        return "setup_or_placement" if machine_related or location else "manual_handling"
    return "general_support"


def _rule_waste(text: str, action: str, activity: str, va_flag: str, location: Optional[str], machine_related: bool) -> Tuple[str, float]:
    # activity comes from the source workbook; use it as a soft constraint instead of the only truth.
    if action in {"wait", "search"} or activity == "W":
        return "Waiting", 0.95
    if action == "inspect" or activity == "I":
        return "Inspection", 0.95
    if action == "communicate":
        return "Extra Processing", 0.75
    if action == "cognitive":
        return "Extra Processing", 0.80
    if action == "move":
        if _contains_any(text, ["with", "to", "from", "back", "another location", "original position"]) or location:
            return "Transportation", 0.84
        return "Motion", 0.84
    if action == "retrieve":
        return "Handling", 0.82
    if action == "handle":
        if activity == "S":
            return "Inventory", 0.92
        if _contains_any(text, ["open", "close", "position", "adjust", "set"]) or machine_related:
            return "Setup", 0.78
        return "Handling", 0.78
    if action in {"start_machine", "process", "cut", "apply", "assemble"}:
        if va_flag.upper() == "VA":
            return "Value Added", 0.90
        if machine_related and action in {"start_machine", "process"}:
            return "Setup", 0.75
        return "Extra Processing", 0.70
    if activity == "S":
        return "Inventory", 0.90
    if activity == "T":
        return "Transportation", 0.75
    if activity == "M":
        return "Handling", 0.65
    if activity == "D":
        return "Extra Processing", 0.65
    return ("Value Added", 0.60) if va_flag.upper() == "VA" else ("Handling", 0.55)


def _review_note(
    text: str,
    action: str,
    obj: Optional[str],
    activity: str,
    va_flag: str,
    waste_pred: str,
) -> Tuple[bool, str]:
    notes: List[str] = []
    if action == "cognitive":
        notes.append("Cognitive or reaction step; review whether it belongs in the physical FPC or should stay as a trigger/support note.")
    if action == "communicate":
        notes.append("Communication step detected; review because it is not an inspection even if the source activity uses I.")
    if obj == "butter" and "knife" in text:
        notes.append("Object corrected to knife-related handling where applicable; check 'butter knife' rows.")
    if va_flag.upper() == "VA" and waste_pred != "Value Added":
        notes.append("Source row is VA but normalized waste is not Value Added; review manually.")
    if activity == "I" and action != "inspect":
        notes.append("Workbook activity says Inspection but verb does not look like inspection.")
    if activity == "W" and action not in {"wait", "search"}:
        notes.append("Workbook activity says Waiting but verb does not look like wait/search.")
    return bool(notes), " ".join(notes)


def normalize_step(step: FPCStep, use_zero_shot: bool = False, use_manufactubert: bool = False) -> NormalizedStep:
    raw_text = _clean_text(getattr(step, "description", ""))
    text = raw_text.lower()
    resources = _resource_list(step)

    actor_spacy, obj_spacy, loc_spacy = _extract_spacy_parts(raw_text)
    action = _detect_action(text)
    actor = _infer_actor(text, resources, actor_spacy)
    obj = _infer_object(text, resources, obj_spacy)
    location = _infer_location(text, resources, loc_spacy)

    equipment, materials, tools = _extract_entities(text, resources)
    resource_entities = resources.copy()
    machine_related = action in {"start_machine", "process"} or bool(equipment) or "toaster" in text
    setup_related = _contains_any(text, ["open", "close", "setup", "prepare", "position", "adjust", "set"])
    delay_signal = action in {"wait", "search"}
    inspection_signal = action == "inspect"
    motion_signal = action in {"move", "retrieve"}
    search_signal = action == "search"

    std_parts = [actor, action, obj, location]
    standardized_text = " | ".join([part for part in std_parts if part])

    waste_pred, waste_score = _rule_waste(text, action, getattr(step, "activity", ""), getattr(step, "va_flag", ""), location, machine_related)

    if use_zero_shot:
        clf = _get_zero_shot_pipeline()
        if clf is not None:
            try:
                out = clf(raw_text, candidate_labels=ZERO_SHOT_LABELS, multi_label=False)
                waste_pred = out["labels"][0]
                waste_score = float(out["scores"][0])
            except Exception:
                pass

    manuf_hint = _semantic_hint_from_action(action, text, location, machine_related)
    if use_manufactubert:
        feat = _get_feature_pipeline()
        if feat is not None:
            try:
                feat(raw_text[:256])
                manuf_hint = f"manufactubert_embedding_ready | {manuf_hint}"
            except Exception:
                manuf_hint = f"heuristic_fallback | {manuf_hint}"
        else:
            manuf_hint = f"heuristic_fallback | {manuf_hint}"

    review_flag, review_note = _review_note(text, action, obj, getattr(step, "activity", ""), getattr(step, "va_flag", ""), waste_pred)

    return NormalizedStep(
        step=getattr(step, "step"),
        raw_description=raw_text,
        standardized_text=standardized_text,
        actor=actor,
        action=action,
        obj=obj,
        location=location,
        equipment=equipment,
        materials=materials,
        tools=tools,
        resource_entities=resource_entities,
        machine_related=machine_related,
        setup_related=setup_related,
        delay_signal=delay_signal,
        inspection_signal=inspection_signal,
        motion_signal=motion_signal,
        search_signal=search_signal,
        duration_sec=int(getattr(step, "duration_sec", 0) or 0),
        resources=resources,
        activity_type=getattr(step, "activity_type", ""),
        activity=getattr(step, "activity", ""),
        va_flag=getattr(step, "va_flag", ""),
        waste_pred=waste_pred,
        waste_score=round(float(waste_score), 3),
        manuf_semantic_hint=manuf_hint,
        review_flag=review_flag,
        review_note=review_note,
    )


def normalize_steps(steps: List[FPCStep], use_zero_shot: bool = False, use_manufactubert: bool = False) -> List[NormalizedStep]:
    return [normalize_step(s, use_zero_shot=use_zero_shot, use_manufactubert=use_manufactubert) for s in steps]


def normalized_to_df(steps: List[NormalizedStep]) -> pd.DataFrame:
    df = pd.DataFrame([asdict(s) for s in steps])
    if df.empty:
        return df

    preferred_order = [
        "step",
        "raw_description",
        "standardized_text",
        "actor",
        "action",
        "obj",
        "location",
        "equipment",
        "materials",
        "tools",
        "resource_entities",
        "machine_related",
        "setup_related",
        "delay_signal",
        "inspection_signal",
        "motion_signal",
        "search_signal",
        "duration_sec",
        "resources",
        "activity_type",
        "activity",
        "va_flag",
        "waste_pred",
        "waste_score",
        "manuf_semantic_hint",
        "review_flag",
        "review_note",
    ]
    return df[[c for c in preferred_order if c in df.columns]]
