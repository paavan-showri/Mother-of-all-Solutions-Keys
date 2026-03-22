from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

from .m01_workbook_loader import FPCStep

try:
    import spacy
    from spacy.language import Language
    from spacy.pipeline import EntityRuler
except Exception:  # pragma: no cover
    spacy = None
    Language = Any  # type: ignore
    EntityRuler = Any  # type: ignore

try:
    from transformers import pipeline as hf_pipeline
except Exception:  # pragma: no cover
    hf_pipeline = None


@dataclass
class NormalizedStep:
    step: int
    raw_description: str
    standardized_text: str
    actor: Optional[str]
    action: str
    action_lemma: str
    obj: Optional[str]
    location: Optional[str]
    machine: Optional[str]
    tool: Optional[str]
    material: Optional[str]
    product: Optional[str]
    quality_signal: Optional[str]
    delay_signal: Optional[str]
    setup_signal: Optional[str]
    tokens: List[str]
    lemmas: List[str]
    noun_chunks: List[str]
    entities: List[Dict[str, str]]
    keyword_hits: List[str]
    waste_hints: List[str]
    duration_sec: int
    resources: List[str]
    activity_type: str
    activity: str
    parser_source: str
    classifier_source: str
    classifier_labels: List[str] = field(default_factory=list)


VERB_PATTERNS = [
    (r"\b(search|look for|locate|find)\b", "search"),
    (r"\b(wait|queue|idle|hold on)\b", "wait"),
    (r"\b(inspect|check|verify|examine|review|measure|test)\b", "inspect"),
    (r"\b(walk|move|carry|transport|bring|go to|return|transfer|reposition|turn back|turn around)\b", "move"),
    (r"\b(retrieve|get|take|pick|collect|fetch|grasp)\b", "retrieve"),
    (r"\b(open|close|position|place|drop|adjust|hold|stage|set)\b", "handle"),
    (r"\b(start|turn on|activate|switch on|cycle start)\b", "start_machine"),
    (r"\b(run|process|machine|toast|drill|weld|mix|assemble|paint)\b", "process"),
    (r"\b(load|unload|insert|mount|remove)\b", "load_unload"),
    (r"\b(cut|slice|trim|split)\b", "cut"),
    (r"\b(spread|butter|apply|coat|dispense)\b", "apply"),
    (r"\b(stack|assemble|press|flip|join)\b", "assemble"),
    (r"\b(pack|label|deliver|serve|ship|tell)\b", "serve"),
    (r"\b(rework|repair|fix|redo)\b", "rework"),
]

ONTOLOGY_PATTERNS: Dict[str, Dict[str, List[str]]] = {
    "EQUIPMENT": {
        "toaster": ["toaster", "oven", "machine", "press", "cnc", "mixer", "furnace", "line"],
    },
    "TOOL": {
        "knife": ["knife", "wrench", "screwdriver", "hammer", "fixture", "gauge", "tool"],
    },
    "MATERIAL": {
        "bread": ["bread", "bag of bread", "slice"],
        "butter": ["butter", "adhesive", "oil", "grease", "resin", "paint"],
    },
    "PRODUCT": {
        "toast": ["toast", "assembly", "part", "wip", "product"],
        "plate": ["plate", "tray", "carton", "container", "bin"],
    },
    "LOCATION": {
        "cabinet": ["cabinet", "rack", "shelf", "tool crib", "storage"],
        "counter": ["counter", "table", "workbench", "station", "cell"],
        "refrigerator": ["refrigerator", "fridge", "cooler"],
        "living_room": ["living room", "shipping", "dock", "customer area"],
    },
    "QUALITY_SIGNAL": {
        "inspection": ["inspect", "verify", "check", "quality", "defect", "porosity", "dimension"],
    },
    "DELAY_SIGNAL": {
        "waiting": ["wait", "queue", "delay", "idle", "hold"],
    },
    "SETUP_SIGNAL": {
        "setup": ["setup", "changeover", "prepare", "stage", "load", "unload"],
    },
}

WASTE_HINT_PATTERNS = {
    "waiting": ["wait", "queue", "delay", "idle", "hold"],
    "motion": ["walk", "reach", "bend", "turn", "search", "look for"],
    "transportation": ["move", "carry", "transport", "bring", "transfer"],
    "inventory": ["store", "inventory", "stock", "queue", "wip"],
    "extra_processing": ["inspect", "reinspect", "verify", "double check", "position again", "adjust"],
    "defects": ["rework", "repair", "fix", "defect", "reject", "scrap"],
    "overproduction": ["extra", "more than needed", "overproduce", "batch"],
    "non_utilized_talent": ["manual decision", "operator decides", "asks supervisor"],
}

HF_LABELS = [
    "waiting",
    "motion",
    "transportation",
    "inventory",
    "extra_processing",
    "defects",
    "overproduction",
    "non_utilized_talent",
    "value_added_operation",
]


def _safe_lower(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip().lower()


class ManufacturingNLPPipeline:
    def __init__(self, enable_transformers: bool = True) -> None:
        self.enable_transformers = enable_transformers
        self.nlp = self._build_spacy_pipeline()
        self.zero_shot = self._build_transformer_pipeline() if enable_transformers else None

    def _build_spacy_pipeline(self) -> Optional[Language]:
        if spacy is None:
            return None
        try:
            nlp = spacy.load("en_core_web_sm")
            source = "spacy_model"
        except Exception:
            nlp = spacy.blank("en")
            if "sentencizer" not in nlp.pipe_names:
                nlp.add_pipe("sentencizer")
            source = "spacy_blank"
        if "entity_ruler" not in nlp.pipe_names:
            ruler = nlp.add_pipe("entity_ruler", config={"overwrite_ents": True})
        else:
            ruler = nlp.get_pipe("entity_ruler")
        assert isinstance(ruler, EntityRuler)
        patterns = []
        for label, mapping in ONTOLOGY_PATTERNS.items():
            for canonical, variants in mapping.items():
                for variant in variants:
                    patterns.append({"label": label, "pattern": variant, "id": canonical})
        ruler.add_patterns(patterns)
        nlp.meta["manufacturing_source"] = source
        return nlp

    def _build_transformer_pipeline(self):
        if hf_pipeline is None:
            return None
        try:
            return hf_pipeline(
                "zero-shot-classification",
                model="facebook/bart-large-mnli",
            )
        except Exception:
            return None

    def _rule_action(self, text: str) -> str:
        for pattern, label in VERB_PATTERNS:
            if re.search(pattern, text):
                return label
        return "other"

    def _waste_hints(self, text: str) -> List[str]:
        hits: List[str] = []
        for waste, patterns in WASTE_HINT_PATTERNS.items():
            if any(p in text for p in patterns):
                hits.append(waste)
        return hits

    def _run_transformer(self, text: str) -> tuple[str, List[str]]:
        if not self.zero_shot or len(text.split()) < 2:
            return "none", []
        try:
            result = self.zero_shot(text, HF_LABELS, multi_label=True)
            labels = [label for label, score in zip(result["labels"], result["scores"]) if score >= 0.35]
            return "huggingface_zero_shot", labels
        except Exception:
            return "none", []

    def normalize_step(self, step: FPCStep) -> NormalizedStep:
        text = _safe_lower(step.description)
        parser_source = "regex_only"
        tokens: List[str] = text.split()
        lemmas: List[str] = tokens[:]
        noun_chunks: List[str] = []
        entities: List[Dict[str, str]] = []

        if self.nlp is not None:
            doc = self.nlp(step.description)
            parser_source = str(self.nlp.meta.get("manufacturing_source", "spacy"))
            tokens = [t.text for t in doc]
            if any(getattr(t, "lemma_", "") for t in doc):
                lemmas = [t.lemma_.lower() if t.lemma_ else t.text.lower() for t in doc]
            noun_chunks = [chunk.text for chunk in doc.noun_chunks] if doc.has_annotation("DEP") else []
            entities = [
                {
                    "label": ent.label_,
                    "text": ent.text,
                    "canonical": ent.ent_id_ or ent.text.lower(),
                }
                for ent in doc.ents
            ]

        def first_entity(label: str) -> Optional[str]:
            for ent in entities:
                if ent["label"] == label:
                    return ent["canonical"]
            return None

        actor = "operator" if any(r.lower() in {"man", "operator", "worker", "personnel"} for r in step.resources) else None
        action = self._rule_action(text)
        obj = first_entity("PRODUCT") or first_entity("MATERIAL")
        location = first_entity("LOCATION")
        machine = first_entity("EQUIPMENT")
        tool = first_entity("TOOL")
        material = first_entity("MATERIAL")
        product = first_entity("PRODUCT")
        quality_signal = first_entity("QUALITY_SIGNAL")
        delay_signal = first_entity("DELAY_SIGNAL")
        setup_signal = first_entity("SETUP_SIGNAL")

        keyword_hits = sorted({ent["canonical"] for ent in entities})
        waste_hints = self._waste_hints(text)
        classifier_source, classifier_labels = self._run_transformer(step.description)

        standardized_parts = [
            actor,
            action,
            machine or tool,
            material or product,
            location,
        ]
        standardized_text = " | ".join([p for p in standardized_parts if p])

        return NormalizedStep(
            step=step.step,
            raw_description=step.description,
            standardized_text=standardized_text,
            actor=actor,
            action=action,
            action_lemma=action,
            obj=obj,
            location=location,
            machine=machine,
            tool=tool,
            material=material,
            product=product,
            quality_signal=quality_signal,
            delay_signal=delay_signal,
            setup_signal=setup_signal,
            tokens=tokens,
            lemmas=lemmas,
            noun_chunks=noun_chunks,
            entities=entities,
            keyword_hits=keyword_hits,
            waste_hints=waste_hints,
            duration_sec=step.duration_sec,
            resources=step.resources,
            activity_type=step.activity_type,
            activity=step.activity,
            parser_source=parser_source,
            classifier_source=classifier_source,
            classifier_labels=classifier_labels,
        )


_DEFAULT_PIPELINE: Optional[ManufacturingNLPPipeline] = None


def get_pipeline(enable_transformers: bool = True) -> ManufacturingNLPPipeline:
    global _DEFAULT_PIPELINE
    if _DEFAULT_PIPELINE is None or _DEFAULT_PIPELINE.enable_transformers != enable_transformers:
        _DEFAULT_PIPELINE = ManufacturingNLPPipeline(enable_transformers=enable_transformers)
    return _DEFAULT_PIPELINE


def normalize_step(step: FPCStep, enable_transformers: bool = True) -> NormalizedStep:
    return get_pipeline(enable_transformers=enable_transformers).normalize_step(step)


def normalize_steps(steps: List[FPCStep], enable_transformers: bool = True) -> List[NormalizedStep]:
    pipe = get_pipeline(enable_transformers=enable_transformers)
    return [pipe.normalize_step(s) for s in steps]
