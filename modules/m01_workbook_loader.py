from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional

import pandas as pd

DEFAULT_FPC_SHEET = "FPC_Current State"
DEFAULT_PRECEDENCE_SHEET = "Precedence Network"
DEFAULT_RESOURCE_SHEET = "Resources"

REQUIRED_FPC_HEADERS = [
    "Step",
    "Description",
    "Activity",
    "Start time",
    "End time",
    "Duration",
    "Activity Type",
    "Resources",
]
REQUIRED_PRECEDENCE_HEADERS = [
    "Task ID",
    "Task Name",
    "Duration",
    "Immediate Predecessors",
    "Resources",
]
REQUIRED_RESOURCE_HEADERS = ["Resource", "Capacity"]


@dataclass
class FPCStep:
    step: int
    description: str
    activity: str
    start_sec: Optional[int]
    end_sec: Optional[int]
    duration_sec: int
    activity_type: str
    resources: List[str]


@dataclass
class Task:
    task_id: int
    name: str
    duration_sec: int
    predecessors: List[int] = field(default_factory=list)
    resources: List[str] = field(default_factory=list)
    internal_external: str = "internal"


def _find_header_row(raw: pd.DataFrame) -> int:
    for i in range(len(raw)):
        vals = [str(x).strip() for x in raw.iloc[i].tolist() if pd.notna(x)]
        if all(h in vals for h in ["Step", "Description", "Activity", "Resources"]):
            return i
    raise ValueError("Could not find the Flow Process Chart header row.")


def to_sec(value) -> Optional[int]:
    if pd.isna(value):
        return None
    if hasattr(value, "hour") and hasattr(value, "minute") and hasattr(value, "second"):
        return int(value.hour) * 3600 + int(value.minute) * 60 + int(value.second)
    if isinstance(value, (int, float)):
        return int(value)
    text = str(value).strip()
    m = re.match(r"^(\d+):(\d{1,2})(?::(\d{1,2}))?$", text)
    if m:
        if m.group(3) is None:
            mm = int(m.group(1)); ss = int(m.group(2))
            return mm * 60 + ss
        hh = int(m.group(1)); mm = int(m.group(2)); ss = int(m.group(3))
        return hh * 3600 + mm * 60 + ss
    return None


def parse_resources(text) -> List[str]:
    if pd.isna(text) or not str(text).strip():
        return []
    parts = re.split(r"\s*&\s*|\s*,\s*|\s+and\s+", str(text).strip(), flags=re.I)
    return [p.strip().title() for p in parts if p and p.strip()]


def parse_predecessors(value) -> List[int]:
    if pd.isna(value):
        return []
    text = str(value).strip()
    if text in {"—", "-", "None", "none", ""}:
        return []
    nums = re.findall(r"\d+", text)
    if "," in text and nums:
        return [int(n) for n in nums]
    if text.isdigit() and len(text) == 5 and text.endswith("100"):
        return [int(text[:2]), int(text[2:])]
    if nums:
        return [int(n) for n in nums]
    return []


def load_current_state_df(file, sheet_name: str = DEFAULT_FPC_SHEET) -> pd.DataFrame:
    raw = pd.read_excel(file, sheet_name=sheet_name, header=None)
    header_row = _find_header_row(raw)
    file.seek(0)
    df = pd.read_excel(file, sheet_name=sheet_name, header=header_row)
    df.columns = [str(c).strip() for c in df.columns]
    missing = [c for c in REQUIRED_FPC_HEADERS if c not in df.columns]
    if missing:
        raise KeyError(f"Missing required headers in {sheet_name}: {missing}")
    return df


def load_current_state_steps(file, sheet_name: str = DEFAULT_FPC_SHEET) -> List[FPCStep]:
    df = load_current_state_df(file, sheet_name=sheet_name)
    steps: List[FPCStep] = []
    for _, row in df.iterrows():
        if pd.isna(row["Step"]):
            continue
        steps.append(FPCStep(
            step=int(row["Step"]),
            description=str(row["Description"]).strip(),
            activity=str(row["Activity"]).strip(),
            start_sec=to_sec(row["Start time"]),
            end_sec=to_sec(row["End time"]),
            duration_sec=to_sec(row["Duration"]) or 0,
            activity_type=str(row["Activity Type"]).strip(),
            resources=parse_resources(row["Resources"]),
        ))
    return steps


def load_precedence_tasks(file, sheet_name: str = DEFAULT_PRECEDENCE_SHEET) -> List[Task]:
    file.seek(0)
    df = pd.read_excel(file, sheet_name=sheet_name)
    df.columns = [str(c).strip() for c in df.columns]
    missing = [c for c in REQUIRED_PRECEDENCE_HEADERS if c not in df.columns]
    if missing:
        raise KeyError(f"Missing required headers in {sheet_name}: {missing}")
    tasks: List[Task] = []
    for _, row in df.iterrows():
        if pd.isna(row["Task ID"]):
            continue
        name = str(row["Task Name"]).strip()
        tasks.append(Task(
            task_id=int(row["Task ID"]),
            name=name,
            duration_sec=to_sec(row["Duration"]) or 0,
            predecessors=parse_predecessors(row["Immediate Predecessors"]),
            resources=parse_resources(row["Resources"]),
            internal_external="external" if name in {"Get butter", "Get knife", "Cut butter"} else "internal",
        ))
    return tasks


def load_resource_capacities(file, sheet_name: str = DEFAULT_RESOURCE_SHEET) -> Dict[str, int]:
    file.seek(0)
    df = pd.read_excel(file, sheet_name=sheet_name)
    df.columns = [str(c).strip() for c in df.columns]
    missing = [c for c in REQUIRED_RESOURCE_HEADERS if c not in df.columns]
    if missing:
        raise KeyError(f"Missing required headers in {sheet_name}: {missing}")
    return {
        str(row["Resource"]).strip().title(): int(row["Capacity"])
        for _, row in df.iterrows() if pd.notna(row["Resource"])
    }
