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
    "Duration (Sec)",
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


def _find_header_row(raw: pd.DataFrame, required_headers: List[str]) -> int:
    for i in range(len(raw)):
        vals = [str(x).strip() for x in raw.iloc[i].tolist() if pd.notna(x)]
        if all(h in vals for h in required_headers):
            return i
    raise ValueError(f"Could not find header row containing: {required_headers}")


def _read_sheet_with_detected_header(file, sheet_name: str, required_headers: List[str]) -> pd.DataFrame:
    file.seek(0)
    raw = pd.read_excel(file, sheet_name=sheet_name, header=None)
    header_row = _find_header_row(raw, required_headers)
    file.seek(0)
    df = pd.read_excel(file, sheet_name=sheet_name, header=header_row)
    df.columns = [str(c).strip() for c in df.columns]
    missing = [c for c in required_headers if c not in df.columns]
    if missing:
        raise KeyError(f"Missing required headers in {sheet_name}: {missing}")
    return df


def to_sec(value) -> Optional[int]:
    if pd.isna(value):
        return None
    if hasattr(value, "hour") and hasattr(value, "minute") and hasattr(value, "second"):
        return int(value.hour) * 3600 + int(value.minute) * 60 + int(value.second)
    if isinstance(value, (int, float)):
        return int(value)
    text = str(value).strip()
    if not text:
        return None
    m = re.match(r"^(\d+):(\d{1,2})(?::(\d{1,2}))?$", text)
    if m:
        if m.group(3) is None:
            mm = int(m.group(1))
            ss = int(m.group(2))
            return mm * 60 + ss
        hh = int(m.group(1))
        mm = int(m.group(2))
        ss = int(m.group(3))
        return hh * 3600 + mm * 60 + ss
    if text.isdigit():
        return int(text)
    return None


def parse_resources(text) -> List[str]:
    if pd.isna(text) or not str(text).strip():
        return []
    parts = re.split(r"\s*&\s*|\s*,\s*|\s+and\s+", str(text).strip(), flags=re.I)
    return [p.strip().title() for p in parts if p and p.strip()]


def parse_task_id(value) -> int:
    text = str(value).strip()
    if not text or text.lower() == 'nan':
        raise ValueError("Task ID cannot be blank.")
    match = re.search(r"\d+", text)
    if not match:
        raise ValueError(f"Could not parse Task ID from value: {value!r}")
    return int(match.group())


def parse_predecessors(value) -> List[int]:
    if pd.isna(value):
        return []
    text = str(value).strip()
    if text in {"—", "-", "None", "none", "", "nan", "NaN"}:
        return []
    return [int(x) for x in re.findall(r"\d+", text)]


def load_current_state_df(file, sheet_name: str = DEFAULT_FPC_SHEET) -> pd.DataFrame:
    return _read_sheet_with_detected_header(file, sheet_name, REQUIRED_FPC_HEADERS)


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
            duration_sec=to_sec(row["Duration (Sec)"]) or 0,
            activity_type=str(row["Activity Type"]).strip(),
            resources=parse_resources(row["Resources"]),
        ))
    return steps


def load_precedence_tasks(file, sheet_name: str = DEFAULT_PRECEDENCE_SHEET) -> List[Task]:
    df = _read_sheet_with_detected_header(file, sheet_name, REQUIRED_PRECEDENCE_HEADERS)
    tasks: List[Task] = []
    seen_ids = set()

    for _, row in df.iterrows():
        if pd.isna(row["Task ID"]):
            continue
        task_id = parse_task_id(row["Task ID"])
        if task_id in seen_ids:
            raise ValueError(f"Duplicate Task ID found in {sheet_name}: {task_id}")
        seen_ids.add(task_id)

        name = str(row["Task Name"]).strip()
        tasks.append(Task(
            task_id=task_id,
            name=name,
            duration_sec=to_sec(row["Duration"]) or 0,
            predecessors=parse_predecessors(row["Immediate Predecessors"]),
            resources=parse_resources(row["Resources"]),
            internal_external="external" if name in {"Get butter", "Get knife", "Cut butter"} else "internal",
        ))

    task_ids = {t.task_id for t in tasks}
    missing_preds = sorted({pred for t in tasks for pred in t.predecessors if pred not in task_ids})
    if missing_preds:
        raise ValueError(
            f"Precedence sheet '{sheet_name}' references predecessor Task IDs that are not defined: {missing_preds}. "
            "Check the 'Immediate Predecessors' column and make sure every predecessor also appears in 'Task ID'."
        )
    return tasks


def load_resource_capacities(file, sheet_name: str = DEFAULT_RESOURCE_SHEET) -> Dict[str, int]:
    df = _read_sheet_with_detected_header(file, sheet_name, REQUIRED_RESOURCE_HEADERS)
    capacities: Dict[str, int] = {}
    for _, row in df.iterrows():
        if pd.isna(row["Resource"]):
            continue
        capacities[str(row["Resource"]).strip().title()] = int(row["Capacity"])
    return capacities
