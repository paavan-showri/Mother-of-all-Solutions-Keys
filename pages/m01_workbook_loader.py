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


def _clean_cell(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def _normalize_headers(cols) -> List[str]:
    return [_clean_cell(c) for c in cols]


def _find_header_row(raw: pd.DataFrame, required_headers: List[str], sheet_name: str) -> int:
    required = set(required_headers)
    for i in range(len(raw)):
        vals = {_clean_cell(x) for x in raw.iloc[i].tolist() if _clean_cell(x)}
        if required.issubset(vals):
            return i
    raise ValueError(
        f"Could not find the header row in worksheet '{sheet_name}'. "
        f"Expected headers: {required_headers}"
    )


def _read_table_with_detected_header(file, sheet_name: str, required_headers: List[str]) -> pd.DataFrame:
    file.seek(0)
    raw = pd.read_excel(file, sheet_name=sheet_name, header=None)
    header_row = _find_header_row(raw, required_headers, sheet_name)
    file.seek(0)
    df = pd.read_excel(file, sheet_name=sheet_name, header=header_row)
    df.columns = _normalize_headers(df.columns)
    missing = [c for c in required_headers if c not in df.columns]
    if missing:
        raise KeyError(f"Missing required headers in '{sheet_name}': {missing}")
    return df


def to_sec(value) -> Optional[int]:
    if pd.isna(value):
        return None
    if hasattr(value, "hour") and hasattr(value, "minute") and hasattr(value, "second"):
        return int(value.hour) * 3600 + int(value.minute) * 60 + int(value.second)
    if isinstance(value, (int, float)) and not pd.isna(value):
        return int(value)
    text = str(value).strip()
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
    return None


def parse_resources(text) -> List[str]:
    if pd.isna(text) or not str(text).strip():
        return []
    parts = re.split(r"\s*&\s*|\s*,\s*|\s+and\s+", str(text).strip(), flags=re.I)
    return [p.strip().title() for p in parts if p and p.strip()]


def _parse_task_id(value) -> Optional[int]:
    if pd.isna(value):
        return None
    text = str(value).strip()
    if not text or text in {"—", "-"}:
        return None
    nums = re.findall(r"\d+", text)
    if not nums:
        return None
    return int(nums[0])


def parse_predecessors(value) -> List[int]:
    if pd.isna(value):
        return []
    text = str(value).strip()
    if text in {"—", "-", "None", "none", ""}:
        return []
    return [int(n) for n in re.findall(r"\d+", text)]


def load_current_state_df(file, sheet_name: str = DEFAULT_FPC_SHEET) -> pd.DataFrame:
    return _read_table_with_detected_header(file, sheet_name, REQUIRED_FPC_HEADERS)


def load_current_state_steps(file, sheet_name: str = DEFAULT_FPC_SHEET) -> List[FPCStep]:
    df = load_current_state_df(file, sheet_name=sheet_name)
    steps: List[FPCStep] = []
    for _, row in df.iterrows():
        if pd.isna(row["Step"]):
            continue
        steps.append(FPCStep(
            step=int(row["Step"]),
            description=_clean_cell(row["Description"]),
            activity=_clean_cell(row["Activity"]),
            start_sec=to_sec(row["Start time"]),
            end_sec=to_sec(row["End time"]),
            duration_sec=to_sec(row["Duration (Sec)"]) or 0,
            activity_type=_clean_cell(row["Activity Type"]),
            resources=parse_resources(row["Resources"]),
        ))
    return steps


def load_precedence_tasks(file, sheet_name: str = DEFAULT_PRECEDENCE_SHEET) -> List[Task]:
    df = _read_table_with_detected_header(file, sheet_name, REQUIRED_PRECEDENCE_HEADERS)

    tasks: List[Task] = []
    seen_ids: set[int] = set()

    for excel_row_number, (_, row) in enumerate(df.iterrows(), start=2):
        task_id = _parse_task_id(row["Task ID"])
        if task_id is None:
            continue

        if task_id in seen_ids:
            raise ValueError(
                f"Duplicate Task ID {task_id} found in worksheet '{sheet_name}'."
            )
        seen_ids.add(task_id)

        name = _clean_cell(row["Task Name"])
        if not name:
            raise ValueError(
                f"Missing Task Name for Task ID {task_id} in worksheet '{sheet_name}'."
            )

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
            f"Worksheet '{sheet_name}' has predecessor IDs that are not present in the Task ID column: {missing_preds}. "
            f"Check the 'Immediate Predecessors' values."
        )

    return tasks


def load_resource_capacities(file, sheet_name: str = DEFAULT_RESOURCE_SHEET) -> Dict[str, int]:
    df = _read_table_with_detected_header(file, sheet_name, REQUIRED_RESOURCE_HEADERS)
    capacities: Dict[str, int] = {}
    for _, row in df.iterrows():
        if pd.isna(row["Resource"]):
            continue
        capacities[_clean_cell(row["Resource"]).title()] = int(row["Capacity"])
    return capacities
