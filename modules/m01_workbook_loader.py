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
    normalized: List[str] = []
    seen: Dict[str, int] = {}
    for c in cols:
        text = _clean_cell(c)
        if text.endswith(".1") and text[:-2] in seen:
            text = f"{text[:-2]}__dup{seen[text[:-2]] + 1}"
        seen[text] = seen.get(text, 0) + 1
        normalized.append(text)
    return normalized


def _ensure_sheet_exists(file, sheet_name: str) -> None:
    file.seek(0)
    xls = pd.ExcelFile(file)
    if sheet_name not in xls.sheet_names:
        raise ValueError(
            f"Worksheet '{sheet_name}' was not found. Available sheets: {xls.sheet_names}"
        )


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
    _ensure_sheet_exists(file, sheet_name)
    file.seek(0)
    raw = pd.read_excel(file, sheet_name=sheet_name, header=None)
    header_row = _find_header_row(raw, required_headers, sheet_name)
    file.seek(0)
    df = pd.read_excel(file, sheet_name=sheet_name, header=header_row)
    df.columns = _normalize_headers(df.columns)
    missing = [c for c in required_headers if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required headers in '{sheet_name}': {missing}")
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
    parts = re.split(r"\s*&\s*|\s*,\s*|\s*;\s*|\s+and\s+", str(text).strip(), flags=re.I)
    cleaned: List[str] = []
    for part in parts:
        p = part.strip()
        if not p:
            continue
        p = re.sub(r":\s*\d+$", "", p)
        cleaned.append(p.title())
    return cleaned


def _parse_task_id(value) -> Optional[int]:
    if pd.isna(value):
        return None
    text = str(value).strip()
    if not text or text in {"—", "-", "None", "none", "nan"}:
        return None
    nums = re.findall(r"\d+", text)
    if not nums:
        return None
    return int(nums[0])


def parse_predecessors(value) -> List[int]:
    if pd.isna(value):
        return []
    text = str(value).strip()
    if text in {"—", "-", "None", "none", "", "nan"}:
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

    task_id_col = "Task ID"
    if "Task ID__dup2" in df.columns:
        # Some workbooks contain two Task ID columns. Prefer the fuller one.
        first_ids = {_parse_task_id(v) for v in df["Task ID"].tolist() if _parse_task_id(v) is not None}
        second_ids = {_parse_task_id(v) for v in df["Task ID__dup2"].tolist() if _parse_task_id(v) is not None}
        if len(second_ids) > len(first_ids):
            task_id_col = "Task ID__dup2"

    tasks: List[Task] = []
    seen_ids: set[int] = set()
    skipped_rows = 0

    for _, row in df.iterrows():
        raw_task_id = row.get(task_id_col)
        task_id = _parse_task_id(raw_task_id)
        name = _clean_cell(row.get("Task Name", ""))
        predecessors = parse_predecessors(row.get("Immediate Predecessors", ""))
        resources = parse_resources(row.get("Resources", ""))
        duration = to_sec(row.get("Duration"))

        # ignore fully blank rows
        if task_id is None and not name and duration is None and not predecessors and not resources:
            continue

        # ignore malformed spillover rows where task id is missing and task name is clearly not a task name
        if task_id is None:
            skipped_rows += 1
            continue

        if not name:
            raise ValueError(
                f"Missing Task Name for Task ID {task_id} in worksheet '{sheet_name}'."
            )

        # skip malformed rows where task name is numeric or obviously shifted from another format
        if re.fullmatch(r"\d+(?:\.\d+)?", name):
            skipped_rows += 1
            continue

        if task_id in seen_ids:
            raise ValueError(
                f"Duplicate Task ID {task_id} found in worksheet '{sheet_name}'."
            )
        seen_ids.add(task_id)

        tasks.append(Task(
            task_id=task_id,
            name=name,
            duration_sec=duration or 0,
            predecessors=predecessors,
            resources=resources,
            internal_external="external" if any(k in name.lower() for k in ["retrieve", "get", "cut butter"]) else "internal",
        ))

    if not tasks:
        raise ValueError(
            f"No valid task rows were found in worksheet '{sheet_name}'. "
            f"Expected columns: {REQUIRED_PRECEDENCE_HEADERS}"
        )

    task_ids = {t.task_id for t in tasks}
    missing_preds = sorted({pred for t in tasks for pred in t.predecessors if pred not in task_ids})
    if missing_preds:
        raise ValueError(
            f"Worksheet '{sheet_name}' has predecessor IDs not present in the Task ID column: {missing_preds}."
        )

    return sorted(tasks, key=lambda t: t.task_id)


def load_resource_capacities(file, sheet_name: str = DEFAULT_RESOURCE_SHEET) -> Dict[str, int]:
    df = _read_table_with_detected_header(file, sheet_name, REQUIRED_RESOURCE_HEADERS)
    capacities: Dict[str, int] = {}
    for _, row in df.iterrows():
        resource = _clean_cell(row.get("Resource"))
        if not resource:
            continue
        try:
            capacity = int(row.get("Capacity"))
        except Exception as exc:  # noqa: BLE001
            raise ValueError(
                f"Invalid capacity for resource '{resource}' in worksheet '{sheet_name}'."
            ) from exc
        capacities[resource.title()] = capacity
    if not capacities:
        raise ValueError(f"No resources found in worksheet '{sheet_name}'.")
    return capacities
