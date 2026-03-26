from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional

import pandas as pd

DEFAULT_FPC_SHEET = "FPC_Current State"
DEFAULT_PRECEDENCE_SHEET = "Precedence Network"
DEFAULT_RESOURCE_SHEET = "Resources"

REQUIRED_FPC_HEADERS = [
    "Step", "Description", "Activity", "Start time", "End time",
    "Duration (Sec)", "Resources",
]
REQUIRED_PRECEDENCE_HEADERS = [
    "Task ID", "Task Name", "Duration", "Immediate Predecessors", "Resources",
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
    va_flag: str = ""


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
    out = []
    for c in cols:
        out.append(_clean_cell(c))
    return out


def _ensure_sheet_exists(file, sheet_name: str) -> None:
    file.seek(0)
    xls = pd.ExcelFile(file)
    if sheet_name not in xls.sheet_names:
        raise ValueError(f"Worksheet '{sheet_name}' was not found. Available sheets: {xls.sheet_names}")


def _find_header_row(raw: pd.DataFrame, required_headers: List[str], sheet_name: str) -> int:
    required = set(required_headers)
    for i in range(len(raw)):
        vals = {_clean_cell(x) for x in raw.iloc[i].tolist() if _clean_cell(x)}
        if required.issubset(vals):
            return i
    raise ValueError(f"Could not find the header row in worksheet '{sheet_name}'. Expected headers: {required_headers}")


def _read_table_with_detected_header(file, sheet_name: str, required_headers: List[str]) -> pd.DataFrame:
    _ensure_sheet_exists(file, sheet_name)
    file.seek(0)
    raw = pd.read_excel(file, sheet_name=sheet_name, header=None)
    header_row = _find_header_row(raw, required_headers, sheet_name)
    file.seek(0)
    df = pd.read_excel(file, sheet_name=sheet_name, header=header_row)
    df.columns = _normalize_headers(df.columns)
    return df


def to_sec(value) -> Optional[int]:
    if pd.isna(value):
        return None
    if hasattr(value, 'hour') and hasattr(value, 'minute') and hasattr(value, 'second'):
        return int(value.hour) * 3600 + int(value.minute) * 60 + int(value.second)
    if isinstance(value, (int, float)) and not pd.isna(value):
        v = float(value)
        if 0 < v < 1:
            return int(round(v * 24 * 3600))
        return int(round(v))
    text = str(value).strip()
    m = re.match(r"^(\d+):(\d{1,2})(?::(\d{1,2}))?$", text)
    if m:
        if m.group(3) is None:
            return int(m.group(1)) * 60 + int(m.group(2))
        return int(m.group(1)) * 3600 + int(m.group(2)) * 60 + int(m.group(3))
    return None


def parse_resources(text) -> List[str]:
    if pd.isna(text) or not str(text).strip():
        return []
    parts = re.split(r"\s*&\s*|\s*,\s*|\s*;\s*|\s+and\s+", str(text).strip(), flags=re.I)
    out = []
    for part in parts:
        p = re.sub(r":\s*\d+$", "", part.strip())
        if p:
            out.append(p.title())
    return out


def _parse_task_id(value) -> Optional[int]:
    if pd.isna(value):
        return None
    nums = re.findall(r"\d+", str(value))
    return int(nums[0]) if nums else None


def parse_predecessors(value) -> List[int]:
    if pd.isna(value):
        return []
    text = str(value).strip()
    if text in {'—', '-', 'None', 'none', '', 'nan'}:
        return []
    return [int(n) for n in re.findall(r"\d+", text)]


def load_current_state_df(file, sheet_name: str = DEFAULT_FPC_SHEET) -> pd.DataFrame:
    return _read_table_with_detected_header(file, sheet_name, REQUIRED_FPC_HEADERS)


def load_current_state_steps(file, sheet_name: str = DEFAULT_FPC_SHEET) -> List[FPCStep]:
    df = load_current_state_df(file, sheet_name)
    steps = []
    for _, row in df.iterrows():
        if pd.isna(row.get('Step')):
            continue
        steps.append(FPCStep(
            step=int(row['Step']),
            description=_clean_cell(row['Description']),
            activity=_clean_cell(row['Activity']).upper(),
            start_sec=to_sec(row.get('Start time')),
            end_sec=to_sec(row.get('End time')),
            duration_sec=to_sec(row.get('Duration (Sec)')) or 0,
            activity_type=_clean_cell(row.get('Activity Type', row.get('VA / NVA / NNVA', ''))),
            resources=parse_resources(row.get('Resources')),
            va_flag=_clean_cell(row.get('VA / NVA / NNVA', '')),
        ))
    return steps


def load_precedence_tasks(file, sheet_name: str = DEFAULT_PRECEDENCE_SHEET) -> List[Task]:
    df = _read_table_with_detected_header(file, sheet_name, REQUIRED_PRECEDENCE_HEADERS)
    tasks = []
    for _, row in df.iterrows():
        task_id = _parse_task_id(row.get('Task ID'))
        if task_id is None:
            continue
        tasks.append(Task(
            task_id=task_id,
            name=_clean_cell(row.get('Task Name')),
            duration_sec=to_sec(row.get('Duration')) or 0,
            predecessors=parse_predecessors(row.get('Immediate Predecessors')),
            resources=parse_resources(row.get('Resources')),
            internal_external=_clean_cell(row.get('Type', 'internal')) or 'internal',
        ))
    return tasks


def load_resource_capacities(file, sheet_name: str = DEFAULT_RESOURCE_SHEET) -> Dict[str, int]:
    df = _read_table_with_detected_header(file, sheet_name, REQUIRED_RESOURCE_HEADERS)
    out = {}
    for _, row in df.iterrows():
        name = _clean_cell(row.get('Resource'))
        if not name:
            continue
        try:
            cap = int(float(row.get('Capacity', 1)))
        except Exception:
            cap = 1
        out[name.title()] = max(cap, 1)
    return out
