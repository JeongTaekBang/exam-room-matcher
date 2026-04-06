"""운영 워크플로우 보조 유틸.

배정 결과 영속화, 작업 이력 기록, 시험 교시 계산을 담당한다.
"""

from __future__ import annotations

import datetime
import json
import logging
import tempfile
from pathlib import Path
from typing import Any

_log = logging.getLogger(__name__)


class StaleFileError(Exception):
    """파일이 다른 세션에서 변경되었음을 나타낸다."""


def _safe_int(value: Any, default: int = 0) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


MIN_PERIOD = 0
MAX_PERIOD = 14


def _clamp_periods(periods: list[int]) -> list[int]:
    """교시를 0~14 범위로 제한한다."""
    return sorted({max(MIN_PERIOD, min(MAX_PERIOD, p)) for p in periods})


def resolve_exam_room(req, exam_day: str) -> str:
    """시험일 요일에 해당하는 강의실을 슬롯에서 결정한다.

    수업시간표에서 시험 요일과 일치하는 슬롯의 강의실을 반환한다.
    일치하는 슬롯이 없으면 req.room(강의실 열)을 폴백으로 사용한다.
    """
    for slot in req.slots:
        if slot.day == exam_day and slot.room:
            return slot.room
    return req.room


def resolve_needed_periods(req, exam_day: str) -> list[int]:
    """요청의 시험 필요 교시를 계산한다.

    우선순위:
    1) 시험 시작/종료교시가 있으면 해당 범위를 사용
    2) 없으면 수업시간표 슬롯에서 시험 요일의 교시를 사용

    반환값은 항상 0~14 범위로 클램프된다.
    """
    if req.exam_start is not None and req.exam_end is not None:
        start_p, end_p = sorted((req.exam_start, req.exam_end))
        return _clamp_periods(list(range(start_p, end_p + 1)))

    periods = set()
    for slot in req.slots:
        if slot.day != exam_day:
            continue
        start_p, end_p = sorted((slot.start, slot.end))
        for p in range(start_p, end_p + 1):
            periods.add(p)
    return _clamp_periods(list(periods))


def load_assignments(path: Path) -> dict:
    """저장된 배정 결과를 읽어 dict로 반환한다."""
    if not path.exists():
        return {}
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        _log.warning("배정 파일 손상: %s — 빈 dict로 로드", path)
        return {}
    except OSError:
        return {}
    if not isinstance(raw, dict):
        return {}

    normalized = {}
    for key, value in raw.items():
        if not isinstance(key, str) or not isinstance(value, dict):
            continue
        periods = value.get("periods", [])
        if not isinstance(periods, list):
            periods = []
        safe_periods = [_safe_int(p, default=-1) for p in periods]
        safe_periods = sorted({p for p in safe_periods if MIN_PERIOD <= p <= MAX_PERIOD})
        if not safe_periods:
            continue

        normalized[key] = {
            "room": str(value.get("room", "")),
            "sheet": str(value.get("sheet", "")),
            "periods": safe_periods,
            "original_room": str(value.get("original_room", "")),
            "students": _safe_int(value.get("students"), default=0),
            "category": str(value.get("category", "")),
            "keep_orig": bool(value.get("keep_orig", True)),
        }
    return normalized


def save_assignments(path: Path, assignments: dict, expected_mtime: float | None = None) -> None:
    """배정 결과를 JSON 파일로 저장한다.

    *expected_mtime* 이 주어지면 저장 전에 파일의 현재 mtime과 비교하여
    외부 변경이 감지되면 :class:`StaleFileError` 를 발생시킨다.
    """
    if expected_mtime is not None and path.exists():
        current_mtime = path.stat().st_mtime
        if abs(current_mtime - expected_mtime) > 0.01:
            raise StaleFileError(f"파일이 외부에서 변경되었습니다: {path.name}")
    path.parent.mkdir(parents=True, exist_ok=True)
    text = json.dumps(assignments, ensure_ascii=False, indent=2)
    # 원자적 저장: 임시 파일 → rename
    fd, tmp = tempfile.mkstemp(dir=path.parent, suffix=".tmp")
    try:
        with open(fd, "w", encoding="utf-8") as f:
            f.write(text)
        Path(tmp).replace(path)
    except BaseException:
        Path(tmp).unlink(missing_ok=True)
        raise


def load_releases(path: Path) -> dict:
    """저장된 강의실 해제 데이터를 읽어 dict로 반환한다."""
    if not path.exists():
        return {}
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    if not isinstance(raw, dict):
        return {}

    normalized = {}
    for key, value in raw.items():
        if not isinstance(key, str) or not isinstance(value, dict):
            continue
        periods = value.get("periods", [])
        if not isinstance(periods, list):
            periods = []
        safe_periods = [_safe_int(p, default=-1) for p in periods]
        safe_periods = sorted({p for p in safe_periods if MIN_PERIOD <= p <= MAX_PERIOD})
        if not safe_periods:
            continue
        normalized[key] = {
            "sheet": str(value.get("sheet", "")),
            "room": str(value.get("room", "")),
            "periods": safe_periods,
        }
    return normalized


def save_releases(path: Path, releases: dict, expected_mtime: float | None = None) -> None:
    """강의실 해제 데이터를 JSON 파일로 저장한다.

    *expected_mtime* 이 주어지면 저장 전에 파일의 현재 mtime과 비교하여
    외부 변경이 감지되면 :class:`StaleFileError` 를 발생시킨다.
    """
    if expected_mtime is not None and path.exists():
        current_mtime = path.stat().st_mtime
        if abs(current_mtime - expected_mtime) > 0.01:
            raise StaleFileError(f"파일이 외부에서 변경되었습니다: {path.name}")
    path.parent.mkdir(parents=True, exist_ok=True)
    text = json.dumps(releases, ensure_ascii=False, indent=2)
    fd, tmp = tempfile.mkstemp(dir=path.parent, suffix=".tmp")
    try:
        with open(fd, "w", encoding="utf-8") as f:
            f.write(text)
        Path(tmp).replace(path)
    except BaseException:
        Path(tmp).unlink(missing_ok=True)
        raise


def releases_to_slot_set(releases: dict) -> set[tuple[str, str, int]]:
    """해제 dict를 {(sheet, room, period)} set으로 변환한다."""
    result = set()
    for value in releases.values():
        sheet = value.get("sheet", "")
        room = value.get("room", "")
        for p in value.get("periods", []):
            result.add((sheet, room, p))
    return result


def append_audit_event(
    path: Path,
    operator: str,
    action: str,
    subject: str = "",
    details: dict | None = None,
) -> None:
    """작업 이력 이벤트를 JSONL 한 줄로 기록한다."""
    event = {
        "timestamp": datetime.datetime.now(datetime.timezone.utc).isoformat(),
        "operator": operator or "unknown",
        "action": action,
        "subject": subject,
        "details": details or {},
    }
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as f:
        f.write(json.dumps(event, ensure_ascii=False) + "\n")


def read_audit_events(path: Path, limit: int = 100) -> list[dict[str, Any]]:
    """작업 이력을 최신순으로 읽는다."""
    if not path.exists():
        return []

    events: list[dict[str, Any]] = []
    try:
        lines = path.read_text(encoding="utf-8").splitlines()
    except OSError:
        return []

    for line in lines:
        line = line.strip()
        if not line:
            continue
        try:
            row = json.loads(line)
        except json.JSONDecodeError:
            continue
        if isinstance(row, dict):
            events.append(row)

    if limit > 0:
        events = events[-limit:]
    events.reverse()
    return events
