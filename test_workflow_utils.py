"""workflow_utils 단위 테스트."""

import datetime
import json

from data_loader import ExamRequest, ScheduleSlot
from workflow_utils import (
    MIN_PERIOD,
    MAX_PERIOD,
    _clamp_periods,
    append_audit_event,
    load_assignments,
    load_releases,
    read_audit_events,
    releases_to_slot_set,
    resolve_needed_periods,
    save_assignments,
    save_releases,
)


def _make_req(**kwargs):
    defaults = dict(
        row=1,
        department="테스트",
        name="테스트과목",
        ban="01",
        professor="교수",
        students=30,
        schedule_raw="화2~3(K106)",
        slots=[ScheduleSlot("화", 2, 3, "K106")],
        room="K106",
        exam_date=datetime.date(2026, 4, 21),
        exam_start=2,
        exam_end=3,
        room_choice="기존 강의실",
        remarks="",
    )
    defaults.update(kwargs)
    return ExamRequest(**defaults)


def test_resolve_needed_periods_from_exam_range():
    req = _make_req(exam_start=5, exam_end=3)
    assert resolve_needed_periods(req, "화") == [3, 4, 5]


def test_resolve_needed_periods_from_slots_unique_sorted():
    req = _make_req(
        exam_start=None,
        exam_end=None,
        slots=[
            ScheduleSlot("화", 4, 5, "K106"),
            ScheduleSlot("화", 5, 6, "K106"),
            ScheduleSlot("수", 1, 2, "K106"),
        ],
    )
    assert resolve_needed_periods(req, "화") == [4, 5, 6]


def test_load_assignments_missing_and_invalid(tmp_path):
    missing = tmp_path / "missing.json"
    assert load_assignments(missing) == {}

    invalid = tmp_path / "invalid.json"
    invalid.write_text("{not-json", encoding="utf-8")
    assert load_assignments(invalid) == {}


def test_save_and_load_assignments_roundtrip(tmp_path):
    path = tmp_path / "assignments.json"
    payload = {
        "테스트과목-01": {
            "room": "K106",
            "sheet": "4.21.(화)",
            "periods": [2, 3],
            "original_room": "K106",
            "students": 30,
            "category": "강의실 변경",
            "keep_orig": True,
        }
    }
    save_assignments(path, payload)
    assert load_assignments(path) == payload


def test_load_assignments_sanitizes_payload(tmp_path):
    path = tmp_path / "broken_assignments.json"
    broken = {
        "과목A-01": {
            "room": "K101",
            "sheet": "4.21.(화)",
            "periods": ["2", "x", -1, 2],
            "students": "41",
            "category": "강의실 변경",
        },
        100: "invalid-row",
    }
    path.write_text(json.dumps(broken, ensure_ascii=False), encoding="utf-8")

    rows = load_assignments(path)
    assert list(rows.keys()) == ["과목A-01"]
    assert rows["과목A-01"]["periods"] == [2]
    assert rows["과목A-01"]["students"] == 41


def test_audit_append_and_read_latest_first(tmp_path):
    path = tmp_path / "audit.jsonl"
    append_audit_event(path, "alice", "assign", "과목A-01", {"room": "K101"})
    append_audit_event(path, "bob", "clear_all", details={"count": 1})

    rows = read_audit_events(path, limit=10)
    assert len(rows) == 2
    assert rows[0]["operator"] == "bob"
    assert rows[0]["action"] == "clear_all"
    assert rows[1]["subject"] == "과목A-01"


# ── 교시 범위 검증 회귀 테스트 ──

def test_clamp_periods_within_range():
    assert _clamp_periods([0, 7, 14]) == [0, 7, 14]


def test_clamp_periods_out_of_range():
    assert _clamp_periods([-2, -1, 0, 14, 15, 20]) == [0, 14]


def test_clamp_periods_empty():
    assert _clamp_periods([]) == []


def test_resolve_needed_periods_clamps_exam_range():
    """exam_start/exam_end가 0~14 밖이면 클램프되어야 한다."""
    req = _make_req(exam_start=-3, exam_end=20)
    result = resolve_needed_periods(req, "화")
    assert result[0] >= MIN_PERIOD
    assert result[-1] <= MAX_PERIOD
    assert result == list(range(0, 15))


def test_resolve_needed_periods_clamps_slot_range():
    """슬롯 교시가 범위 밖이면 클램프되어야 한다."""
    req = _make_req(
        exam_start=None, exam_end=None,
        slots=[ScheduleSlot("화", -1, 16, "K106")],
    )
    result = resolve_needed_periods(req, "화")
    assert result[0] >= MIN_PERIOD
    assert result[-1] <= MAX_PERIOD


def test_load_assignments_clamps_periods(tmp_path):
    """저장본에 범위 밖 교시가 있으면 로딩 시 제거되어야 한다."""
    path = tmp_path / "assignments.json"
    payload = {
        "과목A-01": {
            "room": "K101",
            "sheet": "4.21.(화)",
            "periods": [2, 20, -1],
            "original_room": "K101",
            "students": 30,
            "category": "강의실 변경",
        }
    }
    path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    result = load_assignments(path)
    assert result["과목A-01"]["periods"] == [2]


# ── 강의실 해제 회귀 테스트 ──

def test_load_releases_missing(tmp_path):
    assert load_releases(tmp_path / "missing.json") == {}


def test_save_and_load_releases_roundtrip(tmp_path):
    path = tmp_path / "releases.json"
    payload = {
        "4.23.(목)|N405": {"sheet": "4.23.(목)", "room": "N405", "periods": [8, 9]}
    }
    save_releases(path, payload)
    assert load_releases(path) == payload


def test_load_releases_clamps_periods(tmp_path):
    path = tmp_path / "releases.json"
    payload = {
        "4.21.(화)|K101": {"sheet": "4.21.(화)", "room": "K101", "periods": [-1, 5, 20]}
    }
    path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    result = load_releases(path)
    assert result["4.21.(화)|K101"]["periods"] == [5]


def test_load_releases_skips_empty_periods(tmp_path):
    path = tmp_path / "releases.json"
    payload = {
        "4.21.(화)|K101": {"sheet": "4.21.(화)", "room": "K101", "periods": [-1, 20]}
    }
    path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    result = load_releases(path)
    assert result == {}


def test_load_assignments_skips_empty_periods(tmp_path):
    """periods가 정규화 후 비어 있으면 해당 엔트리를 제거해야 한다."""
    path = tmp_path / "assignments.json"
    payload = {
        "과목A-01": {
            "room": "K101",
            "sheet": "4.21.(화)",
            "periods": [],
            "original_room": "K101",
            "students": 30,
            "category": "강의실 변경",
        },
        "과목B-01": {
            "room": "K102",
            "sheet": "4.21.(화)",
            "periods": [2, 3],
            "original_room": "K102",
            "students": 20,
            "category": "강의실 변경",
        },
    }
    path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    result = load_assignments(path)
    assert "과목A-01" not in result
    assert "과목B-01" in result
    assert result["과목B-01"]["periods"] == [2, 3]


def test_releases_to_slot_set():
    releases = {
        "4.23.(목)|N405": {"sheet": "4.23.(목)", "room": "N405", "periods": [8, 9]},
        "4.21.(화)|K101": {"sheet": "4.21.(화)", "room": "K101", "periods": [3]},
    }
    result = releases_to_slot_set(releases)
    assert result == {
        ("4.23.(목)", "N405", 8),
        ("4.23.(목)", "N405", 9),
        ("4.21.(화)", "K101", 3),
    }
