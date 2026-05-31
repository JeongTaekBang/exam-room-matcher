"""assignment_status.compute_status 단위 테스트.

dashboard.py에서 분리한 핵심 로직. 분반 완료 규칙·수용인원 검증·
ROOM_CHANGE 단일 배정 등 복합 분기를 모두 커버.
"""

import datetime

from data_loader import Category, ExamRequest, ScheduleSlot
from assignment_status import (
    CAT_LABELS,
    LABEL_ROOM_CHANGE,
    LABEL_ROOM_SPLIT,
    LABEL_NORMAL_EXAM,
    _room_cap,
    compute_auto_released,
    compute_status,
    extract_base_key,
    resolve_processed_room,
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


# ──────────────────────────────────────────────
# 자동 완료 분기 (NORMAL_EXAM / NO_EXAM)
# ──────────────────────────────────────────────

class TestAutoComplete:
    def test_normal_exam_is_auto_done(self):
        req = _make_req()
        req.category = Category.NORMAL_EXAM
        assert compute_status(req, {}, None) == "완료"

    def test_no_exam_is_auto_done(self):
        req = _make_req()
        req.category = Category.NO_EXAM
        assert compute_status(req, {}, None) == "완료"

    def test_skip_returns_mihwakjong(self):
        req = _make_req()
        req.category = Category.SKIP
        assert compute_status(req, {}, None) == "미확정"


# ──────────────────────────────────────────────
# ROOM_CHANGE 분기
# ──────────────────────────────────────────────

class TestRoomChange:
    def test_no_assignment_is_unassigned(self):
        req = _make_req(students=30)
        req.category = Category.ROOM_CHANGE
        assert compute_status(req, {}, {"A": 100}) == "미배정"

    def test_capacity_sufficient_is_done(self):
        req = _make_req(students=30)
        req.category = Category.ROOM_CHANGE
        asgn = {req.key: {"room": "A", "category": LABEL_ROOM_CHANGE}}
        assert compute_status(req, asgn, {"A": 100}) == "완료"

    def test_capacity_short_is_unassigned(self):
        """ROOM_CHANGE 단일 배정에서 수용인원 부족 시 '미배정' 반환 (B2 회귀 테스트)."""
        req = _make_req(students=78)
        req.category = Category.ROOM_CHANGE
        asgn = {req.key: {"room": "V111", "category": LABEL_ROOM_CHANGE}}
        assert compute_status(req, asgn, {"V111": 60}) == "미배정"

    def test_no_room_capacity_dict_is_done(self):
        """room_capacity가 None이면 보수적으로 '완료' (검증 스킵)."""
        req = _make_req(students=30)
        req.category = Category.ROOM_CHANGE
        asgn = {req.key: {"room": "A", "category": LABEL_ROOM_CHANGE}}
        assert compute_status(req, asgn, None) == "완료"

    def test_unknown_room_capacity_is_unassigned(self):
        """배정된 강의실이 room_capacity에 없으면 수용인원 0으로 처리 → 미배정."""
        req = _make_req(students=30)
        req.category = Category.ROOM_CHANGE
        asgn = {req.key: {"room": "UNKNOWN", "category": LABEL_ROOM_CHANGE}}
        assert compute_status(req, asgn, {"A": 100}) == "미배정"


# ──────────────────────────────────────────────
# ROOM_SPLIT 분기
# ──────────────────────────────────────────────

class TestRoomSplit:
    def test_no_assignment_is_unassigned(self):
        req = _make_req(students=100)
        req.category = Category.ROOM_SPLIT
        assert compute_status(req, {}, {"A": 60, "B": 60}) == "미배정"

    def test_keep_orig_one_assignment_is_done(self):
        """기존 강의실 유지 분반 — 추가 강의실 1개라도 있으면 '완료'."""
        req = _make_req(students=100)
        req.category = Category.ROOM_SPLIT
        asgn = {
            f"{req.key}+1": {
                "room": "B", "category": LABEL_ROOM_SPLIT, "keep_orig": True,
            },
        }
        assert compute_status(req, asgn, {"A": 60, "B": 50}) == "완료"

    def test_no_keep_orig_capacity_sum_sufficient(self):
        """기존 미유지 분반 — 배정 강의실 합이 학생 이상이면 '완료'."""
        req = _make_req(students=100)
        req.category = Category.ROOM_SPLIT
        asgn = {
            f"{req.key}+1": {
                "room": "B", "category": LABEL_ROOM_SPLIT, "keep_orig": False,
            },
            f"{req.key}+2": {
                "room": "C", "category": LABEL_ROOM_SPLIT, "keep_orig": False,
            },
        }
        assert compute_status(req, asgn, {"B": 60, "C": 50}) == "완료"

    def test_no_keep_orig_capacity_sum_short(self):
        """기존 미유지 분반 — 합이 학생 미만이면 '미배정'."""
        req = _make_req(students=200)
        req.category = Category.ROOM_SPLIT
        asgn = {
            f"{req.key}+1": {
                "room": "B", "category": LABEL_ROOM_SPLIT, "keep_orig": False,
            },
            f"{req.key}+2": {
                "room": "C", "category": LABEL_ROOM_SPLIT, "keep_orig": False,
            },
        }
        assert compute_status(req, asgn, {"B": 60, "C": 50}) == "미배정"

    def test_no_keep_orig_same_room_twice_no_double_count(self):
        """기존 미유지 분반 — 같은 강의실이 여러 키로 등록되어도 한 번만 합산."""
        req = _make_req(students=80)
        req.category = Category.ROOM_SPLIT
        asgn = {
            f"{req.key}+1": {
                "room": "B", "category": LABEL_ROOM_SPLIT, "keep_orig": False,
            },
            f"{req.key}+2": {
                "room": "B", "category": LABEL_ROOM_SPLIT, "keep_orig": False,
            },
        }
        # B 60명 한 번만 합산 → 80명 미만 → 미배정
        assert compute_status(req, asgn, {"B": 60}) == "미배정"

    def test_no_keep_orig_no_capacity_dict(self):
        """기존 미유지 분반 + room_capacity None — 최소 2건이면 '완료'."""
        req = _make_req(students=100)
        req.category = Category.ROOM_SPLIT
        asgn = {
            f"{req.key}+1": {
                "room": "B", "category": LABEL_ROOM_SPLIT, "keep_orig": False,
            },
            f"{req.key}+2": {
                "room": "C", "category": LABEL_ROOM_SPLIT, "keep_orig": False,
            },
        }
        assert compute_status(req, asgn, None) == "완료"

    def test_room_change_category_with_split_storage(self):
        """ROOM_CHANGE 요청이지만 분반으로 저장된 경우 분반 규칙 적용."""
        req = _make_req(students=100)
        req.category = Category.ROOM_CHANGE
        asgn = {
            f"{req.key}+1": {
                "room": "B", "category": LABEL_ROOM_SPLIT, "keep_orig": True,
            },
        }
        # is_split_mode가 True (저장이 분반이라) → keeps_original=True → 완료
        assert compute_status(req, asgn, {"B": 50}) == "완료"


# ──────────────────────────────────────────────
# 키 매칭
# ──────────────────────────────────────────────

class TestKeyMatching:
    def test_split_key_with_plus_suffix(self):
        """+1, +2 분반 키도 base 키로 매칭."""
        req = _make_req(students=80)
        req.category = Category.ROOM_SPLIT
        asgn = {
            f"{req.key}+1": {
                "room": "A", "category": LABEL_ROOM_SPLIT, "keep_orig": False,
            },
            "다른과목-01": {
                "room": "B", "category": LABEL_ROOM_CHANGE,
            },
        }
        # 자기 키만 매칭, 다른과목-01은 매칭 안 됨
        assert compute_status(req, asgn, {"A": 100, "B": 100}) == "완료"


# ──────────────────────────────────────────────
# _room_cap 헬퍼
# ──────────────────────────────────────────────

class TestRoomCap:
    def test_none_dict_returns_zero(self):
        assert _room_cap(None, "A") == 0

    def test_empty_room_returns_zero(self):
        assert _room_cap({"A": 50}, "") == 0

    def test_missing_room_returns_zero(self):
        assert _room_cap({"A": 50}, "B") == 0

    def test_int_value(self):
        assert _room_cap({"A": 50}, "A") == 50

    def test_string_value_parsed(self):
        assert _room_cap({"A": "50"}, "A") == 50

    def test_invalid_value_returns_zero(self):
        assert _room_cap({"A": "abc"}, "A") == 0

    def test_none_value_returns_zero(self):
        assert _room_cap({"A": None}, "A") == 0


# ──────────────────────────────────────────────
# CAT_LABELS 일관성
# ──────────────────────────────────────────────

class TestCatLabels:
    def test_all_categories_have_label(self):
        for cat in Category:
            assert cat in CAT_LABELS
            assert isinstance(CAT_LABELS[cat], str)
            assert CAT_LABELS[cat]  # 빈 문자열 금지

    def test_label_constants_match_dict(self):
        assert LABEL_ROOM_SPLIT == CAT_LABELS[Category.ROOM_SPLIT]
        assert LABEL_ROOM_CHANGE == CAT_LABELS[Category.ROOM_CHANGE]
        assert LABEL_NORMAL_EXAM == CAT_LABELS[Category.NORMAL_EXAM]


# ──────────────────────────────────────────────
# extract_base_key
# ──────────────────────────────────────────────

class TestExtractBaseKey:
    def test_no_split_suffix(self):
        assert extract_base_key("영문학의이해-01") == "영문학의이해-01"

    def test_with_split_suffix(self):
        assert extract_base_key("영문학의이해-01+2") == "영문학의이해-01"
        assert extract_base_key("영문학의이해-01+10") == "영문학의이해-01"

    def test_plus_in_subject_name_no_split(self):
        """과목명에 +가 있지만 분반 suffix 없는 경우."""
        assert extract_base_key("C++-01") == "C++-01"
        assert extract_base_key("A+B+C") == "A+B+C"

    def test_plus_in_subject_name_with_split(self):
        """과목명에 +가 있고 분반 suffix도 있는 경우 — 마지막 + 다음 정수만 분리."""
        assert extract_base_key("C++-01+3") == "C++-01"
        assert extract_base_key("A+B-01+5") == "A+B-01"

    def test_plus_non_digit_suffix(self):
        """+ 다음이 정수가 아니면 base로 인식하지 않음."""
        assert extract_base_key("과목-01+abc") == "과목-01+abc"
        assert extract_base_key("과목-01+") == "과목-01+"

    def test_with_row_dedup_suffix(self):
        """중복 감지 #row suffix + 분반 suffix가 함께 있는 경우."""
        assert extract_base_key("과목-01#123") == "과목-01#123"
        assert extract_base_key("과목-01#123+2") == "과목-01#123"


# ──────────────────────────────────────────────
# compute_auto_released
# ──────────────────────────────────────────────

class TestComputeAutoReleased:
    """자동 해제 도출 — 세 가지 파생 규칙."""

    DATE = datetime.date(2026, 4, 21)
    D2D = {DATE: "화"}
    D2S = {DATE: "4.21.(화)"}
    DAY_TO_SHEETS = {"화": ["4.21.(화)"]}

    def _req(self, category, **kw):
        defaults = dict(
            row=1, department="D", name="N", ban="01", professor="P",
            students=30, schedule_raw="", slots=[], room="",
            exam_date=self.DATE, exam_start=2, exam_end=3,
            room_choice=None, remarks="",
        )
        defaults.update(kw)
        r = ExamRequest(**defaults)
        r.category = category
        return r

    # 규칙 1: ROOM_CHANGE 배정 → 원래 강의실 해제
    def test_room_change_releases_original(self):
        asgn = {
            "과목-01": {
                "room": "B", "sheet": "4.21.(화)", "periods": [2, 3],
                "original_room": "A", "category": LABEL_ROOM_CHANGE,
            }
        }
        result = compute_auto_released(asgn, [], self.D2D, self.D2S, self.DAY_TO_SHEETS)
        assert result == {
            ("4.21.(화)", "A", 2): ("과목-01", "이동"),
            ("4.21.(화)", "A", 3): ("과목-01", "이동"),
        }

    def test_room_change_same_room_no_release(self):
        """같은 강의실로 배정된 경우 해제하지 않음."""
        asgn = {
            "과목-01": {
                "room": "A", "sheet": "4.21.(화)", "periods": [2, 3],
                "original_room": "A", "category": LABEL_ROOM_CHANGE,
            }
        }
        result = compute_auto_released(asgn, [], self.D2D, self.D2S, self.DAY_TO_SHEETS)
        assert result == {}

    # 규칙 1: 분반(keep_orig=False) → 원래 강의실 해제
    def test_split_no_keep_releases_original(self):
        asgn = {
            "과목-01+1": {
                "room": "B", "sheet": "4.21.(화)", "periods": [2],
                "original_room": "A", "category": LABEL_ROOM_SPLIT,
                "keep_orig": False,
            }
        }
        result = compute_auto_released(asgn, [], self.D2D, self.D2S, self.DAY_TO_SHEETS)
        assert result == {("4.21.(화)", "A", 2): ("과목-01+1", "이동")}

    def test_split_keep_orig_no_release(self):
        """기존 강의실 유지 분반은 원래 강의실 해제하지 않음."""
        asgn = {
            "과목-01+1": {
                "room": "B", "sheet": "4.21.(화)", "periods": [2],
                "original_room": "A", "category": LABEL_ROOM_SPLIT,
                "keep_orig": True,
            }
        }
        result = compute_auto_released(asgn, [], self.D2D, self.D2S, self.DAY_TO_SHEETS)
        assert result == {}

    # 규칙 2: NO_EXAM + exam_date 있음 → 시험일 강의실 해제
    def test_no_exam_with_date(self):
        req = self._req(
            Category.NO_EXAM,
            slots=[ScheduleSlot("화", 2, 3, "A")],
            room="A",
        )
        result = compute_auto_released({}, [req], self.D2D, self.D2S, self.DAY_TO_SHEETS)
        assert result == {
            ("4.21.(화)", "A", 2): (req.key, "미실시"),
            ("4.21.(화)", "A", 3): (req.key, "미실시"),
        }

    # 규칙 2: NO_EXAM + exam_date 없음 + 다주차 → 같은 요일 모든 시트
    def test_no_exam_no_date_multi_week(self):
        """exam_date가 없으면 day_to_sheets의 모든 시트에 적용."""
        req = self._req(
            Category.NO_EXAM,
            exam_date=None, exam_start=None, exam_end=None,
            slots=[ScheduleSlot("화", 2, 3, "A")],
        )
        day_to_sheets = {"화": ["4.15.(화)", "4.22.(화)"]}
        result = compute_auto_released({}, [req], {}, {}, day_to_sheets)
        assert result == {
            ("4.15.(화)", "A", 2): (req.key, "미실시"),
            ("4.15.(화)", "A", 3): (req.key, "미실시"),
            ("4.22.(화)", "A", 2): (req.key, "미실시"),
            ("4.22.(화)", "A", 3): (req.key, "미실시"),
        }

    # 규칙 3: NORMAL_EXAM + 시험교시 < 수업교시 → 미사용 교시 부분해제
    def test_normal_exam_partial_release(self):
        """수업 2~5교시, 시험 2~3교시 → 4·5교시 부분해제."""
        req = self._req(
            Category.NORMAL_EXAM,
            slots=[ScheduleSlot("화", 2, 5, "A")],
            exam_start=2, exam_end=3,
        )
        result = compute_auto_released({}, [req], self.D2D, self.D2S, self.DAY_TO_SHEETS)
        assert result == {
            ("4.21.(화)", "A", 4): (req.key, "부분해제"),
            ("4.21.(화)", "A", 5): (req.key, "부분해제"),
        }

    def test_normal_exam_no_partial_when_exam_covers_class(self):
        """시험 교시 = 수업 교시면 부분해제 없음."""
        req = self._req(
            Category.NORMAL_EXAM,
            slots=[ScheduleSlot("화", 2, 3, "A")],
            exam_start=2, exam_end=3,
        )
        result = compute_auto_released({}, [req], self.D2D, self.D2S, self.DAY_TO_SHEETS)
        assert result == {}

    def test_normal_exam_no_exam_start_no_release(self):
        """시험 시작교시 미명시면 부분해제 안 함."""
        req = self._req(
            Category.NORMAL_EXAM,
            slots=[ScheduleSlot("화", 2, 5, "A")],
            exam_start=None, exam_end=None,
        )
        result = compute_auto_released({}, [req], self.D2D, self.D2S, self.DAY_TO_SHEETS)
        assert result == {}

    # 카테고리 외 무시
    def test_skip_category_ignored(self):
        req = self._req(Category.SKIP)
        result = compute_auto_released({}, [req], self.D2D, self.D2S, self.DAY_TO_SHEETS)
        assert result == {}


class TestResolveProcessedRoom:
    """P열(요청사항 처리) 채움용 강의실 결정 문자열."""

    def test_normal_exam_returns_room(self):
        req = _make_req(room="K106")
        req.category = Category.NORMAL_EXAM
        assert resolve_processed_room(req, {}) == "K106"

    def test_no_exam_returns_marker(self):
        req = _make_req()
        req.category = Category.NO_EXAM
        assert resolve_processed_room(req, {}) == "강의실 미사용"

    def test_room_change_assigned(self):
        req = _make_req(name="가", ban="01")  # key = "가-01"
        req.category = Category.ROOM_CHANGE
        assert resolve_processed_room(req, {"가-01": {"room": "N212"}}) == "N212"

    def test_room_change_unassigned_blank(self):
        req = _make_req()
        req.category = Category.ROOM_CHANGE
        assert resolve_processed_room(req, {}) == ""

    def test_room_split_joins_rooms(self):
        req = _make_req(name="가", ban="01")
        req.category = Category.ROOM_SPLIT
        assigns = {"가-01": {"room": "N210"}, "가-01+2": {"room": "N405"}}
        assert resolve_processed_room(req, assigns) == "N210,N405"

    def test_split_dedup_rooms(self):
        req = _make_req(name="가", ban="01")
        req.category = Category.ROOM_SPLIT
        assigns = {"가-01": {"room": "N210"}, "가-01+2": {"room": "N210"}}
        assert resolve_processed_room(req, assigns) == "N210"

    def test_skip_blank(self):
        req = _make_req()
        req.category = Category.SKIP
        assert resolve_processed_room(req, {}) == ""

    def test_assignment_overrides_category_default(self):
        # 배정이 있으면 카테고리 기본값(NORMAL_EXAM의 req.room)보다 우선
        req = _make_req(name="가", ban="01", room="K106")
        req.category = Category.NORMAL_EXAM
        assert resolve_processed_room(req, {"가-01": {"room": "N999"}}) == "N999"
