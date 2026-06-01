"""conflict_check 단위 테스트 — 원본 '요청사항 처리' 강의실 중복 점검."""

import datetime

from data_loader import Category, ExamRequest, ScheduleSlot, normalize_ban
from conflict_check import (
    parse_processed_rooms,
    detect_manual_processed_entries,
    detect_manual_no_use_entries,
    detect_processed_room_conflicts,
    detect_timetable_overlaps,
    _occupant_base,
    _parse_occupant,
    NON_ROOM_MARKERS,
)

D = datetime.date(2026, 6, 17)
D2 = datetime.date(2026, 6, 18)

SHEET = "6.17.(화)"
DTS = {D: SHEET}      # date_to_sheet
DTD = {D: "화"}        # date_to_day


def _make_req(row=1, name="가과목", ban="01", exam_date=D,
              exam_start=2, exam_end=3, processed_room="N212"):
    return ExamRequest(
        row=row, department="학부", name=name, ban=ban, professor="교수",
        students=30, schedule_raw="화2~3(N212)",
        slots=[ScheduleSlot("화", 2, 3, "N212")], room="N212",
        exam_date=exam_date, exam_start=exam_start, exam_end=exam_end,
        room_choice="기존 강의실", remarks="", processed_room=processed_room,
    )


class TestParseProcessedRooms:
    def test_single_room(self):
        assert parse_processed_rooms("N212") == ["N212"]

    def test_multi_room_comma(self):
        assert parse_processed_rooms("N210,N405") == ["N210", "N405"]
        assert parse_processed_rooms("N210, N405") == ["N210", "N405"]

    def test_non_room_markers_excluded(self):
        assert parse_processed_rooms("강의실 미사용") == []
        assert parse_processed_rooms("강의실미사용") == []
        assert parse_processed_rooms("  강의실  미사용 ") == []
        assert parse_processed_rooms("확인필요") == []
        assert parse_processed_rooms("") == []
        assert parse_processed_rooms(None) == []

    def test_gym_is_a_room(self):
        # 체육관/테니스장/헬스장은 실제 시험 장소 — 제외하지 않는다
        assert parse_processed_rooms("테니스장") == ["테니스장"]
        assert parse_processed_rooms("체육관") == ["체육관"]

    def test_mixed_room_and_marker(self):
        assert parse_processed_rooms("N210,확인필요") == ["N210"]

    def test_markers_set_uses_whitespace_stripped_form(self):
        assert "강의실미사용" in NON_ROOM_MARKERS
        assert "" in NON_ROOM_MARKERS


class TestDetectConflicts:
    def test_overlap_same_room_diff_course(self):
        # 일반화학및실험 6~7교시 vs 일반수학 5~6교시, 같은 날·N212 → 6교시 겹침
        a = _make_req(row=127, name="일반화학및실험", exam_start=6, exam_end=7)
        b = _make_req(row=128, name="일반수학1및연습", exam_start=5, exam_end=6)
        conflicts, unjudged = detect_processed_room_conflicts([a, b])
        assert len(conflicts) == 1
        assert unjudged == []
        c = conflicts[0]
        assert c.room == "N212"
        assert c.exam_date == D
        assert {c.req_a.name, c.req_b.name} == {"일반화학및실험", "일반수학1및연습"}

    def test_no_overlap_sequential_use(self):
        # 같은 방·같은 날이지만 교시가 안 겹침(순차 사용) → 충돌 아님
        a = _make_req(row=1, name="가과목", exam_start=1, exam_end=3)
        b = _make_req(row=2, name="나과목", exam_start=4, exam_end=5)
        conflicts, _ = detect_processed_room_conflicts([a, b])
        assert conflicts == []

    def test_same_course_sibling_not_conflict(self):
        # 같은 교과목명 분반끼리 겹쳐도 충돌 아님
        a = _make_req(row=1, name="인간학1", ban="05", exam_start=2, exam_end=3)
        b = _make_req(row=2, name="인간학1", ban="09", exam_start=2, exam_end=3)
        conflicts, _ = detect_processed_room_conflicts([a, b])
        assert conflicts == []

    def test_diff_room_no_conflict(self):
        a = _make_req(row=1, name="가과목", processed_room="N212")
        b = _make_req(row=2, name="나과목", processed_room="N210")
        conflicts, _ = detect_processed_room_conflicts([a, b])
        assert conflicts == []

    def test_diff_date_no_conflict(self):
        a = _make_req(row=1, name="가과목", exam_date=D)
        b = _make_req(row=2, name="나과목", exam_date=D2)
        conflicts, _ = detect_processed_room_conflicts([a, b])
        assert conflicts == []

    def test_multi_room_conflict_on_shared_room_only(self):
        # A는 N210,N405 / B는 N405 → N405에서만 충돌(1건)
        a = _make_req(row=1, name="가과목", processed_room="N210,N405",
                      exam_start=2, exam_end=3)
        b = _make_req(row=2, name="나과목", processed_room="N405",
                      exam_start=2, exam_end=3)
        conflicts, _ = detect_processed_room_conflicts([a, b])
        assert len(conflicts) == 1
        assert conflicts[0].room == "N405"

    def test_non_room_value_no_booking(self):
        a = _make_req(row=1, name="가과목", processed_room="강의실 미사용")
        b = _make_req(row=2, name="나과목", processed_room="강의실 미사용")
        conflicts, unjudged = detect_processed_room_conflicts([a, b])
        assert conflicts == []
        assert unjudged == []

    def test_missing_date_goes_unjudged(self):
        a = _make_req(row=1, name="시창1", exam_date=None,
                      exam_start=None, exam_end=None, processed_room="CH202")
        conflicts, unjudged = detect_processed_room_conflicts([a])
        assert conflicts == []
        assert len(unjudged) == 1
        assert unjudged[0].room == "CH202"
        assert "시험일자" in unjudged[0].reason

    def test_missing_period_goes_unjudged(self):
        # 날짜는 있으나 교시 누락 → unjudged
        a = _make_req(row=1, name="가과목", exam_start=None, exam_end=None)
        b = _make_req(row=2, name="나과목", exam_start=None, exam_end=None)
        conflicts, unjudged = detect_processed_room_conflicts([a, b])
        assert conflicts == []
        assert len(unjudged) == 2
        assert all("교시" in u.reason for u in unjudged)

    def test_boundary_touching_periods_overlap(self):
        # (6,7) 과 (5,6) 은 6교시를 공유 → 겹침
        a = _make_req(row=1, name="가과목", exam_start=6, exam_end=7)
        b = _make_req(row=2, name="나과목", exam_start=5, exam_end=6)
        conflicts, _ = detect_processed_room_conflicts([a, b])
        assert len(conflicts) == 1

    def test_reversed_period_order_handled(self):
        # 시작>종료로 뒤집혀 들어와도 정렬해 비교
        a = _make_req(row=1, name="가과목", exam_start=7, exam_end=6)
        b = _make_req(row=2, name="나과목", exam_start=6, exam_end=5)
        conflicts, _ = detect_processed_room_conflicts([a, b])
        assert len(conflicts) == 1

    def test_three_way_same_slot_produces_pairs(self):
        # 같은 방·같은 교시에 서로 다른 3과목 → 3쌍(C(3,2)) 충돌
        reqs = [
            _make_req(row=1, name="가", exam_start=2, exam_end=3),
            _make_req(row=2, name="나", exam_start=2, exam_end=3),
            _make_req(row=3, name="다", exam_start=2, exam_end=3),
        ]
        conflicts, _ = detect_processed_room_conflicts(reqs)
        assert len(conflicts) == 3

    def test_sorted_output(self):
        # 서로 다른 강의실 충돌 2건이 강의실명 순으로 정렬되는지
        reqs = [
            _make_req(row=10, name="가", processed_room="N212", exam_start=2, exam_end=3),
            _make_req(row=11, name="나", processed_room="N212", exam_start=2, exam_end=3),
            _make_req(row=12, name="다", processed_room="B101", exam_start=2, exam_end=3),
            _make_req(row=13, name="라", processed_room="B101", exam_start=2, exam_end=3),
        ]
        conflicts, _ = detect_processed_room_conflicts(reqs)
        assert len(conflicts) == 2
        assert [c.room for c in conflicts] == ["B101", "N212"]


class TestOccupantBase:
    def test_strips_section(self):
        assert _occupant_base("한국어어휘론-01") == "한국어어휘론"
        assert _occupant_base("I-DESIGN-15") == "I-DESIGN"   # 과목명에 - 포함

    def test_no_section_kept(self):
        assert _occupant_base("I-DESIGN") == "I-DESIGN"      # - 다음이 숫자 아님
        assert _occupant_base("특수예약") == "특수예약"
        assert _occupant_base("") == ""


class TestParseOccupant:
    def test_hyphen_ban(self):
        assert _parse_occupant("한국어어휘론-01") == ("한국어어휘론", "01")
        assert _parse_occupant("I-DESIGN-15") == ("I-DESIGN", "15")
        assert _parse_occupant("마케팅-P1") == ("마케팅", "P1")

    def test_no_hyphen_ban(self):
        assert _parse_occupant("한국어학술작문01") == ("한국어학술작문", "01")

    def test_exam_prefix(self):
        assert _parse_occupant("시험(원래): 뇌와인지-01") == ("뇌와인지", "01")

    def test_name_with_trailing_digit_kept(self):
        # 과목명 끝의 숫자는 분반이 아니라 이름의 일부 — 하이픈 분반만 제거
        assert _parse_occupant("일반생물학1-04") == ("일반생물학1", "04")

    def test_reservation_text_unparsed(self):
        assert _parse_occupant("06-16 [입학팀] 정기회의") == ("06-16 [입학팀] 정기회의", "")
        assert _parse_occupant("특수예약") == ("특수예약", "")


class TestTimetableOverlaps:
    def test_basic_overlap_diff_course(self):
        # 법학개론 시험을 K262(자기 방 N212 아님)에 둠 → 그 방엔 한국어어휘론 수업
        req = _make_req(name="법학개론", processed_room="K262",
                        exam_start=4, exam_end=4)
        tt = {SHEET: {"K262": {4: ("한국어어휘론-01", None)}}}
        ov = detect_timetable_overlaps([req], tt, DTS, DTD)
        assert len(ov) == 1
        assert ov[0].room == "K262" and ov[0].period == 4
        assert ov[0].occupant == "한국어어휘론-01"
        assert ov[0].req.name == "법학개론"

    def test_own_room_excluded(self):
        # 자기 강의실 N212의 자기 수업 교시(2~3)에 시험 → 충돌 아님
        req = _make_req(name="가과목", processed_room="N212",
                        exam_start=2, exam_end=3)
        tt = {SHEET: {"N212": {2: ("다른표기-01", None), 3: ("다른표기-01", None)}}}
        ov = detect_timetable_overlaps([req], tt, DTS, DTD)
        assert ov == []

    def test_released_slot_excluded(self):
        req = _make_req(name="법학개론", processed_room="K262",
                        exam_start=4, exam_end=4)
        tt = {SHEET: {"K262": {4: ("한국어어휘론-01", None)}}}
        released = {(SHEET, "K262", 4)}
        ov = detect_timetable_overlaps([req], tt, DTS, DTD, released)
        assert ov == []

    def test_sibling_occupant_excluded(self):
        # 점유자가 같은 교과목 다른 분반(I-DESIGN-15) → 충돌 아님
        req = _make_req(name="I-DESIGN", processed_room="N210",
                        exam_start=4, exam_end=6)
        tt = {SHEET: {"N210": {4: ("I-DESIGN-15", None),
                               5: ("I-DESIGN-15", None),
                               6: ("I-DESIGN-15", None)}}}
        ov = detect_timetable_overlaps([req], tt, DTS, DTD)
        assert ov == []

    def test_self_occupant_no_hyphen_excluded(self):
        # 점유자가 같은 과목인데 하이픈 없는 표기('한국어학술작문01') → 자기 시험, 충돌 아님
        req = _make_req(name="한국어학술작문", processed_room="NP117",
                        exam_start=5, exam_end=6)
        tt = {SHEET: {"NP117": {5: ("한국어학술작문01", None),
                               6: ("한국어학술작문01", None)}}}
        assert detect_timetable_overlaps([req], tt, DTS, DTD) == []

    def test_self_occupant_exam_prefix_excluded(self):
        # 점유자가 '시험(원래): 뇌와인지-01' → 같은 과목 자기 시험, 충돌 아님
        req = _make_req(name="뇌와인지", processed_room="N218",
                        exam_start=1, exam_end=1)
        tt = {SHEET: {"N218": {1: ("시험(원래): 뇌와인지-01", None)}}}
        assert detect_timetable_overlaps([req], tt, DTS, DTD) == []

    def test_missing_date_or_period_skipped(self):
        r1 = _make_req(name="가", processed_room="K262", exam_date=None)
        r2 = _make_req(name="나", processed_room="K262",
                       exam_start=None, exam_end=None)
        tt = {SHEET: {"K262": {4: ("X-01", None)}}}
        assert detect_timetable_overlaps([r1, r2], tt, DTS, DTD) == []

    def test_room_or_sheet_absent(self):
        req = _make_req(name="법학개론", processed_room="ZZZ",
                        exam_start=4, exam_end=4)
        tt = {SHEET: {"K262": {4: ("X-01", None)}}}
        assert detect_timetable_overlaps([req], tt, DTS, DTD) == []
        # 시트 자체가 없을 때
        assert detect_timetable_overlaps([req], {}, DTS, DTD) == []

    def test_free_slot_no_overlap(self):
        # K262의 4교시가 시간표상 비어 있으면 충돌 아님
        req = _make_req(name="법학개론", processed_room="K262",
                        exam_start=4, exam_end=4)
        tt = {SHEET: {"K262": {2: ("X-01", None)}}}  # 4교시 없음
        assert detect_timetable_overlaps([req], tt, DTS, DTD) == []

    def test_multi_room_each_checked(self):
        req = _make_req(name="법학개론", processed_room="K262,K267",
                        exam_start=4, exam_end=4)
        tt = {SHEET: {"K262": {4: ("A-01", None)}, "K267": {4: ("B-01", None)}}}
        ov = detect_timetable_overlaps([req], tt, DTS, DTD)
        assert {o.room for o in ov} == {"K262", "K267"}
        assert len(ov) == 2

    def test_cell_value_plain_string(self):
        # 셀 값이 (값, rgb) 튜플이 아니라 순수 문자열이어도 처리
        req = _make_req(name="법학개론", processed_room="K262",
                        exam_start=4, exam_end=4)
        tt = {SHEET: {"K262": {4: "한국어어휘론-01"}}}
        ov = detect_timetable_overlaps([req], tt, DTS, DTD)
        assert len(ov) == 1 and ov[0].occupant == "한국어어휘론-01"


class TestTimetableOverlapsOccupantIndex:
    """occupant_index로 '점유자가 실제로 그 방에서 시험을 보는가' 검증."""

    def _req(self):  # K262는 자기 방(N212)이 아님, 6/17 4교시 시험
        return _make_req(name="법학개론", processed_room="K262",
                         exam_start=4, exam_end=4)

    def _tt(self):  # K262 4교시에 한국어어휘론-01 점유
        return {SHEET: {"K262": {4: ("한국어어휘론-01", None)}}}

    def _idx(self, **info):
        base = dict(exam_date=None, exam_start=None, exam_end=None, no_exam=False)
        base.update(info)
        return {("한국어어휘론", normalize_ban("01")): base}

    def test_no_exam_occupant_freed(self):
        ov = detect_timetable_overlaps([self._req()], self._tt(), DTS, DTD,
                                       occupant_index=self._idx(no_exam=True))
        assert ov == []

    def test_different_exam_date_freed(self):
        ov = detect_timetable_overlaps(
            [self._req()], self._tt(), DTS, DTD,
            occupant_index=self._idx(exam_date=D2, exam_start=4, exam_end=4))
        assert ov == []

    def test_no_exam_date_freed(self):
        ov = detect_timetable_overlaps([self._req()], self._tt(), DTS, DTD,
                                       occupant_index=self._idx(exam_date=None))
        assert ov == []

    def test_same_date_overlapping_period_flagged(self):
        ov = detect_timetable_overlaps(
            [self._req()], self._tt(), DTS, DTD,
            occupant_index=self._idx(exam_date=D, exam_start=4, exam_end=5))
        assert len(ov) == 1

    def test_same_date_non_overlapping_period_freed(self):
        ov = detect_timetable_overlaps(
            [self._req()], self._tt(), DTS, DTD,
            occupant_index=self._idx(exam_date=D, exam_start=1, exam_end=2))
        assert ov == []

    def test_same_date_no_periods_conservative_flag(self):
        ov = detect_timetable_overlaps([self._req()], self._tt(), DTS, DTD,
                                       occupant_index=self._idx(exam_date=D))
        assert len(ov) == 1

    def test_occupant_not_in_index_conservative_flag(self):
        ov = detect_timetable_overlaps([self._req()], self._tt(), DTS, DTD,
                                       occupant_index={})
        assert len(ov) == 1

    def test_none_index_is_backward_compatible(self):
        # occupant_index 없으면 시간표 점유를 그대로 신뢰(플래그)
        ov = detect_timetable_overlaps([self._req()], self._tt(), DTS, DTD)
        assert len(ov) == 1


class TestManualProcessedEntries:
    """배정 워크플로 미경유 + P열 직접 입력 탐지."""

    def _req(self, name="가과목", ban="01", processed_room="N212",
             category=Category.ROOM_CHANGE):
        r = _make_req(name=name, ban=ban, processed_room=processed_room)
        r.category = category
        return r

    def test_room_change_no_assignment_detected(self):
        req = self._req(category=Category.ROOM_CHANGE)
        result = detect_manual_processed_entries([req], {})
        assert len(result) == 1
        assert result[0][0] is req
        assert result[0][1] == ["N212"]

    def test_skip_with_processed_room_detected(self):
        req = self._req(category=Category.SKIP, processed_room="N305")
        result = detect_manual_processed_entries([req], {})
        assert [rooms for _, rooms in result] == [["N305"]]

    def test_room_split_multi_room(self):
        req = self._req(category=Category.ROOM_SPLIT, processed_room="N210, N405")
        result = detect_manual_processed_entries([req], {})
        assert result[0][1] == ["N210", "N405"]

    def test_program_assignment_excluded(self):
        # 정확 키 배정이 있으면 프로그램 추적 건 → 제외
        req = self._req(category=Category.ROOM_CHANGE)  # key = "가과목-01"
        result = detect_manual_processed_entries([req], {"가과목-01": {"room": "N999"}})
        assert result == []

    def test_split_assignment_key_excluded(self):
        # 분반 다중 배정 키(+N)도 '배정 있음'으로 간주 → 제외
        req = self._req(category=Category.ROOM_SPLIT)  # key = "가과목-01"
        result = detect_manual_processed_entries([req], {"가과목-01+1": {"room": "N999"}})
        assert result == []

    def test_normal_exam_excluded(self):
        req = self._req(category=Category.NORMAL_EXAM)
        assert detect_manual_processed_entries([req], {}) == []

    def test_no_exam_excluded(self):
        req = self._req(category=Category.NO_EXAM)
        assert detect_manual_processed_entries([req], {}) == []

    def test_marker_and_blank_processed_room_excluded(self):
        for pr in ["강의실 미사용", "확인필요", "", None]:
            req = self._req(category=Category.ROOM_CHANGE, processed_room=pr)
            assert detect_manual_processed_entries([req], {}) == []

    def test_order_preserved_and_filtered(self):
        a = self._req(name="에이", processed_room="A1", category=Category.ROOM_CHANGE)
        b = self._req(name="비", processed_room="B1", category=Category.SKIP)
        c = self._req(name="씨", processed_room="강의실 미사용", category=Category.ROOM_SPLIT)
        result = detect_manual_processed_entries([a, b, c], {})
        assert [req.name for req, _ in result] == ["에이", "비"]


class TestManualNoUseEntries:
    """배정 워크플로 미경유 + P열 '강의실 미사용'(강의실 면제) 탐지."""

    def _req(self, name="가과목", ban="01", processed_room="강의실 미사용",
             category=Category.SKIP):
        r = _make_req(name=name, ban=ban, processed_room=processed_room)
        r.category = category
        return r

    def test_skip_no_use_detected(self):
        req = self._req(category=Category.SKIP)
        assert detect_manual_no_use_entries([req], {}) == [req]

    def test_whitespace_variants_detected(self):
        for pr in ["강의실 미사용", "강의실미사용", "  강의실  미사용 "]:
            req = self._req(processed_room=pr)
            assert detect_manual_no_use_entries([req], {}) == [req]

    def test_real_room_not_detected(self):
        # 실제 강의실은 '미사용'이 아님 → 다른 함수 소관
        req = self._req(processed_room="N212")
        assert detect_manual_no_use_entries([req], {}) == []

    def test_check_needed_and_blank_not_detected(self):
        for pr in ["확인필요", "", None]:
            req = self._req(processed_room=pr)
            assert detect_manual_no_use_entries([req], {}) == []

    def test_program_assignment_excluded(self):
        req = self._req(category=Category.ROOM_CHANGE)  # key = "가과목-01"
        assert detect_manual_no_use_entries([req], {"가과목-01": {"room": "N9"}}) == []
        assert detect_manual_no_use_entries([req], {"가과목-01+1": {"room": "N9"}}) == []

    def test_auto_categories_excluded(self):
        for cat in (Category.NORMAL_EXAM, Category.NO_EXAM):
            req = self._req(category=cat)
            assert detect_manual_no_use_entries([req], {}) == []

    def test_disjoint_from_processed_rooms(self):
        # 같은 입력에서 두 함수가 같은 req를 동시에 잡지 않는다
        room = self._req(name="방", processed_room="N100", category=Category.ROOM_CHANGE)
        nouse = self._req(name="면제", processed_room="강의실 미사용", category=Category.SKIP)
        reqs = [room, nouse]
        r_ids = {id(req) for req, _ in detect_manual_processed_entries(reqs, {})}
        u_ids = {id(req) for req in detect_manual_no_use_entries(reqs, {})}
        assert r_ids == {id(room)}
        assert u_ids == {id(nouse)}
        assert r_ids & u_ids == set()
