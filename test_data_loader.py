"""data_loader 단위 테스트."""

import datetime
import pytest

from data_loader import (
    _parse_date,
    parse_schedule,
    classify_requests,
    build_mappings_from_sheets,
    _infer_year,
    ScheduleSlot,
    Category,
)


class TestParseDate:
    def test_datetime(self):
        assert _parse_date(datetime.datetime(2026, 4, 21)) == datetime.date(2026, 4, 21)

    def test_date(self):
        assert _parse_date(datetime.date(2026, 4, 21)) == datetime.date(2026, 4, 21)

    def test_date_old_year(self):
        assert _parse_date(datetime.date(1, 1, 1)) is None

    def test_string_valid(self):
        assert _parse_date("2026-04-22") == datetime.date(2026, 4, 22)

    def test_string_zero(self):
        assert _parse_date("0000-01-01") is None

    def test_empty(self):
        assert _parse_date("") is None
        assert _parse_date(None) is None

    def test_datetime_old_year(self):
        assert _parse_date(datetime.datetime(1, 1, 1)) is None


class TestParseSchedule:
    def test_single_period(self):
        result = parse_schedule("수1(K107)")
        assert len(result) == 1
        assert result[0] == ScheduleSlot("수", 1, 1, "K107")

    def test_range(self):
        result = parse_schedule("화4~5(K106)")
        assert result[0].start == 4
        assert result[0].end == 5

    def test_multi_slot(self):
        result = parse_schedule("화4~5(K106), 목4(K106)")
        assert len(result) == 2
        assert result[1].day == "목"
        assert result[1].end == 4

    def test_empty_room(self):
        result = parse_schedule("월2~3()")
        assert result[0].room == ""

    def test_empty_string(self):
        assert parse_schedule("") == []
        assert parse_schedule(None) == []


def _make_req(**kwargs):
    from data_loader import ExamRequest
    defaults = dict(
        row=1, department="테스트", name="테스트과목", ban="01",
        professor="교수", students=30, schedule_raw="화2~3(K106)",
        slots=[ScheduleSlot("화", 2, 3, "K106")], room="K106",
        exam_date=datetime.date(2026, 4, 21), exam_start=2, exam_end=3,
        room_choice="기존 강의실", remarks="",
    )
    defaults.update(kwargs)
    return ExamRequest(**defaults)


class TestClassify:
    def test_normal_exam(self):
        req = _make_req()
        classify_requests([req])
        assert req.category == Category.NORMAL_EXAM

    def test_no_exam_keyword(self):
        req = _make_req(exam_date=None, remarks="중간고사 미실시")
        classify_requests([req])
        assert req.category == Category.NO_EXAM

    def test_skip_no_schedule(self):
        req = _make_req(slots=[], schedule_raw="")
        classify_requests([req])
        assert req.category == Category.SKIP

    def test_room_change_with_date(self):
        req = _make_req(room_choice="강의실 변경 요청")
        classify_requests([req])
        assert req.category == Category.ROOM_CHANGE

    def test_room_change_no_date_is_skip(self):
        req = _make_req(room_choice="강의실 변경 요청", exam_date=None)
        classify_requests([req])
        assert req.category == Category.SKIP

    def test_room_split_with_date(self):
        req = _make_req(room_choice="강의실 분반 요청")
        classify_requests([req])
        assert req.category == Category.ROOM_SPLIT

    def test_trailing_space(self):
        req = _make_req(room_choice="강의실 변경 요청 ")
        classify_requests([req])
        assert req.category == Category.ROOM_CHANGE

    def test_leading_space(self):
        req = _make_req(room_choice=" 강의실 분반 요청")
        classify_requests([req])
        assert req.category == Category.ROOM_SPLIT


class TestKeyDedup:
    """중복 키 감지 회귀 테스트."""

    def test_unique_keys_unchanged(self):
        from data_loader import ExamRequest
        reqs = [_make_req(row=2, name="과목A", ban="01"), _make_req(row=3, name="과목B", ban="01")]
        # load_requests의 중복 감지 로직 재현
        seen = {}
        for req in reqs:
            if req.key in seen:
                seen[req.key] += 1
                req.key = f"{req.key}#{req.row}"
            else:
                seen[req.key] = 1
        duped = {k for k, cnt in seen.items() if cnt > 1}
        for req in reqs:
            if req.key in duped:
                req.key = f"{req.key}#{req.row}"

        assert reqs[0].key == "과목A-01"
        assert reqs[1].key == "과목B-01"

    def test_duplicate_keys_get_row_suffix(self):
        from data_loader import ExamRequest
        reqs = [_make_req(row=2, name="과목A", ban="01"), _make_req(row=5, name="과목A", ban="01")]
        seen = {}
        for req in reqs:
            if req.key in seen:
                seen[req.key] += 1
                req.key = f"{req.key}#{req.row}"
            else:
                seen[req.key] = 1
        duped = {k for k, cnt in seen.items() if cnt > 1}
        for req in reqs:
            if req.key in duped:
                req.key = f"{req.key}#{req.row}"

        assert reqs[0].key == "과목A-01#2"
        assert reqs[1].key == "과목A-01#5"
        assert reqs[0].key != reqs[1].key


class TestBuildMappings:
    """동적 날짜 매핑 생성 테스트."""

    def test_basic_mapping(self):
        sheets = ["4.21.(화)", "4.22.(수)", "4.23.(목)", "4.24.(금)", "4.27.(월)"]
        d2d, d2s, order, *_ = build_mappings_from_sheets(sheets, 2026)
        assert d2d[datetime.date(2026, 4, 21)] == "화"
        assert d2d[datetime.date(2026, 4, 27)] == "월"
        assert d2s["화"] == "4.21.(화)"
        assert order == sheets  # 날짜순 정렬

    def test_different_year(self):
        sheets = ["10.15.(수)", "10.16.(목)"]
        d2d, d2s, order, *_ = build_mappings_from_sheets(sheets, 2025)
        assert d2d[datetime.date(2025, 10, 15)] == "수"
        assert d2s["목"] == "10.16.(목)"
        assert len(order) == 2

    def test_non_matching_sheets_ignored(self):
        sheets = ["4.21.(화)", "summary", "기타"]
        d2d, d2s, order, *_ = build_mappings_from_sheets(sheets, 2026)
        assert len(d2d) == 1
        assert len(order) == 1

    def test_classify_with_dynamic_mapping(self):
        """다른 학기 날짜도 동적 매핑으로 올바르게 분류."""
        sheets = ["10.15.(수)"]
        d2d, *_ = build_mappings_from_sheets(sheets, 2025)
        req = _make_req(exam_date=datetime.date(2025, 10, 15))
        classify_requests([req], date_to_day=d2d)
        assert req.category == Category.NORMAL_EXAM

    def test_classify_out_of_range_with_dynamic_mapping(self):
        """동적 매핑 범위 밖 날짜는 SKIP."""
        sheets = ["10.15.(수)"]
        d2d, *_ = build_mappings_from_sheets(sheets, 2025)
        req = _make_req(exam_date=datetime.date(2025, 12, 1))
        classify_requests([req], date_to_day=d2d)
        assert req.category == Category.SKIP


class TestInferYear:
    def test_from_requests(self):
        req = _make_req(exam_date=datetime.date(2025, 10, 15))
        assert _infer_year([req]) == 2025

    def test_none_when_no_dates(self):
        req = _make_req(exam_date=None, remarks="미실시")
        assert _infer_year([req]) is None


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
