"""원본 '요청사항 처리'(P열, col16) 강의실 배정 중복 점검.

요청 엑셀의 ``요청사항 처리`` 열에 사람이 미리 적어둔 실제 강의실 배정 결정을
읽기 전용으로 감사한다. 같은 시험일자·같은 강의실에 교시가 겹치는 서로 다른
교과목이 있으면 이중 배정(double-booking) 충돌로 검출한다.

UI 의존이 없는 순수 로직 — 단위 테스트 가능(assignment_status.py와 동일 패턴).
프로그램 내부 배정(_assignments.json)과는 무관하며 그 상태를 변경하지 않는다.
"""

from __future__ import annotations

import datetime
import re
from collections import defaultdict
from dataclasses import dataclass
from typing import Optional

from data_loader import ExamRequest, normalize_ban

# ``요청사항 처리`` 값 중 실제 강의실이 아닌 마커(공백 제거 후 비교).
# 체육관/테니스장/헬스장 등은 실제 시험 장소이므로 제외 대상이 아니다.
NON_ROOM_MARKERS = {"강의실미사용", "확인필요", ""}


def _strip_ws(s: str) -> str:
    return re.sub(r"\s+", "", s or "")


def parse_processed_rooms(text: str) -> list[str]:
    """``요청사항 처리`` 텍스트에서 배정 강의실 목록을 추출한다.

    콤마로 구분된 다중 배정(예: ``"N210,N405"``)을 분리하고, 강의실이 아닌
    마커(``강의실 미사용``·``확인필요``·빈칸)는 제외한다. 강의실 코드 원문은
    공백만 정리해 보존한다.
    """
    if not text:
        return []
    rooms = []
    for part in text.split(","):
        room = part.strip()
        if not room:
            continue
        if _strip_ws(room) in NON_ROOM_MARKERS:
            continue
        rooms.append(room)
    return rooms


@dataclass
class ProcessedRoomConflict:
    """원본 결정 기준 강의실 이중 배정 한 쌍."""
    exam_date: datetime.date
    room: str
    req_a: ExamRequest
    req_b: ExamRequest


@dataclass
class UnjudgedBooking:
    """강의실은 배정됐으나 날짜 또는 교시가 없어 충돌 판정이 불가한 행."""
    room: str
    req: ExamRequest
    reason: str


@dataclass
class TimetableOverlap:
    """P열 강의실이 기존 수업시간표 점유와 겹친 한 건 (강의실+교시 단위)."""
    exam_date: datetime.date
    room: str
    period: int
    req: ExamRequest
    occupant: str  # 그 슬롯을 점유 중인 기존 수업/예약 텍스트


def _period_range(req: ExamRequest) -> Optional[tuple[int, int]]:
    """시험 교시 구간 (시작, 종료)을 정렬해 반환. 둘 중 하나라도 없으면 None."""
    s, e = req.exam_start, req.exam_end
    if s is None or e is None:
        return None
    return (s, e) if s <= e else (e, s)


def _overlap(a: tuple[int, int], b: tuple[int, int]) -> bool:
    return a[0] <= b[1] and b[0] <= a[1]


def detect_processed_room_conflicts(requests):
    """원본 ``요청사항 처리`` 강의실 배정의 이중 배정 충돌을 검출한다.

    같은 시험일자(``exam_date``) + 같은 강의실(``processed_room``) + 교시
    (``exam_start``~``exam_end``)가 겹치는 **서로 다른 교과목** 쌍을 충돌로 본다.
    같은 교과목명의 분반끼리는 충돌이 아니다(기존 충돌 규칙과 일관).

    강의실은 배정됐으나 ``exam_date`` 또는 시작/종료교시가 없는 행은 충돌로
    세지 않고 ``unjudged`` 목록으로 분리한다(데이터 품질 점검용).

    Returns:
        ``(conflicts, unjudged)``
        — ``conflicts``: ``list[ProcessedRoomConflict]`` (일자→강의실→행 순 정렬)
        — ``unjudged``: ``list[UnjudgedBooking]`` (강의실→행 순 정렬)
    """
    # (exam_date, room) -> list[(req, period_range)]
    groups: dict = defaultdict(list)
    unjudged: list = []

    for req in requests:
        rooms = parse_processed_rooms(getattr(req, "processed_room", ""))
        if not rooms:
            continue
        pr = _period_range(req)
        for room in rooms:
            if req.exam_date is None or pr is None:
                reason = ("시험일자(K열) 없음" if req.exam_date is None
                          else "시작/종료교시(L·M열) 없음")
                unjudged.append(UnjudgedBooking(room=room, req=req, reason=reason))
                continue
            groups[(req.exam_date, room)].append((req, pr))

    conflicts: list = []
    for (exam_date, room), bookings in groups.items():
        n = len(bookings)
        for i in range(n):
            for j in range(i + 1, n):
                req_a, pa = bookings[i]
                req_b, pb = bookings[j]
                if req_a.name == req_b.name:
                    continue  # 같은 교과목 분반 형제는 충돌 아님
                if _overlap(pa, pb):
                    conflicts.append(ProcessedRoomConflict(
                        exam_date=exam_date, room=room,
                        req_a=req_a, req_b=req_b,
                    ))

    conflicts.sort(key=lambda c: (c.exam_date.isoformat(), c.room,
                                  c.req_a.row, c.req_b.row))
    unjudged.sort(key=lambda u: (u.room, u.req.row))
    return conflicts, unjudged


def _cell_text(cell) -> str:
    """timetable_data 셀 값에서 표시 텍스트만 추출 ((value, rgb) 튜플 또는 문자열)."""
    if isinstance(cell, (tuple, list)):
        return str(cell[0]) if cell else ""
    return str(cell)


def _parse_occupant(text) -> tuple[str, str]:
    """시간표 셀 텍스트에서 ``(교과목명, 분반)`` 을 추출한다.

    실제 시간표 셀은 형태가 제각각이라 모두 흡수한다:
    * ``과목명-분반``        (예: ``한국어어휘론-01``, ``I-DESIGN-15``, ``마케팅-P1``)
    * ``과목명분반``         (하이픈 없음, 예: ``한국어학술작문01``)
    * ``시험(원래): 과목명-분반`` (시험 표기 접두)

    과목-분반 꼴이 아니면(예약·회의·외부 일정 등) ``(원문, "")`` 을 반환한다.
    """
    s = re.sub(r"^시험[^:]*:\s*", "", str(text or "")).strip()
    m = re.search(r"-([A-Za-z0-9]{1,3})$", s)
    if m:
        return s[:m.start()].strip(), m.group(1)
    m = re.search(r"(?<=\D)(\d{1,2})$", s)  # 하이픈 없이 끝에 붙은 분반 숫자
    if m:
        return s[:m.start()].strip(), m.group(1)
    return s, ""


def _occupant_base(text) -> str:
    """시간표 점유자 텍스트에서 기본 교과목명만 추출 (분반 제거)."""
    return _parse_occupant(text)[0]


def _occupant_uses_room(occupant: str, date, period: int, index: dict) -> bool:
    """시간표 점유자가 그 시험일·교시에 실제로 그 강의실에서 시험을 보는지 판정.

    정상 수업 시간표는 시험주간 점유의 근거가 아니다(시험만 점유). 점유자가
    아래 중 하나면 **그 방을 안 쓰는 것(비어 있음)** 으로 본다:
    * 미실시/대체 등 사용 안함 선언
    * 그 날 시험 없음(시험일 없음 또는 다른 날) — 다일 수업의 비시험 요일 포함
    * 시험 교시가 겹치지 않음

    점유자를 ``index``(``load_course_exam_index``)에서 못 찾으면 정보 부족이므로
    보수적으로 점유(True)로 본다.
    """
    name, ban = _parse_occupant(occupant)
    info = index.get((name, normalize_ban(ban)))
    if info is None:
        return True  # 정보 없음 → 보수적으로 점유(확인 필요)
    if info.get("no_exam"):
        return False
    exam_date = info.get("exam_date")
    if exam_date is None or exam_date != date:
        return False
    s, e = info.get("exam_start"), info.get("exam_end")
    if s is not None and e is not None:
        lo, hi = (s, e) if s <= e else (e, s)
        return lo <= period <= hi
    return True  # 날짜는 맞으나 교시 정보 없음 → 보수적으로 점유


def detect_timetable_overlaps(requests, timetable_data, date_to_sheet,
                              date_to_day, released_slots=None, occupant_index=None):
    """P열 강의실이 그 시험일·교시에 기존 수업시간표 점유와 겹치는지 교차 검사한다.

    사용자가 ``요청사항 처리``에 직접 적은 강의실이, 그 시험일자에 해당하는
    시간표 시트에서 같은 교시에 이미 점유돼 있으면 충돌로 본다.

    거짓 경고 방지를 위해 다음을 제외한다:
    * **자기 수업 슬롯** — 시험 과목 자신이 그 요일에 쓰는 (강의실, 교시).
    * **해제된 슬롯** — ``released_slots`` (자동/수동 해제)로 비워진 (시트, 강의실, 교시).
    * **같은 교과목 분반**이 점유자인 경우.
    * ``occupant_index`` 가 주어지면, **점유자가 그 시험주간에 실제로 그 방에서
      시험을 보지 않는 경우**(사용 안함·그 날 시험 없음/다른 날·교시 안 겹침).
      → 정상 수업 시간표상 점유라도 시험엔 안 쓰면 그 방은 비어 있는 것으로 본다.
      (공통만 배정하므로 전공이 "사용 안함"이면 그 방을 공통에 배정 가능.)

    Args:
        requests: ExamRequest 리스트(``processed_room`` 포함).
        timetable_data: ``{sheet: {room: {period: (value, rgb)}}}``.
        date_to_sheet: ``{date: sheet}``.
        date_to_day: ``{date: weekday}``.
        released_slots: ``{(sheet, room, period)}`` 해제 슬롯 집합(없으면 빈 집합).
        occupant_index: ``load_course_exam_index`` 결과. 점유자의 실제 시험 사용
            여부를 검증해 거짓양성을 거른다. None이면 시간표 점유를 그대로 신뢰.

    Returns:
        ``list[TimetableOverlap]`` — (일자→강의실→교시→행) 순 정렬.
    """
    released_slots = released_slots or set()
    overlaps: list = []

    for req in requests:
        rooms = parse_processed_rooms(getattr(req, "processed_room", ""))
        if not rooms:
            continue
        pr = _period_range(req)
        if req.exam_date is None or pr is None:
            continue
        sheet = date_to_sheet.get(req.exam_date)
        if not sheet or sheet not in timetable_data:
            continue

        weekday = date_to_day.get(req.exam_date)
        own = set()  # 이 과목 자신이 그 요일에 점유하는 (강의실, 교시)
        for slot in req.slots:
            if slot.day == weekday and slot.room:
                lo, hi = sorted((slot.start, slot.end))
                for p in range(lo, hi + 1):
                    own.add((slot.room, p))

        s, e = pr
        sheet_grid = timetable_data[sheet]
        for room in rooms:
            room_grid = sheet_grid.get(room, {})
            for p in range(s, e + 1):
                if p not in room_grid:
                    continue
                if (room, p) in own:
                    continue
                if (sheet, room, p) in released_slots:
                    continue
                occupant = _cell_text(room_grid[p])
                if _occupant_base(occupant) == req.name:
                    continue  # 같은 교과목(분반)이 그 방을 쓰는 경우 — 충돌 아님
                if occupant_index is not None and not _occupant_uses_room(
                        occupant, req.exam_date, p, occupant_index):
                    continue  # 점유자가 시험주간에 그 방을 실제로 안 씀 → 비어 있음
                overlaps.append(TimetableOverlap(
                    exam_date=req.exam_date, room=room, period=p,
                    req=req, occupant=occupant,
                ))

    overlaps.sort(key=lambda o: (o.exam_date.isoformat(), o.room, o.period, o.req.row))
    return overlaps
