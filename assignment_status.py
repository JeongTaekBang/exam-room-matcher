"""배정 상태 판정 로직.

dashboard.py에서 분리한 UI 의존 없는 순수 함수들. Category 라벨·
수용인원 헬퍼·compute_status를 한 곳에 모아 단위 테스트가 가능하도록 한다.
"""

from data_loader import Category
from workflow_utils import resolve_needed_periods


# ──────────────────────────────────────────────
# Category 라벨 — assignments JSON의 "category" 값으로 직렬화되는 단일 출처
# ──────────────────────────────────────────────

CAT_LABELS = {
    Category.NORMAL_EXAM: "시험 진행",
    Category.NO_EXAM: "미실시/대체과제",
    Category.ROOM_CHANGE: "강의실 변경",
    Category.ROOM_SPLIT: "강의실 분반",
    Category.SKIP: "미확정",
}
LABEL_NORMAL_EXAM = CAT_LABELS[Category.NORMAL_EXAM]
LABEL_NO_EXAM = CAT_LABELS[Category.NO_EXAM]
LABEL_ROOM_CHANGE = CAT_LABELS[Category.ROOM_CHANGE]
LABEL_ROOM_SPLIT = CAT_LABELS[Category.ROOM_SPLIT]
LABEL_SKIP = CAT_LABELS[Category.SKIP]


# ──────────────────────────────────────────────
# 헬퍼
# ──────────────────────────────────────────────

def compute_auto_released(
    assignments: dict,
    requests,
    date_to_day: dict,
    date_to_sheet: dict,
    day_to_sheets: dict,
) -> dict[tuple[str, str, int], tuple[str, str]]:
    """자동 해제 슬롯을 파생한다.

    파생 규칙:
    1. ROOM_CHANGE 배정 + 분반(``keep_orig=False``) → 원래 강의실의 배정 교시를 해제
    2. NO_EXAM 요청 → 수업 강의실의 시험·수업 교시를 해제 (exam_date 없으면 같은
       요일의 모든 시트에 적용 — 다주차 안전)
    3. NORMAL_EXAM + 시험교시 < 수업교시 → 미사용 교시 부분 해제

    Returns:
        ``{(sheet, room, period): (subject_key, label)}`` — label은
        "이동" / "미실시" / "부분해제" 중 하나.
    """
    result: dict[tuple[str, str, int], tuple[str, str]] = {}

    # 1. ROOM_CHANGE 배정 또는 분반(기존 미유지) → 원래 강의실 해제
    for key, a in assignments.items():
        cat = a.get("category", "")
        is_change = cat == LABEL_ROOM_CHANGE
        is_split_no_keep = cat == LABEL_ROOM_SPLIT and not a.get("keep_orig", True)
        if not is_change and not is_split_no_keep:
            continue
        orig = a.get("original_room", "")
        if not orig or orig == a.get("room"):
            continue
        for p in a.get("periods", []):
            result[(a["sheet"], orig, p)] = (key, "이동")

    # 2. NO_EXAM 요청 → 강의실 해제
    for req in requests:
        if req.category != Category.NO_EXAM:
            continue
        if req.exam_date and req.exam_date in date_to_day:
            exam_day = date_to_day[req.exam_date]
            sheet = date_to_sheet.get(req.exam_date, "")
            if not sheet:
                continue
            periods = resolve_needed_periods(req, exam_day)
            if not periods:
                continue
            rooms = {s.room for s in req.slots if s.day == exam_day and s.room}
            if not rooms and req.room:
                rooms = {req.room}
            for room in rooms:
                for p in periods:
                    result[(sheet, room, p)] = (req.key, "미실시")
        else:
            # exam_date 없는 경우: 슬롯의 요일에 해당하는 모든 시트에 적용
            for slot in req.slots:
                if not slot.room:
                    continue
                s, e = max(0, slot.start), min(14, slot.end)
                for sheet in day_to_sheets.get(slot.day, ()):
                    for p in range(s, e + 1):
                        result[(sheet, slot.room, p)] = (req.key, "미실시")

    # 3. NORMAL_EXAM 부분해제: 시험 교시 < 수업 교시 → 미사용 교시 자동 해제
    for req in requests:
        if req.category != Category.NORMAL_EXAM or req.exam_start is None:
            continue
        if not req.exam_date or req.exam_date not in date_to_day:
            continue
        exam_day = date_to_day[req.exam_date]
        sheet = date_to_sheet.get(req.exam_date, "")
        if not sheet:
            continue
        exam_set = set(resolve_needed_periods(req, exam_day))
        for slot in req.slots:
            if slot.day != exam_day or not slot.room:
                continue
            s, e = max(0, slot.start), min(14, slot.end)
            for p in range(s, e + 1):
                if p not in exam_set:
                    result[(sheet, slot.room, p)] = (req.key, "부분해제")

    return result


# '요청사항 처리'(P열)에 기입하는 미실시 표기 — 원본 데이터 표기와 동일하게 유지
PROCESSED_NO_EXAM = "강의실 미사용"


def resolve_processed_room(req, assignments: dict) -> str:
    """요청의 최종 강의실 결정을 '요청사항 처리'(P열) 문자열로 반환한다.

    우선순위:
    1) 프로그램 배정(ROOM_CHANGE/ROOM_SPLIT)이 있으면 배정 강의실
       — 분반(다중 배정)은 강의실을 콤마로 결합 (예: ``"N210,N405"``)
    2) NO_EXAM → ``"강의실 미사용"``
    3) NORMAL_EXAM → 시험 강의실(``req.room``)
    4) 그 외(미배정 변경/분반, 미확정) → ``""``

    빈 문자열은 "프로그램이 아직 결정하지 못함"을 뜻한다. 내보내기 호출자는
    빈 값일 때 원본 셀을 덮어쓰지 않고 보존하여 수동 입력 여지를 남긴다.
    """
    keys = [k for k in assignments if k == req.key or k.startswith(req.key + "+")]
    if keys:
        rooms: list[str] = []
        for k in sorted(keys):
            room = str(assignments[k].get("room", "")).strip()
            if room and room not in rooms:
                rooms.append(room)
        if rooms:
            return ",".join(rooms)
    if req.category == Category.NO_EXAM:
        return PROCESSED_NO_EXAM
    if req.category == Category.NORMAL_EXAM:
        return (req.room or "").strip()
    return ""


def extract_base_key(assignment_key: str) -> str:
    """분반 키(``과목명-분반+N``)에서 기본 과목 키를 추출.

    마지막 ``+`` 다음이 정수일 때만 그 앞부분을 base key로 반환한다.
    과목명에 ``+`` 기호가 포함되어도 안전 (마지막 ``+`` 다음 정수 검증).
    분반 suffix가 없으면 입력을 그대로 반환.

    Examples:
        "영문학의이해-01"      → "영문학의이해-01"
        "영문학의이해-01+2"    → "영문학의이해-01"
        "C++-01"               → "C++-01"   (+ 다음이 정수 아님)
        "C++-01+3"             → "C++-01"
    """
    i = assignment_key.rfind("+")
    if i > 0 and assignment_key[i + 1:].isdigit():
        return assignment_key[:i]
    return assignment_key


def _room_cap(room_capacity: dict | None, room: str) -> int:
    """room_capacity에서 강의실 수용인원을 안전하게 int로 추출."""
    if not room_capacity or not room:
        return 0
    try:
        return int(room_capacity.get(room, 0) or 0)
    except (TypeError, ValueError):
        return 0


# ──────────────────────────────────────────────
# 상태 판정
# ──────────────────────────────────────────────

def compute_status(req, assignments: dict, room_capacity: dict | None = None) -> str:
    """요청의 배정 상태를 판정한다.

    Returns:
        "완료" | "미배정" | "미확정"
    """
    if req.category in (Category.NORMAL_EXAM, Category.NO_EXAM):
        return "완료"
    if req.category in (Category.ROOM_CHANGE, Category.ROOM_SPLIT):
        related = [
            (k, a) for k, a in assignments.items()
            if k == req.key or k.startswith(req.key + "+")
        ]
        if not related:
            return "미배정"

        # 분반으로 저장된 항목이 하나라도 있으면 분반 완료 규칙을 적용한다.
        is_split_mode = (
            req.category == Category.ROOM_SPLIT
            or any(a.get("category") == LABEL_ROOM_SPLIT for _, a in related)
        )
        if not is_split_mode:
            # ROOM_CHANGE 단일 배정. 한 강의실 수용인원이 학생 수 미만이면
            # "미배정"으로 표시 — UI 배정 흐름에서는 cap_short 가드로 막히지만
            # JSON 직접 편집·마이그레이션 등으로 invalid 상태가 들어왔을 때
            # compute_status가 잘못 "완료"로 보고하는 것을 막는다.
            if room_capacity is None:
                return "완료"
            _, only_a = related[0]
            assigned_room = str(only_a.get("room", ""))
            need = int(getattr(req, "students", 0) or 0)
            if _room_cap(room_capacity, assigned_room) < need:
                return "미배정"
            return "완료"

        keeps_original = any(
            bool(a.get("keep_orig", True))
            for _, a in related
            if a.get("category") == LABEL_ROOM_SPLIT
        )

        # 기존 강의실 미유지 분반은 "배정 강의실 수용인원 합"이 수강생 이상일 때 완료.
        if not keeps_original:
            if room_capacity is None:
                # 용량 정보가 없으면 보수적으로 최소 2개 배정일 때만 완료로 본다.
                return "완료" if len(related) >= 2 else "미배정"
            total_cap = 0
            seen_rooms = set()
            for _, a in related:
                room = str(a.get("room", ""))
                if room and room not in seen_rooms:
                    total_cap += _room_cap(room_capacity, room)
                    seen_rooms.add(room)
            if total_cap < int(getattr(req, "students", 0) or 0):
                return "미배정"
        return "완료"
    return "미확정"
