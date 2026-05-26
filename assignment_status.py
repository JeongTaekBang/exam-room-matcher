"""배정 상태 판정 로직.

dashboard.py에서 분리한 UI 의존 없는 순수 함수들. Category 라벨·
수용인원 헬퍼·compute_status를 한 곳에 모아 단위 테스트가 가능하도록 한다.
"""

from data_loader import Category


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
