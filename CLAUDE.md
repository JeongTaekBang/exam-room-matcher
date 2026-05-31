# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Project Does

강의실 배정 프로그램 — 시험 요청사항과 기존 시간표를 시각화하여 사람의 수동 배정을 돕는다.
자동 해제 있음 (NO_EXAM, ROOM_CHANGE 원래강의실, NORMAL_EXAM 부분교시). 요청사항 텍스트는 원문 그대로 표시.

## Folder Structure

```
exam_room_matcher/
├── {년도}/{학기}_{시험종류}/   # 데이터 (예: 2026/1학기_중간고사/)
│   ├── *요청사항*.xlsx
│   ├── *타임테이블*.xlsx
│   ├── _assignments.json      # 배정 결과 (자동 생성)
│   ├── _releases.json         # 해제 결과 (자동 생성)
│   └── _assignment_audit.jsonl
├── data_loader.py             # 엑셀 파싱 + 요청 분류
├── workflow_utils.py          # 영속화 + 감사 로그 + 손상 격리
├── assignment_status.py       # 배정 상태 판정 + Category 라벨
├── conflict_check.py          # 원본 '요청사항 처리'(P열) 강의실 중복 점검
├── dashboard.py               # Streamlit 대시보드 (5탭)
├── doc/                       # 산수·설계·사용 가이드
├── run.bat                    # 더블클릭 실행
├── test_data_loader.py        # 단위 테스트
├── test_workflow_utils.py
├── test_assignment_status.py
└── test_conflict_check.py
```

## Commands

```bash
# 대시보드 실행
streamlit run dashboard.py
# 또는 run.bat 더블클릭

# 테스트
python -X utf8 -m pytest test_data_loader.py -v

# 의존성
pip install openpyxl streamlit pandas pytest
```

## Architecture

`data_loader.py` — 엑셀 2개를 읽어 구조화된 데이터로 변환. 요청 분류는 여기서 수행하지만 배정 자체 로직은 없음. room_choice·NO_EXAM 키워드는 공백을 모두 제거해 정규화 후 비교 (`_normalize_choice`). `load_requests`는 첫 시트(공통)만 읽지만, `load_course_exam_index`는 **모든 시트(공통+전공)** 에서 `(교과목명, 정규화분반)`별 시험일·교시·미실시여부를 모아 ② 시간표 점유자 검증에 쓴다(`normalize_ban`로 분반 표기차 흡수). `load_all` 결과에 `course_exam_index` 포함.

`workflow_utils.py` — 배정/해제 JSON 영속화 (원자적 저장 tempfile→rename), 감사 로그 JSONL, 시험 교시 계산, 동시 작업 보호 (`StaleFileError`), 손상된 JSON의 `.corrupted-*.bak` 자동 격리 (`_quarantine_corrupted`).

`assignment_status.py` — 배정 상태 판정의 순수 함수 (`compute_status`, `_room_cap`, `CAT_LABELS`, `LABEL_*` 상수). dashboard.py에서 분리 — UI 의존성 없이 단위 테스트 가능. assignments JSON의 `category` 필드 값으로 `CAT_LABELS`가 단일 출처. `resolve_processed_room(req, assignments)` — 요청의 최종 강의실 결정을 P열(`요청사항 처리`) 문자열로 환원(배정→강의실, 분반→콤마결합, NO_EXAM→`강의실 미사용`, NORMAL_EXAM→`req.room`, 그 외→빈 문자열). 내보내기에서 P열 자동 채움에 사용.

`conflict_check.py` — 요청 엑셀 P열(16번째, `요청사항 처리`)에 사람이 미리 적은 배정 결정을 읽기 전용으로 감사하는 순수 함수. 프로그램 내부 배정(`_assignments.json`)과 무관. 두 검사:
- `detect_processed_room_conflicts` — **P열↔P열**: 같은 시험일자+강의실에 교시(`exam_start`~`exam_end`)가 겹치는 다른 교과목을 이중 배정으로 검출. 같은 과목명 분반은 형제로 제외, 비강의실 마커(`강의실 미사용`·`확인필요`·빈칸)는 제외, 날짜/교시 누락 행은 `unjudged`로 분리.
- `detect_timetable_overlaps` — **P열↔기존 시간표**: P열 강의실이 그 시험일·교시에 `timetable_data` 점유와 겹치는지 교차 검사. **시험주간 점유는 정상 수업이 아니라 '시험' 기준**이다 — `occupant_index`(`load_course_exam_index`)로 점유자가 그 날 실제로 그 방에서 시험을 보는지 검증해, 사용 안함·시험일 없음·다른 날·교시 안 겹침이면 비운 것으로 본다(공통만 배정하므로 전공이 사용 안함이면 그 방을 공통에 배정 가능). 자기 수업 슬롯(`req.slots`)·해제 슬롯(`all_released_slots`)·같은 교과목 분반도 제외.
- 공통 헬퍼 `parse_processed_rooms`(콤마 분리 + 마커 제외), `_parse_occupant`(시간표 셀의 `과목명-분반`/`과목명분반`/`시험(원래): 과목명-분반` 등 다양한 표기에서 `(과목명, 분반)` 추출).

`dashboard.py` — Streamlit 5탭 (업무 흐름 순서):
1. **기존 시간표** — 원본 격자 (갈색/청록)
2. **배정 작업** — 진행 요약 + 필터/검색 + 히트맵 + 수급현황 + 미배정→배정(이동/분반/기존유지) + 해제 + 점유현황 격자 + 작업현황
3. **배정 현황** — 진행 요약 + 필터/검색 + 검수 큐 + 완료 목록 + 내보내기/이력 (배정 결과를 P열에 채운 **요청 엑셀 사본 다운로드** 포함 — `generate_request_with_processed`, 버튼 클릭 시 생성)
4. **결과 검증** — 상단: 원본 P열(`요청사항 처리`) 강의실 점검(읽기 전용 감사 — ①P열 이중배정 ②기존 시간표 겹침 ③날짜/교시 누락) + 요청+배정 오버레이 + 충돌 감지 (빨간 경고)
5. **통계** — 분류별/일별/가동률

## Key Concepts

- **분류**: N열(강의실선택) + K열(시험일자) + O열(요청사항 키워드) 기반 5개 카테고리 (NORMAL_EXAM / NO_EXAM / ROOM_CHANGE / ROOM_SPLIT / SKIP). NO_EXAM 판정 시 요청사항 텍스트에서 "미실시"/"대체과제" 등 키워드를 매칭함 (자유 텍스트 파싱이 아닌 사전 정의 키워드 기반)
- **강의실 결정**: 수업시간표(콤마 구분)에서 시험일 요일 슬롯의 강의실을 `req.room`으로 설정. 강의실 열(10번째) 폴백
- **날짜 매핑**: 시간표 시트 이름(예: "4.21.(화)")에서 날짜↔요일↔시트 매핑을 동적 생성. 연도는 요청 데이터에서 추출
- **충돌 감지**: 같은 강의실+교시에 2개+ 과목 → 빨간 셀 (요청 간, 요청↔배정, 배정 간 모두 감지). 같은 과목의 분반은 충돌 아님 — 1단계(미배정 요청끼리)는 `req.name` 기준, 2단계(배정 간)는 `+N` suffix를 제거한 base_key 기준으로 형제 필터링
- **자동 해제**: NO_EXAM → 전체 교시 해제, ROOM_CHANGE → 원래 강의실 해제, NORMAL_EXAM → 시험 교시 < 수업 교시일 때 미사용 교시 부분 해제. 결과 검증에서 핑크색 표시
- **배정 모드**: 이동(다른 강의실로) / 분반(기존 유지+추가, 기존 미유지+추가) / 기존 강의실 유지. ROOM_CHANGE/ROOM_SPLIT 모두 3가지 모드 선택 가능. 분반은 다중 배정(+N 키)
- **시간표 외 강의실**: 수업시간표 슬롯에는 있지만 시간표 엑셀에 없는 강의실도 결과 검증(초록색), 점유 현황, 해제 대상에 포함
- **배정**: session_state 기반, JSON 파일로 영속화. 원자적 저장(tempfile→rename). 저장 시 파일 mtime을 비교하여 다른 세션의 동시 변경을 감지(StaleFileError), 실패 시 롤백
- **timetable_data**: `{sheet: {room: {period: (value, color_rgb)}}}` — 원본 셀 색상 보존

## Safety Guards

- **동시 작업 보호**: 저장 시 파일 mtime 비교 → 외부 변경 감지 시 저장 차단 + 새로고침 안내 (`workflow_utils.StaleFileError`)
- **전체 초기화**: 체크박스 확인 후에만 버튼 활성화 (오클릭 방지)
- **취소 2단계 확인**: 배정 작업 현황의 `✕`와 해제 목록의 `취소` 버튼은 한 번 누르면 `✓ 정말?`로 바뀌고 다시 눌러야 실제 삭제 (60초 TTL, 한 번에 하나만 pending)
- **수동 배정 교시 경고**: 조건 직접 검색에서 선택 교시가 요구 교시와 다를 때 `st.warning` 표시
- **완료 지표 분리**: "완료" metric 아래 caption으로 `자동 N / 수동 M` 표시 (자동=NORMAL_EXAM+NO_EXAM, 수동=ROOM_CHANGE/SPLIT 배정 충족). 검수 큐 중복 건수는 help 텍스트로 표시
- **손상 JSON 격리**: `_assignments.json`/`_releases.json` 파싱 실패 시 원본을 `{name}.corrupted-YYYYMMDD_HHMMSS.bak`으로 rename하고 빈 dict 로드. 다음 저장이 정상 파일을 새로 만든다
- **수용인원 검증**: `compute_status`는 ROOM_CHANGE 단일 배정 시에도 배정 강의실 수용인원이 학생 수 미만이면 "미배정" 반환 (UI 가드 우회 케이스 방어)

## Constraints

- 자동 배정 없음 — 시각화/충돌감지/자동해제 전용
- 요청사항 텍스트는 원문 그대로 표시 (단, NO_EXAM 분류 시 사전 정의 키워드 매칭 사용)
- 분류/해제 판단은 구조화된 숫자 컬럼 기반 (요청사항 텍스트에 의존하지 않음)
- Windows + 한글 경로 → `python -X utf8` 필수

## Date/Day/Sheet 매핑

- `date_to_sheet`: 날짜→시트 1:1. 항상 안전. 실호출 경로는 이 매핑을 사용
- `day_to_sheets`: 요일→시트 목록(시간순). 다주차 안전. exam_date가 없는 NO_EXAM 자동 해제에서 같은 요일의 모든 시트에 적용할 때 사용
- `day_to_sheet`: 요일→마지막 시트(하위 호환). `_resolve_sheet`의 폴백 경로용. 호출처가 모두 사전에 `exam_date in DATE_TO_DAY`를 검증하므로 실질 데드 경로지만 안전망으로 유지
