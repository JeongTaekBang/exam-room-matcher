# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Project Does

강의실 배정 프로그램 — 시험 요청사항과 기존 시간표를 시각화하여 사람의 수동 배정을 돕는다.
자동 해제 있음 (NO_EXAM, ROOM_CHANGE 원래강의실, NORMAL_EXAM 부분교시). 요청사항 텍스트는 원문 그대로 표시.

## Folder Structure

```
timetable_matcher/
├── {년도}/{학기}_{시험종류}/   # 데이터 (예: 2026/1학기_중간고사/)
│   ├── *요청사항*.xlsx
│   └── *타임테이블*.xlsx
├── data_loader.py             # 엑셀 파싱 + 요청 분류 (표시용)
├── dashboard.py               # Streamlit 대시보드 (5탭)
├── run.bat                    # 더블클릭 실행
└── test_data_loader.py        # 단위 테스트
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

`data_loader.py` — 엑셀 2개를 읽어 구조화된 데이터로 변환. 배정 로직 없음.

`workflow_utils.py` — 배정/해제 JSON 영속화, 감사 로그(JSONL), 시험 교시 계산, 동시 작업 보호(StaleFileError).

`dashboard.py` — Streamlit 5탭 (업무 흐름 순서):
1. **기존 시간표** — 원본 격자 (갈색/청록)
2. **배정 작업** — 진행 요약 + 필터/검색 + 히트맵 + 수급현황 + 미배정→배정(이동/분반/기존유지) + 해제 + 점유현황 격자 + 작업현황
3. **배정 현황** — 진행 요약 + 필터/검색 + 검수 큐 + 완료 목록 + 내보내기/이력
4. **결과 검증** — 요청+배정 오버레이 + 충돌 감지 (빨간 경고)
5. **통계** — 분류별/일별/가동률

## Key Concepts

- **분류**: N열(강의실선택) + K열(시험일자) + O열(요청사항 키워드) 기반 5개 카테고리 (NORMAL_EXAM / NO_EXAM / ROOM_CHANGE / ROOM_SPLIT / SKIP). NO_EXAM 판정 시 요청사항 텍스트에서 "미실시"/"대체과제" 등 키워드를 매칭함 (자유 텍스트 파싱이 아닌 사전 정의 키워드 기반)
- **강의실 결정**: 수업시간표(콤마 구분)에서 시험일 요일 슬롯의 강의실을 `req.room`으로 설정. 강의실 열(10번째) 폴백
- **날짜 매핑**: 시간표 시트 이름(예: "4.21.(화)")에서 날짜↔요일↔시트 매핑을 동적 생성. 연도는 요청 데이터에서 추출
- **충돌 감지**: 같은 강의실+교시에 2개+ 과목 → 빨간 셀 (요청 간, 요청↔배정, 배정 간 모두 감지). ROOM_SPLIT은 `keep_orig=True`일 때 원래 강의실도 충돌 집계에 포함. 같은 과목의 분반 배정끼리는 충돌 아님
- **자동 해제**: NO_EXAM → 전체 교시 해제, ROOM_CHANGE → 원래 강의실 해제, NORMAL_EXAM → 시험 교시 < 수업 교시일 때 미사용 교시 부분 해제. 결과 검증에서 핑크색 표시
- **배정 모드**: 이동(다른 강의실로) / 분반(기존 유지+추가, 기존 미유지+추가) / 기존 강의실 유지. ROOM_CHANGE/ROOM_SPLIT 모두 3가지 모드 선택 가능. 분반은 다중 배정(+N 키)
- **시간표 외 강의실**: 수업시간표 슬롯에는 있지만 시간표 엑셀에 없는 강의실도 결과 검증(초록색), 점유 현황, 해제 대상에 포함
- **배정**: session_state 기반, JSON 파일로 영속화. 원자적 저장(tempfile→rename). 저장 시 파일 mtime을 비교하여 다른 세션의 동시 변경을 감지(StaleFileError), 실패 시 롤백
- **timetable_data**: `{sheet: {room: {period: (value, color_rgb)}}}` — 원본 셀 색상 보존

## Safety Guards

- **동시 작업 보호**: 저장 시 파일 mtime 비교 → 외부 변경 감지 시 저장 차단 + 새로고침 안내 (`workflow_utils.StaleFileError`)
- **전체 초기화**: 체크박스 확인 후에만 버튼 활성화 (오클릭 방지)
- **수동 배정 교시 경고**: 조건 직접 검색에서 선택 교시가 요구 교시와 다를 때 `st.warning` 표시
- **완료 지표 보정**: "완료" 항목 중 검수 큐에 해당하는 건수를 metric help 텍스트로 표시

## Constraints

- 자동 배정 없음 — 시각화/충돌감지/자동해제 전용
- 요청사항 텍스트는 원문 그대로 표시 (단, NO_EXAM 분류 시 사전 정의 키워드 매칭 사용)
- 분류/해제 판단은 구조화된 숫자 컬럼 기반 (요청사항 텍스트에 의존하지 않음)
- Windows + 한글 경로 → `python -X utf8` 필수

## Known Limitations

- **요일→시트 1:1 매핑**: `day_to_sheet`가 동일 요일을 마지막 시트로 덮어씀. 다주차 시험(예: 화요일 2회)에서는 한쪽만 매핑됨. 확장 시 1:N 구조로 변경 필요 (`data_loader.py:160-163`)
