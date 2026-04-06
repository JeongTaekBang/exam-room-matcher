# 강의실 배정 프로그램

시험 요청사항과 기존 강의실 시간표를 시각화하여 운영자의 수동 배정을 돕는 Streamlit 대시보드.

## 주요 기능

- **5탭 대시보드**: 기존 시간표 / 배정 작업 / 배정 현황 / 결과 검증 / 통계
- **배정 모드**: 이동 (다른 강의실로) / 분반 (기존 유지 + 추가) / 기존 강의실 유지
- **자동 해제**: 미실시(NO_EXAM), 강의실 이동 원래강의실, 부분 교시 미사용 → 자동 해제 + 빈 강의실 풀 반환
- **충돌 감지**: 같은 강의실 + 교시에 2개+ 과목 겹침 시 빨간 셀 경고
- **시각화**: 점유 현황 격자, 수급 현황 (공급 vs 수요), 잔여 슬롯 히트맵
- **시간표 외 강의실**: 수업시간표에만 있는 강의실도 검증/해제/점유 현황에 표시
- **다주차 안전**: 날짜 → 시트 직접 매핑으로 동일 요일 다중 시트 지원
- **안정성**: 원자적 저장, 동시 작업 보호 (mtime 비교), 실패 시 롤백

## 전체 흐름

```
입력 (엑셀 2개)
  요청사항 + 타임테이블
        |
        v
  데이터 파싱/분류 (data_loader.py)
  5개 카테고리: 시험진행 / 미실시 / 강의실변경 / 강의실분반 / 미확정
        |
        v
  AS-IS: 기존 시간표 (원본 그대로)
        |
        v
  자동 해제 (미실시/부분교시/이동 원래강의실)
        |
        v
  배정 작업 (운영자 수동)
  히트맵 → 수급 현황 → 과목 선택 → 배정 방식 → 빈 강의실 → 배정
        |
        v
  TO-BE: 결과 검증 (오버레이 + 충돌 감지)
        |
        v
  내보내기 (엑셀/CSV 다운로드)
```

## 설치 및 실행

```bash
# 의존성 설치
pip install openpyxl streamlit pandas pytest

# 대시보드 실행
streamlit run dashboard.py

# 또는 Windows에서 더블클릭
run.bat
```

## 데이터 구조

```
exam_room_matcher/
├── {년도}/{학기}_{시험종류}/     # 데이터 폴더
│   ├── *요청사항*.xlsx          # 시험 요청 엑셀
│   ├── *타임테이블*.xlsx        # 강의실 시간표 엑셀
│   ├── _assignments.json       # 배정 결과 (자동 생성)
│   ├── _releases.json          # 해제 결과 (자동 생성)
│   └── _assignment_audit.jsonl # 작업 이력 (자동 생성)
├── data_loader.py              # 엑셀 파싱 + 요청 분류
├── workflow_utils.py           # 영속화 + 감사 로그
├── dashboard.py                # Streamlit 대시보드
├── test_data_loader.py         # 테스트
├── test_workflow_utils.py      # 테스트
└── run.bat                     # 실행 스크립트
```

## 테스트

```bash
python -X utf8 -m pytest test_data_loader.py test_workflow_utils.py -v
```

> Windows + 한글 경로 환경에서는 `python -X utf8` 필수.

## 라이선스

MIT
