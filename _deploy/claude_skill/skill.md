# Airbnb 시장 분석 도구 스킬 v3

Airbnb 숙소 데이터를 수집하고 시장 분석 리포트를 생성합니다.

## 스킬 시작 시 반드시 먼저 할 것

**스킬이 발동되면 가장 먼저 AskUserQuestion 도구로 아래 질문을 사용자에게 던져라. 답변을 받기 전까지 어떤 작업도 시작하지 마라.**

질문 내용:
- **A — 기본 수집** (빠름): 가격·위치·평점 등 목록 데이터만 Excel 추출. 약 30~60초.
- **B — 상세 수집** (느림): 각 숙소 URL 방문해 소개글·편의시설·호스트 정보 추가 수집. 숙소당 ~2초 × 최대 80개 = 약 3~5분.
- **M — 시장 분석 리포트** (느림): 평일/주말 2개 날짜창 자동 수집 → 통계 분석 → 손님용 Excel + 내부용 Excel + HTML 동시 생성. 약 10~20분.

---

## ⚠ 핵심 정보 (이 컴퓨터 전용 — 다른 컴퓨터 이전 시 반드시 수정)

**프로젝트 경로**: `<<프로젝트_절대경로>>`  
예: `P:/0_지키기/02_PROJECT/99_Working/63_airbnb_price/`

**Python 절대 경로**: `<<Python_절대경로>>`  
확인 방법: Git Bash에서 `which python` 또는 PowerShell에서 `(Get-Command python).Source`

---

## A 모드 실행

```bash
cd "<<프로젝트_절대경로>>"
PYTHONIOENCODING=utf-8 <<Python_절대경로>> \
  export_excel.py <여행지> <체크인YYYY-MM-DD> <체크아웃YYYY-MM-DD>
```

예시:
```bash
PYTHONIOENCODING=utf-8 /c/Users/사용자명/AppData/Local/Programs/Python/Python313/python.exe \
  export_excel.py 홍대 2026-08-01 2026-08-02
```

산출물: `output/raw/airbnb_{지역}_{체크인}_{체크아웃}.xlsx` (14컬럼)

---

## B 모드 실행

```bash
cd "<<프로젝트_절대경로>>"
PYTHONIOENCODING=utf-8 <<Python_절대경로>> \
  export_excel_detail.py <여행지> <체크인YYYY-MM-DD> <체크아웃YYYY-MM-DD>
```

산출물: `output/raw/airbnb_detail_{지역}_{날짜}.xlsx` (22컬럼)

추가 컬럼: 최대인원, 침대종류, 소개글, 편의시설, 호스트명, 슈퍼호스트, 하우스룰, 취소정책

---

## M 모드 실행 (시장 분석 리포트 v3)

```bash
cd "<<프로젝트_절대경로>>"
PYTHONIOENCODING=utf-8 <<Python_절대경로>> \
  market_report.py <지역> [--beds N] [--baths N] [--mode both|client|internal] \
  [--occ-low 0.40] [--occ-base 0.60] [--occ-high 0.70] \
  [--cleaning-fee 80000] [--avg-nights 2.0] [--monthly-cost 0]
```

**기본 실행 예시:**
```bash
# 충신동 3침실 2욕실 시장 분석
PYTHONIOENCODING=utf-8 <<Python_절대경로>> \
  market_report.py 충신동 --beds 3 --baths 2

# 홍대 전체 시장
PYTHONIOENCODING=utf-8 <<Python_절대경로>> \
  market_report.py 홍대

# 손님용만 + 수익 시뮬레이터 커스텀
PYTHONIOENCODING=utf-8 <<Python_절대경로>> \
  market_report.py 이태원 --beds 2 --mode client \
  --cleaning-fee 100000 --monthly-cost 500000
```

산출물 위치: `output/market/`
```
market_report_{지역}_{날짜}_손님용.xlsx   ← 5시트 (손님/클라이언트용)
market_report_{지역}_{날짜}_내부용.xlsx   ← 6시트 (내부 분석용)
market_report_{지역}_{날짜}.html          ← Toss 스타일 HTML
```

### 손님용 시트 구성 (5시트)
| 시트 | 내용 |
|------|------|
| 📊 시장 개요 | KPI 3개 박스 + 서브 지표 |
| 📈 시장 분석 | 가격 구간 분포 + 평일/주말 비교 |
| 🏆 경쟁 현황 | 평점 Top 30 숙소 |
| 💡 추천 객단가 | 스타터/표준/프리미엄 3단계 × 평일/주말 |
| 💰 수익 시뮬레이터 | 보수/기준/공격 시나리오 × 전체 손익 + 시장 매력도 100점 |

### 내부용 시트 구성 (6시트)
| 시트 | 내용 |
|------|------|
| 📋 대시보드 | 전체 KPI + 개인/상업용 비교 + 이상치 목록 |
| 📂 원본 데이터 | 24컬럼 전체 (자동필터) |
| 🔍 세그먼트 분석 | 침실 × 구분 교차표 + 슈퍼호스트 비교 |
| 📊 가격 분포 | P10~P90 + 이상치 + 숙소별 주말 프리미엄 |
| 🗃️ 수집 메타 | 수집 일시, 날짜창, 지오코딩 결과 |
| 📈 수요 추정 | 리뷰 기반 점유율 추정 (저신뢰도 3시나리오) |

---

## 파일 구조

```
프로젝트/
├── airbnb_fetch.py           ← 크롤러 (수정 금지)
├── export_excel.py           ← A 모드
├── export_excel_detail.py    ← B 모드
├── market_report.py          ← M 모드 v3
├── app.py                    ← 웹앱 (선택)
├── requirements.txt
├── templates/index.html
└── output/
    ├── market/               ← M 모드 산출물
    └── raw/                  ← A/B 모드 산출물
```

---

## 웹앱 실행 (선택)

```bash
FLASK_PORT=5001 PYTHONIOENCODING=utf-8 <<Python_절대경로>> app.py
```
→ http://localhost:5001

---

## 주의사항

- `python` 명령이 안 되면 절대 경로 사용 (`which python` 또는 `where python`)
- `PYTHONIOENCODING=utf-8` 없으면 Windows에서 한글 깨짐
- 결과 0개면 지역명 더 넓게 (예: "충무로" → "서울 중구")
- Excel 파일 열려 있으면 저장 오류 → 닫고 재실행
- M 모드 점유율 추정은 저신뢰도 — 참고용으로만 활용
