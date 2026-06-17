# Airbnb 시장 분석 도구 — 설치 및 사용 가이드 v3

## 전체 흐름 요약

```
지역명 입력
    ↓
[A 모드] 기본 목록 수집 → Excel (14컬럼, 30~60초)
[B 모드] 상세 수집     → Excel (22컬럼, 숙소당 ~2초)
[M 모드] 시장 분석     → 손님용 Excel + 내부용 Excel + HTML (10~20분)
    ↓
output/ 폴더에 저장
  ├── output/market/  ← M 모드 산출물
  └── output/raw/     ← A/B 모드 수집 파일
```

---

## 폴더 구조

```
project/
├── airbnb_fetch.py           ← Airbnb 크롤러 + 지오코더 (수정 금지)
├── export_excel.py           ← A 모드 CLI
├── export_excel_detail.py    ← B 모드 CLI
├── market_report.py          ← M 모드 CLI (시장 분석 리포트)
├── app.py                    ← Flask 웹앱 (선택)
├── requirements.txt          ← 의존성
├── templates/
│   └── index.html            ← 웹 UI
└── output/                   ← 산출물 (자동 생성)
    ├── market/               ← M 모드 xlsx + html
    └── raw/                  ← A/B 모드 xlsx
```

---

## 1단계 — Python 설치

- Python **3.10 이상** 필수 (3.13 권장): https://www.python.org/downloads/
- 설치 시 **"Add Python to PATH"** 반드시 체크

설치 확인:
```bash
python --version
```

---

## 2단계 — 프로젝트 배치

`project/` 폴더를 원하는 위치에 복사합니다.

예: `C:\Users\사용자명\airbnb_price\`

---

## 3단계 — 의존성 설치

```bash
cd C:\Users\사용자명\airbnb_price
pip install -r requirements.txt
```

설치 패키지: `flask`, `curl_cffi`, `xlsxwriter`, `requests`, `gunicorn`

---

## 4단계 — 실행

> **Windows CMD/PowerShell 사용 시 한글 깨짐 방지:**
> ```powershell
> $env:PYTHONIOENCODING="utf-8"
> ```
> Git Bash: `PYTHONIOENCODING=utf-8 python ...`

---

### A 모드 — 기본 목록 수집 (빠름, ~30~60초)

```bash
python export_excel.py <지역> <체크인YYYY-MM-DD> <체크아웃YYYY-MM-DD>
```

예시:
```bash
python export_excel.py 충무로 2026-07-07 2026-07-09
```

산출물: `airbnb_충무로_2026-07-07_2026-07-09.xlsx` (14컬럼)

| 컬럼 | 내용 |
|------|------|
| 번호, 숙소명, 숙소유형, 건물유형 | 기본 정보 |
| 침실, 침대, 욕실, 평점 | 숙소 스펙 |
| 1박가, 총가 | 가격 |
| 링크, 위도, 경도, 지역쿼리 | 위치/URL |

---

### B 모드 — 상세 수집 (느림, 숙소당 ~2초)

```bash
python export_excel_detail.py <지역> <체크인YYYY-MM-DD> <체크아웃YYYY-MM-DD>
```

예시:
```bash
python export_excel_detail.py 충무로 2026-07-07 2026-07-09
```

산출물: `airbnb_detail_충무로_2026-07-07_2026-07-09.xlsx` (22컬럼)

A 모드 14컬럼 + 추가 8컬럼:

| 추가 컬럼 | 내용 |
|-----------|------|
| 최대인원 | 숙소 최대 게스트 수 |
| 침대종류 | 침실별 침대 유형 |
| 소개글 | 숙소 전체 소개 텍스트 |
| 편의시설 | 카테고리별 전체 편의시설 |
| 호스트명 | 호스트 이름 |
| 슈퍼호스트 | 슈퍼호스트 / 일반 |
| 하우스룰 | 체크인 시간, 반려동물, 흡연 규정 |
| 취소정책 | 환불 정책 (날짜 미선택 시 빈 값 — 정상) |

---

### M 모드 — 시장 분석 리포트 v3 (느림, ~10~20분)

```bash
python market_report.py <지역> [옵션]
```

**기본 옵션:**

| 옵션 | 기본값 | 설명 |
|------|--------|------|
| `--beds N` | 없음 | 침실 수 필터 |
| `--baths N` | 없음 | 욕실 수 필터 |
| `--mode` | both | both / client / internal |

**v3 수익 시뮬레이터 옵션:**

| 옵션 | 기본값 | 설명 |
|------|--------|------|
| `--occ-low` | 0.40 | 보수 점유율 (40%) |
| `--occ-base` | 0.60 | 기준 점유율 (60%) |
| `--occ-high` | 0.70 | 공격 점유율 (70%) |
| `--cleaning-fee` | 80000 | 청소비 / 회 (원) |
| `--avg-nights` | 2.0 | 평균 체류일 (박) |
| `--monthly-cost` | 0 | 월 고정 운영비 (원, 청소비 제외) |

**사용 예시:**

```bash
# 기본 실행 (전체 시장)
python market_report.py 홍대

# 침실/욕실 필터 + 수익 시뮬레이터 커스텀
python market_report.py 충신동 --beds 3 --baths 2 --cleaning-fee 100000 --monthly-cost 500000

# 손님용 리포트 + HTML만 생성
python market_report.py 이태원 --beds 2 --mode client

# 내부 분석만
python market_report.py 해운대 --mode internal
```

**산출물** (`output/market/` 폴더에 자동 저장):

```
market_report_{지역}_{날짜}_손님용.xlsx
market_report_{지역}_{날짜}_내부용.xlsx
market_report_{지역}_{날짜}.html
```

---

## M 모드 산출물 상세

### 손님용 Excel (5 시트)

| 시트 | 내용 |
|------|------|
| 📊 시장 개요 | KPI 3개 + 서브 지표 (클라이언트 첫 화면) |
| 📈 시장 분석 | 가격 구간 분포 + 평일/주말 비교 + 침실별 평균가 |
| 🏆 경쟁 현황 | 평점 Top 30 (상위 5개 골드 강조) |
| 💡 추천 객단가 | 스타터(P40) / 표준(P55) / 프리미엄(P70) × 평일/주말 |
| 💰 수익 시뮬레이터 | 보수/기준/공격 시나리오 × 전체 손익 구조 + 시장 매력도 점수 |

### 내부용 Excel (6 시트)

| 시트 | 내용 |
|------|------|
| 📋 대시보드 | 전체 KPI + 개인/상업용 비교 + 이상치 목록 |
| 📂 원본 데이터 | 전체 24컬럼 (자동필터 포함) |
| 🔍 세그먼트 분석 | 침실 × 개인/상업용 교차표 + 슈퍼호스트 vs 일반 |
| 📊 가격 분포 | P10~P90 전체 + 이상치 목록 + 숙소별 주말 프리미엄 |
| 🗃️ 수집 메타 | 수집 일시, 날짜창, 지오코딩 결과, 방법론 |
| 📈 수요 추정 | 리뷰 기반 점유율 추정 (저신뢰도, 3개 시나리오) |

### HTML (Toss 스타일, 자기완결형)

- 외부 의존성 없음 — 오프라인 동작, 이메일 첨부 가능
- 섹션: Hero KPI → 시장현황 → 가격 구간 분포 → 경쟁현황 → 추천 객단가 → **시장 매력도 점수** → **월 수익 시뮬레이터**
- 모바일 반응형

---

## M 모드 분석 방법론

### 데이터 수집
- 3주 후 **평일창** (월→화) + **주말창** (금→토) 2개 날짜로 자동 수집
- Airbnb 검색 결과 → `data-deferred-state-0` JSON 파싱 (curl_cffi, Chrome 120 핑거프린트)
- 각 숙소 상세 URL 방문 → 22컬럼 수집, 숙소당 2초 딜레이

### 가격 통계
- **이상치 제거**: IQR × 1.5 (Q1 − 1.5×IQR ~ Q3 + 1.5×IQR)
- **퍼센타일**: P10 / P25 / P40 / P50 / P55 / P70 / P75 / P90 (선형보간)
- **주말 프리미엄**: 양쪽 창에 공통으로 나타난 숙소만으로 계산

### 분류 로직
상업용/개인 스코어링 (2점 이상 = 상업용):
- `property_type`에 호텔/게스트하우스 키워드: +3점
- `host_name`에 상업 키워드: +2점
- 소개글에 운영 키워드 2개+: +2점
- 숙소명에 호텔/게스트하우스: +1점

### 추천 객단가 3단계
| 단계 | 기준 | 전략 |
|------|------|------|
| 🌱 스타터 | P40 | 초기 리뷰 확보 |
| ⭐ 표준 | P55 | 리뷰 10개+ 기준 |
| 👑 프리미엄 | P70 | 슈퍼호스트 목표 |

### 시장 매력도 점수 (v3)
4개 지표 × 25점 = 100점:
1. **수요 안정성** — 평균 리뷰 수 (40개+ = 25점)
2. **주말 수요 탄력성** — 주말 프리미엄 (30%+ = 25점)
3. **가격 성장 여력** — (P70−P40)/P40 (70%+ = 25점)
4. **진입 품질 허들** — 경쟁자 평점 역점수 (낮을수록 차별화 쉬움)

### 수익 시뮬레이터 (v3)
- 블렌디드 ADR = 평일 ADR × 5/7 + 주말 ADR × 2/7
- 호스트 수수료 3% 차감
- 임대료/대출 상환 미포함 — 별도 차감 필요

---

## 웹앱 실행 (선택)

```bash
python app.py
```

→ http://localhost:5001 (지도 UI + 검색 + Excel 다운로드)

---

## Claude Code 스킬 등록 (선택)

Claude Code를 사용하는 경우 자연어로 실행 가능합니다.

1. `claude_skill/skill.md`를 아래 경로에 복사:
   ```
   C:\Users\사용자명\.claude\skills\airbnb-price-excel\skill.md
   ```

2. `skill.md` 안의 두 경로를 실제 환경에 맞게 수정:
   - **프로젝트 경로**: `project/` 폴더를 복사한 실제 위치
   - **Python 경로**: `where python` 으로 확인한 절대 경로

---

## 주의사항

| 항목 | 내용 |
|------|------|
| Python 경로 | `python` 명령이 안 되면 절대 경로 사용 (`where python` 으로 확인) |
| 한글 깨짐 | Windows CMD: `chcp 65001` 또는 `PYTHONIOENCODING=utf-8` |
| 결과 0개 | 지역명을 더 넓게 입력 (예: "충무로" → "서울 중구") |
| 소요 시간 | A모드 ~1분 / B모드 80개 기준 ~3분 / M모드 ~15분 |
| 엑셀 열려있음 | 파일 이미 열려있으면 저장 오류 → 닫고 재실행 |
| 취소정책 | 날짜 선택 전에는 Airbnb가 미노출 → 빈 값 정상 |
| 점유율 추정 | 리뷰 기반 저신뢰도 추정 — 의사결정 보조용으로만 활용 |
