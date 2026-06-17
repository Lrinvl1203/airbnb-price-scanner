# market_report.py 로직 및 산출 방법 문서

최종 업데이트: 2026-06-17  
버전: v1.0

---

## 1. 전체 흐름

```
입력 (CLI)
  지역명 + [침실 수] + [욕실 수] + [모드]
       ↓
지오코딩 (OSM Nominatim)
  지역명 → 위도/경도
       ↓
데이터 수집 × 2회 (B 모드)
  날짜창 1: 평일 (3주 후 월→화)
  날짜창 2: 주말 (3주 후 금→토)
       ↓
URL 기준 머지
  동일 숙소의 평일가 + 주말가 통합
       ↓
침실/욕실 필터 적용
       ↓
상업용 vs 개인 분류
       ↓
통계 산출
       ↓
출력 3종
  손님용.xlsx  내부용.xlsx  .html
```

---

## 2. 날짜창 선택 로직

```python
base   = today + timedelta(weeks=3)
monday = base - timedelta(days=base.weekday())   # 해당 주 월요일
friday = monday + timedelta(days=4)               # 같은 주 금요일

평일창: monday → monday+1 (1박)
주말창: friday → friday+1 (1박)
```

**선택 이유:**
- 3주 후 = 예약 가능성이 충분히 열려있는 가까운 미래
- 같은 주의 월·금 비교 → 계절 편차 최소화
- 1박 기준 = Airbnb의 가장 기본 단위

---

## 3. 데이터 수집 (B 모드)

`airbnb_fetch.crawl_airbnb()` → 최대 80개 숙소 목록 수집  
`export_excel_detail.fetch_detail()` → 각 숙소 URL 방문, 상세 8컬럼 추가  
숙소당 2초 딜레이 (봇 차단 방지)

수집 컬럼 (22개):
- 기본 14: 번호, 숙소명, 숙소유형, 건물유형, 침실, 침대, 욕실, 평점, 1박가, 총가, 링크, 위도, 경도, 지역쿼리
- 상세 8: 최대인원, 침대종류, 소개글, 편의시설, 호스트명, 슈퍼호스트, 하우스룰, 취소정책

---

## 4. URL 기준 머지

```python
all_urls = set(wd_by_url) | set(we_by_url)

for url in all_urls:
    base = wd_by_url.get(url) or we_by_url.get(url)  # 평일 우선
    listing['price_weekday'] = wd_price or None
    listing['price_weekend'] = we_price or None
    listing['price_per_night'] = price_weekday or price_weekend
```

결과: 숙소별로 평일가/주말가를 동시에 보유. 한 창에만 나타난 숙소는 해당 가격만 보유.

---

## 5. 상업용 vs 개인 분류

스코어링 방식 (2점 이상 → 상업용):

| 신호 | 점수 | 판단 기준 |
|------|------|-----------|
| `property_type`에 상업 키워드 | **+3** | 호텔, hotel, 게스트하우스, guesthouse, 호스텔, hostel, b&b 등 |
| `host_name`에 상업 키워드 | **+2** | 스테이, 레지던스, 리조트, inn, stay, 펜션 등 |
| `description`에 운영 키워드 2개+ | **+2** | 프론트, 리셉션, 체크인 카운터, 24시간 운영 |
| `title`에 상업 키워드 | **+1** | 호텔, 게스트하우스, hotel 등 |

**설계 의도:** 단일 신호 오류 방지. property_type이 "아파트"여도 호스트명에 "호텔"이 2개 이상 신호와 결합되면 상업용으로 분류.

---

## 6. 가격 통계 산출

### 6-1. 이상치 탐지 (IQR 방법)

```python
q1  = percentile(prices, 25)
q3  = percentile(prices, 75)
iqr = q3 - q1

outlier_low  = q1 - 1.5 * iqr   # 이 미만은 이상치
outlier_high = q3 + 1.5 * iqr   # 이 초과는 이상치
```

이상치를 **제거한** 클린 데이터셋으로 `mean_clean` 산출 → 손님용 리포트의 "평균 1박가" 기준값으로 사용.  
이상치는 내부용 리포트에서 별도 플래그(`⚠ 이상치`) 표시.

### 6-2. 퍼센타일 계산 (선형보간)

```python
def _percentile(data, p):
    idx = (len(sorted_data) - 1) * p / 100
    lo  = int(idx)
    hi  = min(lo + 1, len(data) - 1)
    return data[lo] + (data[hi] - data[lo]) * (idx - lo)
```

산출 퍼센타일: P10, P25, **P40**, P50, **P55**, **P70**, P75, P90

### 6-3. 주말 프리미엄

```python
# 양쪽 창 모두 존재하는 숙소만으로 계산 (공정한 비교)
pairs = [(l.price_weekday, l.price_weekend) for l in listings
         if l.price_weekday and l.price_weekend]

wd_avg  = mean(pair[0] for pair in pairs)
we_avg  = mean(pair[1] for pair in pairs)
premium = (we_avg - wd_avg) / wd_avg
```

**한 창에만 있는 숙소를 제외하는 이유:** 평일창에만 나타나는 숙소 ≠ 주말 비예약 → 단순 가용 차이일 수 있으므로, 동일 숙소 비교만 유효.

---

## 7. 추천 객단가 산출

| 단계 | 퍼센타일 | 설명 |
|------|----------|------|
| 🌱 스타터 | P40 | 시장 하위 40% → 초기 리뷰 확보 우선 |
| ⭐ 표준 | P55 | 시장 중심 포지션 → 리뷰 10개+ 기준 |
| 👑 프리미엄 | P70 | 상위 30% → 슈퍼호스트 등급 목표 |

```python
weekday_price = round(p_value / 1000) * 1000   # 천원 단위 반올림
weekend_price = round(p_value * (1 + premium) / 1000) * 1000
```

**반올림 이유:** 실제 호스팅 환경에서 ₩67,834보다 ₩68,000이 자연스럽고 신뢰감을 준다.

---

## 8. 가격 구간 정의

| 구간 | 범위 |
|------|------|
| ~₩3만 | 0 ~ 30,000 |
| ₩3만~₩5만 | 30,000 ~ 50,000 |
| ₩5만~₩8만 | 50,000 ~ 80,000 |
| ₩8만~₩12만 | 80,000 ~ 120,000 |
| ₩12만~₩20만 | 120,000 ~ 200,000 |
| ₩20만+ | 200,000 ~ ∞ |

**설계 의도:** 서울 기준 단기숙소 시장 실제 분포 반영. 5만 이하는 세분화(게스트룸/공유), 5~8만이 개인 1~2룸 코어, 8만+ 이상이 프리미엄/대형.

---

## 9. 출력 포맷

### 손님용 Excel (4 시트)
| 시트 | 목적 |
|------|------|
| 📊 시장 개요 | KPI 3개 박스 + 서브 지표 (클라이언트 첫 화면) |
| 📈 시장 분석 | 가격 구간 분포(데이터바) + 평일/주말 비교 + 침실별 평균가 |
| 🏆 경쟁 현황 | 평점 Top 30, 상위 5개 골드 강조, 평점 4.8+ 초록 조건부 서식 |
| 💡 추천 객단가 | 3단계 × 평일/주말 가격 테이블 + 이상치 제거 샘플 수 근거 |

### 내부용 Excel (5 시트)
| 시트 | 목적 |
|------|------|
| 📋 대시보드 | 전체 KPI + 개인/상업용 비교 + 이상치 목록 |
| 📂 원본 데이터 | 24컬럼 (22 기본 + 분류 + 이상치 플래그) + 자동필터 |
| 🔍 세그먼트 분석 | 침실 수 × 개인/상업용 교차표 + 슈퍼호스트 vs 일반 비교 |
| 📊 가격 분포 | P10~P90 전체 + 이상치 목록(URL 포함) + 숙소별 주말 프리미엄 |
| 🗃️ 수집 메타 | 수집 일시/날짜창/지오코딩/방법론/분류 로직 버전 |

### HTML (Toss 스타일, 자기완결형)
- 외부 의존성 없음 (CDN 불필요, 오프라인 동작)
- 섹션: Hero(KPI 3개) → 시장현황(4 메트릭 + 비교바) → 가격 구간 분포(CSS 바) → 경쟁현황 Top 15 → 추천객단가 3티어
- 반응형 (모바일 지원)

---

## 10. 알려진 한계

| 항목 | 내용 |
|------|------|
| 점유율 | Airbnb가 직접 노출하지 않음 → 이 버전에서는 미포함 |
| 샘플 수 | 특정 조건(3침실+2욕실 등)은 숙소 수 적을 수 있음 → 지역 확장 권장 |
| 분류 정확도 | 스코어링 기반이므로 엣지케이스(개인이 "Stay" 이름 사용 등) 오분류 가능 |
| 계절성 | 단일 날짜창 2개 수집 → 성수기/비수기 차이 미반영 |
| 취소정책 | 날짜 선택 전 Airbnb가 미노출 → 공란 정상 |

---

## 11. CLI 레퍼런스

```bash
python market_report.py <지역> [--beds N] [--baths N] [--mode both|client|internal]

# 예시
python market_report.py 충신동 --beds 3 --baths 2
python market_report.py 홍대                           # 전체 시장
python market_report.py 이태원 --beds 2 --mode client  # 손님용 + HTML만
```

출력 파일명 패턴:
```
market_report_{지역}_{YYYY-MM-DD}_손님용.xlsx
market_report_{지역}_{YYYY-MM-DD}_내부용.xlsx
market_report_{지역}_{YYYY-MM-DD}.html
```
