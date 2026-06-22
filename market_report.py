"""
Airbnb 시장 분석 리포트 생성기
사용: python market_report.py <지역> [--beds N] [--baths N] [--mode both|client|internal]
출력: market_report_<지역>_<날짜>_손님용.xlsx + 내부용.xlsx
"""
from __future__ import annotations

import argparse
import math
import re
import shutil
import statistics
import sys
import tempfile
import time
import zipfile
from datetime import date, timedelta
from pathlib import Path

import xlsxwriter

sys.path.insert(0, str(Path(__file__).parent))
from airbnb_fetch import crawl_airbnb, geocode_region
from export_excel import _fix_colors, C_DARK_BLUE, C_MID_BLUE, C_LIGHT_BLUE, C_WHITE, C_LINK
from export_excel_detail import fetch_detail, DETAIL_COLS
from curl_cffi import requests as cf_requests

# ── 추가 색상 ──────────────────────────────────────────────────────
C_GREEN        = "#70AD47"
C_GOLD         = "#FFB900"
C_GOLD_BG      = "#FFF2CC"
C_RED          = "#C00000"
C_ORANGE       = "#ED7D31"
C_ORANGE_BG    = "#FCE4D6"
C_LIGHT_GREEN  = "#E2EFDA"
C_LIGHT_RED    = "#FFDAD5"
C_GRAY         = "#808080"
C_LIGHT_GRAY   = "#F2F2F2"
C_ACCENT_BLUE  = "#4472C4"
C_ACCENT_BG    = "#DEEAF1"
C_DARK_TEXT    = "#1F2D3D"
C_SECTION_BG   = "#D6E4F0"
C_YELLOW       = "#FFC000"
C_YELLOW_BG    = "#FFEB9C"

# ── 가격 구간 ──────────────────────────────────────────────────────
PRICE_BUCKETS = [
    (0,       30_000,       "~₩3만"),
    (30_000,  50_000,       "₩3만~₩5만"),
    (50_000,  80_000,       "₩5만~₩8만"),
    (80_000,  120_000,      "₩8만~₩12만"),
    (120_000, 200_000,      "₩12만~₩20만"),
    (200_000, float("inf"), "₩20만+"),
]

# ── 상업용 분류 키워드 ─────────────────────────────────────────────
_COMMERCIAL_PROP_TYPES = [
    "호텔", "hotel", "게스트하우스", "guesthouse", "guest house",
    "호스텔", "hostel", "bed and breakfast", "b&b", "부티크 호텔",
    "부티크호텔", "boutique hotel",
]
_COMMERCIAL_HOST_KW = [
    "호텔", "게스트하우스", "호스텔", "스테이", "레지던스", "리조트",
    "펜션", "모텔", "hotel", "hostel", "guesthouse", "inn", "stay",
    "하우스", "숙소", "residence", "게스트 하우스",
]
_COMMERCIAL_DESC_KW = [
    "프론트", "리셉션", "체크인 카운터", "24시간 운영",
    "프론트 데스크", "front desk", "reception",
]


# ══════════════════════════════════════════════════════════════════
# 1. 날짜창 계산
# ══════════════════════════════════════════════════════════════════

def get_date_windows(target: date | None = None) -> list[tuple[str, str, str]]:
    """지정 날짜가 속한 주의 평일(월→화)과 주말(금→토) 반환.
    target 미지정 시 오늘 기준 3주 후 주간 사용."""
    base   = target if target else (date.today() + timedelta(weeks=3))
    monday = base - timedelta(days=base.weekday())
    friday = monday + timedelta(days=4)
    return [
        (monday.isoformat(), (monday + timedelta(days=1)).isoformat(), "평일"),
        (friday.isoformat(), (friday + timedelta(days=1)).isoformat(), "주말"),
    ]


# ══════════════════════════════════════════════════════════════════
# 2. 분류 로직
# ══════════════════════════════════════════════════════════════════

def classify_host_type(listing: dict) -> str:
    """상업용(호텔·게스트하우스 등) vs 개인 분류. 스코어 2점 이상 = 상업용."""
    score = 0

    prop = (listing.get("property_type") or "").lower()
    if any(k in prop for k in _COMMERCIAL_PROP_TYPES):
        score += 3

    host = listing.get("host_name") or listing.get("title") or ""
    if any(k in host for k in _COMMERCIAL_HOST_KW):
        score += 2

    desc = listing.get("description") or ""
    hit  = sum(1 for k in _COMMERCIAL_DESC_KW if k in desc)
    if hit >= 2:
        score += 2

    title = listing.get("title") or ""
    if any(k in title for k in ["호텔", "게스트하우스", "호스텔", "hotel"]):
        score += 1

    return "상업용" if score >= 2 else "개인"


# ══════════════════════════════════════════════════════════════════
# 3. 데이터 수집 및 머지
# ══════════════════════════════════════════════════════════════════

def fetch_calendar_occ(
    listing_url: str,
    session: cf_requests.Session,
    months: int = 3,
) -> float | None:
    """Airbnb 캘린더 API로 향후 N개월 예약률(0~1) 수집. 실패 시 None."""
    m = re.search(r"/rooms/(\d+)", listing_url)
    if not m:
        return None
    listing_id = m.group(1)
    today = date.today()
    try:
        r = session.get(
            "https://www.airbnb.co.kr/api/v2/calendar_months",
            params={
                "key":        "d306zoyjsyarp7ifhu67rjxn52tv0t20",
                "listing_id": listing_id,
                "month":      today.month,
                "year":       today.year,
                "count":      months,
                "_format":    "with_conditions",
                "locale":     "ko",
                "currency":   "KRW",
            },
            timeout=10,
        )
        if r.status_code != 200:
            return None
        data = r.json()
        total = 0
        booked = 0
        for month_data in data.get("calendar_months", []):
            for day in month_data.get("days", []):
                try:
                    d = date.fromisoformat(day["date"])
                except Exception:
                    continue
                if d < today:
                    continue
                total += 1
                if not day.get("available", True):
                    booked += 1
        return booked / total if total > 0 else None
    except Exception:
        return None


def _fetch_window(
    query: str,
    checkin: str,
    checkout: str,
    label: str,
    geo: dict,
    session: cf_requests.Session,
) -> list[dict]:
    print(f"\n  [{label}] 크롤링: {checkin} ~ {checkout}")
    listings = crawl_airbnb(query, checkin, checkout, geo=geo, max_results=80)
    print(f"  [{label}] {len(listings)}개 수집 → 상세 크롤링 + 예약률 수집...")
    for i, lst in enumerate(listings, 1):
        detail = fetch_detail(lst["url"], session)
        lst.update(detail)
        lst["calendar_occ"] = fetch_calendar_occ(lst["url"], session)
        ok = bool(detail.get("description") and
                  not detail["description"].startswith("오류"))
        occ_str = f"{lst['calendar_occ']*100:.0f}%" if lst.get("calendar_occ") is not None else "?"
        print(f"    [{i:>2}/{len(listings)}] {'✅' if ok else '⚠'} 예약률:{occ_str}", end="\r")
        if i < len(listings):
            time.sleep(2.0)
    print()
    return listings


def collect_data(
    query: str,
    geo: dict,
    beds: int | None,
    baths: float | None,
    target_date: date | None = None,
) -> dict:
    """두 날짜창 수집 → 머지 → 필터 → 분류 결과 반환."""
    windows = get_date_windows(target_date)
    session = cf_requests.Session()
    session.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        ),
        "Accept-Language": "ko-KR,ko;q=0.9",
    })

    wd_in, wd_out, wd_label = windows[0]
    we_in, we_out, we_label = windows[1]

    wd_raw = _fetch_window(query, wd_in, wd_out, wd_label, geo, session)
    we_raw = _fetch_window(query, we_in, we_out, we_label, geo, session)

    # ── 머지 ─────────────────────────────────────────────────────
    wd_by_url = {l["url"]: l for l in wd_raw}
    we_by_url = {l["url"]: l for l in we_raw}
    all_urls  = set(wd_by_url) | set(we_by_url)

    merged: list[dict] = []
    for url in all_urls:
        wd = wd_by_url.get(url)
        we = we_by_url.get(url)
        base = (wd or we).copy()
        base["price_weekday"] = wd["price_per_night"] if wd else None
        base["price_weekend"] = we["price_per_night"] if we else None
        base["price_per_night"] = base["price_weekday"] or base["price_weekend"]
        merged.append(base)

    # ── 침실/욕실 필터 ────────────────────────────────────────────
    if beds is not None:
        merged = [l for l in merged if (l.get("bedrooms") or 0) == beds]
    if baths is not None:
        merged = [l for l in merged if abs((l.get("bathrooms") or 0) - baths) < 0.5]

    # ── 분류 ─────────────────────────────────────────────────────
    for lst in merged:
        lst["host_type"] = classify_host_type(lst)

    # 분류 후 예약률 집계
    occ_vals = [l["calendar_occ"] for l in merged if l.get("calendar_occ") is not None]
    avg_calendar_occ = statistics.mean(occ_vals) if occ_vals else None

    return {
        "listings": merged,
        "windows":  windows,
        "wd_count": len(wd_raw),
        "we_count": len(we_raw),
        "common_count": len(set(wd_by_url) & set(we_by_url)),
        "geo": geo,
        "avg_calendar_occ":    avg_calendar_occ,
        "calendar_occ_count":  len(occ_vals),
    }


# ══════════════════════════════════════════════════════════════════
# 4. 통계 계산
# ══════════════════════════════════════════════════════════════════

def _percentile(data: list[float], p: float) -> float:
    if not data:
        return 0.0
    sd = sorted(data)
    idx = (len(sd) - 1) * p / 100
    lo  = int(idx)
    hi  = min(lo + 1, len(sd) - 1)
    return sd[lo] + (sd[hi] - sd[lo]) * (idx - lo)


def compute_stats(listings: list[dict]) -> dict:
    prices  = [l["price_per_night"] for l in listings if l.get("price_per_night")]
    ratings = [l["rating"] for l in listings if l.get("rating")]
    if not prices:
        return {"count": 0}

    q1  = _percentile(prices, 25)
    q3  = _percentile(prices, 75)
    iqr = q3 - q1
    lo_fence = q1 - 1.5 * iqr
    hi_fence = q3 + 1.5 * iqr
    clean = [p for p in prices if lo_fence <= p <= hi_fence]

    return {
        "count":         len(prices),
        "clean_count":   len(clean),
        "outlier_count": len(prices) - len(clean),
        "mean":          statistics.mean(prices),
        "mean_clean":    statistics.mean(clean) if clean else 0,
        "median":        statistics.median(prices),
        "min":           min(prices),
        "max":           max(prices),
        "stdev":         statistics.stdev(prices) if len(prices) > 1 else 0,
        "p10":  _percentile(prices, 10),
        "p25":  q1,
        "p40":  _percentile(prices, 40),
        "p50":  _percentile(prices, 50),
        "p55":  _percentile(prices, 55),
        "p70":  _percentile(prices, 70),
        "p75":  q3,
        "p90":  _percentile(prices, 90),
        "outlier_low":  lo_fence,
        "outlier_high": hi_fence,
        "avg_rating":   statistics.mean(ratings) if ratings else 0,
        "rating_count": len(ratings),
    }


def compute_weekend_premium(listings: list[dict]) -> float:
    """공통 숙소(평일+주말 가격 모두 있음) 기준 주말 프리미엄."""
    pairs = [
        (l["price_weekday"], l["price_weekend"])
        for l in listings
        if l.get("price_weekday") and l.get("price_weekend")
    ]
    if not pairs:
        return 0.0
    wd_avg = statistics.mean(p[0] for p in pairs)
    we_avg = statistics.mean(p[1] for p in pairs)
    return (we_avg - wd_avg) / wd_avg if wd_avg > 0 else 0.0


def price_distribution(listings: list[dict]) -> list[dict]:
    prices_all = [l["price_per_night"] for l in listings if l.get("price_per_night")]
    result = []
    for lo, hi, label in PRICE_BUCKETS:
        indiv = [l for l in listings
                 if l.get("price_per_night") and lo <= l["price_per_night"] < hi
                 and l.get("host_type") == "개인"]
        comm  = [l for l in listings
                 if l.get("price_per_night") and lo <= l["price_per_night"] < hi
                 and l.get("host_type") == "상업용"]
        count = len(indiv) + len(comm)
        result.append({
            "label":  label,
            "count":  count,
            "indiv":  len(indiv),
            "comm":   len(comm),
            "pct":    count / len(prices_all) * 100 if prices_all else 0,
        })
    return result


def is_outlier(listing: dict, stats: dict) -> bool:
    p = listing.get("price_per_night") or 0
    return p < stats.get("outlier_low", 0) or p > stats.get("outlier_high", float("inf"))


# ══════════════════════════════════════════════════════════════════
# 4-V3. v3 추가 분석 함수들
# ══════════════════════════════════════════════════════════════════

def compute_effective_adr(listings: list[dict]) -> dict:
    """총가(= 1박가 + 청소비 + Airbnb 서비스료) 기준 체감 ADR."""
    totals = [l.get("total_price") or l.get("price_per_night")
              for l in listings if l.get("price_per_night")]
    totals = [t for t in totals if t and t > 0]
    if not totals:
        return {"mean": 0, "median": 0}
    return {
        "mean":   statistics.mean(totals),
        "median": statistics.median(totals),
    }


def _haversine_km(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    R = 6371.0
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = (math.sin(dlat / 2) ** 2
         + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2))
         * math.sin(dlon / 2) ** 2)
    return R * 2 * math.asin(math.sqrt(a))


def compute_comp_similarity(
    listings: list[dict],
    target_beds: int | None,
    target_baths: float | None,
    center_lat: float,
    center_lon: float,
    market_adr: float,
) -> list[dict]:
    """각 숙소에 유사도 점수(0-100) 부여 후 내림차순 정렬.

    가중치: 거리 30% + 침실 25% + 욕실 15% + 가격 15% + 품질 15%
    """
    result = []
    for l in listings:
        lat = float(l.get("latitude")  or center_lat)
        lon = float(l.get("longitude") or center_lon)
        dist_km = _haversine_km(center_lat, center_lon, lat, lon)
        dist_score = max(0.0, (1.0 - dist_km / 3.0)) * 100.0

        beds_l = float(l.get("bedrooms")  or 0)
        beds_t = float(target_beds        or beds_l)
        room_score = max(0.0, 100.0 - abs(beds_l - beds_t) * 33.0)

        baths_l = float(l.get("bathrooms") or 0)
        baths_t = float(target_baths       or baths_l)
        bath_score = max(0.0, 100.0 - abs(baths_l - baths_t) * 50.0)

        price = float(l.get("price_per_night") or market_adr or 1)
        ratio = price / (market_adr or price)
        price_score = max(0.0, 100.0 - abs(1.0 - ratio) * 100.0)

        rating = float(l.get("rating") or 0)
        quality_score = max(0.0, min(100.0, (rating - 3.0) / 2.0 * 100.0))

        sim = (dist_score  * 0.30
               + room_score  * 0.25
               + bath_score  * 0.15
               + price_score * 0.15
               + quality_score * 0.15)

        item = l.copy()
        item["similarity_score"] = round(sim, 1)
        item["dist_km"]          = round(dist_km, 2)
        result.append(item)

    return sorted(result, key=lambda x: x["similarity_score"], reverse=True)


def compute_demand_signal(listings: list[dict]) -> dict:
    """리뷰 수 기반 수요 시그널 및 예약률 추정 (저신뢰도).

    가정: 평균 등록 기간 18개월, 리뷰 작성률 50% / 70% / 90%
    """
    reviews = [l.get("review_count") or 0 for l in listings]
    total = len(reviews)
    if total == 0:
        return {}

    avg_rc = statistics.mean(reviews)
    ASSUMED_MONTHS = 18
    monthly_avg_reviews = avg_rc / ASSUMED_MONTHS

    occ_est = {}
    for key, rate in [("low", 0.50), ("mid", 0.70), ("high", 0.90)]:
        monthly_bookings = monthly_avg_reviews / rate
        occ = min(monthly_bookings / 30.0, 0.95)
        occ_est[key] = round(occ * 100, 1)

    return {
        "avg_review_count":  round(avg_rc, 1),
        "pct_10plus":        sum(1 for r in reviews if r >= 10) / total * 100,
        "pct_20plus":        sum(1 for r in reviews if r >= 20) / total * 100,
        "pct_no_review":     sum(1 for r in reviews if r == 0)  / total * 100,
        "occ_estimates":     occ_est,
        "assumed_months":    ASSUMED_MONTHS,
        "confidence":        "저신뢰도 추정 (등록일 미확인, 작성률 가정)",
    }


def compute_market_score(
    listings: list[dict],
    stats: dict,
    premium: float,
    demand: dict,
) -> dict:
    """4개 지표 × 25점 = 100점 시장 매력도 점수.

    1. 수요 안정성    — 평균 리뷰 수
    2. 주말 수요 탄력성 — 주말 프리미엄
    3. 가격 성장 여력  — (P70 - P40) / P40
    4. 진입 품질 허들  — 경쟁자 평균 평점 역점수
    """
    # 1. 수요 안정성
    avg_rc = demand.get("avg_review_count", 0)
    if   avg_rc >= 40: d_score = 25
    elif avg_rc >= 20: d_score = 18
    elif avg_rc >= 10: d_score = 12
    else:              d_score = 6

    # 2. 주말 수요 탄력성
    prem_pct = premium * 100
    if   prem_pct >= 30: w_score = 25
    elif prem_pct >= 15: w_score = 18
    elif prem_pct >=  5: w_score = 10
    else:                w_score = 4

    # 3. 가격 성장 여력 (P70 - P40 스프레드)
    p40 = stats.get("p40", 0)
    p70 = stats.get("p70", 0)
    spread_pct = (p70 - p40) / p40 * 100 if p40 > 0 else 0
    if   spread_pct >= 70: g_score = 25
    elif spread_pct >= 40: g_score = 18
    elif spread_pct >= 20: g_score = 12
    else:                  g_score = 6

    # 4. 진입 품질 허들 (낮은 경쟁자 품질 = 차별화 쉬움 = 높은 점수)
    avg_rating = stats.get("avg_rating", 0)
    if   avg_rating < 4.50:  q_score = 25
    elif avg_rating < 4.70:  q_score = 18
    elif avg_rating < 4.85:  q_score = 12
    else:                    q_score = 6

    total = d_score + w_score + g_score + q_score

    if   total >= 80: judgment = "적극 진입"
    elif total >= 60: judgment = "조건부 진입"
    elif total >= 40: judgment = "보수적 접근"
    else:             judgment = "진입 미권장"

    return {
        "total":     total,
        "judgment":  judgment,
        "demand":    d_score,
        "weekend":   w_score,
        "growth":    g_score,
        "quality":   q_score,
        "warning":   len(listings) < 15,  # 샘플 적으면 신뢰도 낮음
    }


def build_revenue_scenarios(
    adr_clean:            float,
    premium:              float,
    occ_low:              float,
    occ_base:             float,
    occ_high:             float,
    cleaning_fee_per_stay: int,
    avg_stay_nights:      float,
    monthly_ops_cost:     int,
) -> list[dict]:
    """3개 예약률 시나리오 × 손익 구조 계산 (가정 기반 시뮬레이터).

    블렌디드 ADR = 평일 ADR × 5/7 + 주말 ADR × 2/7
    호스트 수수료: 3% (Airbnb 기본)
    청소비: booking_days / avg_stay_nights × cleaning_fee_per_stay
    임대료: 미포함 (물건마다 다름)
    """
    HOST_FEE  = 0.03
    MONTH_DAY = 30
    WD_W = 5 / 7
    WE_W = 2 / 7
    we_adr     = adr_clean * (1 + premium)
    blend_adr  = adr_clean * WD_W + we_adr * WE_W

    results = []
    for label, occ in [("보수적", occ_low), ("기준", occ_base), ("공격적", occ_high)]:
        booking_days   = occ * MONTH_DAY
        checkouts      = max(1.0, booking_days / avg_stay_nights)
        gross          = blend_adr * booking_days
        platform_fee   = gross * HOST_FEE
        net_revenue    = gross - platform_fee
        cleaning_cost  = checkouts * cleaning_fee_per_stay
        total_ops      = cleaning_cost + monthly_ops_cost
        op_profit      = net_revenue - total_ops
        results.append({
            "label":         label,
            "occ_pct":       round(occ * 100),
            "booking_days":  round(booking_days, 1),
            "blend_adr":     round(blend_adr),
            "gross":         round(gross),
            "platform_fee":  round(platform_fee),
            "net_revenue":   round(net_revenue),
            "cleaning_cost": round(cleaning_cost),
            "monthly_ops":   round(monthly_ops_cost),
            "total_ops":     round(total_ops),
            "op_profit":     round(op_profit),
        })
    return results


# ══════════════════════════════════════════════════════════════════
# 5. 포맷 팩토리 (공통)
# ══════════════════════════════════════════════════════════════════

def _make_formats(wb: xlsxwriter.Workbook) -> dict:
    """재사용 포맷 딕셔너리."""
    f = {}

    # 배너/헤더
    f["banner"] = wb.add_format({
        "bold": True, "font_size": 18,
        "font_color": C_WHITE, "bg_color": C_DARK_BLUE,
        "align": "center", "valign": "vcenter",
    })
    f["subtitle"] = wb.add_format({
        "bold": False, "font_size": 11,
        "font_color": C_WHITE, "bg_color": C_MID_BLUE,
        "align": "center", "valign": "vcenter",
    })
    f["section"] = wb.add_format({
        "bold": True, "font_size": 11,
        "font_color": C_WHITE, "bg_color": C_MID_BLUE,
        "valign": "vcenter",
    })
    f["header"] = wb.add_format({
        "bold": True, "font_size": 10,
        "font_color": C_WHITE, "bg_color": C_MID_BLUE,
        "align": "center", "valign": "vcenter",
        "border": 1, "border_color": C_DARK_BLUE, "text_wrap": True,
    })
    f["header_dark"] = wb.add_format({
        "bold": True, "font_size": 10,
        "font_color": C_WHITE, "bg_color": C_DARK_BLUE,
        "align": "center", "valign": "vcenter",
        "border": 1, "border_color": C_DARK_BLUE, "text_wrap": True,
    })

    # KPI 박스
    f["kpi_label"] = wb.add_format({
        "bold": True, "font_size": 9,
        "font_color": C_DARK_BLUE, "bg_color": C_ACCENT_BG,
        "align": "center", "valign": "vcenter",
        "top": 2, "left": 2, "right": 2, "top_color": C_MID_BLUE,
        "left_color": C_MID_BLUE, "right_color": C_MID_BLUE,
    })
    f["kpi_value"] = wb.add_format({
        "bold": True, "font_size": 22,
        "font_color": C_DARK_BLUE, "bg_color": C_WHITE,
        "align": "center", "valign": "vcenter",
        "bottom": 2, "left": 2, "right": 2, "bottom_color": C_MID_BLUE,
        "left_color": C_MID_BLUE, "right_color": C_MID_BLUE,
    })
    f["kpi_label_green"] = wb.add_format({
        "bold": True, "font_size": 9,
        "font_color": "#276221", "bg_color": C_LIGHT_GREEN,
        "align": "center", "valign": "vcenter",
        "top": 2, "left": 2, "right": 2,
        "top_color": C_GREEN, "left_color": C_GREEN, "right_color": C_GREEN,
    })
    f["kpi_value_green"] = wb.add_format({
        "bold": True, "font_size": 22,
        "font_color": "#276221", "bg_color": C_WHITE,
        "align": "center", "valign": "vcenter",
        "bottom": 2, "left": 2, "right": 2,
        "bottom_color": C_GREEN, "left_color": C_GREEN, "right_color": C_GREEN,
    })

    # 데이터 셀
    f["text"]      = wb.add_format({"valign": "vcenter", "border": 1, "border_color": "#BDD7EE"})
    f["text_alt"]  = wb.add_format({"valign": "vcenter", "bg_color": C_LIGHT_BLUE, "border": 1, "border_color": "#BDD7EE"})
    f["wrap"]      = wb.add_format({"valign": "top", "text_wrap": True, "border": 1, "border_color": "#BDD7EE"})
    f["wrap_alt"]  = wb.add_format({"valign": "top", "text_wrap": True, "bg_color": C_LIGHT_BLUE, "border": 1, "border_color": "#BDD7EE"})
    f["num"]       = wb.add_format({"num_format": "#,##0", "valign": "vcenter", "border": 1, "border_color": "#BDD7EE"})
    f["num_alt"]   = wb.add_format({"num_format": "#,##0", "valign": "vcenter", "bg_color": C_LIGHT_BLUE, "border": 1, "border_color": "#BDD7EE"})
    f["num_c"]     = wb.add_format({"num_format": "#,##0", "align": "center", "valign": "vcenter", "border": 1, "border_color": "#BDD7EE"})
    f["num_c_alt"] = wb.add_format({"num_format": "#,##0", "align": "center", "valign": "vcenter", "bg_color": C_LIGHT_BLUE, "border": 1, "border_color": "#BDD7EE"})
    f["dec"]       = wb.add_format({"num_format": "0.00", "valign": "vcenter", "border": 1, "border_color": "#BDD7EE"})
    f["dec_alt"]   = wb.add_format({"num_format": "0.00", "valign": "vcenter", "bg_color": C_LIGHT_BLUE, "border": 1, "border_color": "#BDD7EE"})
    f["dec_c"]     = wb.add_format({"num_format": "0.00", "align": "center", "valign": "vcenter", "border": 1, "border_color": "#BDD7EE"})
    f["dec_c_alt"] = wb.add_format({"num_format": "0.00", "align": "center", "valign": "vcenter", "bg_color": C_LIGHT_BLUE, "border": 1, "border_color": "#BDD7EE"})
    f["pct"]       = wb.add_format({"num_format": "0.0%", "align": "center", "valign": "vcenter", "border": 1, "border_color": "#BDD7EE"})
    f["pct_alt"]   = wb.add_format({"num_format": "0.0%", "align": "center", "valign": "vcenter", "bg_color": C_LIGHT_BLUE, "border": 1, "border_color": "#BDD7EE"})
    f["link"]      = wb.add_format({"font_color": C_LINK, "underline": True, "valign": "vcenter", "border": 1, "border_color": "#BDD7EE"})
    f["link_alt"]  = wb.add_format({"font_color": C_LINK, "underline": True, "valign": "vcenter", "bg_color": C_LIGHT_BLUE, "border": 1, "border_color": "#BDD7EE"})
    f["center"]    = wb.add_format({"align": "center", "valign": "vcenter", "border": 1, "border_color": "#BDD7EE"})
    f["center_alt"]= wb.add_format({"align": "center", "valign": "vcenter", "bg_color": C_LIGHT_BLUE, "border": 1, "border_color": "#BDD7EE"})

    # 강조
    f["gold_text"]  = wb.add_format({"bold": True, "valign": "vcenter", "bg_color": C_GOLD_BG, "border": 1, "border_color": C_YELLOW})
    f["gold_num"]   = wb.add_format({"bold": True, "num_format": "#,##0", "valign": "vcenter", "bg_color": C_GOLD_BG, "border": 1, "border_color": C_YELLOW})
    f["gold_dec"]   = wb.add_format({"bold": True, "num_format": "0.00", "align": "center", "valign": "vcenter", "bg_color": C_GOLD_BG, "border": 1, "border_color": C_YELLOW})
    f["gold_center"]= wb.add_format({"bold": True, "align": "center", "valign": "vcenter", "bg_color": C_GOLD_BG, "border": 1, "border_color": C_YELLOW})
    f["red_text"]   = wb.add_format({"font_color": C_RED, "valign": "vcenter", "bg_color": C_LIGHT_RED, "border": 1, "border_color": C_RED})
    f["red_num"]    = wb.add_format({"font_color": C_RED, "num_format": "#,##0", "valign": "vcenter", "bg_color": C_LIGHT_RED, "border": 1, "border_color": C_RED})

    # 서브 KPI 테이블
    f["sub_key"]   = wb.add_format({"bold": True, "font_size": 9, "align": "right", "valign": "vcenter", "bg_color": C_SECTION_BG})
    f["sub_val"]   = wb.add_format({"font_size": 10, "align": "left", "valign": "vcenter"})
    f["sub_key_hl"]= wb.add_format({"bold": True, "font_size": 9, "align": "right", "valign": "vcenter", "bg_color": C_ACCENT_BG})
    f["sub_val_hl"]= wb.add_format({"bold": True, "font_size": 10, "font_color": C_DARK_BLUE, "align": "left", "valign": "vcenter", "bg_color": C_ACCENT_BG})

    # 추천 객단가 전용
    f["rec_tier"] = wb.add_format({
        "bold": True, "font_size": 12, "align": "center", "valign": "vcenter",
        "font_color": C_WHITE, "bg_color": C_DARK_BLUE,
        "border": 1, "border_color": C_DARK_BLUE,
    })
    f["rec_label"] = wb.add_format({
        "bold": True, "font_size": 9, "align": "center", "valign": "vcenter",
        "bg_color": C_ACCENT_BG,
        "border": 1, "border_color": C_MID_BLUE,
    })
    f["rec_price"] = wb.add_format({
        "bold": True, "font_size": 18, "num_format": "#,##0",
        "align": "center", "valign": "vcenter",
        "font_color": C_DARK_BLUE, "bg_color": C_WHITE,
        "border": 2, "border_color": C_MID_BLUE,
    })
    f["rec_desc"]  = wb.add_format({
        "font_size": 9, "font_color": C_GRAY, "align": "center", "valign": "vcenter",
        "text_wrap": True, "bg_color": C_WHITE,
        "bottom": 2, "left": 2, "right": 2,
        "bottom_color": C_MID_BLUE, "left_color": C_MID_BLUE, "right_color": C_MID_BLUE,
    })

    return f


# ══════════════════════════════════════════════════════════════════
# 6. 손님용 리포트 빌더
# ══════════════════════════════════════════════════════════════════

def _write_kpi_box(ws, row, col, label, value, f_label, f_value, span=2):
    ws.merge_range(row, col, row,     col + span - 1, label, f_label)
    ws.merge_range(row + 1, col, row + 1, col + span - 1, value, f_value)


def _build_cover_sheet(wb, f, listings, query, beds, baths, windows, stats, premium, avg_calendar_occ=None):
    ws = wb.add_worksheet("📊 시장 개요")

    # 열 너비 (A=1.5, B-C=15, D=1.5, E-F=15, G=1.5, H-I=15, J=1.5)
    ws.set_column(0, 0, 1.5)   # 좌 여백
    ws.set_column(1, 2, 15)    # KPI 1
    ws.set_column(3, 3, 1.5)   # 구분
    ws.set_column(4, 5, 15)    # KPI 2
    ws.set_column(6, 6, 1.5)   # 구분
    ws.set_column(7, 8, 15)    # KPI 3
    ws.set_column(9, 9, 1.5)   # 우 여백

    # ── 배너 ──
    ws.merge_range(0, 0, 0, 9, "에어비앤비 시장 분석 리포트", f["banner"])
    ws.set_row(0, 50)

    cond_txt = ""
    if beds is not None:
        cond_txt += f" | 침실 {beds}개"
    if baths is not None:
        cond_txt += f" | 욕실 {baths}개"
    subtitle = f"지역: {query}  |  분석일: {date.today()}{cond_txt}  |  수집 기간: {windows[0][0]} ~ {windows[1][1]}"
    ws.merge_range(1, 0, 1, 9, subtitle, f["subtitle"])
    ws.set_row(1, 30)
    ws.set_row(2, 12)  # spacer

    indiv_cnt = sum(1 for l in listings if l.get("host_type") == "개인")
    comm_cnt  = sum(1 for l in listings if l.get("host_type") == "상업용")
    total_cnt = len(listings)
    avg_price = stats.get("mean_clean", stats.get("mean", 0))
    avg_rating= stats.get("avg_rating", 0)

    # ── KPI 박스 3개 ──
    ws.set_row(3, 22)
    ws.set_row(4, 55)

    _write_kpi_box(ws, 3, 1, "수집 숙소",    f"{total_cnt}개",    f["kpi_label"], f["kpi_value"])
    _write_kpi_box(ws, 3, 4, "평균 1박가",   f"₩{avg_price:,.0f}", f["kpi_label"], f["kpi_value"])
    _write_kpi_box(ws, 3, 7, "평균 평점",    f"★ {avg_rating:.2f}", f["kpi_label_green"], f["kpi_value_green"])

    ws.set_row(5, 12)  # spacer

    # ── 서브 KPI 테이블 ──
    ws.set_row(6, 22)
    ws.merge_range(6, 1, 6, 8, "세부 지표", f["section"])

    sub_kpis = [
        ("개인 숙소",       f"{indiv_cnt}개",                     False),
        ("상업용 숙소",     f"{comm_cnt}개",                      False),
        ("최저 1박가",      f"₩{stats.get('min', 0):,.0f}",      False),
        ("최고 1박가",      f"₩{stats.get('max', 0):,.0f}",      False),
        ("중앙값 1박가",    f"₩{stats.get('median', 0):,.0f}",   True),
        ("주말 프리미엄",   f"+{premium*100:.1f}%",               True),
        ("평점 집계 숙소",  f"{stats.get('rating_count', 0)}개", False),
        ("이상치 제외 수",  f"{stats.get('outlier_count', 0)}개",False),
        ("시장 예약률 (90일)", f"{avg_calendar_occ*100:.1f}%" if avg_calendar_occ is not None else "수집 중단", True),
    ]

    for i, (k, v, hl) in enumerate(sub_kpis):
        r = 7 + i
        ws.write(r, 1, k, f["sub_key_hl"] if hl else f["sub_key"])
        ws.write(r, 2, v, f["sub_val_hl"] if hl else f["sub_val"])
        # 오른쪽에 두 번째 컬럼 (3-4번째 서브KPI)
        if i < len(sub_kpis) // 2:
            pass
        ws.set_row(r, 18)

    # 2컬럼 레이아웃으로 재배치
    ws.set_row(7, 22)
    ws.merge_range(7, 4, 7, 8, "가격 분포 요약", f["section"])
    price_rows = [
        ("P25 (하위 25%)",  f"₩{stats.get('p25', 0):,.0f}"),
        ("P50 (중앙값)",    f"₩{stats.get('p50', 0):,.0f}"),
        ("P75 (상위 25%)",  f"₩{stats.get('p75', 0):,.0f}"),
        ("P90 (상위 10%)",  f"₩{stats.get('p90', 0):,.0f}"),
    ]
    for i, (k, v) in enumerate(price_rows):
        r = 8 + i
        ws.write(r, 4, k, f["sub_key"])
        ws.write(r, 5, v, f["sub_val"])
        ws.set_row(r, 18)


def _build_market_sheet(wb, f, listings, stats, premium):
    ws = wb.add_worksheet("📈 시장 분석")

    ws.set_column(0, 0, 16)   # 가격구간
    ws.set_column(1, 1, 12)   # 전체
    ws.set_column(2, 2, 12)   # 개인
    ws.set_column(3, 3, 12)   # 상업용
    ws.set_column(4, 4, 10)   # 비율
    ws.set_column(5, 5, 2)
    ws.set_column(6, 6, 18)
    ws.set_column(7, 7, 14)
    ws.set_column(8, 8, 14)

    row = 0
    ws.merge_range(row, 0, row, 8, "가격 구간 분포", f["section"])
    ws.set_row(row, 24); row += 1

    headers = ["가격 구간", "전체 숙소", "개인", "상업용", "비율(%)"]
    for c, h in enumerate(headers):
        ws.write(row, c, h, f["header"])
    ws.set_row(row, 20); row += 1

    dist = price_distribution(listings)
    for i, d in enumerate(dist):
        alt = i % 2 == 1
        tf  = f["text_alt"] if alt else f["text"]
        nf  = f["num_c_alt"] if alt else f["num_c"]
        pf  = f["pct_alt"] if alt else f["pct"]
        ws.write(row, 0, d["label"],           tf)
        ws.write(row, 1, d["count"],           nf)
        ws.write(row, 2, d["indiv"],           nf)
        ws.write(row, 3, d["comm"],            nf)
        ws.write(row, 4, d["pct"] / 100,      pf)
        ws.set_row(row, 18); row += 1

    # 데이터바 (전체 숙소 열)
    dist_start = 2
    dist_end   = row - 1
    ws.conditional_format(dist_start, 1, dist_end, 1, {
        "type": "data_bar",
        "bar_color": "#2E75B6",
        "bar_border_color": "#1F4E79",
        "data_bar_2010": True,
    })

    row += 1
    ws.merge_range(row, 0, row, 8, "평일 vs 주말 비교", f["section"])
    ws.set_row(row, 24); row += 1

    ws.write(row, 0, "구분",     f["header"])
    ws.write(row, 1, "평균 1박가", f["header"])
    ws.write(row, 2, "중앙값",    f["header"])
    ws.write(row, 3, "샘플 수",   f["header"])
    ws.set_row(row, 20); row += 1

    wd_prices = [l["price_weekday"] for l in listings if l.get("price_weekday")]
    we_prices = [l["price_weekend"] for l in listings if l.get("price_weekend")]

    def _row_vals(label, prices, alt):
        tf = f["text_alt"] if alt else f["text"]
        nf = f["num_alt"] if alt else f["num"]
        cf = f["center_alt"] if alt else f["center"]
        ws.write(row, 0, label, tf)
        ws.write(row, 1, int(statistics.mean(prices)) if prices else 0, nf)
        ws.write(row, 2, int(statistics.median(prices)) if prices else 0, nf)
        ws.write(row, 3, len(prices), cf)
        ws.set_row(row, 18)

    _row_vals("평일 (1박)", wd_prices, False); row += 1
    _row_vals("주말 (1박)", we_prices, True);  row += 1
    ws.write(row, 0, "주말 프리미엄", f["text"])
    ws.merge_range(row, 1, row, 3,
                   f"+{premium*100:.1f}% (주말이 평일 대비 {premium*100:.1f}% 더 비쌈)",
                   f["text"])
    ws.set_row(row, 18); row += 1

    # 침실 수별 평균가
    row += 1
    ws.merge_range(row, 0, row, 8, "침실 수별 평균 1박가", f["section"])
    ws.set_row(row, 24); row += 1

    ws.write(row, 0, "침실 수", f["header"])
    ws.write(row, 1, "숙소 수",  f["header"])
    ws.write(row, 2, "평균가",   f["header"])
    ws.write(row, 3, "중앙값",   f["header"])
    ws.set_row(row, 20); row += 1

    beds_map: dict[int, list] = {}
    for l in listings:
        b = int(l.get("bedrooms") or 0)
        beds_map.setdefault(b, []).append(l.get("price_per_night", 0))

    for i, (b, prices_b) in enumerate(sorted(beds_map.items())):
        alt = i % 2 == 1
        tf = f["text_alt"] if alt else f["text"]
        nf = f["num_alt"] if alt else f["num"]
        cf = f["center_alt"] if alt else f["center"]
        label = f"{b}침실" if b > 0 else "미분류"
        ws.write(row, 0, label, tf)
        ws.write(row, 1, len(prices_b), cf)
        ws.write(row, 2, int(statistics.mean(prices_b)), nf)
        ws.write(row, 3, int(statistics.median(prices_b)), nf)
        ws.set_row(row, 18); row += 1


def _build_competition_sheet(wb, f, listings):
    ws = wb.add_worksheet("🏆 경쟁 현황")

    ws.set_column(0, 0, 5)    # 순위
    ws.set_column(1, 1, 40)   # 숙소명
    ws.set_column(2, 2, 8)    # 구분
    ws.set_column(3, 3, 6)    # 침실
    ws.set_column(4, 4, 6)    # 욕실
    ws.set_column(5, 5, 8)    # 평점
    ws.set_column(6, 6, 13)   # 평일가
    ws.set_column(7, 7, 13)   # 주말가
    ws.set_column(8, 8, 10)   # 슈퍼호스트

    row = 0
    ws.merge_range(row, 0, row, 8, "경쟁 숙소 현황 (평점 기준 Top 30)", f["section"])
    ws.set_row(row, 24); row += 1

    headers = ["순위", "숙소명", "구분", "침실", "욕실", "평점", "평일가", "주말가", "슈퍼호스트"]
    for c, h in enumerate(headers):
        ws.write(row, c, h, f["header"])
    ws.set_row(row, 22); row += 1

    # 평점 기준 정렬, 최대 30개
    sorted_lst = sorted(
        [l for l in listings if l.get("rating")],
        key=lambda x: x.get("rating", 0),
        reverse=True,
    )[:30]

    for rank, l in enumerate(sorted_lst, 1):
        is_gold = rank <= 5
        alt     = rank % 2 == 0 and not is_gold

        tf = f["gold_text"]   if is_gold else (f["text_alt"]   if alt else f["text"])
        nf = f["gold_num"]    if is_gold else (f["num_alt"]    if alt else f["num"])
        df = f["gold_dec"]    if is_gold else (f["dec_c_alt"]  if alt else f["dec_c"])
        cf = f["gold_center"] if is_gold else (f["center_alt"] if alt else f["center"])

        wd_p = l.get("price_weekday") or l.get("price_per_night") or 0
        we_p = l.get("price_weekend") or 0

        ws.write(row, 0, rank,                                 cf)
        ws.write(row, 1, l.get("title", "")[:45],             tf)
        ws.write(row, 2, l.get("host_type", "개인"),           cf)
        ws.write(row, 3, int(l.get("bedrooms") or 0),         cf)
        ws.write(row, 4, l.get("bathrooms") or 0,             df)
        ws.write(row, 5, l.get("rating") or 0,                df)
        ws.write(row, 6, wd_p,                                 nf)
        ws.write(row, 7, we_p if we_p else "-",               nf if we_p else tf)
        ws.write(row, 8, l.get("superhost", "일반"),           cf)
        ws.set_row(row, 18); row += 1

    # 평점 조건부 서식
    data_start = 2
    data_end   = row - 1
    ws.conditional_format(data_start, 5, data_end, 5, {
        "type": "cell", "criteria": ">=", "value": 4.8,
        "format": wb.add_format({"bg_color": C_LIGHT_GREEN, "font_color": "#276221"}),
    })
    ws.conditional_format(data_start, 5, data_end, 5, {
        "type": "cell", "criteria": "between", "minimum": 4.5, "maximum": 4.799,
        "format": wb.add_format({"bg_color": C_YELLOW_BG, "font_color": "#7F5000"}),
    })

    ws.freeze_panes(2, 2)


def _build_recommendation_sheet(wb, f, stats, premium):
    ws = wb.add_worksheet("💡 추천 객단가")

    ws.set_column(0, 0, 2)
    ws.set_column(1, 4, 18)
    ws.set_column(5, 5, 2)
    ws.set_column(6, 9, 14)

    row = 0
    ws.merge_range(row, 0, row, 9, "추천 객단가 — 포지셔닝 전략", f["banner"])
    ws.set_row(row, 40); row += 1

    subtitle = f"기준: 이상치 제거 후 {stats.get('clean_count', 0)}개 숙소 데이터  |  분석일: {date.today()}"
    ws.merge_range(row, 0, row, 9, subtitle, f["subtitle"])
    ws.set_row(row, 24); row += 2

    tiers = [
        ("🌱 스타터",    stats.get("p40", 0), "초기 진입 전략\n리뷰 확보 우선"),
        ("⭐ 표준",      stats.get("p55", 0), "리뷰 10개+ 기준\n시장 적정 포지션"),
        ("👑 프리미엄",  stats.get("p70", 0), "슈퍼호스트 목표\n차별화 포지션"),
    ]

    # 평일 추천가 헤더
    ws.merge_range(row, 1, row, 4, "평일 추천가", f["section"])
    ws.set_row(row, 22); row += 1

    for col_start, (tier, base_price, desc) in enumerate(tiers):
        c = 1 + col_start
        ws.write(row, c, tier, f["rec_tier"])
    ws.set_row(row, 28); row += 1

    for col_start, (_, p, _) in enumerate(tiers):
        c = 1 + col_start
        ws.write(row, c, "평일 1박", f["rec_label"])
    ws.set_row(row, 20); row += 1

    for col_start, (_, p, _) in enumerate(tiers):
        c = 1 + col_start
        ws.write(row, c, round(p / 1000) * 1000, f["rec_price"])  # 천원 단위 반올림
    ws.set_row(row, 55); row += 1

    for col_start, (_, _, desc) in enumerate(tiers):
        c = 1 + col_start
        ws.write(row, c, desc, f["rec_desc"])
    ws.set_row(row, 40); row += 2

    # 주말 추천가
    ws.merge_range(row, 1, row, 4, f"주말 추천가 (평일 대비 +{premium*100:.1f}%)", f["section"])
    ws.set_row(row, 22); row += 1

    for col_start, (tier, _, _) in enumerate(tiers):
        c = 1 + col_start
        ws.write(row, c, tier, f["rec_tier"])
    ws.set_row(row, 28); row += 1

    for col_start, (_, p, _) in enumerate(tiers):
        c = 1 + col_start
        ws.write(row, c, "주말 1박", f["rec_label"])
    ws.set_row(row, 20); row += 1

    for col_start, (_, p, _) in enumerate(tiers):
        c = 1 + col_start
        we_price = p * (1 + premium)
        ws.write(row, c, round(we_price / 1000) * 1000, f["rec_price"])
    ws.set_row(row, 55); row += 1

    for col_start, (_, _, desc) in enumerate(tiers):
        c = 1 + col_start
        ws.write(row, c, desc, f["rec_desc"])
    ws.set_row(row, 40); row += 2

    # 주의사항
    notes = [
        "※ 본 추천가는 수집 시점 기준이며 계절·이벤트·Airbnb 정책 변동에 따라 달라질 수 있습니다.",
        "※ 이상치 제거(IQR 1.5× 방법) 후 클린 데이터 기준으로 산출되었습니다.",
        "※ 초기 리뷰 5개 이상 확보 후 단계적으로 가격을 올리는 전략을 권장합니다.",
    ]
    for note in notes:
        ws.merge_range(row, 1, row, 8, note,
                       wb.add_format({"font_size": 9, "font_color": C_GRAY, "text_wrap": True}))
        ws.set_row(row, 16); row += 1


def _build_revenue_sheet(wb, f, scenarios, adr_clean, premium, market_score,
                          cleaning_fee, avg_stay, monthly_ops):
    """💰 수익 시뮬레이터 시트 (손님용 Sheet 5 / 내부용 참조)."""
    ws = wb.add_worksheet("💰 수익 시뮬레이터")
    ws.set_tab_color(C_GREEN)
    ws.set_column("A:A", 22)
    ws.set_column("B:D", 18)

    r = 0
    ws.merge_range(r, 0, r, 3, "💰 월 수익 시뮬레이터", f["banner"])
    ws.set_row(r, 30)

    r += 1
    note = ("⚠ 이 수치는 업계 평균 예약률 기반 가정치입니다. 실제 매출과 다를 수 있으며, "
            "임대료·대출 상환은 포함되지 않습니다.")
    warn_fmt = wb.add_format({
        "bold": True, "font_size": 9, "font_color": C_RED,
        "bg_color": C_LIGHT_RED, "border": 1, "border_color": C_RED,
        "text_wrap": True, "valign": "vcenter",
    })
    ws.merge_range(r, 0, r, 3, note, warn_fmt)
    ws.set_row(r, 36)

    r += 2
    ws.merge_range(r, 0, r, 3, "■ 가정값 요약", f["section"])
    r += 1
    lbl = wb.add_format({"bold": True, "font_size": 10, "bg_color": C_ACCENT_BG,
                          "border": 1, "border_color": C_MID_BLUE})
    val = wb.add_format({"font_size": 10, "num_format": "#,##0",
                          "border": 1, "border_color": C_MID_BLUE})
    val_pct = wb.add_format({"font_size": 10, "num_format": '0.0"%"',
                               "border": 1, "border_color": C_MID_BLUE})
    assumptions = [
        ("평일 기준 ADR (이상치 제거)",  adr_clean,   "원"),
        ("주말 프리미엄",                premium*100, "%"),
        ("청소비 / 회",                  cleaning_fee, "원"),
        ("평균 체류일 (가정)",            avg_stay,    "박"),
        ("월 고정 운영비 (청소비 제외)", monthly_ops,  "원"),
        ("호스트 수수료 (Airbnb)",        3,           "%"),
    ]
    for label, v, unit in assumptions:
        ws.write(r, 0, label, lbl)
        ws.write(r, 1, v, val_pct if unit == "%" else val)
        ws.write(r, 2, unit,
                 wb.add_format({"font_size": 10, "border": 1,
                                 "border_color": C_MID_BLUE}))
        ws.write(r, 3, "", wb.add_format({"border": 1, "border_color": C_MID_BLUE}))
        r += 1

    r += 1
    ws.merge_range(r, 0, r, 3, "■ 시나리오별 월 손익", f["section"])
    r += 1
    col_hdrs = ["항목", "보수적", "기준", "공격적"]
    for c, h in enumerate(col_hdrs):
        ws.write(r, c, h, f["header_dark"])
    r += 1

    rows_def = [
        ("예약률 (가정)", "occ_pct",      "pct"),
        ("예약일수 / 월", "booking_days", "days"),
        ("블렌디드 ADR",  "blend_adr",    "krw"),
        ("━ 총 매출",     "gross",        "krw"),
        ("  플랫폼 수수료 (-3%)", "platform_fee", "krw"),
        ("  청소비",      "cleaning_cost","krw"),
        ("  월 운영비",   "monthly_ops",  "krw"),
        ("━ 영업이익 *",  "op_profit",    "krw_bold"),
    ]
    alt_bg = wb.add_format({"bg_color": C_LIGHT_GRAY, "border": 1,
                              "border_color": C_MID_BLUE, "font_size": 10})
    def _cell(is_bold=False, bg=C_WHITE):
        return wb.add_format({"bold": is_bold, "num_format": "#,##0",
                               "bg_color": bg, "border": 1,
                               "border_color": C_MID_BLUE, "font_size": 10,
                               "align": "right"})
    for i, (label, key, fmt) in enumerate(rows_def):
        bg = C_LIGHT_GREEN if key == "op_profit" else (C_LIGHT_GRAY if i % 2 else C_WHITE)
        is_profit = key == "op_profit"
        row_lbl_fmt = wb.add_format({
            "bold": is_profit, "font_size": 10,
            "bg_color": bg, "border": 1, "border_color": C_MID_BLUE,
        })
        ws.write(r, 0, label, row_lbl_fmt)
        for c, sc in enumerate(scenarios, 1):
            v = sc[key]
            if fmt == "pct":
                cell_fmt = wb.add_format({
                    "bold": is_profit, "num_format": '0"%"',
                    "bg_color": bg, "border": 1,
                    "border_color": C_MID_BLUE, "align": "right", "font_size": 10,
                })
            elif fmt == "days":
                cell_fmt = wb.add_format({
                    "bold": is_profit, "num_format": "0.0",
                    "bg_color": bg, "border": 1,
                    "border_color": C_MID_BLUE, "align": "right", "font_size": 10,
                })
            else:
                cell_fmt = wb.add_format({
                    "bold": is_profit,
                    "num_format": "#,##0",
                    "bg_color": C_LIGHT_GREEN if is_profit else bg,
                    "border": 1, "border_color": C_MID_BLUE,
                    "align": "right", "font_size": 10,
                })
            ws.write(r, c, v, cell_fmt)
        r += 1

    r += 1
    footnote_fmt = wb.add_format({
        "font_size": 9, "italic": True, "font_color": C_GRAY,
        "text_wrap": True,
    })
    ws.merge_range(r, 0, r+2, 3,
        "* 영업이익 = 순 매출 - 청소비 - 월 운영비 (임대료·대출 상환·세금 별도)\n"
        "* 블렌디드 ADR = 평일 ADR × 5/7 + 주말 ADR × 2/7\n"
        "* 업계 기준 예약률: 보수 40%, 기준 60%, 공격 70% — 지역·계절에 따라 다름",
        footnote_fmt)
    ws.set_row(r, 15); ws.set_row(r+1, 15); ws.set_row(r+2, 15)

    # 시장 매력도 점수
    r += 4
    ws.merge_range(r, 0, r, 3, "■ 시장 매력도 점수 (4요소 × 25점)", f["section"])
    r += 1
    score_hdrs = ["평가 항목", "점수", "만점", "근거"]
    for c, h in enumerate(score_hdrs):
        ws.write(r, c, h, f["header"])
    r += 1
    score_rows = [
        ("수요 안정성",    market_score["demand"],  25, "시장 평균 리뷰 수"),
        ("주말 수요 탄력성", market_score["weekend"], 25, "주말 프리미엄(%)"),
        ("가격 성장 여력", market_score["growth"],  25, "P70-P40 스프레드"),
        ("진입 품질 허들", market_score["quality"], 25, "경쟁자 평균 평점(역점수)"),
        ("합계",           market_score["total"],  100, market_score["judgment"]),
    ]
    for i, (lbl, sc, mx, note) in enumerate(score_rows):
        is_total = lbl == "합계"
        bg = C_LIGHT_GREEN if is_total else (C_LIGHT_GRAY if i % 2 else C_WHITE)
        rf = wb.add_format({"bold": is_total, "font_size": 10,
                             "bg_color": bg, "border": 1, "border_color": C_MID_BLUE})
        nf = wb.add_format({"bold": is_total, "num_format": "0",
                             "bg_color": bg, "border": 1,
                             "border_color": C_MID_BLUE, "align": "center", "font_size": 10})
        ws.write(r, 0, lbl, rf)
        ws.write(r, 1, sc,  nf)
        ws.write(r, 2, mx,  nf)
        ws.write(r, 3, note, rf)
        r += 1
    if market_score.get("warning"):
        r += 1
        ws.merge_range(r, 0, r, 3,
            "⚠ 샘플 수가 적어(< 15개) 시장 매력도 점수의 신뢰도가 낮습니다.",
            warn_fmt)


def _build_demand_sheet(wb, f, demand, listings):
    """📈 수요 추정 시트 (내부용)."""
    ws = wb.add_worksheet("📈 수요 추정")
    ws.set_tab_color(C_ACCENT_BLUE)
    ws.set_column("A:A", 26)
    ws.set_column("B:C", 20)
    ws.set_column("D:D", 36)

    r = 0
    ws.merge_range(r, 0, r, 3, "📈 리뷰 기반 수요 추정 (내부용)", f["banner"])
    ws.set_row(r, 30)

    r += 1
    warn_txt = ("⚠ 이 섹션의 예약률 수치는 저신뢰도 추정치입니다.\n"
                "가정: 평균 등록 기간 18개월, 리뷰 작성률 50%/70%/90% 범위\n"
                "실제 예약률과 ±20%p 이상 오차가 발생할 수 있습니다.")
    wfmt = wb.add_format({
        "bold": True, "font_size": 9, "font_color": C_RED,
        "bg_color": C_LIGHT_RED, "border": 1, "border_color": C_RED,
        "text_wrap": True, "valign": "vcenter",
    })
    ws.merge_range(r, 0, r+1, 3, warn_txt, wfmt)
    ws.set_row(r, 18); ws.set_row(r+1, 18)

    r += 3
    ws.merge_range(r, 0, r, 3, "■ 시장 리뷰 수 분포 (수요 시그널)", f["section"])
    r += 1
    lbl = wb.add_format({"bold": True, "font_size": 10, "bg_color": C_ACCENT_BG,
                          "border": 1, "border_color": C_MID_BLUE})
    val = wb.add_format({"font_size": 10, "num_format": "#,##0.0",
                          "border": 1, "border_color": C_MID_BLUE, "align": "right"})
    val_pct = wb.add_format({"font_size": 10, "num_format": '0.0"%"',
                               "border": 1, "border_color": C_MID_BLUE, "align": "right"})
    val_int = wb.add_format({"font_size": 10, "num_format": "#,##0",
                               "border": 1, "border_color": C_MID_BLUE, "align": "right"})

    signal_rows = [
        ("분석 대상 숙소 수",    len(listings),                   None),
        ("평균 리뷰 수",         demand.get("avg_review_count", 0), "float"),
        ("리뷰 10개+ 비중",       demand.get("pct_10plus", 0),       "pct"),
        ("리뷰 20개+ 비중",       demand.get("pct_20plus", 0),       "pct"),
        ("리뷰 0개 (신규 숙소)", demand.get("pct_no_review", 0),   "pct"),
    ]
    for label, v, ftype in signal_rows:
        ws.write(r, 0, label, lbl)
        if ftype == "pct":
            ws.write(r, 1, v, val_pct)
        elif ftype == "float":
            ws.write(r, 1, v, val)
        else:
            ws.write(r, 1, v, val_int)
        ws.write(r, 2, "", wb.add_format({"border": 1, "border_color": C_MID_BLUE}))
        ws.write(r, 3, "", wb.add_format({"border": 1, "border_color": C_MID_BLUE}))
        r += 1

    r += 1
    ws.merge_range(r, 0, r, 3, "■ 예약률 추정 범위 (저신뢰도)", f["section"])
    r += 1
    hdrs = ["시나리오", "리뷰 작성률 가정", "추정 예약률", "해석"]
    for c, h in enumerate(hdrs):
        ws.write(r, c, h, f["header"])
    r += 1

    occ = demand.get("occ_estimates", {})
    occ_rows = [
        ("낮은 추정 (보수)", "50%",  occ.get("low", 0),  "리뷰 작성률이 높을 때"),
        ("중간 추정 (기준)", "70%",  occ.get("mid", 0),  "업계 평균 작성률"),
        ("높은 추정 (공격)", "90%",  occ.get("high", 0), "리뷰 작성률이 낮을 때"),
    ]
    for i, (lbl_txt, rate, est, interp) in enumerate(occ_rows):
        bg = C_LIGHT_GRAY if i % 2 else C_WHITE
        rf = wb.add_format({"font_size": 10, "bg_color": bg, "border": 1,
                             "border_color": C_MID_BLUE})
        nf = wb.add_format({"font_size": 10, "num_format": '0.0"%"',
                             "bg_color": bg, "border": 1,
                             "border_color": C_MID_BLUE, "align": "right"})
        ws.write(r, 0, lbl_txt, rf)
        ws.write(r, 1, rate,    rf)
        ws.write(r, 2, est,     nf)
        ws.write(r, 3, interp,  rf)
        r += 1

    r += 2
    foot = wb.add_format({"font_size": 9, "italic": True, "font_color": C_GRAY,
                            "text_wrap": True})
    ws.merge_range(r, 0, r+1, 3,
        f"계산식: 월 평균리뷰({demand.get('avg_review_count',0):.1f}개) "
        f"÷ 가정 등록기간({demand.get('assumed_months',18)}개월) "
        f"÷ 리뷰작성률 ÷ 30일 = 예약률\n"
        "※ 실제 체크아웃 데이터 없이는 정확한 예약률 산출 불가. "
        "의사결정 보조 지표로만 활용하세요.",
        foot)

    r += 3
    ws.merge_range(r, 0, r, 3, "■ 숙소별 리뷰 수 현황", f["section"])
    r += 1
    hdr_items = ["숙소명", "리뷰 수", "구분", "평점"]
    for c, h in enumerate(hdr_items):
        ws.write(r, c, h, f["header"])
    r += 1
    for l in sorted(listings, key=lambda x: x.get("review_count") or 0, reverse=True):
        alt_bg = C_LIGHT_GRAY if r % 2 == 0 else C_WHITE
        rf = wb.add_format({"font_size": 10, "bg_color": alt_bg, "border": 1,
                             "border_color": C_MID_BLUE})
        nf = wb.add_format({"font_size": 10, "num_format": "#,##0",
                             "bg_color": alt_bg, "border": 1,
                             "border_color": C_MID_BLUE, "align": "right"})
        ws.write(r, 0, (l.get("title") or "")[:40], rf)
        ws.write(r, 1, l.get("review_count") or 0, nf)
        ws.write(r, 2, l.get("host_type", ""), rf)
        ws.write(r, 3, l.get("rating") or 0,
                 wb.add_format({"font_size": 10, "num_format": "0.00",
                                 "bg_color": alt_bg, "border": 1,
                                 "border_color": C_MID_BLUE, "align": "right"}))
        r += 1


def _build_guide_sheet(wb, f, is_internal: bool = False) -> None:
    """📖 용어·로직 안내 시트 — 모든 항목의 계산 방법과 용어 설명."""
    ws = wb.add_worksheet("📖 용어·로직 안내")
    ws.set_column("A:A", 2)
    ws.set_column("B:B", 26)
    ws.set_column("C:E", 28)

    note_fmt = wb.add_format({
        "font_size": 9, "italic": True, "font_color": C_GRAY,
        "text_wrap": True, "valign": "top",
    })

    def _section(row, title):
        ws.merge_range(row, 1, row, 4, title, f["section"])
        ws.set_row(row, 22)
        return row + 1

    def _item(row, label, desc, highlight=False):
        lf = f["sub_key_hl"] if highlight else f["sub_key"]
        vf = f["sub_val_hl"] if highlight else wb.add_format(
            {"font_size": 10, "align": "left", "valign": "vcenter", "text_wrap": True})
        ws.write(row, 1, label, lf)
        ws.merge_range(row, 2, row, 4, desc, vf)
        ws.set_row(row, 32)
        return row + 1

    def _note(row, text):
        ws.merge_range(row, 1, row, 4, text, note_fmt)
        ws.set_row(row, 18)
        return row + 1

    # 배너
    ws.merge_range(0, 0, 0, 4, "📖 용어 및 로직 설명", f["banner"])
    ws.set_row(0, 40)
    ws.merge_range(1, 0, 1, 4,
                   "이 시트는 리포트 각 항목의 계산 방법과 용어를 설명합니다. 의사결정 시 참고하세요.",
                   f["subtitle"])
    ws.set_row(1, 24)

    r = 3

    # ── 섹션 1: 데이터 수집 ──────────────────────────────────────────
    r = _section(r, "📡  데이터 수집 방식")
    r = _item(r, "날짜창 자동 설정",
              "입력한 체크인 날짜가 속한 주의 평일(월요일→화요일)과 주말(금요일→토요일) "
              "1박 창 2개를 자동으로 수집합니다.")
    r = _item(r, "중복 URL 머지",
              "두 날짜창에서 동일 숙소(URL 기준)를 하나로 합쳐 평일가·주말가를 각각 보존합니다.")
    r = _item(r, "상세 크롤링",
              "각 숙소 상세 페이지를 방문해 소개글·편의시설·호스트 정보를 추가로 수집합니다.")
    r += 1

    # ── 섹션 2: 가격 용어 ──────────────────────────────────────────
    r = _section(r, "💰  가격 용어")
    r = _item(r, "1박가 (ADR)",
              "Airbnb 검색 결과에 표시되는 1박 기준 숙박료. 청소비·서비스료 미포함.")
    r = _item(r, "총가 (Total Price)",
              "1박가 × 숙박 박수. Airbnb 청소비·서비스료는 별도입니다.")
    r = _item(r, "블렌디드 ADR",
              "평일 ADR × 5/7 + 주말 ADR × 2/7 — 주 7일 기준 가중 평균 일당 요금.", True)
    r = _item(r, "주말 프리미엄",
              "(공통 숙소의 주말 평균가 − 평일 평균가) ÷ 평일 평균가. "
              "주말 수요 탄력성 지표입니다.")
    r += 1

    # ── 섹션 3: 통계 지표 ──────────────────────────────────────────
    r = _section(r, "📊  통계 지표")
    r = _item(r, "이상치 제거 (IQR)",
              "Q3 + 1.5×IQR 초과 또는 Q1 − 1.5×IQR 미만인 가격을 이상치로 분류합니다. "
              "클린 평균 산출 시 이상치를 제외합니다.")
    r = _item(r, "클린 평균",
              "이상치를 제외한 후 계산한 평균가. 시장 대표 가격으로 신뢰도가 높습니다.", True)
    r = _item(r, "중앙값 (P50)",
              "전체 숙소를 가격 순으로 정렬했을 때 정중앙에 위치하는 값. 극단값에 강건합니다.")
    r = _item(r, "P10 ~ P90 (퍼센타일)",
              "전체 숙소를 가격 순으로 정렬했을 때 하위 N%에 해당하는 가격. "
              "예: P25 = 하위 25%에 해당하는 가격.")
    r = _item(r, "표준편차 (Stdev)",
              "가격 분포의 퍼짐 정도. 클수록 시장 가격이 다양하게 형성되어 있습니다.")
    r += 1

    # ── 섹션 4: 추천 객단가 ────────────────────────────────────────
    r = _section(r, "💡  추천 객단가 포지셔닝")
    r = _item(r, "스타터 = P40",
              "초기 진입 단계. 리뷰 0개 상태에서 예약을 빠르게 늘리기 위한 진입가. "
              "시장 하위 40% 수준.")
    r = _item(r, "표준 = P55",
              "리뷰 10개 이상 확보 후 적정 포지션. 시장 중간보다 약간 위 수준.", True)
    r = _item(r, "프리미엄 = P70",
              "슈퍼호스트·차별화 요소 확보 후 목표 가격. 시장 상위 30% 수준.")
    r = _note(r, "※ 초기 리뷰 5개 이상 확보 후 단계적으로 가격을 올리는 전략을 권장합니다.")
    r += 1

    # ── 섹션 5: 수익 시뮬레이터 ───────────────────────────────────
    r = _section(r, "💰  수익 시뮬레이터 가정")
    r = _item(r, "호스트 수수료 3%",
              "Airbnb 기본 호스트 수수료. 총 매출의 3%를 공제합니다.")
    r = _item(r, "청소비 계산",
              "청소 횟수 = 예약일수 ÷ 평균 체류일수. 총 청소비 = 횟수 × 1회 청소비.")
    r = _item(r, "영업이익 계산",
              "영업이익 = 순매출(총매출 − 수수료) − 청소비 − 월 고정비. "
              "임대료·대출 상환·세금은 포함되지 않습니다.", True)
    r = _item(r, "블렌디드 ADR 계산",
              "평일 ADR × 5/7 + 주말 ADR × 2/7 (주 7일 중 평일 5일, 주말 2일 가정).")
    r = _note(r, "※ 임대료, 대출 상환, 감가상각, 세금은 포함되지 않습니다. 의사결정 보조 지표로만 활용하세요.")
    r = _item(r, "예약률 시나리오 (가정)",
              "하/기/상 세 수치는 '이 정도면 얼마나 벌까?'를 시뮬레이션하는 가정값입니다. "
              "M 모드 캘린더 수집으로 실제 시장 예약률도 함께 표시됩니다.")
    r += 1

    if is_internal:
        # ── 섹션 6: 시장 매력도 점수 ───────────────────────────────
        r = _section(r, "🏆  시장 매력도 점수 (4요소 × 25점 = 100점)")
        r = _item(r, "수요 안정성 (25점)",
                  "시장 평균 리뷰 수 기반. ≥40개=25점 / ≥20개=18점 / ≥10개=12점 / 미만=6점.")
        r = _item(r, "주말 수요 탄력성 (25점)",
                  "주말 프리미엄 비율 기반. ≥30%=25점 / ≥15%=18점 / ≥5%=10점 / 미만=4점.", True)
        r = _item(r, "가격 성장 여력 (25점)",
                  "(P70 − P40) ÷ P40 스프레드 기반. ≥70%=25점 / ≥40%=18점 / ≥20%=12점 / 미만=6점.")
        r = _item(r, "진입 품질 허들 (25점)",
                  "경쟁자 평균 평점의 역점수. 낮은 경쟁 품질=차별화 용이. "
                  "<4.50=25점 / <4.70=18점 / <4.85=12점 / 이상=6점.")
        r = _note(r, "※ 수집 숙소 15개 미만 시 신뢰도 경고가 표시됩니다.")
        r += 1

        # ── 섹션 7: 개인·상업용 분류 ──────────────────────────────
        r = _section(r, "🏠  개인·상업용 분류 기준")
        r = _item(r, "분류 로직",
                  "건물유형·호스트명·소개글 키워드에 스코어를 부여합니다. "
                  "합산 점수 2점 이상이면 '상업용'으로 분류합니다.", True)
        r = _item(r, "건물 유형 (+3점)",
                  "호텔, 게스트하우스, 호스텔, B&B, 부티크호텔 등 해당 시 3점 추가.")
        r = _item(r, "호스트명 (+2점)",
                  "이름에 '호텔', '스테이', '레지던스', '리조트', '펜션', '모텔' 등 포함 시 2점 추가.")
        r = _item(r, "소개글 (+2점)",
                  "'프론트 데스크', '리셉션', '체크인 카운터' 등 2개 이상 포함 시 2점 추가.")
        r += 1

        # ── 섹션 8: 수요 추정 ──────────────────────────────────────
        r = _section(r, "📈  예약률 추정 방법 (저신뢰도)")
        r = _item(r, "계산식",
                  "추정 예약률 = (월 평균리뷰 ÷ 가정 등록기간) ÷ 리뷰작성률 ÷ 30일", True)
        r = _item(r, "등록기간 가정 (18개월)",
                  "숙소의 실제 등록일을 알 수 없어 평균 18개월로 가정합니다.")
        r = _item(r, "리뷰 작성률 시나리오",
                  "예약 대비 리뷰 작성 비율을 50%(낙관) / 70%(기준) / 90%(보수) 세 시나리오로 계산합니다.")
        r = _note(r,
                  "※ 실제 체크아웃 데이터 없이는 정확한 예약률 산출이 불가합니다. "
                  "실제와 ±20%p 이상 오차가 발생할 수 있습니다. 참고 지표로만 활용하세요.")


def build_client_report(
    listings: list[dict],
    query: str,
    beds: int | None,
    baths: float | None,
    windows: list,
    stats: dict,
    premium: float,
    out_path: Path,
    scenarios: list[dict] | None = None,
    market_score: dict | None = None,
    cleaning_fee: int = 80000,
    avg_stay: float = 2.0,
    monthly_ops: int = 0,
    avg_calendar_occ: float | None = None,
) -> None:
    wb = xlsxwriter.Workbook(str(out_path))
    f  = _make_formats(wb)

    _build_cover_sheet(wb, f, listings, query, beds, baths, windows, stats, premium, avg_calendar_occ=avg_calendar_occ)
    _build_market_sheet(wb, f, listings, stats, premium)
    _build_competition_sheet(wb, f, listings)
    _build_recommendation_sheet(wb, f, stats, premium)
    if scenarios and market_score:
        adr_clean = stats.get("mean_clean") or stats.get("mean") or 0
        _build_revenue_sheet(wb, f, scenarios, adr_clean, premium, market_score,
                             cleaning_fee, avg_stay, monthly_ops)
    _build_guide_sheet(wb, f, is_internal=False)

    wb.close()
    _fix_colors(out_path)
    print(f"✅ 손님용 저장: {out_path.name}")


# ══════════════════════════════════════════════════════════════════
# 7. 내부용 리포트 빌더
# ══════════════════════════════════════════════════════════════════

def _build_dashboard_sheet(wb, f, listings, stats, premium, windows, avg_calendar_occ=None):
    ws = wb.add_worksheet("📋 대시보드")

    ws.set_column(0, 0, 22)
    ws.set_column(1, 1, 20)
    ws.set_column(2, 2, 22)
    ws.set_column(3, 3, 20)

    row = 0
    ws.merge_range(row, 0, row, 3, "내부 분석 대시보드", f["banner"])
    ws.set_row(row, 40); row += 1

    ws.merge_range(row, 0, row, 3, f"수집 창: {windows[0][0]}(평일) / {windows[1][0]}(주말)  |  생성: {date.today()}", f["subtitle"])
    ws.set_row(row, 24); row += 2

    # ── 전체 KPI ──
    ws.merge_range(row, 0, row, 3, "전체 시장 지표", f["section"])
    ws.set_row(row, 22); row += 1

    indiv_lst = [l for l in listings if l.get("host_type") == "개인"]
    comm_lst  = [l for l in listings if l.get("host_type") == "상업용"]

    kpis = [
        ("전체 숙소 수",      f"{stats.get('count', 0)}개",                              True),
        ("개인 숙소",         f"{len(indiv_lst)}개",                                    False),
        ("상업용 숙소",       f"{len(comm_lst)}개",                                     False),
        ("이상치 숙소",       f"{stats.get('outlier_count', 0)}개 (IQR 1.5× 기준)",     False),
        ("평균 1박가 (전체)", f"₩{stats.get('mean', 0):,.0f}",                          True),
        ("평균 1박가 (클린)", f"₩{stats.get('mean_clean', 0):,.0f}",                    True),
        ("중앙값 1박가",      f"₩{stats.get('median', 0):,.0f}",                        False),
        ("최저가",            f"₩{stats.get('min', 0):,.0f}",                           False),
        ("최고가",            f"₩{stats.get('max', 0):,.0f}",                           False),
        ("표준편차",          f"₩{stats.get('stdev', 0):,.0f}",                         False),
        ("주말 프리미엄",     f"+{premium*100:.1f}%",                                   True),
        ("평균 평점",         f"{stats.get('avg_rating', 0):.2f}",                      False),
        ("시장 예약률 (90일)", f"{avg_calendar_occ*100:.1f}%" if avg_calendar_occ is not None else "-", True),
    ]

    for i, (k, v, hl) in enumerate(kpis):
        c_off = (i % 2) * 2
        r_off = i // 2
        ws.write(row + r_off, c_off,     k, f["sub_key_hl"] if hl else f["sub_key"])
        ws.write(row + r_off, c_off + 1, v, f["sub_val_hl"] if hl else f["sub_val"])

    row += (len(kpis) + 1) // 2 + 2

    # ── 개인 vs 상업용 비교 ──
    ws.merge_range(row, 0, row, 3, "개인 vs 상업용 비교", f["section"])
    ws.set_row(row, 22); row += 1

    for c, h in enumerate(["구분", "숙소 수", "평균가", "중앙값"]):
        ws.write(row, c, h, f["header"])
    ws.set_row(row, 20); row += 1

    for i, (label, subset) in enumerate([("개인", indiv_lst), ("상업용", comm_lst)]):
        prices_sub = [l.get("price_per_night", 0) for l in subset if l.get("price_per_night")]
        alt = i % 2 == 1
        tf = f["text_alt"] if alt else f["text"]
        nf = f["num_alt"] if alt else f["num"]
        ws.write(row, 0, label, tf)
        ws.write(row, 1, len(subset), nf)
        ws.write(row, 2, int(statistics.mean(prices_sub)) if prices_sub else 0, nf)
        ws.write(row, 3, int(statistics.median(prices_sub)) if prices_sub else 0, nf)
        ws.set_row(row, 18); row += 1

    row += 1

    # ── 이상치 목록 ──
    ws.merge_range(row, 0, row, 3, "이상치 숙소 목록", f["section"])
    ws.set_row(row, 22); row += 1

    lo = stats.get("outlier_low", 0)
    hi = stats.get("outlier_high", float("inf"))
    outliers = [l for l in listings if is_outlier(l, stats)]

    ws.write(row, 0, f"기준: ₩{lo:,.0f} 미만 또는 ₩{hi:,.0f} 초과",
             wb.add_format({"font_size": 9, "font_color": C_GRAY}))
    ws.set_row(row, 16); row += 1

    for c, h in enumerate(["숙소명", "1박가", "구분", "평점"]):
        ws.write(row, c, h, f["header"])
    ws.set_row(row, 20); row += 1

    for l in outliers[:20]:
        p = l.get("price_per_night", 0)
        tf = f["red_text"] if p > hi else f["text"]
        nf = f["red_num"]  if p > hi else f["num"]
        ws.write(row, 0, (l.get("title") or "")[:40], tf)
        ws.write(row, 1, p,                             nf)
        ws.write(row, 2, l.get("host_type", ""),        tf)
        ws.write(row, 3, l.get("rating") or 0,          f["dec_c"])
        ws.set_row(row, 18); row += 1


def _build_rawdata_sheet(wb, f, listings, stats):
    ws = wb.add_worksheet("📂 원본 데이터")

    base_cols = [
        ("번호",          "idx",             5),
        ("숙소명",        "title",           40),
        ("구분",          "host_type",       8),
        ("이상치",        "_outlier",        7),
        ("숙소유형",      "room_type",       13),
        ("건물유형",      "property_type",   13),
        ("침실",          "bedrooms",        7),
        ("침대",          "beds",            7),
        ("욕실",          "bathrooms",       7),
        ("평점",          "rating",          8),
        ("평일가",        "price_weekday",   14),
        ("주말가",        "price_weekend",   14),
        ("1박가",         "price_per_night", 14),
        ("예약률",        "calendar_occ",    9),
        ("최대인원",      "max_guests",      8),
        ("침대종류",      "bed_types",       22),
        ("소개글",        "description",     40),
        ("편의시설",      "amenities",       35),
        ("호스트명",      "host_name",       14),
        ("슈퍼호스트",    "superhost",       9),
        ("하우스룰",      "house_rules",     28),
        ("취소정책",      "cancellation",    20),
        ("링크",          "url",             50),
        ("위도",          "latitude",        12),
        ("경도",          "longitude",       12),
    ]

    # 배너
    ws.merge_range(0, 0, 0, len(base_cols) - 1, "원본 데이터 (전체)", f["section"])
    ws.set_row(0, 24)

    for ci, (h, _, w) in enumerate(base_cols):
        ws.set_column(ci, ci, w)
        ws.write(1, ci, h, f["header"])
    ws.set_row(1, 22)

    lo = stats.get("outlier_low", 0)
    hi = stats.get("outlier_high", float("inf"))

    for ri, lst in enumerate(listings):
        row = 2 + ri
        alt = ri % 2 == 1
        p   = lst.get("price_per_night") or 0
        oo  = p < lo or p > hi

        for ci, (_, field, _) in enumerate(base_cols):
            if oo:
                base_fmt = f["red_text"]
                n_fmt    = f["red_num"]
            else:
                base_fmt = f["text_alt"] if alt else f["text"]
                n_fmt    = f["num_alt"]  if alt else f["num"]

            if field == "idx":
                ws.write(row, ci, ri + 1, n_fmt)
            elif field == "_outlier":
                ws.write(row, ci, "⚠ 이상치" if oo else "", base_fmt)
            elif field == "url":
                lnk_fmt = wb.add_format({"font_color": C_LINK, "underline": True,
                                         "bg_color": "#FFDAD5" if oo else ("#EBF3FB" if alt else "#FFFFFF"),
                                         "border": 1, "border_color": "#BDD7EE"})
                ws.write_url(row, ci, str(lst.get(field, "")), lnk_fmt, str(lst.get(field, "")))
            elif field in ("price_weekday", "price_weekend", "price_per_night"):
                v = lst.get(field)
                ws.write(row, ci, v if v else "", n_fmt)
            elif field in ("bedrooms", "beds", "max_guests"):
                ws.write(row, ci, int(lst.get(field) or 0), n_fmt)
            elif field in ("bathrooms", "rating"):
                ws.write(row, ci, lst.get(field) or 0,
                         wb.add_format({"num_format": "0.0", "bg_color": "#FFDAD5" if oo else ("#EBF3FB" if alt else "#FFFFFF"), "border": 1, "border_color": "#BDD7EE"}))
            elif field in ("latitude", "longitude"):
                ws.write(row, ci, lst.get(field) or 0,
                         wb.add_format({"num_format": "0.0000", "bg_color": "#FFDAD5" if oo else ("#EBF3FB" if alt else "#FFFFFF"), "border": 1, "border_color": "#BDD7EE"}))
            elif field == "calendar_occ":
                v = lst.get(field)
                pf = wb.add_format({"num_format": "0.0%", "bg_color": "#FFDAD5" if oo else ("#EBF3FB" if alt else "#FFFFFF"), "border": 1, "border_color": "#BDD7EE", "align": "center"})
                ws.write(row, ci, v if v is not None else "", pf)
            else:
                ws.write(row, ci, str(lst.get(field) or ""), base_fmt)

        ws.set_row(row, 40 if any(lst.get(c[1]) for c in base_cols if c[1] in ("description", "amenities")) else 18)

    ws.freeze_panes(2, 2)
    ws.autofilter(1, 0, 1 + len(listings), len(base_cols) - 1)


def _build_segment_sheet(wb, f, listings):
    ws = wb.add_worksheet("🔍 세그먼트 분석")

    ws.set_column(0, 0, 16)
    ws.set_column(1, 6, 14)

    row = 0
    ws.merge_range(row, 0, row, 6, "개인 vs 상업용 × 침실별 교차 분석", f["section"])
    ws.set_row(row, 24); row += 1

    ws.write(row, 0, "침실 수", f["header_dark"])
    for c, h in enumerate(["개인 숙소 수", "개인 평균가", "상업용 숙소 수", "상업용 평균가", "전체", "전체 평균가"]):
        ws.write(row, c + 1, h, f["header"])
    ws.set_row(row, 22); row += 1

    beds_set = sorted({int(l.get("bedrooms") or 0) for l in listings})
    for i, b in enumerate(beds_set):
        alt = i % 2 == 1
        subset = [l for l in listings if int(l.get("bedrooms") or 0) == b]
        indiv  = [l for l in subset if l.get("host_type") == "개인"]
        comm   = [l for l in subset if l.get("host_type") == "상업용"]
        ip     = [l.get("price_per_night", 0) for l in indiv if l.get("price_per_night")]
        cp     = [l.get("price_per_night", 0) for l in comm  if l.get("price_per_night")]
        ap     = ip + cp

        tf = f["text_alt"] if alt else f["text"]
        nf = f["num_alt"]  if alt else f["num"]
        cf = f["center_alt"] if alt else f["center"]

        ws.write(row, 0, f"{b}침실" if b else "미분류", tf)
        ws.write(row, 1, len(indiv), cf)
        ws.write(row, 2, int(statistics.mean(ip)) if ip else 0, nf)
        ws.write(row, 3, len(comm), cf)
        ws.write(row, 4, int(statistics.mean(cp)) if cp else 0, nf)
        ws.write(row, 5, len(subset), cf)
        ws.write(row, 6, int(statistics.mean(ap)) if ap else 0, nf)
        ws.set_row(row, 18); row += 1

    # 슈퍼호스트 비교
    row += 2
    ws.merge_range(row, 0, row, 6, "슈퍼호스트 vs 일반 가격 비교", f["section"])
    ws.set_row(row, 22); row += 1

    for c, h in enumerate(["구분", "숙소 수", "평균가", "중앙값", "P75"]):
        ws.write(row, c, h, f["header"])
    ws.set_row(row, 20); row += 1

    groups = [
        ("슈퍼호스트", [l for l in listings if l.get("superhost") == "슈퍼호스트"]),
        ("일반 호스트", [l for l in listings if l.get("superhost") != "슈퍼호스트"]),
    ]
    for i, (label, subset) in enumerate(groups):
        prices_s = [l.get("price_per_night", 0) for l in subset if l.get("price_per_night")]
        alt = i % 2 == 1
        tf = f["text_alt"] if alt else f["text"]
        nf = f["num_alt"]  if alt else f["num"]
        cf = f["center_alt"] if alt else f["center"]
        ws.write(row, 0, label, tf)
        ws.write(row, 1, len(subset), cf)
        ws.write(row, 2, int(statistics.mean(prices_s))   if prices_s else 0, nf)
        ws.write(row, 3, int(statistics.median(prices_s)) if prices_s else 0, nf)
        ws.write(row, 4, int(_percentile(prices_s, 75))   if prices_s else 0, nf)
        ws.set_row(row, 18); row += 1


def _build_pricedist_sheet(wb, f, listings, stats):
    ws = wb.add_worksheet("📊 가격 분포")

    ws.set_column(0, 0, 14)
    ws.set_column(1, 1, 14)
    ws.set_column(2, 6, 12)

    row = 0
    ws.merge_range(row, 0, row, 6, "가격 분포 상세 분석", f["section"])
    ws.set_row(row, 24); row += 1

    # 퍼센타일 테이블
    ws.write(row, 0, "퍼센타일", f["header_dark"])
    ws.write(row, 1, "가격",     f["header"])
    ws.write(row, 2, "전체 평균 대비", f["header"])
    ws.set_row(row, 22); row += 1

    mean_val = stats.get("mean", 1) or 1
    percentile_rows = [
        ("P10 (하위 10%)", stats.get("p10", 0)),
        ("P25 (하위 25%)", stats.get("p25", 0)),
        ("P40 (추천-스타터)", stats.get("p40", 0)),
        ("P50 (중앙값)",   stats.get("p50", 0)),
        ("P55 (추천-표준)", stats.get("p55", 0)),
        ("P70 (추천-프리미엄)", stats.get("p70", 0)),
        ("P75 (상위 25%)", stats.get("p75", 0)),
        ("P90 (상위 10%)", stats.get("p90", 0)),
    ]
    for i, (label, val) in enumerate(percentile_rows):
        alt = i % 2 == 1
        tf = f["text_alt"] if alt else f["text"]
        nf = f["num_alt"]  if alt else f["num"]
        ratio = val / mean_val if mean_val else 0
        pf = f["pct_alt"] if alt else f["pct"]
        ws.write(row, 0, label, tf)
        ws.write(row, 1, int(val), nf)
        ws.write(row, 2, ratio, pf)
        ws.set_row(row, 18); row += 1

    row += 1
    ws.merge_range(row, 0, row, 6,
                   f"이상치 기준: ₩{stats.get('outlier_low', 0):,.0f} 미만 또는 ₩{stats.get('outlier_high', 0):,.0f} 초과",
                   f["section"])
    ws.set_row(row, 22); row += 1

    # 이상치 목록
    for c, h in enumerate(["숙소명", "1박가", "구분", "침실", "평점", "링크"]):
        ws.write(row, c, h, f["header"])
    ws.set_row(row, 20); row += 1

    outliers = [l for l in listings if is_outlier(l, stats)]
    for l in outliers:
        p  = l.get("price_per_night", 0)
        lo = stats.get("outlier_low", 0)
        hi = stats.get("outlier_high", float("inf"))
        if p > hi:
            fmt_t, fmt_n = f["red_text"], f["red_num"]
        else:
            fmt_t, fmt_n = f["text"], f["num"]

        ws.write(row, 0, (l.get("title") or "")[:35], fmt_t)
        ws.write(row, 1, p, fmt_n)
        ws.write(row, 2, l.get("host_type", ""), fmt_t)
        ws.write(row, 3, int(l.get("bedrooms") or 0), fmt_t)
        ws.write(row, 4, l.get("rating") or 0, f["dec_c"])
        ws.write_url(row, 5, l.get("url", ""), f["link"], "링크")
        ws.set_row(row, 18); row += 1

    # 주말 프리미엄 개별 목록
    row += 1
    ws.merge_range(row, 0, row, 6, "숙소별 주말 프리미엄", f["section"])
    ws.set_row(row, 22); row += 1

    for c, h in enumerate(["숙소명", "평일가", "주말가", "프리미엄(%)", "구분"]):
        ws.write(row, c, h, f["header"])
    ws.set_row(row, 20); row += 1

    both = [l for l in listings if l.get("price_weekday") and l.get("price_weekend")]
    both_sorted = sorted(both, key=lambda x: (x["price_weekend"] / x["price_weekday"]) - 1, reverse=True)

    for i, l in enumerate(both_sorted):
        alt = i % 2 == 1
        wd = l["price_weekday"]
        we = l["price_weekend"]
        prem = (we - wd) / wd if wd > 0 else 0
        tf = f["text_alt"] if alt else f["text"]
        nf = f["num_alt"]  if alt else f["num"]
        pf = f["pct_alt"]  if alt else f["pct"]
        ws.write(row, 0, (l.get("title") or "")[:35], tf)
        ws.write(row, 1, wd,   nf)
        ws.write(row, 2, we,   nf)
        ws.write(row, 3, prem, pf)
        ws.write(row, 4, l.get("host_type", ""), tf)
        ws.set_row(row, 18); row += 1


def _build_meta_sheet(wb, f, query, beds, baths, geo, collected):
    ws = wb.add_worksheet("🗃️ 수집 메타")

    ws.set_column(0, 0, 22)
    ws.set_column(1, 1, 35)

    row = 0
    ws.merge_range(row, 0, row, 1, "수집 메타데이터", f["section"])
    ws.set_row(row, 24); row += 1

    rows = [
        ("생성 일시",       str(date.today()),                                True),
        ("검색 지역",       query,                                             True),
        ("침실 필터",       str(beds) if beds is not None else "전체",        False),
        ("욕실 필터",       str(baths) if baths is not None else "전체",      False),
        ("지오코딩 위도",   str(geo.get("lat", "")),                          False),
        ("지오코딩 경도",   str(geo.get("lon", "")),                          False),
        ("",               "",                                                 False),
        ("평일 수집 날짜",  f"{collected['windows'][0][0]} ~ {collected['windows'][0][1]}", True),
        ("평일 수집 수",    f"{collected['wd_count']}개",                     False),
        ("주말 수집 날짜",  f"{collected['windows'][1][0]} ~ {collected['windows'][1][1]}", True),
        ("주말 수집 수",    f"{collected['we_count']}개",                     False),
        ("공통 숙소 수",    f"{collected['common_count']}개 (주말 프리미엄 계산 기준)", True),
        ("최종 분석 수",    f"{len(collected['listings'])}개",                True),
        ("",               "",                                                 False),
        ("분류 로직",       "상업용 스코어링 v1.0 (property_type+host_name+소개글)", False),
        ("이상치 방법",     "IQR × 1.5 (Q1-1.5IQR ~ Q3+1.5IQR)",            False),
        ("추천가 기준",     "P40(스타터) / P55(표준) / P70(프리미엄)",        False),
        ("데이터 소스",     "Airbnb.com (비공식 크롤링, 개인 분석 목적)",     False),
    ]

    for k, v, hl in rows:
        if not k:
            ws.set_row(row, 8); row += 1; continue
        ws.write(row, 0, k, f["sub_key_hl"] if hl else f["sub_key"])
        ws.write(row, 1, v, f["sub_val_hl"] if hl else f["sub_val"])
        ws.set_row(row, 18); row += 1


def build_internal_report(
    listings: list[dict],
    query: str,
    beds: int | None,
    baths: float | None,
    windows: list,
    stats: dict,
    premium: float,
    collected: dict,
    out_path: Path,
    demand: dict | None = None,
    avg_calendar_occ: float | None = None,
) -> None:
    wb = xlsxwriter.Workbook(str(out_path))
    f  = _make_formats(wb)

    _build_dashboard_sheet(wb, f, listings, stats, premium, windows, avg_calendar_occ=avg_calendar_occ)
    _build_rawdata_sheet(wb, f, listings, stats)
    _build_segment_sheet(wb, f, listings)
    _build_pricedist_sheet(wb, f, listings, stats)
    _build_meta_sheet(wb, f, query, beds, baths, collected["geo"], collected)
    if demand:
        _build_demand_sheet(wb, f, demand, listings)
    _build_guide_sheet(wb, f, is_internal=True)

    wb.close()
    _fix_colors(out_path)
    print(f"✅ 내부용 저장: {out_path.name}")


# ══════════════════════════════════════════════════════════════════
# 8. HTML 리포트 빌더 (Toss 스타일)
# ══════════════════════════════════════════════════════════════════

_HTML_CSS = """
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
:root {
  --blue: #3182F6; --blue-2: #5BA3F8; --blue-bg: #EBF3FE;
  --navy: #1A1E2F; --navy-2: #1B3A8A;
  --bg: #F2F4F6; --card: #FFFFFF;
  --text-1: #191F28; --text-2: #6B7684; --text-3: #B0B8C1;
  --border: #E5E8EB;
  --green: #00B84A; --green-bg: #E5F8EE;
  --orange: #FF6B00; --orange-bg: #FFF1E5;
  --gold: #FFB800; --gold-bg: #FFF8E1;
  --r-lg: 20px; --r-md: 16px; --r-sm: 10px;
  --shadow: 0 4px 24px rgba(0,0,0,0.10);
  --shadow-sm: 0 2px 12px rgba(0,0,0,0.06);
}
body {
  font-family: 'Pretendard Variable', Pretendard, -apple-system, BlinkMacSystemFont,
               'Apple SD Gothic Neo', 'Noto Sans KR', sans-serif;
  background: var(--bg); color: var(--text-1);
  line-height: 1.6; -webkit-font-smoothing: antialiased;
}
.container { max-width: 960px; margin: 0 auto; padding: 0 24px; }

/* ── Hero ── */
.hero {
  background: linear-gradient(145deg, #0F1629 0%, #1B3A8A 60%, #2A1F5A 100%);
  padding: 56px 0 110px;
}
.hero-tag {
  display: inline-block; background: rgba(49,130,246,0.25); color: #90C2FF;
  font-size: 12px; font-weight: 600; padding: 5px 14px; border-radius: 100px;
  margin-bottom: 18px; letter-spacing: 0.04em;
}
.hero-title {
  font-size: 38px; font-weight: 900; color: #fff; margin-bottom: 6px; line-height: 1.2;
}
.hero-cond { font-size: 22px; font-weight: 500; color: rgba(255,255,255,0.55); }
.hero-meta { font-size: 13px; color: rgba(255,255,255,0.38); margin-top: 10px; margin-bottom: 0; }

/* ── KPI row (floats over hero) ── */
.kpi-row {
  display: grid; grid-template-columns: repeat(3, 1fr); gap: 16px;
  margin-top: -64px; margin-bottom: 0;
}
.kpi-card {
  background: var(--card); border-radius: var(--r-md);
  padding: 24px 22px; box-shadow: var(--shadow);
}
.kpi-label { font-size: 12px; font-weight: 500; color: var(--text-2); margin-bottom: 8px; }
.kpi-val {
  font-size: 30px; font-weight: 900; color: var(--text-1); line-height: 1;
}
.kpi-unit { font-size: 16px; font-weight: 600; color: var(--text-2); margin-left: 2px; }
.kpi-sub { font-size: 11px; color: var(--text-3); margin-top: 6px; }

/* ── Main ── */
main { padding: 48px 0 80px; }
.section { margin-bottom: 44px; }
.section-title {
  font-size: 20px; font-weight: 800; color: var(--text-1); margin-bottom: 20px;
}
.card {
  background: var(--card); border-radius: var(--r-md);
  padding: 28px; box-shadow: var(--shadow-sm);
}

/* ── Metric grid ── */
.metric-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 14px; margin-bottom: 20px; }
.metric-item {
  background: var(--card); border-radius: var(--r-md);
  padding: 20px 18px; box-shadow: var(--shadow-sm);
}
.metric-label { font-size: 12px; color: var(--text-2); margin-bottom: 6px; }
.metric-value { font-size: 26px; font-weight: 800; color: var(--text-1); }
.metric-value.blue  { color: var(--blue); }
.metric-value.green { color: var(--green); }

/* ── Compare bars ── */
.compare-row {
  display: flex; align-items: center; gap: 16px;
  padding: 16px 0; border-bottom: 1px solid var(--border);
}
.compare-row:last-child { border-bottom: none; }
.compare-label { font-size: 13px; font-weight: 600; color: var(--text-2); width: 72px; flex-shrink: 0; }
.compare-bar-wrap {
  flex: 1; height: 10px; background: var(--border); border-radius: 5px; overflow: hidden;
}
.compare-bar { height: 100%; border-radius: 5px; background: var(--blue); transition: width 0.6s ease; }
.compare-val { font-size: 15px; font-weight: 800; color: var(--text-1); flex-shrink: 0; text-align: right; }
.compare-premium { font-size: 12px; color: var(--orange); font-weight: 700; }

/* ── Distribution ── */
.dist-row {
  display: flex; align-items: center; gap: 14px;
  padding: 11px 0; border-bottom: 1px solid var(--border);
}
.dist-row:last-child { border-bottom: none; }
.dist-label { font-size: 13px; color: var(--text-2); width: 112px; flex-shrink: 0; font-weight: 500; }
.dist-bar-wrap {
  flex: 1; height: 26px; background: var(--bg); border-radius: 5px;
  overflow: hidden; display: flex;
}
.dist-bar-indiv { height: 100%; background: linear-gradient(90deg, #3182F6, #5BA3F8); }
.dist-bar-comm  { height: 100%; background: linear-gradient(90deg, #B8D0F0, #D8ECF8); }
.dist-count { font-size: 13px; font-weight: 700; color: var(--text-1); width: 38px; text-align: right; flex-shrink: 0; }
.dist-pct   { font-size: 12px; color: var(--text-3); width: 42px; text-align: right; flex-shrink: 0; }
.legend { display: flex; gap: 18px; margin-top: 18px; }
.legend-item { display: flex; align-items: center; gap: 7px; font-size: 12px; color: var(--text-2); }
.legend-dot { width: 12px; height: 12px; border-radius: 3px; }

/* ── Chip ── */
.chip {
  display: inline-block; font-size: 10px; font-weight: 700;
  padding: 2px 7px; border-radius: 5px; line-height: 1.5;
}
.chip-blue   { background: var(--blue-bg); color: var(--blue); }
.chip-orange { background: var(--orange-bg); color: var(--orange); }
.chip-green  { background: var(--green-bg); color: var(--green); }
.chip-gray   { background: var(--bg); color: var(--text-2); }

/* ── Competition list ── */
.listing-list {
  display: flex; flex-direction: column; gap: 1px;
  background: var(--border); border-radius: var(--r-md); overflow: hidden;
}
.listing-row {
  display: grid; grid-template-columns: 36px 1fr auto;
  align-items: center; gap: 16px; padding: 14px 20px;
  background: var(--card); transition: background 0.15s;
}
.listing-row:hover { background: #FAFBFC; }
.rank-badge {
  width: 28px; height: 28px; border-radius: 8px;
  display: flex; align-items: center; justify-content: center;
  font-size: 12px; font-weight: 900; flex-shrink: 0;
}
.r1 { background: #FFD700; color: #7A5F00; }
.r2 { background: #D8D8D8; color: #555; }
.r3 { background: #F0C080; color: #7A4F00; }
.rN { background: var(--bg); color: var(--text-3); }
.listing-name {
  font-size: 13px; font-weight: 700; color: var(--text-1);
  white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
}
.listing-meta { font-size: 11px; color: var(--text-2); margin-top: 3px; }
.listing-price { text-align: right; flex-shrink: 0; }
.price-main { font-size: 15px; font-weight: 800; color: var(--text-1); }
.price-we   { font-size: 11px; color: var(--text-3); margin-top: 2px; }
.star { color: var(--gold); }

/* ── Tier cards ── */
.tier-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; }
.tier-card {
  background: var(--card); border-radius: var(--r-lg);
  padding: 28px 22px 24px; box-shadow: var(--shadow-sm);
  text-align: center;
}
.tier-card.hl {
  background: linear-gradient(145deg, #1A1E2F, #1B3A8A 50%, #3182F6);
  box-shadow: 0 8px 32px rgba(49,130,246,0.30);
}
.tier-badge { font-size: 14px; font-weight: 700; color: var(--text-2); margin-bottom: 18px; }
.tier-card.hl .tier-badge { color: rgba(255,255,255,0.65); }
.tier-wd-label { font-size: 11px; color: var(--text-3); margin-bottom: 4px; }
.tier-card.hl .tier-wd-label { color: rgba(255,255,255,0.45); }
.tier-price {
  font-size: 34px; font-weight: 900; color: var(--text-1); line-height: 1; margin-bottom: 14px;
}
.tier-card.hl .tier-price { color: #fff; }
.tier-we {
  font-size: 13px; font-weight: 700; color: var(--orange);
  padding: 7px 12px; background: var(--orange-bg); border-radius: 8px; margin-bottom: 14px;
}
.tier-card.hl .tier-we { background: rgba(255,255,255,0.12); color: #FFD090; }
.tier-desc { font-size: 11px; color: var(--text-3); line-height: 1.6; }
.tier-card.hl .tier-desc { color: rgba(255,255,255,0.45); }

/* ── Footer ── */
footer {
  background: var(--navy); color: rgba(255,255,255,0.38);
  padding: 32px 0; font-size: 12px; line-height: 1.9;
}
footer strong { color: rgba(255,255,255,0.65); }

/* ── Responsive ── */
@media (max-width: 680px) {
  .kpi-row, .tier-grid { grid-template-columns: 1fr; }
  .metric-grid { grid-template-columns: repeat(2, 1fr); }
  .hero-title { font-size: 26px; }
  .kpi-row { margin-top: -48px; }
}
"""


def _build_score_html(ms: dict) -> str:
    """시장 매력도 점수 HTML 섹션."""
    total = ms.get("total", 0)
    judgment = ms.get("judgment", "")
    color_map = {"적극 진입": "#00B84A", "조건부 진입": "#3182F6",
                 "보수적 접근": "#FF6B00", "진입 미권장": "#FF3B3B"}
    jcolor = color_map.get(judgment, "#3182F6")
    warn = ('<p style="font-size:11px;color:#FF3B3B;margin-top:12px">'
            '⚠ 샘플 수 부족(15개 미만)으로 신뢰도가 낮습니다.</p>'
            if ms.get("warning") else "")
    factors = [
        ("수요 안정성", ms.get("demand", 0), "시장 평균 리뷰 수"),
        ("주말 수요 탄력성", ms.get("weekend", 0), "주말 프리미엄 크기"),
        ("가격 성장 여력", ms.get("growth", 0), "P70-P40 가격 스프레드"),
        ("진입 품질 허들", ms.get("quality", 0), "경쟁자 평점 역점수"),
    ]
    rows = ""
    for name, score, note in factors:
        pct = score / 25 * 100
        rows += (
            f'<div style="display:flex;align-items:center;gap:12px;margin-bottom:12px">'
            f'<div style="width:120px;font-size:13px;font-weight:600;color:var(--text-1);flex-shrink:0">{name}</div>'
            f'<div style="flex:1;background:#EEF2F7;border-radius:6px;height:10px;overflow:hidden">'
            f'<div style="width:{pct:.0f}%;height:100%;background:linear-gradient(90deg,#3182F6,#5BA3F8);border-radius:6px"></div></div>'
            f'<div style="width:40px;text-align:right;font-size:13px;font-weight:700;color:var(--text-1)">{score}<span style="font-size:10px;color:var(--text-3)">/25</span></div>'
            f'<div style="font-size:11px;color:var(--text-3);width:130px;flex-shrink:0">{note}</div>'
            f'</div>\n'
        )
    return (
        '<div class="section">\n'
        '<h2 class="section-title">시장 매력도 점수</h2>\n'
        '<div class="card">\n'
        f'<div style="display:flex;align-items:center;gap:20px;margin-bottom:28px">'
        f'<div style="font-size:56px;font-weight:900;color:{jcolor};line-height:1">{total}</div>'
        f'<div><div style="font-size:11px;color:var(--text-3)">100점 만점</div>'
        f'<div style="font-size:20px;font-weight:800;color:{jcolor};margin-top:2px">{judgment}</div>'
        f'<div style="font-size:11px;color:var(--text-3);margin-top:4px">4개 지표 × 25점</div></div>'
        f'</div>\n'
        + rows +
        warn +
        '<p style="font-size:11px;color:var(--text-3);margin-top:14px;line-height:1.8">'
        '※ 수요 안정성: 평균 리뷰 수(40개+ = 만점) &nbsp;·&nbsp; '
        '주말 수요 탄력성: 프리미엄 30%+ = 만점<br>'
        '※ 가격 성장 여력: (P70-P40)/P40 &ge;70% = 만점 &nbsp;·&nbsp; '
        '진입 품질 허들: 경쟁자 평점 낮을수록 유리</p>\n'
        '</div>\n</div>\n'  # card, section
    )


def _build_simulator_html(scenarios: list[dict], eff_adr: dict | None) -> str:
    """수익 시뮬레이터 HTML 섹션."""
    def fmt_krw(v: int) -> str:
        return f"₩{v:,}"

    eff_str = ""
    if eff_adr and eff_adr.get("mean"):
        eff_str = (
            f'<div style="background:#FFF8E1;border-radius:12px;padding:14px 18px;margin-bottom:20px;'
            f'font-size:13px;color:#7A4F00">'
            f'<strong>💡 체감 ADR (총가 기준):</strong> 평균 {fmt_krw(int(eff_adr["mean"]))} '
            f'/ 중앙값 {fmt_krw(int(eff_adr["median"]))}'
            f'<br><span style="font-size:11px;opacity:0.7">총가 = 1박가 + 청소비 + Airbnb 서비스료 (게스트 실부담)</span></div>'
        )

    cols = ""
    for sc in scenarios:
        op = sc["op_profit"]
        op_color = "#00B84A" if op >= 0 else "#FF3B3B"
        label_color = {"보수적": "#6B7684", "기준": "#3182F6", "공격적": "#FF6B00"}.get(sc["label"], "#3182F6")
        cols += (
            f'<div style="background:var(--card);border-radius:16px;padding:24px 20px;'
            f'box-shadow:0 2px 12px rgba(0,0,0,0.06);text-align:center">\n'
            f'<div style="font-size:13px;font-weight:700;color:{label_color};margin-bottom:16px">'
            f'{sc["label"]} ({sc["occ_pct"]:.0f}%)</div>\n'
            f'<div style="font-size:11px;color:var(--text-3);margin-bottom:4px">월 영업이익</div>\n'
            f'<div style="font-size:30px;font-weight:900;color:{op_color};line-height:1;margin-bottom:16px">'
            f'{fmt_krw(op)}</div>\n'
            + "".join(
                f'<div style="display:flex;justify-content:space-between;font-size:12px;'
                f'color:var(--text-2);border-top:1px solid var(--border);padding:7px 0">'
                f'<span>{lbl}</span><span style="font-weight:600;color:var(--text-1)">{val}</span></div>\n'
                for lbl, val in [
                    ("예약일수", f"{sc['booking_days']:.1f}일"),
                    ("총 매출", fmt_krw(sc["gross"])),
                    ("플랫폼 수수료", f"−{fmt_krw(sc['platform_fee'])}"),
                    ("청소비", f"−{fmt_krw(sc['cleaning_cost'])}"),
                    ("운영비", f"−{fmt_krw(sc['monthly_ops'])}"),
                ]
            )
            + '</div>\n'
        )

    return (
        '<div class="section">\n'
        '<h2 class="section-title">월 수익 시뮬레이터</h2>\n'
        '<div class="card">\n'
        '<div style="background:#FFF1E5;border-radius:10px;padding:12px 16px;margin-bottom:20px;'
        'font-size:12px;color:#7A3800;font-weight:600">'
        '⚠ 업계 평균 기반 가정치입니다. 임대료·대출 상환은 포함되지 않습니다.</div>\n'
        + eff_str
        + f'<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:16px">\n{cols}</div>\n'
        '<p style="font-size:11px;color:var(--text-3);margin-top:14px;line-height:1.8">'
        '* 영업이익 = 총 매출 − 플랫폼 수수료(3%) − 청소비 − 월 운영비<br>'
        '* 블렌디드 ADR = 평일 ADR × 5/7 + 주말 ADR × 2/7</p>\n'
        '</div>\n</div>\n'  # card, section
    )


def build_html_report(
    listings: list[dict],
    query: str,
    beds: int | None,
    baths: float | None,
    windows: list,
    stats: dict,
    premium: float,
    out_path: Path,
    market_score: dict | None = None,
    scenarios: list[dict] | None = None,
    demand: dict | None = None,
    eff_adr: dict | None = None,
) -> None:
    """Toss 스타일 자기완결형 HTML 시장 분석 리포트 생성."""
    if not listings:
        print("⚠ HTML: 숙소 없음, 건너뜀")
        return

    today = date.today().isoformat()
    indiv_lst = [l for l in listings if l.get("host_type") == "개인"]
    comm_lst  = [l for l in listings if l.get("host_type") == "상업용"]
    total_cnt = len(listings)
    avg_price = stats.get("mean_clean", stats.get("mean", 0))
    avg_rating = stats.get("avg_rating", 0)
    median    = stats.get("median", 0)
    p40 = stats.get("p40", 0)
    p55 = stats.get("p55", 0)
    p70 = stats.get("p70", 0)

    cond_parts = []
    if beds  is not None: cond_parts.append(f"{beds}침실")
    if baths is not None: cond_parts.append(f"{int(baths) if baths == int(baths) else baths}욕실")
    cond_str = " ".join(cond_parts) if cond_parts else "전체"

    wd_label = f"{windows[0][0]}(평일)"
    we_label = f"{windows[1][0]}(주말)"

    # 가격 구간 분포 (빈 구간 제외)
    dist = price_distribution(listings)
    max_dist = max((d["count"] for d in dist), default=1) or 1

    def dist_row_html(d: dict) -> str:
        if d["count"] == 0:
            return ""
        bw   = d["count"] / max_dist * 100
        iw   = d["indiv"] / d["count"] * bw if d["count"] else 0
        cw   = bw - iw
        return (
            f'<div class="dist-row">'
            f'<span class="dist-label">{d["label"]}</span>'
            f'<div class="dist-bar-wrap">'
            f'<div class="dist-bar-indiv" style="width:{iw:.1f}%"></div>'
            f'<div class="dist-bar-comm"  style="width:{cw:.1f}%"></div>'
            f'</div>'
            f'<span class="dist-count">{d["count"]}</span>'
            f'<span class="dist-pct">{d["pct"]:.1f}%</span>'
            f'</div>'
        )

    dist_html = "".join(dist_row_html(d) for d in dist)

    # 평일/주말 비교
    wd_prices = [l["price_weekday"] for l in listings if l.get("price_weekday")]
    we_prices = [l["price_weekend"]  for l in listings if l.get("price_weekend")]
    wd_avg = int(statistics.mean(wd_prices)) if wd_prices else 0
    we_avg = int(statistics.mean(we_prices)) if we_prices else 0
    max_p  = max(wd_avg, we_avg) or 1

    # 경쟁 현황 (평점 상위 15개)
    top = sorted(
        [l for l in listings if l.get("rating")],
        key=lambda x: x.get("rating", 0), reverse=True,
    )[:15]

    def listing_row_html(rank: int, l: dict) -> str:
        r_cls  = ["r1","r2","r3"][rank-1] if rank <= 3 else "rN"
        ht     = l.get("host_type","개인")
        c_cls  = "chip-blue" if ht == "개인" else "chip-orange"
        sup    = '<span class="chip chip-green">슈퍼호스트</span>&nbsp;' if l.get("superhost") == "슈퍼호스트" else ""
        wd_p   = l.get("price_weekday") or l.get("price_per_night") or 0
        we_p   = l.get("price_weekend")
        we_str = f'<div class="price-we">주말 ₩{we_p:,}</div>' if we_p else ""
        beds_n  = int(l.get("bedrooms")  or 0)
        baths_n = float(l.get("bathrooms") or 0)
        baths_s = f"{int(baths_n)}" if baths_n == int(baths_n) else f"{baths_n}"
        return (
            f'<div class="listing-row">'
            f'<div class="rank-badge {r_cls}">{rank}</div>'
            f'<div>'
            f'<div class="listing-name">{l.get("title","")[:42]}</div>'
            f'<div class="listing-meta">'
            f'{sup}<span class="chip {c_cls}">{ht}</span>'
            f'&nbsp;·&nbsp;{beds_n}침실 {baths_s}욕실'
            f'&nbsp;·&nbsp;<span class="star">★</span> {l.get("rating",0):.2f}'
            f'</div>'
            f'</div>'
            f'<div class="listing-price">'
            f'<div class="price-main">₩{wd_p:,}</div>'
            f'{we_str}'
            f'</div>'
            f'</div>'
        )

    listing_html = "".join(listing_row_html(i+1, l) for i, l in enumerate(top))

    def rnd1k(v: float) -> int:
        return round(v / 1000) * 1000

    p40_wd = rnd1k(p40);          p40_we = rnd1k(p40 * (1 + premium))
    p55_wd = rnd1k(p55);          p55_we = rnd1k(p55 * (1 + premium))
    p70_wd = rnd1k(p70);          p70_we = rnd1k(p70 * (1 + premium))

    # ── HTML 조립 ──────────────────────────────────────────────
    html = (
        '<!DOCTYPE html>\n'
        '<html lang="ko">\n'
        '<head>\n'
        '<meta charset="UTF-8">\n'
        '<meta name="viewport" content="width=device-width, initial-scale=1.0">\n'
        f'<title>에어비앤비 시장 분석 — {query}</title>\n'
        f'<style>{_HTML_CSS}</style>\n'
        '</head>\n'
        '<body>\n'

        # Hero
        '<header class="hero">\n'
        '<div class="container">\n'
        '<div class="hero-tag">에어비앤비 시장 분석 리포트</div>\n'
        f'<h1 class="hero-title">{query} <span class="hero-cond">{cond_str} 시장</span></h1>\n'
        f'<p class="hero-meta">분석일 {today} &nbsp;·&nbsp; 수집창 {wd_label} / {we_label} &nbsp;·&nbsp; 총 {total_cnt}개 숙소</p>\n'
        '</div>\n'
        '</header>\n'

        # KPI row
        '<div class="container">\n'
        '<div class="kpi-row">\n'
        '<div class="kpi-card">\n'
        '<div class="kpi-label">수집 숙소</div>\n'
        f'<div class="kpi-val">{total_cnt}<span class="kpi-unit">개</span></div>\n'
        f'<div class="kpi-sub">개인 {len(indiv_lst)}개 &nbsp;·&nbsp; 상업용 {len(comm_lst)}개</div>\n'
        '</div>\n'
        '<div class="kpi-card">\n'
        '<div class="kpi-label">평균 1박가 (이상치 제거)</div>\n'
        f'<div class="kpi-val">₩{avg_price:,.0f}</div>\n'
        f'<div class="kpi-sub">중앙값 ₩{median:,.0f}</div>\n'
        '</div>\n'
        '<div class="kpi-card">\n'
        '<div class="kpi-label">평균 평점</div>\n'
        f'<div class="kpi-val" style="color:#3182F6">{"★ " + f"{avg_rating:.2f}" if avg_rating else "—"}</div>\n'
        f'<div class="kpi-sub">{stats.get("rating_count",0)}개 숙소 기준</div>\n'
        '</div>\n'
        '</div>\n'  # kpi-row
        '</div>\n'  # container

        # Main
        '<main>\n<div class="container">\n'

        # Section 1: 시장 현황
        '<div class="section">\n'
        '<h2 class="section-title">시장 현황</h2>\n'
        '<div class="metric-grid">\n'
        '<div class="metric-item"><div class="metric-label">개인 숙소</div>'
        f'<div class="metric-value blue">{len(indiv_lst)}<span style="font-size:14px;font-weight:500;color:var(--text-2)"> 개</span></div></div>\n'
        '<div class="metric-item"><div class="metric-label">상업용 숙소</div>'
        f'<div class="metric-value">{len(comm_lst)}<span style="font-size:14px;font-weight:500;color:var(--text-2)"> 개</span></div></div>\n'
        '<div class="metric-item"><div class="metric-label">주말 프리미엄</div>'
        f'<div class="metric-value {"green" if premium > 0 else ""}">+{premium*100:.1f}<span style="font-size:14px;font-weight:500;color:var(--text-2)">%</span></div></div>\n'
        '<div class="metric-item"><div class="metric-label">P75 (상위 25%)</div>'
        f'<div class="metric-value" style="font-size:20px">₩{stats.get("p75",0):,.0f}</div></div>\n'
        '</div>\n'  # metric-grid

        # Compare bars
        '<div class="card">\n'
        '<h3 style="font-size:14px;font-weight:700;color:var(--text-1);margin-bottom:18px">평일 vs 주말 평균 1박가</h3>\n'
        '<div class="compare-row">\n'
        '<span class="compare-label">평일</span>\n'
        '<div class="compare-bar-wrap">'
        f'<div class="compare-bar" style="width:{wd_avg/max_p*100:.1f}%"></div>'
        '</div>\n'
        f'<span class="compare-val">₩{wd_avg:,}</span>\n'
        '</div>\n'
        '<div class="compare-row">\n'
        '<span class="compare-label">주말</span>\n'
        '<div class="compare-bar-wrap">'
        f'<div class="compare-bar" style="width:{we_avg/max_p*100:.1f}%;background:linear-gradient(90deg,#1B3A8A,#3182F6)"></div>'
        '</div>\n'
        f'<span class="compare-val">₩{we_avg:,} <span class="compare-premium">+{premium*100:.1f}%</span></span>\n'
        '</div>\n'
        '</div>\n'  # card
        '</div>\n'  # section

        # Section 2: 가격 구간 분포
        '<div class="section">\n'
        '<h2 class="section-title">가격 구간 분포</h2>\n'
        '<div class="card">\n'
        + dist_html +
        '<div class="legend">\n'
        '<div class="legend-item"><div class="legend-dot" style="background:linear-gradient(90deg,#3182F6,#5BA3F8)"></div>개인 숙소</div>\n'
        '<div class="legend-item"><div class="legend-dot" style="background:linear-gradient(90deg,#B8D0F0,#D8ECF8)"></div>상업용 숙소</div>\n'
        '</div>\n'
        '</div>\n'  # card
        '</div>\n'  # section

        # Section 3: 경쟁 현황
        f'<div class="section">\n'
        f'<h2 class="section-title">경쟁 현황 <span style="font-size:14px;font-weight:500;color:var(--text-2)">평점 상위 {len(top)}개</span></h2>\n'
        '<div class="listing-list">\n'
        + listing_html +
        '</div>\n'
        '</div>\n'  # section

        # Section 4: 추천 객단가
        '<div class="section">\n'
        '<h2 class="section-title">추천 객단가</h2>\n'
        '<div class="tier-grid">\n'

        '<div class="tier-card">\n'
        '<div class="tier-badge">🌱 스타터</div>\n'
        '<div class="tier-wd-label">평일 1박</div>\n'
        f'<div class="tier-price">₩{p40_wd:,}</div>\n'
        f'<div class="tier-we">주말 ₩{p40_we:,}</div>\n'
        '<div class="tier-desc">리뷰 확보 초기 전략<br>시장 하위 40% 수준</div>\n'
        '</div>\n'

        '<div class="tier-card hl">\n'
        '<div class="tier-badge">⭐ 표준</div>\n'
        '<div class="tier-wd-label">평일 1박</div>\n'
        f'<div class="tier-price">₩{p55_wd:,}</div>\n'
        f'<div class="tier-we">주말 ₩{p55_we:,}</div>\n'
        '<div class="tier-desc">리뷰 10개+ 기준 적정가<br>시장 중심 포지션</div>\n'
        '</div>\n'

        '<div class="tier-card">\n'
        '<div class="tier-badge">👑 프리미엄</div>\n'
        '<div class="tier-wd-label">평일 1박</div>\n'
        f'<div class="tier-price">₩{p70_wd:,}</div>\n'
        f'<div class="tier-we">주말 ₩{p70_we:,}</div>\n'
        '<div class="tier-desc">슈퍼호스트 등급 목표가<br>시장 상위 30% 수준</div>\n'
        '</div>\n'

        '</div>\n'  # tier-grid
        f'<p style="font-size:11px;color:var(--text-3);margin-top:18px;line-height:1.9">'
        f'※ 이상치 제거(IQR×1.5) 후 {stats.get("clean_count", total_cnt)}개 숙소 기준 · 주말 프리미엄 +{premium*100:.1f}% 자동 적용<br>'
        f'※ 리뷰 누적 후 단계적 인상 권장 (스타터 → 표준 → 프리미엄)'
        f'</p>\n'
        '</div>\n'  # section

        # Section 5: 시장 매력도 점수 (v3)
        + (_build_score_html(market_score) if market_score else '')

        # Section 6: 수익 시뮬레이터 (v3)
        + (_build_simulator_html(scenarios, eff_adr) if scenarios else '')

        + '</div>\n</main>\n'  # container, main

        # Footer
        + '<footer>\n<div class="container">\n'
        f'<strong>데이터 출처:</strong> Airbnb.com (비공식 크롤링, 개인 분석 목적) &nbsp;·&nbsp; '
        f'<strong>생성:</strong> {today} &nbsp;·&nbsp; '
        f'<strong>지역:</strong> {query} &nbsp;·&nbsp; '
        f'<strong>수집창:</strong> {wd_label} / {we_label}<br>'
        f'본 리포트는 수집 시점 기준이며, 시장 상황 변동에 따라 달라질 수 있습니다.'
        '\n</div>\n</footer>\n'

        '</body>\n</html>'
    )

    out_path.write_text(html, encoding="utf-8")
    print(f"✅ HTML  저장: {out_path.name}")


# ══════════════════════════════════════════════════════════════════
# 9. 메인
# ══════════════════════════════════════════════════════════════════

def run(
    query: str,
    beds:          int   | None = None,
    baths:         float | None = None,
    output_mode:   str         = "both",
    occ_low:       float       = 0.40,
    occ_base:      float       = 0.60,
    occ_high:      float       = 0.70,
    cleaning_fee:  int         = 80_000,
    avg_nights:    float       = 2.0,
    monthly_cost:  int         = 0,
    checkin:       str  | None = None,
) -> None:
    today = date.today().isoformat()

    target_date: date | None = None
    if checkin:
        try:
            target_date = date.fromisoformat(checkin)
        except ValueError:
            print(f"⚠ --checkin 날짜 형식 오류 ({checkin}), 자동 날짜 사용")

    print(f"\n{'='*60}")
    print(f" Airbnb 시장 분석 리포트 v3 — {query}")
    print(f"{'='*60}")
    if target_date: print(f" 기준 날짜: {target_date.isoformat()} 주간")
    if beds  is not None: print(f" 침실 필터: {beds}개")
    if baths is not None: print(f" 욕실 필터: {baths}개")
    print(f" 예약률 가정: {occ_low*100:.0f}% / {occ_base*100:.0f}% / {occ_high*100:.0f}%")
    print(f" 청소비: ₩{cleaning_fee:,} / 평균 체류: {avg_nights}박")

    print(f"\n[1/5] 지오코딩: {query}")
    geo = geocode_region(query)
    print(f"      → lat={geo.get('lat')}, lon={geo.get('lon')}")

    print("\n[2/5] 데이터 수집 (2개 날짜창 × B 모드)")
    collected = collect_data(query, geo, beds, baths, target_date)
    listings  = collected["listings"]

    if not listings:
        print("❌ 수집된 숙소가 없습니다. 지역명을 더 넓게 입력해보세요.")
        return

    print(f"\n[3/5] 통계 분석 ({len(listings)}개 숙소)")
    stats   = compute_stats(listings)
    premium = compute_weekend_premium(listings)
    windows = collected["windows"]
    avg_calendar_occ = collected.get("avg_calendar_occ")

    print(f"      평균가: ₩{stats.get('mean', 0):,.0f}")
    print(f"      중앙값: ₩{stats.get('median', 0):,.0f}")
    print(f"      주말 프리미엄: +{premium*100:.1f}%")
    print(f"      이상치: {stats.get('outlier_count', 0)}개 제거")
    if avg_calendar_occ is not None:
        print(f"      시장 평균 예약률 (90일): {avg_calendar_occ*100:.1f}%")

    print(f"\n[4/5] v3 분석 (체감 ADR · 수요 추정 · 시장 매력도 · 수익 시뮬레이터)")
    eff_adr      = compute_effective_adr(listings)
    demand       = compute_demand_signal(listings)
    center_lat   = float(geo.get("lat") or 0)
    center_lon   = float(geo.get("lon") or 0)
    adr_clean    = stats.get("mean_clean") or stats.get("mean") or 0
    comps        = compute_comp_similarity(listings, beds, baths, center_lat, center_lon, adr_clean)
    market_score = compute_market_score(listings, stats, premium, demand)
    scenarios    = build_revenue_scenarios(
        adr_clean             = adr_clean,
        premium               = premium,
        occ_low               = occ_low,
        occ_base              = occ_base,
        occ_high              = occ_high,
        cleaning_fee_per_stay = cleaning_fee,
        avg_stay_nights       = avg_nights,
        monthly_ops_cost      = monthly_cost,
    )

    print(f"      체감 ADR (총가): ₩{eff_adr.get('mean', 0):,.0f}")
    print(f"      수요 추정 예약률 (기준 70% 작성률): {demand.get('occ_estimates', {}).get('mid', 0):.1f}%")
    print(f"      시장 매력도: {market_score['total']}점 / 100점 ({market_score['judgment']})")
    print(f"      기준 시나리오 월 영업이익: ₩{scenarios[1]['op_profit']:,}")

    today_str = date.today().strftime("%Y%m%d")
    base = Path(__file__).parent / "output" / f"{today_str}_{query}"
    base.mkdir(parents=True, exist_ok=True)
    print("\n[5/5] 리포트 생성 (손님용 + 내부용 + HTML)")

    if output_mode in ("both", "client"):
        out_client = base / f"{query}_시장분석_손님용.xlsx"
        build_client_report(
            listings, query, beds, baths, windows, stats, premium, out_client,
            scenarios=scenarios, market_score=market_score,
            cleaning_fee=cleaning_fee, avg_stay=avg_nights,
            monthly_ops=monthly_cost,
            avg_calendar_occ=avg_calendar_occ,
        )

    if output_mode in ("both", "internal"):
        out_internal = base / f"{query}_시장분석_내부용.xlsx"
        build_internal_report(
            listings, query, beds, baths, windows, stats, premium, collected, out_internal,
            demand=demand,
            avg_calendar_occ=avg_calendar_occ,
        )

    if output_mode in ("both", "client"):
        out_html = base / f"{query}_시장분석.html"
        build_html_report(listings, query, beds, baths, windows, stats, premium, out_html,
                          market_score=market_score, scenarios=scenarios,
                          demand=demand, eff_adr=eff_adr)

    print(f"\n{'='*60}")
    print(f" 완료!  →  output/{today_str}_{query}/")
    if output_mode in ("both", "client"):
        print(f"  📊 {query}_시장분석_손님용.xlsx")
    if output_mode in ("both", "internal"):
        print(f"  📋 {query}_시장분석_내부용.xlsx")
    if output_mode in ("both", "client"):
        print(f"  🌐 {query}_시장분석.html")
    print(f"{'='*60}\n")


def main() -> None:
    parser = argparse.ArgumentParser(description="Airbnb 시장 분석 리포트 생성기 v3")
    parser.add_argument("query",          help="검색 지역 (예: 충신동)")
    parser.add_argument("--beds",  type=int,   default=None)
    parser.add_argument("--baths", type=float, default=None)
    parser.add_argument("--mode",  choices=["both", "client", "internal"], default="both")
    parser.add_argument("--occ-low",      type=float, default=0.40)
    parser.add_argument("--occ-base",     type=float, default=0.60)
    parser.add_argument("--occ-high",     type=float, default=0.70)
    parser.add_argument("--cleaning-fee", type=int,   default=80_000)
    parser.add_argument("--avg-nights",   type=float, default=2.0)
    parser.add_argument("--monthly-cost", type=int,   default=0)
    parser.add_argument("--checkin",      type=str,   default=None,
                        help="기준 체크인 날짜 YYYY-MM-DD (해당 주의 평일/주말 창 사용)")
    args = parser.parse_args()

    run(
        query        = args.query,
        beds         = args.beds,
        baths        = args.baths,
        output_mode  = args.mode,
        occ_low      = args.occ_low,
        occ_base     = args.occ_base,
        occ_high     = args.occ_high,
        cleaning_fee = args.cleaning_fee,
        avg_nights   = args.avg_nights,
        monthly_cost = args.monthly_cost,
        checkin      = args.checkin,
    )


if __name__ == "__main__":
    main()
