from __future__ import annotations

import math
from datetime import date, timedelta

from flask import Flask, jsonify, render_template, request

from airbnb_fetch import AirbnbError, crawl_airbnb, geocode_region

app = Flask(__name__)

# 한국 영토 바운딩박스 (지오코딩 실패 시 폴백 필터)
KR_LAT = (33.0, 38.9)
KR_LON = (124.5, 131.0)


def haversine(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    R = 6371.0
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = (
        math.sin(dlat / 2) ** 2
        + math.cos(math.radians(lat1))
        * math.cos(math.radians(lat2))
        * math.sin(dlon / 2) ** 2
    )
    return R * 2 * math.asin(math.sqrt(a))


def with_coords(listings: list[dict]) -> list[dict]:
    return [l for l in listings if abs(l.get("latitude", 0)) > 0.1]


def korea_filter(listings: list[dict]) -> list[dict]:
    return [
        l for l in with_coords(listings)
        if KR_LAT[0] <= l["latitude"] <= KR_LAT[1]
        and KR_LON[0] <= l["longitude"] <= KR_LON[1]
    ]


def radius_filter(
    listings: list[dict], clat: float, clon: float, radius_km: float
) -> list[dict]:
    return [
        l for l in with_coords(listings)
        if haversine(clat, clon, l["latitude"], l["longitude"]) <= radius_km
    ]


@app.route("/")
def index():
    today = date.today()
    return render_template(
        "index.html",
        default_checkin=(today + timedelta(days=7)).isoformat(),
        default_checkout=(today + timedelta(days=8)).isoformat(),
    )


@app.route("/api/search", methods=["POST"])
def api_search():
    body     = request.get_json(force=True) or {}
    query    = (body.get("query") or "").strip()
    checkin  = (body.get("checkin") or "").strip()
    checkout = (body.get("checkout") or "").strip()

    if not query:
        return jsonify({"error": "지역명을 입력해주세요."}), 400
    if not checkin or not checkout:
        return jsonify({"error": "체크인/체크아웃 날짜를 선택해주세요."}), 400
    if checkin >= checkout:
        return jsonify({"error": "체크아웃은 체크인보다 늦어야 합니다."}), 400

    # 1) Airbnb 수집 — 한국 결과가 부족하면 " 서울" 붙여 재시도
    def _kr_count(lst: list[dict]) -> int:
        return sum(1 for l in lst if KR_LAT[0] <= l.get("latitude", 0) <= KR_LAT[1]
                   and KR_LON[0] <= l.get("longitude", 0) <= KR_LON[1])

    # 시도할 쿼리 목록: 원본 → "원본+서울" (이미 대도시명 포함 시 하나만)
    q_candidates: list[str] = [query]
    if not any(c in query for c in ["서울", "부산", "대구", "인천", "제주", "광주", "대전"]):
        q_candidates.append(query + " 서울")

    all_listings: list[dict] = []
    last_error: str = ""
    for q_try in dict.fromkeys(q_candidates):
        try:
            raw = crawl_airbnb(q_try, checkin, checkout, max_results=60)
        except AirbnbError as exc:
            last_error = str(exc)
            continue
        except Exception as exc:
            last_error = f"수집 중 오류: {exc}"
            continue

        if _kr_count(raw) >= 3:
            all_listings = raw
            break
        elif not all_listings:
            all_listings = raw  # 부족하더라도 일단 저장, 다음 쿼리로 개선 시도

    if not all_listings and last_error:
        return jsonify({"error": last_error}), 500

    # 2) 필터링: 지오코딩 성공 → 반경 필터 / 실패 → 한국 범위 필터
    geo = geocode_region(query)
    meta: dict = {}
    listings: list[dict] = []

    if geo:
        clat, clon = geo["lat"], geo["lon"]
        meta = {"center_lat": clat, "center_lon": clon}
        for radius in [30, 60, 150]:
            listings = radius_filter(all_listings, clat, clon, radius)
            if len(listings) >= 3:
                break
    else:
        # 지오코딩 실패 → 한국 좌표 범위로 필터
        listings = korea_filter(all_listings)

    if not listings:
        return jsonify({
            "error": (
                "해당 지역에서 숙소를 찾지 못했습니다. "
                "더 넓은 지역명(예: '서울 마포구', '홍대 연남', '서울')을 시도해보세요."
            )
        }), 404

    # 3) 지도 중심
    if not meta:
        lats = [l["latitude"] for l in listings]
        lons = [l["longitude"] for l in listings]
        meta = {
            "center_lat": sum(lats) / len(lats),
            "center_lon": sum(lons) / len(lons),
        }

    notice = ""
    if len(listings) < 5:
        notice = (
            f"'{query}' 인근 Airbnb 매물이 {len(listings)}개로 적습니다. "
            "더 많은 결과를 원하면 '강남', '마포구', '서울' 같이 더 넓은 지역명을 검색해 보세요."
        )

    return jsonify({"listings": listings, "count": len(listings), "meta": meta, "notice": notice})


if __name__ == "__main__":
    app.run(debug=True, port=5001)
