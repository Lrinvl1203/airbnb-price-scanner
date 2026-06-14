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


@app.route("/api/diag")
def diag():
    """Vercel 환경 진단 — 스크래퍼 동작 확인용 임시 엔드포인트."""
    import sys, time as _time
    result: dict = {"python": sys.version, "steps": []}

    # 1) curl_cffi import
    try:
        from curl_cffi import requests as _cr
        result["steps"].append("curl_cffi: OK")
    except Exception as e:
        result["steps"].append(f"curl_cffi IMPORT FAIL: {e}")
        return jsonify(result)

    # 2) Nominatim geocoding — raw 응답 디버깅
    try:
        import urllib.request as _ur
        import urllib.parse as _up
        for gq in ["홍대", "제주도", "부산 해운대"]:
            raw_url = f"https://nominatim.openstreetmap.org/search?q={_up.quote(gq)}&format=json&limit=3&countrycodes=kr"
            _req = _ur.Request(raw_url, headers={"User-Agent": "AirbnbPriceScanner/1.0", "Accept": "application/json"})
            try:
                with _ur.urlopen(_req, timeout=8) as _rr:
                    _data = __import__("json").loads(_rr.read().decode())
                items_info = [(d.get("lat","?"), d.get("lon","?"), d.get("class","?"), d.get("place_rank","?"), d.get("display_name","")[:30]) for d in _data[:2]]
                result["steps"].append(f"nominatim '{gq}': {items_info}")
            except Exception as _e:
                result["steps"].append(f"nominatim '{gq}' FAIL: {_e}")
        geo = geocode_region("홍대")
    except Exception as e:
        result["steps"].append(f"geocode FAIL: {e}")

    # 3) keyword+bbox 테스트 — 서울(홍대), 제주도, 부산
    try:
        from airbnb_fetch import AirbnbClient, _normalize
        import math as _math
        client = AirbnbClient(delay_min=0, delay_max=0)

        def _make_test_bbox(clat: float, clon: float, km: float) -> dict:
            dlat = km / 111.0
            dlon = km / (111.0 * _math.cos(_math.radians(clat)))
            return {"lat": clat, "lon": clon,
                    "bb_minlat": clat - dlat, "bb_maxlat": clat + dlat,
                    "bb_minlon": clon - dlon, "bb_maxlon": clon + dlon}

        def _test(q: str, clat: float, clon: float, km: float) -> str:
            g = _make_test_bbox(clat, clon, km)
            sr, _ = client._fetch_query(q, "2026-06-21", "2026-06-22", geo=g)
            near = [_normalize(x, q) for x in sr]
            close = sum(1 for x in near if x and haversine(clat, clon, x.latitude, x.longitude) <= km * 2)
            first = next((x for x in near if x), None)
            first_loc = f"lat={first.latitude:.2f},lon={first.longitude:.2f}" if first else "none"
            return f"{len(sr)} items, {close} near, first={first_loc}"

        for (q, clat, clon, km) in [
            ("홍대",   37.5503, 126.9254, 10),
            ("제주도",  33.4996, 126.5312, 25),
            ("부산 해운대", 35.1631, 129.1631, 10),
        ]:
            t0 = _time.time()
            info = _test(q, clat, clon, km)
            result["steps"].append(f"'{q}' +bbox{km}km ({_time.time()-t0:.1f}s): {info}")
    except Exception as e:
        result["steps"].append(f"airbnb FAIL: {e}")

    return jsonify(result)


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
    try:
        d_checkin  = date.fromisoformat(checkin)
        d_checkout = date.fromisoformat(checkout)
    except ValueError:
        return jsonify({"error": "날짜 형식이 올바르지 않습니다 (YYYY-MM-DD)."}), 400
    if d_checkin >= d_checkout:
        return jsonify({"error": "체크아웃은 체크인보다 늦어야 합니다."}), 400
    today = date.today()
    if d_checkin < today:
        return jsonify({"error": "체크인은 오늘 이후여야 합니다."}), 400
    if (d_checkout - today).days > 730:
        return jsonify({"error": "2년 이내 날짜만 검색할 수 있습니다."}), 400

    # 1) 지오코딩 먼저 — 반경 스케일 결정
    _CODE_VER = "v5-20260614"  # 코드 버전 확인 (핫 Lambda 진단용)
    from airbnb_fetch import _build_geo_candidates, _geocode_one, _KR_REGION_FALLBACK
    _cands = _build_geo_candidates(query)
    _geo_steps: list[str] = [f"cands={_cands}"]
    for _cq in _cands:
        _g = _geocode_one(_cq)
        _geo_steps.append(f"geocode({_cq!r})→{_g}")
        if _g:
            break
    geo = geocode_region(query)
    clat_geo = geo["lat"] if geo else None
    clon_geo = geo["lon"] if geo else None

    # 2) Airbnb 수집 — geo가 있으면 좌표 bbox 검색 (IP 무관), 없으면 키워드 폴백
    def _near_geo(lst: list[dict]) -> int:
        """geo 중심 100km 이내 매물 수 (일본 등 원격지 필터)."""
        if clat_geo is None:
            return sum(1 for l in lst if KR_LAT[0] <= l.get("latitude", 0) <= KR_LAT[1]
                       and KR_LON[0] <= l.get("longitude", 0) <= KR_LON[1])
        return sum(1 for l in with_coords(lst)
                   if haversine(clat_geo, clon_geo, l["latitude"], l["longitude"]) <= 100)

    def _make_bbox(clat: float, clon: float, radius_km: float) -> dict:
        """중심점 + 반경(km)에서 Airbnb bbox dict 생성."""
        dlat = radius_km / 111.0
        dlon = radius_km / (111.0 * math.cos(math.radians(clat)))
        return {
            "lat": clat, "lon": clon,
            "bb_minlat": clat - dlat, "bb_maxlat": clat + dlat,
            "bb_minlon": clon - dlon, "bb_maxlon": clon + dlon,
        }

    all_listings: list[dict] = []
    last_error: str = ""
    _geo_label = ('Jeju' if geo and geo['lat'] < 34
                  else 'Seoul' if geo and geo['lat'] > 37
                  else 'Other' if geo else 'None')
    _dbg: list[str] = [
        f"code={_CODE_VER}",
        f"geo={_geo_label} lat={clat_geo} lon={clon_geo}",
    ] + _geo_steps

    if geo:
        # bbox 검색: IP와 무관하게 좌표로 검색 → 일본 결과 없음
        bbox_radii = [10, 25, 50]
        for bk in bbox_radii:
            expanded_geo = _make_bbox(clat_geo, clon_geo, bk)
            try:
                raw = crawl_airbnb(query, checkin, checkout, geo=expanded_geo, max_results=120)
            except AirbnbError as exc:
                last_error = str(exc)
                _dbg.append(f"bk={bk} AirbnbError: {exc}")
                continue
            except Exception as exc:
                last_error = f"수집 중 오류: {exc}"
                _dbg.append(f"bk={bk} Exception: {exc}")
                continue
            near_cnt = _near_geo(raw)
            first_lat = raw[0]["latitude"] if raw else None
            _dbg.append(f"bk={bk}: raw={len(raw)}, near={near_cnt}, first_lat={first_lat}")
            if near_cnt >= 3:
                all_listings = raw
                break
            elif not all_listings:
                all_listings = raw
    else:
        # geo 없음 → 키워드 검색 (폴백)
        q_candidates: list[str] = [query]
        if not any(c in query for c in ["서울", "부산", "대구", "인천", "제주", "광주", "대전"]):
            q_candidates.append(query + " 서울")
        for q_try in dict.fromkeys(q_candidates):
            try:
                raw = crawl_airbnb(q_try, checkin, checkout, max_results=120)
            except AirbnbError as exc:
                last_error = str(exc)
                continue
            except Exception as exc:
                last_error = f"수집 중 오류: {exc}"
                continue
            if _near_geo(raw) >= 3:
                all_listings = raw
                break
            elif not all_listings:
                all_listings = raw

    if not all_listings and last_error:
        return jsonify({"error": last_error}), 500

    # 3) 필터링: 지오코딩 성공 → 반경 필터 / 실패 → 한국 범위 필터
    meta: dict = {}
    listings: list[dict] = []

    if geo:
        clat, clon = geo["lat"], geo["lon"]
        # bbox 크기로 검색 지역 스케일 추정 → 반경 단계 결정
        bb_span = max(
            geo["bb_maxlat"] - geo["bb_minlat"],
            geo["bb_maxlon"] - geo["bb_minlon"],
        )
        if bb_span < 0.05:          # 동(洞) 수준 (~5km)
            radii = [5, 10, 25]
        elif bb_span < 0.15:        # 구(區) 수준 (~15km)
            radii = [15, 30, 60]
        else:                        # 시/도 수준
            radii = [30, 60, 150]

        used_radius = radii[-1]
        for radius in radii:
            listings = radius_filter(all_listings, clat, clon, radius)
            if len(listings) >= 3:
                used_radius = radius
                break
        meta = {"center_lat": clat, "center_lon": clon, "radius_km": used_radius}
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

    return jsonify({"listings": listings, "count": len(listings), "meta": meta, "notice": notice, "_dbg": _dbg})


if __name__ == "__main__":
    import os
    port = int(os.getenv("FLASK_PORT", "5001"))
    app.run(debug=os.getenv("FLASK_DEBUG") == "1", port=port)
