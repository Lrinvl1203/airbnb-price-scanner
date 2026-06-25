from __future__ import annotations

import io
import math
import os
import tempfile
import urllib.parse
import urllib.request
from datetime import date, timedelta
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file

from booking_fetch import BookingError, crawl_airbnb, geocode_region

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


@app.route("/api/subway")
def api_subway():
    """Overpass API 프록시 — 지하철/철도역 조회 (CORS 우회)."""
    try:
        s = float(request.args["s"])
        w = float(request.args["w"])
        n = float(request.args["n"])
        e = float(request.args["e"])
    except (KeyError, ValueError):
        return jsonify({"error": "bbox params s/w/n/e required"}), 400

    query = (
        f'[out:json][timeout:10];'
        f'node["railway"="station"]({s:.5f},{w:.5f},{n:.5f},{e:.5f});'
        f'out body;'
    )
    url = "https://overpass-api.de/api/interpreter?data=" + urllib.parse.quote(query)
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Booking.comPriceScanner/1.0"})
        with urllib.request.urlopen(req, timeout=12) as r:
            body = r.read().decode("utf-8")
        return app.response_class(body, mimetype="application/json")
    except Exception as exc:
        return jsonify({"error": str(exc)}), 502


@app.route("/")
def index():
    today = date.today()
    return render_template(
        "index.html",
        default_checkin=(today + timedelta(days=7)).isoformat(),
        default_checkout=(today + timedelta(days=8)).isoformat(),
    )


@app.route("/legacy")
def legacy():
    today = date.today()
    return render_template(
        "index_leaflet.html",
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
    geo = geocode_region(query)
    clat_geo = geo["lat"] if geo else None
    clon_geo = geo["lon"] if geo else None

    # 2) Booking.com 수집 — geo가 있으면 좌표 bbox 검색 (IP 무관), 없으면 키워드 폴백
    def _near_geo(lst: list[dict]) -> int:
        """geo 중심 100km 이내 매물 수 (일본 등 원격지 필터)."""
        if clat_geo is None:
            return sum(1 for l in lst if KR_LAT[0] <= l.get("latitude", 0) <= KR_LAT[1]
                       and KR_LON[0] <= l.get("longitude", 0) <= KR_LON[1])
        return sum(1 for l in with_coords(lst)
                   if haversine(clat_geo, clon_geo, l["latitude"], l["longitude"]) <= 100)

    def _make_bbox(clat: float, clon: float, radius_km: float) -> dict:
        """중심점 + 반경(km)에서 Booking.com bbox dict 생성."""
        dlat = radius_km / 111.0
        dlon = radius_km / (111.0 * math.cos(math.radians(clat)))
        return {
            "lat": clat, "lon": clon,
            "bb_minlat": clat - dlat, "bb_maxlat": clat + dlat,
            "bb_minlon": clon - dlon, "bb_maxlon": clon + dlon,
        }

    all_listings: list[dict] = []
    last_error: str = ""

    if geo:
        # 지역 bbox 크기로 행정 수준 판별 → 검색/필터 반경 결정
        bb_span = max(
            geo["bb_maxlat"] - geo["bb_minlat"],
            geo["bb_maxlon"] - geo["bb_minlon"],
        )
        if bb_span < 0.05:          # 동(洞) 수준 ≈ 1~3km
            search_radii = [3, 5, 8]
            filter_radii = [2, 3, 5]
        elif bb_span < 0.20:        # 구(區)/읍(邑) 수준 ≈ 5~15km
            search_radii = [8, 15, 25]
            filter_radii = [8, 12, 20]
        elif bb_span < 0.60:        # 시(市) 수준 ≈ 20~50km
            search_radii = [20, 40, 70]
            filter_radii = [20, 40, 70]
        else:                        # 도(道)/광역시 수준
            search_radii = [40, 80, 150]
            filter_radii = [40, 80, 150]

        for bk in search_radii:
            expanded_geo = _make_bbox(clat_geo, clon_geo, bk)
            try:
                raw = crawl_airbnb(query, checkin, checkout, geo=expanded_geo, max_results=120)
            except BookingError as exc:
                last_error = str(exc)
                continue
            except Exception as exc:
                last_error = f"수집 중 오류: {exc}"
                continue
            near_cnt = _near_geo(raw)
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
            except BookingError as exc:
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
        used_radius = filter_radii[-1]
        for radius in filter_radii:
            listings = radius_filter(all_listings, clat, clon, radius)
            if len(listings) >= 3:
                used_radius = radius
                break
        meta = {"center_lat": clat, "center_lon": clon, "radius_km": used_radius}
    else:
        # 지오코딩 실패 → 한국 좌표 범위로 필터
        listings = korea_filter(all_listings)

    if not listings:
        if geo:
            return jsonify({
                "error": (
                    "해당 지역에서 숙소를 찾지 못했습니다. "
                    "더 넓은 지역명(예: '서울 마포구', '홍대 연남', '서울')을 시도해보세요."
                )
            }), 404
        else:
            return jsonify({
                "error": (
                    f"'{query}' 지역을 찾을 수 없습니다. "
                    "한국의 도시명이나 지역명을 정확히 입력해 주세요. (예: 홍대, 강남, 제주시)"
                )
            }), 404

    # geocode 실패한 keyword 검색의 경우 결과가 너무 적으면 오류 처리
    if not geo and len(listings) < 5:
        return jsonify({
            "error": (
                f"'{query}' 지역을 찾을 수 없습니다. "
                "한국의 도시명이나 지역명을 정확히 입력해 주세요. (예: 홍대, 강남, 제주시)"
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
            f"'{query}' 인근 Booking.com 매물이 {len(listings)}개로 적습니다. "
            "더 많은 결과를 원하면 '강남', '마포구', '서울' 같이 더 넓은 지역명을 검색해 보세요."
        )

    return jsonify({"listings": listings, "count": len(listings), "meta": meta, "notice": notice})


@app.route("/api/export", methods=["POST"])
def api_export():
    from export_excel import build_excel

    body     = request.get_json(force=True) or {}
    listings = body.get("listings", [])
    query    = (body.get("query") or "booking").strip()
    checkin  = (body.get("checkin") or "").strip()
    checkout = (body.get("checkout") or "").strip()

    if not listings:
        return jsonify({"error": "다운로드할 데이터가 없습니다."}), 400
    if not checkin or not checkout:
        return jsonify({"error": "날짜 정보가 필요합니다."}), 400

    tmp = Path(tempfile.mktemp(suffix=".xlsx"))
    try:
        build_excel(listings, query, checkin, checkout, tmp)
        data = tmp.read_bytes()
    finally:
        if tmp.exists():
            tmp.unlink(missing_ok=True)

    safe_q   = "".join(c if c.isalnum() or c in "-_" else "_" for c in query)
    filename = f"booking_{safe_q}_{checkin}_{checkout}.xlsx"

    return send_file(
        io.BytesIO(data),
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    port = int(os.getenv("FLASK_PORT", "5001"))
    app.run(debug=os.getenv("FLASK_DEBUG") == "1", port=port)
