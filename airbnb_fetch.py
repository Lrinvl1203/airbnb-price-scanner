from __future__ import annotations

import base64
import json
import random
import re
import time
import urllib.parse
import urllib.request
from dataclasses import asdict, dataclass
from typing import Any

try:
    from curl_cffi import requests as curl_requests
except ImportError:
    curl_requests = None

BASE = "https://www.airbnb.co.kr"


@dataclass
class AirbnbListing:
    listing_id: str
    title: str
    url: str
    region_query: str
    room_type: str
    property_type: str
    price_per_night: int
    currency: str = "KRW"
    bedrooms: int = 0
    beds: int = 0
    bathrooms: float = 0.0
    rating: float = 0.0
    review_count: int = 0
    latitude: float = 0.0
    longitude: float = 0.0


class AirbnbError(RuntimeError):
    pass


# ────────────────────────────────────────────
#  쿼리 자동 보정
# ────────────────────────────────────────────

# 한국 행정구역 접미사 (단독 입력 시 "서울" 컨텍스트 추가)
_KR_SUFFIX = re.compile(r"[가-힣]+(동|구|읍|면|리|로|가|길|역|동네|마을)$")
# 제주/부산/대구 등 대도시는 그 자체로 충분
_KR_CITY = re.compile(r"(서울|부산|대구|인천|광주|대전|울산|세종|수원|성남|제주|전주|창원|청주|춘천|강릉)")


def normalize_query(query: str) -> str:
    """
    Airbnb 검색용 쿼리는 원본을 그대로 반환.
    (Airbnb는 한국어 지역명을 자체 인식하므로 수정하지 않음)
    """
    return query.strip()


def _geocode_one(q: str) -> dict | None:
    """Nominatim 단일 쿼리 (한국 우선 → 전체)."""
    for extra in [{"countrycodes": "kr"}, {}]:
        params: dict[str, Any] = {
            "q": q,
            "format": "json",
            "limit": "1",
            **extra,
        }
        url = "https://nominatim.openstreetmap.org/search?" + urllib.parse.urlencode(params)
        req = urllib.request.Request(
            url,
            headers={
                "User-Agent": "AirbnbPriceScanner/1.0 personal-project",
                "Accept": "application/json",
            },
        )
        try:
            with urllib.request.urlopen(req, timeout=8) as r:
                data = json.loads(r.read().decode("utf-8"))
            if not data:
                continue
            item = data[0]
            bb = item.get("boundingbox") or []
            return {
                "lat": float(item["lat"]),
                "lon": float(item["lon"]),
                "bb_minlat": float(bb[0]) if len(bb) >= 4 else float(item["lat"]) - 0.02,
                "bb_maxlat": float(bb[1]) if len(bb) >= 4 else float(item["lat"]) + 0.02,
                "bb_minlon": float(bb[2]) if len(bb) >= 4 else float(item["lon"]) - 0.02,
                "bb_maxlon": float(bb[3]) if len(bb) >= 4 else float(item["lon"]) + 0.02,
            }
        except Exception:
            pass
    return None


def geocode_region(query: str) -> dict | None:
    """
    Nominatim(OSM)으로 지역 중심 좌표 + 바운딩박스 반환.
    여러 쿼리 변형을 순서대로 시도해서 첫 번째 성공 결과를 반환.
    """
    candidates = _build_geo_candidates(query)
    for q in candidates:
        result = _geocode_one(q)
        if result:
            return result
    return None


def _build_geo_candidates(query: str) -> list[str]:
    """지오코딩 시도 순서: 보정된 쿼리 → 원본 → 서울 추가."""
    normalized = normalize_query(query)
    seen: list[str] = [normalized]
    if query not in seen:
        seen.append(query)
    # "서울" 붙인 버전
    if "서울" not in query and "Seoul" not in query:
        seen.append(query + " 서울")
    return seen


def expand_bbox(geo: dict, delta: float = 0.08) -> dict:
    """바운딩박스를 delta 도만큼 확장 (작은 동네 검색 시 결과 보강)."""
    return {
        **geo,
        "bb_minlat": geo["bb_minlat"] - delta,
        "bb_maxlat": geo["bb_maxlat"] + delta,
        "bb_minlon": geo["bb_minlon"] - delta,
        "bb_maxlon": geo["bb_maxlon"] + delta,
    }


# ────────────────────────────────────────────
#  파싱 유틸
# ────────────────────────────────────────────

def _decode_id(raw: str) -> str:
    try:
        decoded = base64.b64decode(raw + "==").decode("utf-8")
        return decoded.split(":", 1)[-1]
    except Exception:
        return raw


def _parse_price_str(text: Any) -> int | None:
    if text is None:
        return None
    s = re.sub(r"[₩원,\s]", "", str(text))
    m = re.search(r"\d+", s)
    return int(m.group()) if m else None


def _night_price(item: dict) -> int | None:
    sdp     = item.get("structuredDisplayPrice") or {}
    primary = sdp.get("primaryLine") or {}
    qualifier = primary.get("qualifier", "")

    if qualifier in ("박", "1박"):
        return _parse_price_str(primary.get("price"))

    # "N박 x ₩XXX" 패턴
    exp = sdp.get("explanationData") or {}
    for group in exp.get("priceDetails") or []:
        for it in group.get("items") or []:
            desc = it.get("description", "")
            m1 = re.search(r"(\d+)박\s*[x×]\s*₩([\d,]+)", desc)
            m2 = re.search(r"₩([\d,]+)\s*[x×]\s*(\d+)박", desc)
            if m1:
                return int(m1.group(2).replace(",", ""))
            if m2:
                return int(m2.group(1).replace(",", ""))

    # fallback: 총액 ÷ 박 수
    total = _parse_price_str(primary.get("price"))
    if total:
        params = item.get("listingParamOverrides") or {}
        checkin, checkout = params.get("checkin", ""), params.get("checkout", "")
        if checkin and checkout:
            try:
                from datetime import date
                nights = (date.fromisoformat(checkout) - date.fromisoformat(checkin)).days
                if nights > 0:
                    return round(total / nights)
            except Exception:
                pass
        return total
    return None


def _room_info(item: dict) -> dict[str, Any]:
    info: dict[str, Any] = {"bedrooms": 0, "beds": 0, "bathrooms": 0.0}
    for msg in (item.get("structuredContent") or {}).get("primaryLine") or []:
        body  = msg.get("body", "")
        mtype = msg.get("type", "")
        if mtype == "BEDINFO" or "침실" in body:
            m = re.search(r"침실\s*(\d+)", body)
            if m:
                info["bedrooms"] = int(m.group(1))
        if ("침대" in body or mtype == "BEDINFO") and "침실" not in body:
            m = re.search(r"(\d+)", body)
            if m:
                info["beds"] = int(m.group(1))
        if mtype == "BATHROOMINFO" or "욕실" in body:
            m = re.search(r"욕실\s*([\d.]+)", body)
            if m:
                info["bathrooms"] = float(m.group(1))
    return info


def _infer_type(title_text: str) -> tuple[str, str]:
    t = title_text
    if "아파트" in t or "apartment" in t.lower():
        return "집 전체", "아파트"
    if "게스트하우스" in t or "guesthouse" in t.lower():
        return "개인실", "게스트하우스"
    if "호텔" in t or "hotel" in t.lower() or "모텔" in t:
        return "호텔 객실", "호텔"
    if "스튜디오" in t or "studio" in t.lower():
        return "집 전체", "스튜디오"
    if "별장" in t or "빌라" in t or "villa" in t.lower():
        return "집 전체", "빌라"
    if "독채" in t or "주택" in t or "house" in t.lower():
        return "집 전체", "주택"
    if "한옥" in t:
        return "집 전체", "한옥"
    if "펜션" in t:
        return "집 전체", "펜션"
    m = re.search(r"의\s*(\S+)$", t.split("(")[0].strip())
    prop = m.group(1) if m else "기타"
    return "집 전체", prop


def _normalize(item: dict, region_query: str) -> AirbnbListing | None:
    demand = item.get("demandStayListing") or {}
    lid    = _decode_id(demand.get("id", ""))
    if not lid:
        return None

    name_obj = item.get("nameLocalized") or {}
    title    = (
        name_obj.get("localizedStringWithTranslationPreference")
        or item.get("subtitle", "")
        or item.get("title", "")
    )

    loc   = demand.get("location") or {}
    coord = loc.get("coordinate") or {}
    lat   = float(coord.get("latitude") or 0)
    lon   = float(coord.get("longitude") or 0)

    price = _night_price(item)
    if not price:
        return None

    info = _room_info(item)

    rating_str = item.get("avgRatingLocalized") or ""
    rating = 0.0
    if rating_str and rating_str not in ("신규", "New"):
        m2 = re.search(r"[\d.]+", rating_str)
        if m2:
            rating = float(m2.group())

    room_type, property_type = _infer_type(item.get("title", ""))

    return AirbnbListing(
        listing_id=lid,
        title=title,
        url=f"{BASE}/rooms/{lid}",
        region_query=region_query,
        room_type=room_type,
        property_type=property_type,
        price_per_night=price,
        bedrooms=info["bedrooms"],
        beds=info["beds"],
        bathrooms=info["bathrooms"],
        rating=rating,
        review_count=0,
        latitude=lat,
        longitude=lon,
    )


def _extract(data: dict) -> tuple[list[dict], list[str]]:
    try:
        results = (
            data["niobeClientData"][0][1]
            ["data"]["presentation"]["staysSearch"]["results"]
        )
        sr      = results.get("searchResults") or []
        cursors = results.get("paginationInfo", {}).get("pageCursors") or []
    except (KeyError, IndexError, TypeError):
        sr, cursors = [], []
    return sr, cursors


# ────────────────────────────────────────────
#  HTTP 클라이언트
# ────────────────────────────────────────────

class AirbnbClient:
    def __init__(self, delay_min: float = 1.3, delay_max: float = 2.5) -> None:
        if curl_requests is None:
            raise AirbnbError("curl_cffi 를 설치하세요: pip install curl_cffi")
        self.session = curl_requests.Session(impersonate="chrome124")
        self.delay_min = delay_min
        self.delay_max = delay_max
        self.session.headers.update({
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            "Accept": "text/html,application/xhtml+xml,*/*;q=0.9",
            "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8",
            "Referer": BASE + "/",
        })

    def _pause(self) -> None:
        time.sleep(random.uniform(self.delay_min, self.delay_max))

    def _fetch_raw(self, url: str, params: dict) -> dict:
        self._pause()
        try:
            r = self.session.get(url, params=params, timeout=25)
        except Exception as exc:
            raise AirbnbError(f"네트워크 오류: {exc}") from exc

        if r.status_code >= 400:
            raise AirbnbError(f"HTTP {r.status_code} — 접근이 차단되었거나 일시적 오류입니다.")

        for pat in [
            r'<script[^>]+id="data-deferred-state-0"[^>]*>(.*?)</script>',
            r'<script id="__NEXT_DATA__"[^>]*>(.*?)</script>',
        ]:
            m = re.search(pat, r.text, re.DOTALL)
            if m:
                try:
                    return json.loads(m.group(1))
                except json.JSONDecodeError as exc:
                    raise AirbnbError(f"JSON 파싱 실패: {exc}") from exc

        raise AirbnbError(
            "Airbnb 페이지 데이터를 파싱하지 못했습니다. "
            "CAPTCHA 또는 페이지 구조 변경일 수 있습니다."
        )

    def _fetch_bounds(
        self,
        geo: dict,
        checkin: str,
        checkout: str,
        cursor: str | None = None,
    ) -> tuple[list[dict], list[str]]:
        """바운딩박스 기반 지도 검색."""
        bb_span = max(
            abs(geo["bb_maxlat"] - geo["bb_minlat"]),
            abs(geo["bb_maxlon"] - geo["bb_minlon"]),
        )
        if bb_span < 0.05:
            zoom = 14
        elif bb_span < 0.2:
            zoom = 13
        elif bb_span < 0.5:
            zoom = 12
        elif bb_span < 1.0:
            zoom = 11
        else:
            zoom = 10
        params: dict[str, Any] = {
            "ne_lat":  geo["bb_maxlat"],
            "ne_lng":  geo["bb_maxlon"],
            "sw_lat":  geo["bb_minlat"],
            "sw_lng":  geo["bb_minlon"],
            "zoom":    zoom,
            "search_type": "user_map_move",
            "tab_id":  "home_tab",
            "checkin": checkin,
            "checkout": checkout,
            "adults":  "1",
            "price_filter_input_type": "0",
            "channel": "EXPLORE",
        }
        if cursor:
            params["cursor"] = cursor
        data = self._fetch_raw(f"{BASE}/s/homes", params)
        return _extract(data)

    def _fetch_query(
        self,
        query: str,
        checkin: str,
        checkout: str,
        cursor: str | None = None,
        geo: dict | None = None,
    ) -> tuple[list[dict], list[str]]:
        """키워드 기반 검색. geo가 있으면 bbox 파라미터도 함께 전송해 지역 필터링."""
        params: dict[str, Any] = {
            "tab_id": "home_tab",
            "query":  query,
            "checkin": checkin,
            "checkout": checkout,
            "adults": "1",
            "source": "structured_search_input_header",
            "search_type": "autocomplete_click",
            "locale": "ko",
            "currency": "KRW",
        }
        if geo:
            params["ne_lat"]  = geo["bb_maxlat"]
            params["ne_lng"]  = geo["bb_maxlon"]
            params["sw_lat"]  = geo["bb_minlat"]
            params["sw_lng"]  = geo["bb_minlon"]
            bb_span = max(abs(geo["bb_maxlat"] - geo["bb_minlat"]),
                         abs(geo["bb_maxlon"] - geo["bb_minlon"]))
            params["zoom"] = (14 if bb_span < 0.05 else
                              13 if bb_span < 0.2 else
                              12 if bb_span < 0.5 else
                              11 if bb_span < 1.0 else 10)
        if cursor:
            params["cursor"] = cursor
        data = self._fetch_raw(f"{BASE}/s/{urllib.parse.quote(query)}/homes", params)
        return _extract(data)

    def crawl(
        self,
        query: str,
        checkin: str,
        checkout: str,
        geo: dict | None = None,
        max_results: int = 60,
        max_pages: int = 4,
    ) -> list[AirbnbListing]:
        results: list[AirbnbListing] = []
        seen: set[str] = set()

        def collect(sr: list[dict]) -> bool:
            for item in sr:
                lst = _normalize(item, query)
                if lst and lst.listing_id not in seen:
                    seen.add(lst.listing_id)
                    results.append(lst)
                if len(results) >= max_results:
                    return True  # done
            return False

        # keyword+bbox 병합 검색: geo가 있으면 bbox 파라미터를 keyword 요청에 포함
        # (bbox-only 검색은 Vercel IP에서 Airbnb homepage로 redirect됨)
        sr, all_cursors = self._fetch_query(query, checkin, checkout, geo=geo)
        if collect(sr):
            return results
        for cursor in all_cursors[1:max_pages]:
            sr, _ = self._fetch_query(query, checkin, checkout, cursor=cursor, geo=geo)
            if collect(sr) or not sr:
                break

        return results


# ────────────────────────────────────────────
#  공개 API
# ────────────────────────────────────────────

def crawl_airbnb(
    query: str,
    checkin: str,
    checkout: str,
    geo: dict | None = None,
    max_results: int = 60,
) -> list[dict]:
    client = AirbnbClient()
    listings = client.crawl(query, checkin, checkout, geo=geo, max_results=max_results)
    return [asdict(lst) for lst in listings]
