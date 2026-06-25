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
    guest_favorite: str = ""
    search_badges: str = ""


def _walk(obj: Any, path: tuple[str, ...] = ()):
    yield path, obj
    if isinstance(obj, dict):
        for key, value in obj.items():
            yield from _walk(value, path + (str(key),))
    elif isinstance(obj, list):
        for idx, value in enumerate(obj):
            yield from _walk(value, path + (str(idx),))


def _clean_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, dict):
        for key in ("text", "title", "body", "name", "localizedString", "localizedStringWithTranslationPreference"):
            text = _clean_text(value.get(key))
            if text:
                return text
        return ""
    text = str(value)
    text = re.sub(r"<[^>]+>", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _extract_review_count_from_search(item: dict) -> int:
    best = 0
    for path, value in _walk(item):
        key = ".".join(path).lower()
        if isinstance(value, int) and 0 < value < 200_000:
            if any(token in key for token in ("reviewcount", "reviewscount", "review_count", "reviews_count")):
                return value
        text = _clean_text(value)
        if not text:
            continue
        for pattern in (
            r"(?:후기|리뷰)\s*([\d,]+)\s*개",
            r"([\d,]+)\s*(?:개의\s*)?(?:후기|리뷰)",
            r"([\d,]+)\s*reviews?",
        ):
            match = re.search(pattern, text, re.I)
            if match:
                try:
                    best = max(best, int(match.group(1).replace(",", "")))
                except ValueError:
                    pass
    return best


def _extract_search_badges(item: dict) -> tuple[str, str]:
    badges: list[str] = []
    for path, value in _walk(item):
        key = ".".join(path).lower()
        if not any(token in key for token in ("badge", "label", "title", "message", "subtitle")):
            continue
        text = _clean_text(value)
        if 2 <= len(text) <= 80 and text not in badges:
            if any(term in text.lower() for term in ("guest favorite", "superhost", "게스트", "슈퍼호스트", "인기")):
                badges.append(text)
    joined = " | ".join(badges[:8])
    guest_favorite = "Y" if any(term in joined.lower() for term in ("guest favorite", "게스트 선호")) else ""
    return guest_favorite, joined


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
    """Nominatim 단일 쿼리 (한국 우선 → 전체). 행정구역(place_rank≤25)만 수락."""
    for extra in [{"countrycodes": "kr"}, {}]:
        params: dict[str, Any] = {
            "q": q,
            "format": "json",
            "limit": "5",  # 여러 결과 중 행정구역 우선 선택
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
            # 행정구역/지역 랜드마크만 사용 — 식당/상점 POI 제외
            # place_rank: 1-25=행정구역, 26-30=장소/섬, 31+=POI
            # class: boundary/place/natural/railway/highway = 지역 관련
            #        amenity/shop/tourism 등 = 상업 POI (제외)
            for item in data:
                rank = int(item.get("place_rank") or 99)
                osm_class = item.get("class", "") or ""
                ok_class = osm_class in ("boundary", "place", "natural", "railway", "highway")
                if rank > 25 and not ok_class:
                    continue  # 식당, 상점 등 POI 스킵
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


# 한국 주요 지역 하드코딩 좌표 (Nominatim 실패 시 최종 폴백)
# 키: 쿼리에 이 문자열이 포함되면 해당 좌표 반환 (더 구체적인 키가 먼저 오도록 정렬)
_KR_REGION_FALLBACK: dict[str, dict] = {
    # 수도권 동네 (서울 키보다 먼저)
    "해운대": {"lat": 35.1631, "lon": 129.1631,
               "bb_minlat": 35.10, "bb_maxlat": 35.23, "bb_minlon": 129.10, "bb_maxlon": 129.25},
    "강남": {"lat": 37.5172, "lon": 127.0473,
             "bb_minlat": 37.48, "bb_maxlat": 37.55, "bb_minlon": 127.00, "bb_maxlon": 127.10},
    "홍대": {"lat": 37.5503, "lon": 126.9254,
             "bb_minlat": 37.52, "bb_maxlat": 37.58, "bb_minlon": 126.90, "bb_maxlon": 126.96},
    "마포": {"lat": 37.5638, "lon": 126.9084,
             "bb_minlat": 37.52, "bb_maxlat": 37.60, "bb_minlon": 126.87, "bb_maxlon": 126.95},
    "명동": {"lat": 37.5636, "lon": 126.9869,
             "bb_minlat": 37.54, "bb_maxlat": 37.58, "bb_minlon": 126.97, "bb_maxlon": 127.00},
    "이태원": {"lat": 37.5340, "lon": 126.9947,
               "bb_minlat": 37.51, "bb_maxlat": 37.55, "bb_minlon": 126.97, "bb_maxlon": 127.02},
    "대치": {"lat": 37.4967, "lon": 127.0613,
             "bb_minlat": 37.47, "bb_maxlat": 37.52, "bb_minlon": 127.04, "bb_maxlon": 127.08},
    "압구정": {"lat": 37.5272, "lon": 127.0286,
               "bb_minlat": 37.51, "bb_maxlat": 37.54, "bb_minlon": 127.01, "bb_maxlon": 127.05},
    "삼청": {"lat": 37.5825, "lon": 126.9810,
             "bb_minlat": 37.56, "bb_maxlat": 37.60, "bb_minlon": 126.96, "bb_maxlon": 127.00},
    # 광역시/도
    "제주": {"lat": 33.4996, "lon": 126.5312,
             "bb_minlat": 33.10, "bb_maxlat": 33.90, "bb_minlon": 126.10, "bb_maxlon": 127.00},
    "부산": {"lat": 35.1796, "lon": 129.0756,
             "bb_minlat": 34.85, "bb_maxlat": 35.45, "bb_minlon": 128.75, "bb_maxlon": 129.35},
    "대구": {"lat": 35.8714, "lon": 128.6014,
             "bb_minlat": 35.70, "bb_maxlat": 36.05, "bb_minlon": 128.45, "bb_maxlon": 128.80},
    "인천": {"lat": 37.4563, "lon": 126.7052,
             "bb_minlat": 37.30, "bb_maxlat": 37.65, "bb_minlon": 126.40, "bb_maxlon": 126.90},
    "광주": {"lat": 35.1595, "lon": 126.8526,
             "bb_minlat": 35.05, "bb_maxlat": 35.30, "bb_minlon": 126.75, "bb_maxlon": 127.00},
    "대전": {"lat": 36.3504, "lon": 127.3845,
             "bb_minlat": 36.20, "bb_maxlat": 36.50, "bb_minlon": 127.25, "bb_maxlon": 127.55},
    "경주": {"lat": 35.8562, "lon": 129.2247,
             "bb_minlat": 35.70, "bb_maxlat": 36.05, "bb_minlon": 129.10, "bb_maxlon": 129.40},
    "속초": {"lat": 38.2070, "lon": 128.5918,
             "bb_minlat": 38.15, "bb_maxlat": 38.27, "bb_minlon": 128.55, "bb_maxlon": 128.65},
    "강릉": {"lat": 37.7519, "lon": 128.8761,
             "bb_minlat": 37.65, "bb_maxlat": 37.85, "bb_minlon": 128.80, "bb_maxlon": 128.95},
    "서울": {"lat": 37.5665, "lon": 126.9780,
             "bb_minlat": 37.40, "bb_maxlat": 37.70, "bb_minlon": 126.80, "bb_maxlon": 127.20},
}


def geocode_region(query: str) -> dict | None:
    """
    Nominatim(OSM)으로 지역 중심 좌표 + 바운딩박스 반환.
    여러 쿼리 변형을 순서대로 시도, 모두 실패하면 하드코딩 폴백 사용.
    """
    candidates = _build_geo_candidates(query)
    for q in candidates:
        result = _geocode_one(q)
        if result:
            return result
    # Nominatim 완전 실패 → 주요 지역 하드코딩 폴백
    for key, geo in _KR_REGION_FALLBACK.items():
        if key in query:
            return geo
    return None


# 명시적 광역 지명 — 서울 폴백 없이 그 자체로 지오코딩
_EXPLICIT_REGION = re.compile(
    r"(제주|부산|대구|인천|광주|대전|울산|세종|전주|창원|청주|춘천|강릉|속초|경주|수원|성남|고양|용인|안산|화성|평택|의정부|포항|김해|전남|경남|경북|전북|충남|충북|강원|경기|해운대|해변|해수욕|강남|홍대|명동|이태원|대치|압구정|삼청)"
)


def _build_geo_candidates(query: str) -> list[str]:
    """지오코딩 시도 순서: 보정된 쿼리 → 원본 → (모호한 경우만) 서울 추가."""
    normalized = normalize_query(query)
    seen: list[str] = [normalized]
    if query not in seen:
        seen.append(query)
    # 명시적 지역/도시명이 없는 모호한 쿼리만 서울 컨텍스트 추가
    # (제주도, 부산 등은 서울 폴백 없이 직접 지오코딩)
    has_region = (
        "서울" in query
        or "Seoul" in query
        or bool(_EXPLICIT_REGION.search(query))
    )
    if not has_region:
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
    guest_favorite, search_badges = _extract_search_badges(item)

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
        review_count=_extract_review_count_from_search(item),
        latitude=lat,
        longitude=lon,
        guest_favorite=guest_favorite,
        search_badges=search_badges,
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
