from __future__ import annotations

import hashlib
import html as html_lib
import json
import os
import random
import re
import shutil
import sys
import time
import urllib.parse
import urllib.request
import webbrowser
from dataclasses import asdict, dataclass
from datetime import date
from pathlib import Path
from typing import Any, Iterable

try:
    from curl_cffi import requests as curl_requests
except ImportError:  # pragma: no cover - handled at runtime for packaged users
    curl_requests = None


BASE = "https://www.booking.com"
SEARCH_PATH = "/searchresults.ko.html"
PAGE_SIZE = 25


def _runtime_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


@dataclass
class BookingListing:
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


class BookingError(RuntimeError):
    pass


# Compatibility names used by the reused analysis modules.
AirbnbListing = BookingListing
AirbnbError = BookingError


_KR_REGION_FALLBACK: dict[str, dict[str, float]] = {
    "홍대": {"lat": 37.5503, "lon": 126.9254, "bb_minlat": 37.52, "bb_maxlat": 37.58, "bb_minlon": 126.90, "bb_maxlon": 126.96},
    "마포": {"lat": 37.5638, "lon": 126.9084, "bb_minlat": 37.52, "bb_maxlat": 37.60, "bb_minlon": 126.87, "bb_maxlon": 126.95},
    "강남": {"lat": 37.5172, "lon": 127.0473, "bb_minlat": 37.48, "bb_maxlat": 37.55, "bb_minlon": 127.00, "bb_maxlon": 127.10},
    "명동": {"lat": 37.5636, "lon": 126.9869, "bb_minlat": 37.54, "bb_maxlat": 37.58, "bb_minlon": 126.97, "bb_maxlon": 127.00},
    "충신동": {"lat": 37.5743, "lon": 127.0061, "bb_minlat": 37.56, "bb_maxlat": 37.59, "bb_minlon": 126.99, "bb_maxlon": 127.02},
    "황학동": {"lat": 37.5675, "lon": 127.0211, "bb_minlat": 37.55, "bb_maxlat": 37.58, "bb_minlon": 127.00, "bb_maxlon": 127.04},
    "이태원": {"lat": 37.5340, "lon": 126.9947, "bb_minlat": 37.51, "bb_maxlat": 37.55, "bb_minlon": 126.97, "bb_maxlon": 127.02},
    "성수": {"lat": 37.5446, "lon": 127.0557, "bb_minlat": 37.52, "bb_maxlat": 37.57, "bb_minlon": 127.03, "bb_maxlon": 127.08},
    "서울": {"lat": 37.5665, "lon": 126.9780, "bb_minlat": 37.40, "bb_maxlat": 37.70, "bb_minlon": 126.80, "bb_maxlon": 127.20},
    "해운대": {"lat": 35.1631, "lon": 129.1631, "bb_minlat": 35.10, "bb_maxlat": 35.23, "bb_minlon": 129.10, "bb_maxlon": 129.25},
    "부산": {"lat": 35.1796, "lon": 129.0756, "bb_minlat": 34.85, "bb_maxlat": 35.45, "bb_minlon": 128.75, "bb_maxlon": 129.35},
    "제주": {"lat": 33.4996, "lon": 126.5312, "bb_minlat": 33.10, "bb_maxlat": 33.90, "bb_minlon": 126.10, "bb_maxlon": 127.00},
    "서귀포": {"lat": 33.2539, "lon": 126.5590, "bb_minlat": 33.10, "bb_maxlat": 33.40, "bb_minlon": 126.35, "bb_maxlon": 126.75},
    "대구": {"lat": 35.8714, "lon": 128.6014, "bb_minlat": 35.70, "bb_maxlat": 36.05, "bb_minlon": 128.45, "bb_maxlon": 128.80},
    "인천": {"lat": 37.4563, "lon": 126.7052, "bb_minlat": 37.30, "bb_maxlat": 37.65, "bb_minlon": 126.40, "bb_maxlon": 126.90},
    "광주": {"lat": 35.1595, "lon": 126.8526, "bb_minlat": 35.05, "bb_maxlat": 35.30, "bb_minlon": 126.75, "bb_maxlon": 127.00},
    "대전": {"lat": 36.3504, "lon": 127.3845, "bb_minlat": 36.20, "bb_maxlat": 36.50, "bb_minlon": 127.25, "bb_maxlon": 127.55},
    "경주": {"lat": 35.8562, "lon": 129.2247, "bb_minlat": 35.70, "bb_maxlat": 36.05, "bb_minlon": 129.10, "bb_maxlon": 129.40},
    "속초": {"lat": 38.2070, "lon": 128.5918, "bb_minlat": 38.15, "bb_maxlat": 38.27, "bb_minlon": 128.55, "bb_maxlon": 128.65},
    "강릉": {"lat": 37.7519, "lon": 128.8761, "bb_minlat": 37.65, "bb_maxlat": 37.85, "bb_minlon": 128.80, "bb_maxlon": 128.95},
}


def normalize_query(query: str) -> str:
    return query.strip()


def _geocode_one(query: str) -> dict[str, float] | None:
    params: dict[str, Any] = {
        "q": query,
        "format": "json",
        "limit": "5",
        "countrycodes": "kr",
    }
    url = "https://nominatim.openstreetmap.org/search?" + urllib.parse.urlencode(params)
    req = urllib.request.Request(
        url,
        headers={
            "User-Agent": "BookingPriceScanner/1.0 personal-project",
            "Accept": "application/json",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=8) as response:
            data = json.loads(response.read().decode("utf-8"))
    except Exception:
        return None
    for item in data or []:
        rank = int(item.get("place_rank") or 99)
        osm_class = item.get("class", "") or ""
        if rank > 25 and osm_class not in ("boundary", "place", "natural", "railway", "highway"):
            continue
        bb = item.get("boundingbox") or []
        lat = float(item["lat"])
        lon = float(item["lon"])
        return {
            "lat": lat,
            "lon": lon,
            "bb_minlat": float(bb[0]) if len(bb) >= 4 else lat - 0.02,
            "bb_maxlat": float(bb[1]) if len(bb) >= 4 else lat + 0.02,
            "bb_minlon": float(bb[2]) if len(bb) >= 4 else lon - 0.02,
            "bb_maxlon": float(bb[3]) if len(bb) >= 4 else lon + 0.02,
        }
    return None


def geocode_region(query: str) -> dict[str, float] | None:
    normalized = normalize_query(query)
    candidates = [normalized]
    if "서울" not in normalized and not any(key in normalized for key in _KR_REGION_FALLBACK):
        candidates.append(f"{normalized} 서울")
    for candidate in candidates:
        geo = _geocode_one(candidate)
        if geo:
            return geo
    for key, geo in _KR_REGION_FALLBACK.items():
        if key in normalized:
            return geo.copy()
    return None


def _clean_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (int, float)):
        return str(value)
    if isinstance(value, dict):
        for key in ("text", "value", "label", "title", "name", "localizedString"):
            if key in value:
                text = _clean_text(value.get(key))
                if text:
                    return text
        return ""
    text = html_lib.unescape(str(value))
    text = re.sub(r"<[^>]+>", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _walk(obj: Any, path: tuple[str, ...] = ()) -> Iterable[tuple[tuple[str, ...], Any]]:
    yield path, obj
    if isinstance(obj, dict):
        for key, value in obj.items():
            yield from _walk(value, path + (str(key),))
    elif isinstance(obj, list):
        for idx, value in enumerate(obj):
            yield from _walk(value, path + (str(idx),))


def _looks_like_blocked(html: str, status_code: int) -> bool:
    lower = html.lower()
    return (
        status_code in (202, 403, 429)
        or "awswaf" in lower
        or "aws-waf" in lower
        or "window.aws" in lower
        or "px-captcha" in lower
        or "captcha" in lower
        or "are you a human" in lower
        or "automated requests" in lower
        or "robot" in lower and "booking" in lower
    )


def _looks_like_search_results(html: str) -> bool:
    lower = html.lower()
    return (
        'data-testid="property-card"' in lower
        or 'data-testid="property-card-container"' in lower
        or 'data-testid="title-link"' in lower
        or "/hotel/" in lower and ("price-and-discounted-price" in lower or "property-card" in lower)
    )


def _browser_fallback_enabled() -> bool:
    raw = os.environ.get("BOOKING_BROWSER_FALLBACK", "0").strip().lower()
    return raw not in ("0", "false", "no", "off")


def _find_chrome_executable() -> str | None:
    env_path = os.environ.get("BOOKING_CHROME_PATH")
    candidates = [
        env_path,
        shutil.which("chrome"),
        shutil.which("chrome.exe"),
        shutil.which("msedge"),
        shutil.which("msedge.exe"),
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    ]
    for candidate in candidates:
        if candidate and Path(candidate).exists():
            return str(candidate)
    return None


def _browser_profile_dir() -> Path:
    raw = os.environ.get("BOOKING_BROWSER_PROFILE")
    if raw:
        return Path(raw)
    return _runtime_dir() / "booking_browser_profile"


def _manual_pages_dir() -> Path:
    raw = os.environ.get("BOOKING_MANUAL_PAGES")
    if raw:
        return Path(raw)
    return _runtime_dir() / "booking_manual_pages"


def _safe_filename_part(text: str) -> str:
    cleaned = re.sub(r"[^\w가-힣.-]+", "_", text.strip(), flags=re.UNICODE).strip("._")
    return cleaned[:80] or "query"


def _manual_page_paths(query: str, checkin: str, checkout: str, page: int) -> dict[str, Path]:
    stem = f"{_safe_filename_part(query)}_{checkin}_{checkout}_p{page}"
    base = _manual_pages_dir()
    return {
        "html": base / f"{stem}.html",
        "htm": base / f"{stem}.htm",
        "url": base / f"{stem}.url",
        "txt": base / f"{stem}_README.txt",
    }


def _read_text_file(path: Path) -> str:
    data = path.read_bytes()
    for encoding in ("utf-8", "utf-8-sig", "cp949", "euc-kr", "latin-1"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace")


def _read_manual_html(query: str, checkin: str, checkout: str, page: int) -> str | None:
    paths = _manual_page_paths(query, checkin, checkout, page)
    candidates = [paths["html"], paths["htm"]]
    if not any(path.exists() for path in candidates):
        wildcard = f"*_{checkin}_{checkout}_p{page}.htm*"
        candidates.extend(sorted(_manual_pages_dir().glob(wildcard), key=lambda p: p.stat().st_mtime, reverse=True))
    seen_paths: set[Path] = set()
    for path in candidates:
        if path in seen_paths or not path.exists() or not path.is_file():
            continue
        seen_paths.add(path)
        html = _read_text_file(path)
        if _looks_like_blocked(html, 200) and not _looks_like_search_results(html):
            raise BookingError(
                "Booking.com manual capture required.\n"
                f"Saved manual page is still a Booking.com verification page: {path}\n"
                "Open the URL in a normal browser, complete verification, and wait until real property cards are visible.\n"
                f"Save the real search-results page as: {path}\n"
                "Then run the same analysis again."
            )
        print(f"[Booking.com] saved search HTML: {path}")
        return html
    return None


def _write_manual_request(query: str, checkin: str, checkout: str, page: int, url: str, reason: str) -> str:
    paths = _manual_page_paths(query, checkin, checkout, page)
    paths["html"].parent.mkdir(parents=True, exist_ok=True)
    paths["url"].write_text(f"[InternetShortcut]\nURL={url}\n", encoding="utf-8")
    instructions = (
        "Booking.com manual capture request\n"
        "==================================\n\n"
        f"Reason: {reason}\n"
        f"URL: {url}\n\n"
        "1. Open the .url shortcut in a normal Chrome/Edge browser, not an automated browser.\n"
        "2. If Booking.com asks for verification, complete it manually.\n"
        "3. Wait until real property search results are visible.\n"
        "4. Press Ctrl+S and save the complete web page as this exact file:\n"
        f"   {paths['html']}\n"
        "5. Run Booking.com Market Studio again with the same region/date settings.\n"
    )
    paths["txt"].write_text(instructions, encoding="utf-8")
    if os.environ.get("BOOKING_OPEN_MANUAL_URL", "1").strip().lower() not in ("0", "false", "no", "off"):
        try:
            os.startfile(str(paths["url"]))
        except Exception:
            webbrowser.open(url)
    return (
        "Booking.com blocked automated requests. A normal-browser capture request was created.\n"
        f"Open URL shortcut: {paths['url']}\n"
        f"Save the real search-results page as: {paths['html']}\n"
        "Then run the same analysis again."
    )


def prepare_manual_capture_request(
    query: str,
    checkin: str,
    checkout: str,
    page: int = 1,
    reason: str = "manual capture requested",
) -> dict[str, str]:
    url = BookingClient(delay_min=0, delay_max=0)._search_url(query, checkin, checkout, page)
    paths = _manual_page_paths(query, checkin, checkout, page)
    _write_manual_request(query, checkin, checkout, page, url, reason)
    return {
        "url": url,
        "shortcut": str(paths["url"]),
        "html": str(paths["html"]),
        "readme": str(paths["txt"]),
    }


def _parse_price_value(value: Any, key_hint: str = "") -> tuple[int, bool] | None:
    low_hint = key_hint.lower()
    total_hint = any(token in low_hint for token in ("total", "gross", "allinclusive", "stayprice", "all_inclusive"))
    price_hint = any(token in low_hint for token in ("price", "amount", "rate", "charge", "gross", "value"))
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        amount = int(round(float(value)))
        if price_hint and 10_000 <= amount <= 5_000_000:
            return amount, total_hint
        return None

    text = _clean_text(value)
    if not text:
        return None
    lower = text.lower()
    has_currency = "\u20a9" in text or "원" in text or "krw" in lower or "won" in lower
    if not has_currency and not price_hint:
        return None
    numbers = re.findall(r"\d[\d,]{2,}", text)
    if not numbers:
        return None
    parsed = [int(num.replace(",", "")) for num in numbers]
    parsed = [num for num in parsed if 10_000 <= num <= 5_000_000]
    if not parsed:
        return None
    return min(parsed), total_hint or any(token in lower for token in ("총", "total", "entire stay"))


def _find_price(obj: Any, nights: int) -> int | None:
    candidates: list[tuple[int, int]] = []
    for path, value in _walk(obj):
        if isinstance(value, (dict, list)):
            continue
        key_hint = ".".join(path)
        parsed = _parse_price_value(value, key_hint)
        if not parsed:
            continue
        amount, is_total = parsed
        low_hint = key_hint.lower()
        priority = 2
        if any(token in low_hint for token in ("pernight", "nightly", "unitprice", "night_price")):
            priority = 0
        elif is_total:
            priority = 1
        if is_total and nights > 1:
            amount = max(1, round(amount / nights))
        candidates.append((priority, amount))
    if not candidates:
        return None
    candidates.sort(key=lambda item: (item[0], item[1]))
    return candidates[0][1]


def _valid_title(text: str) -> bool:
    if not text or len(text) < 2 or len(text) > 160:
        return False
    lowered = text.lower()
    bad_tokens = (
        "booking.com",
        "개인정보",
        "쿠키",
        "로그인",
        "필터",
        "정렬",
        "결제",
        "검색 결과",
        "지도에서 보기",
        "숙소를 검색",
    )
    if any(token in lowered for token in bad_tokens):
        return False
    if re.fullmatch(r"[\d\s.,\u20a9원%-]+", text):
        return False
    return True


def _extract_title(node: dict[str, Any]) -> str:
    direct_keys = (
        "hotel_name",
        "hotelName",
        "propertyName",
        "accommodationName",
        "displayName",
        "name",
        "title",
        "cardTitle",
    )
    for key in direct_keys:
        text = _clean_text(node.get(key))
        if _valid_title(text):
            return text
    for nested_key in ("property", "hotel", "accommodation", "basicPropertyData", "summary", "location"):
        nested = node.get(nested_key)
        if isinstance(nested, dict):
            text = _extract_title(nested)
            if text:
                return text
    return ""


def _extract_url(node: dict[str, Any]) -> str:
    for _, value in _walk(node):
        if isinstance(value, (dict, list)):
            continue
        text = _clean_text(value)
        if not text:
            continue
        if "booking.com/hotel/" in text or text.startswith("/hotel/") or "/hotel/" in text:
            if text.startswith("/"):
                return urllib.parse.urljoin(BASE, text)
            if text.startswith("http"):
                return text
    return ""


def _extract_id(node: dict[str, Any], title: str, url: str) -> str:
    id_keys = ("hotel_id", "hotelId", "accommodationId", "propertyId", "ufi", "id")
    for key in id_keys:
        raw = node.get(key)
        if raw is None:
            continue
        text = _clean_text(raw)
        if text and len(text) <= 80 and not text.startswith("urn:"):
            return text
    if url:
        m = re.search(r"/hotel/[^/]+/([^/?#]+)", url)
        if m:
            return urllib.parse.unquote(m.group(1))
    digest = hashlib.sha1(f"{title}|{url}".encode("utf-8", "ignore")).hexdigest()
    return digest[:16]


def _extract_rating(node: dict[str, Any]) -> float:
    candidates: list[float] = []
    for path, value in _walk(node):
        if isinstance(value, (dict, list)):
            continue
        key_hint = ".".join(path).lower()
        text = _clean_text(value)
        if not any(token in key_hint for token in ("rating", "reviewscore", "score")):
            if not re.search(r"(평점|후기 평점|scored|score|review)", text, re.I):
                continue
        match = re.search(r"\d+(?:\.\d+)?", text)
        if not match:
            continue
        score = float(match.group())
        if 0 < score <= 10:
            if score > 5:
                score = score / 2
            candidates.append(score)
    return round(max(candidates), 2) if candidates else 0.0


def _extract_review_count(node: dict[str, Any]) -> int:
    count_candidates: list[int] = []
    for path, value in _walk(node):
        if isinstance(value, (dict, list)):
            continue
        key_hint = ".".join(path).lower()
        text = _clean_text(value)
        key_is_review = any(token in key_hint for token in ("review", "reviews"))
        if key_is_review and any(token in key_hint for token in ("count", "total", "amount", "number")):
            try:
                count = int(float(str(value).replace(",", "")))
                if count > 0:
                    count_candidates.append(count)
                    continue
            except Exception:
                pass
        if not key_is_review and not re.search(r"(후기|리뷰|reviews?)", text, re.I):
            continue
        numbers = re.findall(r"\d[\d,]*", text)
        if numbers:
            count_candidates.append(max(int(num.replace(",", "")) for num in numbers))
    return max(count_candidates) if count_candidates else 0


def _extract_coord(node: dict[str, Any], names: tuple[str, ...]) -> float:
    for path, value in _walk(node):
        key = path[-1].lower() if path else ""
        if key not in names:
            continue
        try:
            return float(value)
        except Exception:
            continue
    return 0.0


def _extract_room_info(node: dict[str, Any]) -> tuple[int, int, float]:
    text = " ".join(_clean_text(value) for _, value in _walk(node) if not isinstance(value, (dict, list)))
    bedrooms = beds = 0
    bathrooms = 0.0
    patterns = [
        (r"(?:침실|bedroom)s?\s*(\d+)", "bedrooms"),
        (r"(?:침대|bed)s?\s*(\d+)", "beds"),
        (r"(?:욕실|bathroom)s?\s*([\d.]+)", "bathrooms"),
    ]
    for pattern, field in patterns:
        m = re.search(pattern, text, re.I)
        if not m:
            continue
        if field == "bedrooms":
            bedrooms = int(float(m.group(1)))
        elif field == "beds":
            beds = int(float(m.group(1)))
        else:
            bathrooms = float(m.group(1))
    return bedrooms, beds, bathrooms


def _property_type(title: str, text: str = "") -> tuple[str, str]:
    lower = f"{title} {text}".lower()
    if any(token in lower for token in ("hostel", "호스텔")):
        return "호텔 객실", "호스텔"
    if any(token in lower for token in ("guesthouse", "guest house", "게스트하우스", "게스트 하우스")):
        return "호텔 객실", "게스트하우스"
    if any(token in lower for token in ("apartment", "아파트", "residence", "레지던스", "studio")):
        return "숙소 전체", "아파트"
    if any(token in lower for token in ("villa", "빌라", "펜션", "pension")):
        return "숙소 전체", "펜션/빌라"
    if any(token in lower for token in ("motel", "모텔")):
        return "호텔 객실", "모텔"
    if any(token in lower for token in ("resort", "리조트")):
        return "호텔 객실", "리조트"
    return "호텔 객실", "호텔"


def _make_listing(node: dict[str, Any], query: str, nights: int) -> BookingListing | None:
    title = _extract_title(node)
    if not title:
        return None
    price = _find_price(node, nights)
    if not price:
        return None
    url = _extract_url(node)
    listing_id = _extract_id(node, title, url)
    bedrooms, beds, bathrooms = _extract_room_info(node)
    room_type, prop_type = _property_type(title, _clean_text(node))
    return BookingListing(
        listing_id=listing_id,
        title=title,
        url=url or BASE,
        region_query=query,
        room_type=room_type,
        property_type=prop_type,
        price_per_night=price,
        bedrooms=bedrooms,
        beds=beds,
        bathrooms=bathrooms,
        rating=_extract_rating(node),
        review_count=_extract_review_count(node),
        latitude=_extract_coord(node, ("latitude", "lat")),
        longitude=_extract_coord(node, ("longitude", "lng", "lon")),
    )


def _json_payloads_from_html(html: str) -> list[Any]:
    payloads: list[Any] = []
    patterns = [
        r'<script[^>]+id="__NEXT_DATA__"[^>]*>(.*?)</script>',
        r'<script[^>]+type="application/json"[^>]*>(.*?)</script>',
        r'<script[^>]+type="application/ld\+json"[^>]*>(.*?)</script>',
    ]
    for pattern in patterns:
        for match in re.finditer(pattern, html, re.DOTALL | re.I):
            raw = html_lib.unescape(match.group(1).strip())
            if not raw:
                continue
            try:
                payloads.append(json.loads(raw))
            except Exception:
                continue
    return payloads


def _extract_from_json(data: Any, query: str, nights: int) -> list[BookingListing]:
    listings: list[BookingListing] = []
    seen: set[str] = set()
    for _, node in _walk(data):
        if not isinstance(node, dict):
            continue
        listing = _make_listing(node, query, nights)
        if not listing or listing.listing_id in seen:
            continue
        seen.add(listing.listing_id)
        listings.append(listing)
    listings.sort(key=lambda item: item.price_per_night)
    return listings


def _html_to_text(fragment: str) -> str:
    text = re.sub(r"<script.*?</script>", " ", fragment, flags=re.DOTALL | re.I)
    text = re.sub(r"<style.*?</style>", " ", text, flags=re.DOTALL | re.I)
    text = re.sub(r"<[^>]+>", "\n", text)
    text = html_lib.unescape(text)
    lines = [re.sub(r"\s+", " ", line).strip() for line in text.splitlines()]
    return "\n".join(line for line in lines if line)


def _extract_card_url(chunk: str) -> str:
    patterns = [
        r'<a[^>]+data-testid=["\']title-link["\'][^>]+href=["\']([^"\']+)["\']',
        r'<a[^>]+href=["\']([^"\']*/hotel/[^"\']+)["\']',
    ]
    for pattern in patterns:
        match = re.search(pattern, chunk, re.I | re.S)
        if match:
            return urllib.parse.urljoin(BASE, html_lib.unescape(match.group(1)))
    return ""


def _extract_card_title(chunk: str, text: str) -> str:
    patterns = [
        r'data-testid=["\']title["\'][^>]*>(.*?)</',
        r"<h3[^>]*>(.*?)</h3>",
        r"<h2[^>]*>(.*?)</h2>",
    ]
    for pattern in patterns:
        for match in re.finditer(pattern, chunk, re.I | re.S):
            title = _html_to_text(match.group(1)).replace("\n", " ").strip()
            if _valid_title(title):
                return title
    for line in text.splitlines():
        if _valid_title(line) and len(line) <= 90 and not re.search(r"(후기|리뷰|평점|요금|예약)", line):
            return line
    return ""


def _extract_card_price(text: str, nights: int) -> int | None:
    amounts: list[int] = []
    price_re = re.compile(r"(?:\u20a9|KRW)\s*[\d,]+|[\d,]+\s*(?:원|KRW)", re.I)
    for match in price_re.finditer(text):
        parsed = _parse_price_value(match.group(0))
        if not parsed:
            continue
        amount, is_total = parsed
        if is_total and nights > 1:
            amount = max(1, round(amount / nights))
        amounts.append(amount)
    if not amounts:
        return None
    return min(amounts)


def _extract_card_rating(text: str) -> tuple[float, int]:
    rating = 0.0
    review_count = 0
    for line in text.splitlines():
        stripped = line.strip()
        if re.fullmatch(r"\d+(?:\.\d+)?", stripped):
            score = float(stripped)
            if 0 < score <= 10:
                rating = score / 2 if score > 5 else score
                break
    rating_patterns = [
        r"(?:scored|score|평점)\s*(\d+(?:\.\d+)?)",
        r"(\d+(?:\.\d+)?)\s*/\s*10",
        r"(\d+(?:\.\d+)?)\s*점",
    ]
    if not rating:
        for pattern in rating_patterns:
            match = re.search(pattern, text, re.I)
            if match:
                rating = float(match.group(1))
                if rating > 5:
                    rating = rating / 2
                break
    review_patterns = [
        r"(?:후기|리뷰)\s*(\d[\d,]*)\s*개",
        r"(\d[\d,]*)\s*개\s*(?:후기|리뷰)",
        r"(\d[\d,]*)\s*reviews?",
        r"(?:후기|리뷰|reviews?)\s*[:(]?\s*(\d[\d,]*)",
    ]
    for pattern in review_patterns:
        match = re.search(pattern, text, re.I)
        if match:
            review_count = int(match.group(1).replace(",", ""))
            break
    return round(rating, 2), review_count


def _extract_coords_from_text(text: str) -> tuple[float, float]:
    patterns = [
        r"data-atlas-latlng=[\"'](-?\d+(?:\.\d+)?),\s*(-?\d+(?:\.\d+)?)[\"']",
        r"latitude=(-?\d+(?:\.\d+)?).*?longitude=(-?\d+(?:\.\d+)?)",
        r"lat=(-?\d+(?:\.\d+)?).*?(?:lng|lon)=(-?\d+(?:\.\d+)?)",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.I | re.S)
        if match:
            try:
                return float(match.group(1)), float(match.group(2))
            except ValueError:
                return 0.0, 0.0
    return 0.0, 0.0


def _extract_from_soup_cards(html: str, query: str, nights: int) -> list[BookingListing]:
    try:
        from bs4 import BeautifulSoup
    except Exception:
        return []

    soup = BeautifulSoup(html, "html.parser")
    cards = soup.select('[data-testid="property-card"]')
    if not cards:
        cards = soup.select('[data-testid="property-card-container"]')

    listings: list[BookingListing] = []
    seen: set[str] = set()
    for card in cards:
        text = card.get_text("\n", strip=True)
        if not text:
            continue

        title_el = card.select_one('[data-testid="title"]') or card.select_one("h3") or card.select_one("h2")
        title = _clean_text(title_el.get_text(" ", strip=True) if title_el else "")
        if not _valid_title(title):
            title = _extract_card_title(str(card), text)
        if not title:
            continue

        link_el = card.select_one('a[data-testid="title-link"]') or card.select_one('a[href*="/hotel/"]')
        raw_href = link_el.get("href") if link_el else ""
        url = urllib.parse.urljoin(BASE, html_lib.unescape(raw_href)) if raw_href else ""

        price_el = card.select_one('[data-testid="price-and-discounted-price"]')
        price_text = price_el.get_text(" ", strip=True) if price_el else text
        price = _extract_card_price(price_text, nights) or _extract_card_price(text, nights)
        if not price:
            continue

        rating_text = ""
        rating_el = card.select_one('[data-testid="review-score"]') or card.select_one('[aria-label*="Scored"]')
        if rating_el:
            rating_text = rating_el.get_text(" ", strip=True) or str(rating_el.get("aria-label") or "")
        rating, review_count = _extract_card_rating("\n".join([rating_text, text]))

        raw_card = str(card)
        lat, lon = _extract_coords_from_text(raw_card + " " + url)
        bedrooms, beds, bathrooms = _extract_room_info({"text": text})
        room_type, prop_type = _property_type(title, text)
        listing_id = _extract_id({}, title, url)
        if listing_id in seen:
            continue
        seen.add(listing_id)
        listings.append(BookingListing(
            listing_id=listing_id,
            title=title,
            url=url or BASE,
            region_query=query,
            room_type=room_type,
            property_type=prop_type,
            price_per_night=price,
            bedrooms=bedrooms,
            beds=beds,
            bathrooms=bathrooms,
            rating=rating,
            review_count=review_count,
            latitude=lat,
            longitude=lon,
        ))
    listings.sort(key=lambda item: item.price_per_night)
    return listings


def _booking_card_chunks(html: str) -> list[str]:
    starts: list[int] = []
    for match in re.finditer(r'data-testid=["\']property-card["\']', html, re.I):
        starts.append(max(0, html.rfind("<", 0, match.start())))
    if not starts:
        for match in re.finditer(r'/hotel/', html, re.I):
            starts.append(max(0, match.start() - 5000))
    chunks: list[str] = []
    for idx, start in enumerate(starts):
        end = starts[idx + 1] if idx + 1 < len(starts) else min(len(html), start + 22_000)
        chunk = html[start:end]
        if re.search(r"(?:\u20a9|KRW)\s*[\d,]+|[\d,]+\s*(?:원|KRW)", chunk, re.I):
            chunks.append(chunk)
    return chunks


def _extract_from_html_cards(html: str, query: str, nights: int) -> list[BookingListing]:
    soup_listings = _extract_from_soup_cards(html, query, nights)
    if soup_listings:
        return soup_listings

    listings: list[BookingListing] = []
    seen: set[str] = set()
    for chunk in _booking_card_chunks(html):
        text = _html_to_text(chunk)
        title = _extract_card_title(chunk, text)
        if not title:
            continue
        price = _extract_card_price(text, nights)
        if not price:
            continue
        url = _extract_card_url(chunk)
        listing_id = _extract_id({}, title, url)
        if listing_id in seen:
            continue
        seen.add(listing_id)
        rating, review_count = _extract_card_rating(text)
        lat, lon = _extract_coords_from_text(chunk + " " + url)
        bedrooms, beds, bathrooms = _extract_room_info({"text": text})
        room_type, prop_type = _property_type(title, text)
        listings.append(BookingListing(
            listing_id=listing_id,
            title=title,
            url=url or BASE,
            region_query=query,
            room_type=room_type,
            property_type=prop_type,
            price_per_night=price,
            bedrooms=bedrooms,
            beds=beds,
            bathrooms=bathrooms,
            rating=rating,
            review_count=review_count,
            latitude=lat,
            longitude=lon,
        ))
    listings.sort(key=lambda item: item.price_per_night)
    return listings


class BookingClient:
    def __init__(self, delay_min: float = 1.4, delay_max: float = 2.8) -> None:
        if curl_requests is None:
            raise BookingError("curl_cffi is required. Install with: pip install curl_cffi")
        self.session = curl_requests.Session(impersonate="chrome124")
        self.delay_min = delay_min
        self.delay_max = delay_max
        self.session.headers.update({
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8",
            "Referer": BASE + "/",
        })

    def _pause(self) -> None:
        time.sleep(random.uniform(self.delay_min, self.delay_max))

    def _search_url(self, query: str, checkin: str, checkout: str, page: int) -> str:
        params: dict[str, Any] = {
            "ss": query,
            "checkin": checkin,
            "checkout": checkout,
            "group_adults": "1",
            "no_rooms": "1",
            "group_children": "0",
            "selected_currency": "KRW",
            "lang": "ko",
            "order": "popularity",
        }
        offset = max(0, page - 1) * PAGE_SIZE
        if offset:
            params["offset"] = str(offset)
        return BASE + SEARCH_PATH + "?" + urllib.parse.urlencode(params)

    def _fetch_html(
        self,
        url: str,
        query: str | None = None,
        checkin: str | None = None,
        checkout: str | None = None,
        page: int = 1,
    ) -> str:
        self._pause()
        try:
            response = self.session.get(url, timeout=30)
        except Exception as exc:
            raise BookingError(f"Network error while requesting Booking.com: {exc}") from exc
        if _looks_like_blocked(response.text, response.status_code):
            if _browser_fallback_enabled():
                return self._fetch_html_with_browser(url, response.status_code)
            if query and checkin and checkout:
                raise BookingError(_write_manual_request(
                    query,
                    checkin,
                    checkout,
                    page,
                    url,
                    f"HTTP {response.status_code} verification page",
                ))
            raise BookingError(
                "Booking.com returned a verification page "
                f"(HTTP {response.status_code}). Save a manual search-results HTML page or set "
                "BOOKING_BROWSER_FALLBACK=1 to try the browser fallback."
            )
        if response.status_code >= 400:
            raise BookingError(f"Booking.com HTTP {response.status_code}")
        return response.text

    def _fetch_html_with_browser(self, url: str, status_code: int | None = None) -> str:
        chrome_path = _find_chrome_executable()
        if not chrome_path:
            raise BookingError(
                "Booking.com returned a verification page"
                + (f" (HTTP {status_code})" if status_code else "")
                + ", and Chrome/Edge was not found for browser fallback."
            )
        try:
            from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
            from playwright.sync_api import sync_playwright
        except Exception as exc:
            raise BookingError(
                "Booking.com returned a verification page"
                + (f" (HTTP {status_code})" if status_code else "")
                + ", and Playwright is not installed for browser fallback. "
                "Install it with: pip install playwright"
            ) from exc

        wait_seconds = int(os.environ.get("BOOKING_BROWSER_WAIT", "180"))
        profile_dir = _browser_profile_dir()
        profile_dir.mkdir(parents=True, exist_ok=True)

        print(
            "\n[Booking.com] HTTP request hit verification. "
            "Complete the Chrome window if needed; the scanner continues when results are visible."
        )
        print(f"[Booking.com] browser profile: {profile_dir}")

        try:
            with sync_playwright() as p:
                context = p.chromium.launch_persistent_context(
                    user_data_dir=str(profile_dir),
                    executable_path=chrome_path,
                    headless=False,
                    locale="ko-KR",
                    viewport={"width": 1365, "height": 900},
                    args=["--start-maximized"],
                )
                page = context.pages[0] if context.pages else context.new_page()
                page.set_default_timeout(15_000)
                try:
                    page.goto(url, wait_until="domcontentloaded", timeout=60_000)
                except PlaywrightTimeoutError:
                    pass

                last_html = ""
                told_waiting = False
                deadline = time.time() + wait_seconds
                while time.time() < deadline:
                    try:
                        html = page.content()
                    except Exception:
                        html = ""
                    last_html = html or last_html
                    if html and not _looks_like_blocked(html, 200):
                        try:
                            page.wait_for_load_state("networkidle", timeout=10_000)
                        except PlaywrightTimeoutError:
                            pass
                        for _ in range(4):
                            try:
                                page.mouse.wheel(0, 900)
                                page.wait_for_timeout(700)
                            except Exception:
                                break
                        html = page.content()
                        context.close()
                        return html
                    if not told_waiting:
                        print(f"[Booking.com] waiting up to {wait_seconds}s for a real search-results page.")
                        told_waiting = True
                    page.wait_for_timeout(2_000)

                context.close()
        except BookingError:
            raise
        except Exception as exc:
            raise BookingError(f"Browser fallback failed: {exc}") from exc

        raise BookingError(
            "Booking.com browser fallback timed out before a search result page was available. "
            "Run again and complete the Chrome verification window."
        )

    def _parse_html(self, html: str, query: str, nights: int) -> list[BookingListing]:
        listings: list[BookingListing] = []
        seen: set[str] = set()
        for listing in _extract_from_html_cards(html, query, nights):
            if listing.listing_id not in seen:
                seen.add(listing.listing_id)
                listings.append(listing)
        if len(listings) < 3:
            for payload in _json_payloads_from_html(html):
                for listing in _extract_from_json(payload, query, nights):
                    if listing.listing_id not in seen:
                        seen.add(listing.listing_id)
                        listings.append(listing)
        if len(listings) < 3:
            for listing in _extract_from_html_cards(html, query, nights):
                if listing.listing_id not in seen:
                    seen.add(listing.listing_id)
                    listings.append(listing)
        return sorted(listings, key=lambda item: item.price_per_night)

    def crawl(
        self,
        query: str,
        checkin: str,
        checkout: str,
        geo: dict | None = None,
        max_results: int = 60,
        max_pages: int = 4,
    ) -> list[BookingListing]:
        del geo  # Booking.com search is destination-string driven.
        nights = max(1, (date.fromisoformat(checkout) - date.fromisoformat(checkin)).days)
        results: list[BookingListing] = []
        seen: set[str] = set()
        for page in range(1, max_pages + 1):
            last_error: Exception | None = None
            page_listings: list[BookingListing] = []
            manual_html = _read_manual_html(query, checkin, checkout, page)
            if manual_html:
                page_listings = self._parse_html(manual_html, query, nights)
                if not page_listings:
                    raise BookingError(
                        "Saved manual Booking.com page was found, but no property cards/prices were parsed. "
                        "Make sure you saved the actual search-results page after all results loaded."
                    )
            if not page_listings:
                url = self._search_url(query, checkin, checkout, page)
                try:
                    html = self._fetch_html(url, query=query, checkin=checkin, checkout=checkout, page=page)
                    page_listings = self._parse_html(html, query, nights)
                except BookingError as exc:
                    last_error = exc
            if not page_listings:
                if page == 1 and last_error:
                    raise last_error
                break
            for listing in page_listings:
                if listing.listing_id in seen:
                    continue
                seen.add(listing.listing_id)
                results.append(listing)
                if len(results) >= max_results:
                    return results
        return results


# Compatibility class name used by gui_app.py monkey-patching.
AirbnbClient = BookingClient


def crawl_booking(
    query: str,
    checkin: str,
    checkout: str,
    geo: dict | None = None,
    max_results: int = 60,
    max_pages: int = 4,
) -> list[dict]:
    client = BookingClient()
    listings = client.crawl(
        query,
        checkin,
        checkout,
        geo=geo,
        max_results=max_results,
        max_pages=max_pages,
    )
    return [asdict(listing) for listing in listings]


def crawl_airbnb(
    query: str,
    checkin: str,
    checkout: str,
    geo: dict | None = None,
    max_results: int = 60,
    max_pages: int = 4,
) -> list[dict]:
    return crawl_booking(
        query,
        checkin,
        checkout,
        geo=geo,
        max_results=max_results,
        max_pages=max_pages,
    )


def fetch_detail(url: str, session: Any) -> dict:
    fields = {
        "max_guests": "",
        "bed_types": "",
        "description": "",
        "amenities": "",
        "host_name": "Booking.com",
        "superhost": "",
        "house_rules": "",
        "cancellation": "",
    }
    detail_enabled = os.environ.get("BOOKING_DETAIL_FETCH", "0").strip().lower()
    if detail_enabled in ("0", "false", "no", "off"):
        fields["description"] = "Booking.com detail fetch skipped; using search result card data."
        return fields
    try:
        response = session.get(url, timeout=25, impersonate="chrome124")
        if _looks_like_blocked(response.text, response.status_code):
            fields["description"] = "Booking.com verification page"
            return fields
        if response.status_code >= 400:
            fields["description"] = f"HTTP {response.status_code}"
            return fields
        html = response.text
        meta = re.search(r'<meta[^>]+name=["\']description["\'][^>]+content=["\']([^"\']+)["\']', html, re.I)
        if meta:
            fields["description"] = _clean_text(meta.group(1))[:900]
        payloads = _json_payloads_from_html(html)
        texts: list[str] = []
        amenities: list[str] = []
        policies: list[str] = []
        for payload in payloads:
            for path, value in _walk(payload):
                if isinstance(value, (dict, list)):
                    continue
                key_hint = ".".join(path).lower()
                text = _clean_text(value)
                if not text:
                    continue
                if "description" in key_hint and len(text) > 40:
                    texts.append(text)
                if "amenit" in key_hint and 1 < len(text) < 80:
                    amenities.append(text)
                if any(token in key_hint for token in ("policy", "policies", "cancellation", "checkin", "checkout")) and len(text) < 160:
                    policies.append(text)
        if texts and not fields["description"]:
            fields["description"] = max(texts, key=len)[:900]
        if amenities:
            fields["amenities"] = " | ".join(dict.fromkeys(amenities))[:1200]
        if policies:
            policy_text = " | ".join(dict.fromkeys(policies))
            fields["house_rules"] = policy_text[:600]
            if "취소" in policy_text or "cancel" in policy_text.lower():
                fields["cancellation"] = policy_text[:200]
        return fields
    except Exception as exc:
        fields["description"] = f"detail error: {exc}"
        return fields
