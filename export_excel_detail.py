"""
Airbnb 시세 상세 분석 → Excel (B 모드)
각 숙소 URL을 직접 방문해 소개글·편의시설·호스트 정보·하우스 룰을 추가 열로 저장.

사용: python export_excel_detail.py <지역> <체크인YYYY-MM-DD> <체크아웃YYYY-MM-DD>
"""
from __future__ import annotations

import json
import re
import sys
import time
from datetime import date, datetime
from pathlib import Path

import xlsxwriter
from curl_cffi import requests as cf_requests

_BASE_DIR = (
    Path(sys.executable).parent if getattr(sys, "frozen", False)
    else Path(__file__).parent
)
sys.path.insert(0, str(_BASE_DIR))
from airbnb_fetch import crawl_airbnb, geocode_region
from export_excel import _fix_colors, C_DARK_BLUE, C_MID_BLUE, C_LIGHT_BLUE, C_LINK, C_WHITE, COLS


# ── 상세 열 정의 ──────────────────────────────────────────────
DETAIL_COLS = [
    # (헤더,         키,                너비)
    ("최대인원",    "max_guests",       8),
    ("침대종류",    "bed_types",       22),
    ("소개글",      "description",     45),
    ("편의시설",    "amenities",       40),
    ("호스트명",    "host_name",       14),
    ("슈퍼호스트",  "superhost",        9),
    ("하우스룰",    "house_rules",     32),
    ("취소정책",    "cancellation",    25),
]


def _get_deferred_data(html: str) -> dict | None:
    """data-deferred-state-0 스크립트에서 JSON 파싱."""
    m = re.search(
        r'id="data-deferred-state-0"[^>]*>(.*?)</script>',
        html, re.DOTALL
    )
    if not m:
        return None
    try:
        return json.loads(m.group(1).strip())
    except Exception:
        return None


def fetch_detail(url: str, session: cf_requests.Session) -> dict:
    """숙소 상세 페이지 → 추가 정보 dict 반환."""
    empty: dict = {col[1]: "" for col in DETAIL_COLS}

    try:
        r = session.get(url, timeout=20, impersonate="chrome120")
        if r.status_code != 200:
            empty["description"] = f"HTTP {r.status_code}"
            return empty

        html = r.text
        raw = _get_deferred_data(html)
        if not raw:
            empty["description"] = "JSON 없음"
            return empty

        # ── JSON 경로 탐색 ────────────────────────────────────
        niobe_val = raw.get("niobeClientData", [[None, {}]])[0][1]
        data_root = niobe_val.get("data", {})

        # presentation → sections
        pdp_sections = (
            data_root
            .get("presentation", {})
            .get("stayProductDetailPage", {})
            .get("sections", {})
            .get("sections", [])
        )
        sec_map = {s["sectionId"]: s.get("section", {}) for s in pdp_sections if "sectionId" in s}

        # node → pdpPresentation (편의시설·최대인원)
        pdp_pres = (
            data_root
            .get("node", {})
            .get("pdpPresentation", {})
        )

        # ── 소개글 ──────────────────────────────────────────
        html_desc = (
            sec_map.get("DESCRIPTION_DEFAULT", {})
            .get("htmlDescription", {})
        )
        if isinstance(html_desc, dict):
            raw_text = html_desc.get("htmlText", "")
        else:
            # longDescriptionHtml fallback
            raw_text = (
                pdp_pres
                .get("descriptions", {})
                .get("longDescriptionHtml", {})
                .get("localizedStringWithTranslationPreference", "")
            )
        description = re.sub(r"<[^>]+>", "\n", raw_text).strip()
        description = re.sub(r"\n{3,}", "\n\n", description)

        # ── 최대 인원 ────────────────────────────────────────
        max_guests = str(pdp_pres.get("personCapacity") or "")

        # ── 침대 종류 ────────────────────────────────────────
        arrangements = (
            sec_map.get("SLEEPING_ARRANGEMENT_WITH_IMAGES", {})
            .get("arrangementDetails", [])
        )
        bed_parts = []
        for arr in (arrangements or []):
            subtitle = arr.get("subtitle", "")
            title = arr.get("title", "")
            if subtitle:
                bed_parts.append(f"{title}: {subtitle}" if title else subtitle)
        bed_types = " | ".join(bed_parts)

        # ── 편의시설 ─────────────────────────────────────────
        amen_groups = (
            pdp_pres
            .get("amenities", {})
            .get("seeAllAmenitiesGroups", [])
        )
        amen_parts = []
        for g in (amen_groups or []):
            cat = g.get("title", "")
            names = [
                a.get("title", "")
                for a in g.get("amenities", [])
                if a.get("available") and a.get("title")
            ]
            if names:
                amen_parts.append(f"[{cat}] {', '.join(names)}" if cat else ", ".join(names))
        amenities = " | ".join(amen_parts)

        # ── 호스트 정보 ──────────────────────────────────────
        card = sec_map.get("MEET_YOUR_HOST", {}).get("cardData", {}) or {}
        host_name = card.get("name", "")
        superhost = "슈퍼호스트" if card.get("isSuperhost") else ("일반" if host_name else "")

        # ── 하우스 룰 ────────────────────────────────────────
        rules_list = sec_map.get("POLICIES_DEFAULT", {}).get("houseRules", []) or []
        house_rules = " | ".join(r.get("title", "") for r in rules_list if r.get("title"))

        # ── 취소 정책 ────────────────────────────────────────
        # BOOK_IT_SIDEBAR.cancellationPolicies → 첫 번째 항목 title
        cancel_policies = (
            sec_map.get("BOOK_IT_SIDEBAR", {}).get("cancellationPolicies") or
            sec_map.get("BOOK_IT_FLOATING_FOOTER", {}).get("cancellationPolicies") or []
        )
        if cancel_policies and isinstance(cancel_policies, list):
            first = cancel_policies[0] if cancel_policies else {}
            cancellation = first.get("title") or first.get("label") or ""
        else:
            # houseRulesSections 내 취소 정책 탐색
            hr_sections = sec_map.get("POLICIES_DEFAULT", {}).get("houseRulesSections", []) or []
            cancellation = ""
            for section in hr_sections:
                if "취소" in (section.get("title") or ""):
                    items = section.get("items", [])
                    if items:
                        cancellation = " / ".join(
                            i.get("title", "") for i in items if i.get("title")
                        )
                    break

        return {
            "max_guests":   max_guests,
            "bed_types":    bed_types[:250],
            "description":  description[:900],
            "amenities":    amenities[:1200],
            "host_name":    host_name[:50],
            "superhost":    superhost,
            "house_rules":  house_rules[:600],
            "cancellation": cancellation[:200],
        }

    except Exception as e:
        empty["description"] = f"오류: {e}"
        return empty


def build_excel_detail(
    listings: list[dict],
    query: str,
    checkin: str,
    checkout: str,
    out_path: Path,
) -> None:
    nights = (date.fromisoformat(checkout) - date.fromisoformat(checkin)).days
    all_cols = COLS + DETAIL_COLS

    wb = xlsxwriter.Workbook(str(out_path))
    ws = wb.add_worksheet("숙소 목록 (상세)")

    # ── 포맷 ─────────────────────────────────────────────────
    fmt_banner = wb.add_format({
        "bold": True, "font_size": 12,
        "font_color": C_WHITE, "bg_color": C_DARK_BLUE,
        "align": "center", "valign": "vcenter",
    })
    fmt_header = wb.add_format({
        "bold": True, "font_size": 10,
        "font_color": C_WHITE, "bg_color": C_MID_BLUE,
        "align": "center", "valign": "vcenter",
        "border": 1, "border_color": C_DARK_BLUE, "text_wrap": True,
    })
    fmt_num    = wb.add_format({"num_format": "#,##0", "valign": "top", "border": 1, "border_color": "#BDD7EE"})
    fmt_num_a  = wb.add_format({"num_format": "#,##0", "valign": "top", "bg_color": C_LIGHT_BLUE, "border": 1, "border_color": "#BDD7EE"})
    fmt_dec    = wb.add_format({"num_format": "0.00",  "valign": "top", "border": 1, "border_color": "#BDD7EE"})
    fmt_dec_a  = wb.add_format({"num_format": "0.00",  "valign": "top", "bg_color": C_LIGHT_BLUE, "border": 1, "border_color": "#BDD7EE"})
    fmt_coord  = wb.add_format({"num_format": "0.0000","valign": "top", "border": 1, "border_color": "#BDD7EE"})
    fmt_coord_a= wb.add_format({"num_format": "0.0000","valign": "top", "bg_color": C_LIGHT_BLUE, "border": 1, "border_color": "#BDD7EE"})
    fmt_text   = wb.add_format({"valign": "top", "border": 1, "border_color": "#BDD7EE"})
    fmt_text_a = wb.add_format({"valign": "top", "bg_color": C_LIGHT_BLUE, "border": 1, "border_color": "#BDD7EE"})
    fmt_wrap   = wb.add_format({"valign": "top", "text_wrap": True, "border": 1, "border_color": "#BDD7EE"})
    fmt_wrap_a = wb.add_format({"valign": "top", "text_wrap": True, "bg_color": C_LIGHT_BLUE, "border": 1, "border_color": "#BDD7EE"})
    fmt_link   = wb.add_format({"font_color": C_LINK, "underline": True, "valign": "top", "border": 1, "border_color": "#BDD7EE"})
    fmt_link_a = wb.add_format({"font_color": C_LINK, "underline": True, "valign": "top", "bg_color": C_LIGHT_BLUE, "border": 1, "border_color": "#BDD7EE"})

    FMT: dict[str, tuple] = {
        "idx":             (fmt_num,   fmt_num_a),
        "title":           (fmt_wrap,  fmt_wrap_a),
        "room_type":       (fmt_text,  fmt_text_a),
        "property_type":   (fmt_text,  fmt_text_a),
        "bedrooms":        (fmt_num,   fmt_num_a),
        "beds":            (fmt_num,   fmt_num_a),
        "bathrooms":       (fmt_dec,   fmt_dec_a),
        "rating":          (fmt_dec,   fmt_dec_a),
        "price_per_night": (fmt_num,   fmt_num_a),
        "total_price":     (fmt_num,   fmt_num_a),
        "url":             (fmt_link,  fmt_link_a),
        "latitude":        (fmt_coord, fmt_coord_a),
        "longitude":       (fmt_coord, fmt_coord_a),
        "region_query":    (fmt_text,  fmt_text_a),
        "max_guests":      (fmt_text,  fmt_text_a),
        "bed_types":       (fmt_wrap,  fmt_wrap_a),
        "description":     (fmt_wrap,  fmt_wrap_a),
        "amenities":       (fmt_wrap,  fmt_wrap_a),
        "host_name":       (fmt_text,  fmt_text_a),
        "superhost":       (fmt_text,  fmt_text_a),
        "house_rules":     (fmt_wrap,  fmt_wrap_a),
        "cancellation":    (fmt_wrap,  fmt_wrap_a),
    }

    # ── 열 너비 ──────────────────────────────────────────────
    for ci, (_, _, w) in enumerate(all_cols):
        ws.set_column(ci, ci, w)

    # ── 배너 ─────────────────────────────────────────────────
    banner_txt = (
        f"여행지: {query}   |   체크인: {checkin}   |   "
        f"체크아웃: {checkout}   |   숙박: {nights}박   |   "
        f"수집 숙소: {len(listings)}개  [상세 분석 포함]"
    )
    ws.merge_range(0, 0, 0, len(all_cols) - 1, banner_txt, fmt_banner)
    ws.set_row(0, 24)

    # ── 헤더 ─────────────────────────────────────────────────
    for ci, (header, _, _) in enumerate(all_cols):
        ws.write(1, ci, header, fmt_header)
    ws.set_row(1, 22)

    # ── 데이터 ───────────────────────────────────────────────
    for ri, lst in enumerate(listings):
        row = 2 + ri
        alt = ri % 2 == 1
        total = lst["price_per_night"] * nights

        field_map: dict[str, object] = {
            "idx":             ri + 1,
            "title":           lst["title"],
            "room_type":       lst["room_type"],
            "property_type":   lst["property_type"],
            "bedrooms":        lst.get("bedrooms") or 0,
            "beds":            lst.get("beds") or 0,
            "bathrooms":       lst.get("bathrooms") or 0.0,
            "rating":          lst.get("rating") or 0.0,
            "price_per_night": lst["price_per_night"],
            "total_price":     total,
            "url":             lst["url"],
            "latitude":        lst["latitude"],
            "longitude":       lst["longitude"],
            "region_query":    lst["region_query"],
            # 상세 필드 (fetch_detail 로 채워짐)
            **{col[1]: lst.get(col[1], "") for col in DETAIL_COLS},
        }

        for ci, (_, field, _) in enumerate(all_cols):
            fmt = FMT[field][1 if alt else 0]
            val = field_map[field]
            if field == "url":
                ws.write_url(row, ci, str(val), fmt, str(val))
            else:
                ws.write(row, ci, val, fmt)

        ws.set_row(row, 60)  # 상세 텍스트용 높은 행 높이

    ws.freeze_panes(2, 1)
    ws.autofilter(1, 0, 1 + len(listings), len(all_cols) - 1)

    wb.close()
    _fix_colors(out_path)
    print(f"✅ 저장: {out_path}")


def run(query: str, checkin: str, checkout: str) -> None:
    print(f"[1/4] 지오코딩: {query}")
    geo = geocode_region(query)
    print(f"      → {geo}")

    print(f"[2/4] Airbnb 크롤링: {query} ({checkin} ~ {checkout})")
    listings = crawl_airbnb(query, checkin, checkout, geo=geo, max_results=80)
    print(f"      → {len(listings)}개 수집")

    if not listings:
        print("❌ 수집된 숙소가 없습니다.")
        return

    print(f"[3/4] 숙소 상세 페이지 크롤링 ({len(listings)}개) ...")
    session = cf_requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept-Language": "ko-KR,ko;q=0.9",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    })

    for i, lst in enumerate(listings, 1):
        url = lst["url"]
        print(f"  [{i:>2}/{len(listings)}] {lst['title'][:40]} ...", end=" ", flush=True)
        detail = fetch_detail(url, session)
        lst.update(detail)

        ok = bool(detail.get("description") and not detail["description"].startswith("오류"))
        print("✅" if ok else f"⚠ {detail.get('description', '')[:30]}")

        if i < len(listings):
            time.sleep(2.0)

    ci_tag    = checkin.replace("-", "")[2:]
    co_tag    = checkout.replace("-", "")[2:]
    hhmm      = datetime.now().strftime("%H%M")
    folder_ts = datetime.now().strftime("%y%m%d_%H%M")
    out_dir   = _BASE_DIR / "output" / f"{folder_ts}_{query}"
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path  = out_dir / f"{query}_상세_{ci_tag}-{co_tag}_{hhmm}.xlsx"
    print(f"[4/4] Excel 생성: {out_path.name}")
    build_excel_detail(listings, query, checkin, checkout, out_path)


def main() -> None:
    query    = sys.argv[1] if len(sys.argv) > 1 else "홍대"
    checkin  = sys.argv[2] if len(sys.argv) > 2 else "2026-09-08"
    checkout = sys.argv[3] if len(sys.argv) > 3 else "2026-09-09"
    run(query, checkin, checkout)


if __name__ == "__main__":
    main()
