"""
Airbnb 검색 결과 → Excel 추출 스크립트
사용: python export_excel.py <지역> <체크인YYYY-MM-DD> <체크아웃YYYY-MM-DD>
"""
from __future__ import annotations

import sys
from datetime import date, datetime
from pathlib import Path

import re
import shutil
import tempfile
import zipfile

import xlsxwriter

sys.path.insert(0, str(Path(__file__).parent))
from airbnb_fetch import crawl_airbnb, geocode_region

# 이모지·특수기호 제거 (Airbnb 숙소명에 포함될 수 있음)
_EMOJI_RE = re.compile(
    "["
    "\U0001F000-\U0001FFFF"   # 보조 다국어 면 (이모지 대부분)
    "\U00002600-\U000027BF"   # 기타 기호
    "\U0001F900-\U0001F9FF"   # 보충 기호
    "\U0000200D"              # ZWJ
    "\U0000FE0F"              # 변형 선택자
    "]+",
    flags=re.UNICODE,
)

def _strip_emoji(text: object) -> str:
    if not isinstance(text, str):
        return str(text) if text is not None else ""
    return _EMOJI_RE.sub("", text).strip()


def _fix_colors(xlsx_path: Path) -> None:
    """
    styles.xml 3단계 패치 (xlsxwriter 색상 렌더링 버그 수정):
    1) fgColor alpha 00 → FF  (투명 → 불투명)
    2) solid fill에 bgColor indexed="64" 추가  (Excel 필수)
    3) fillId>1인 xf 항목에 applyFill="1" 추가  (없으면 Excel이 fill 무시)
    """
    tmp = Path(tempfile.mktemp(suffix=".xlsx"))
    with zipfile.ZipFile(xlsx_path, "r") as zin, \
         zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/styles.xml":
                xml = data.decode("utf-8")

                # 1) alpha 00 → FF
                xml = re.sub(
                    r'((?:fg|bg)Color rgb=")00([0-9A-Fa-f]{6}")',
                    r"\1FF\2",
                    xml,
                )
                # 2) solid fill에 bgColor indexed="64" 추가 (없을 경우만)
                xml = re.sub(
                    r'(<fgColor rgb="FF[0-9A-Fa-f]{6}"/>)(?!<bgColor)',
                    r'\1<bgColor indexed="64"/>',
                    xml,
                )
                # 3) cellXfs: fillId > 1인 <xf>에 applyFill="1" 추가
                def _add_apply_fill(m: re.Match) -> str:
                    tag = m.group(0)
                    fid = re.search(r'fillId="(\d+)"', tag)
                    if fid and int(fid.group(1)) > 1 and 'applyFill' not in tag:
                        tag = tag.replace('fillId=', 'applyFill="1" fillId=')
                    return tag
                xml = re.sub(r'<xf [^>]+>', _add_apply_fill, xml)

                data = xml.encode("utf-8")
            zout.writestr(item, data)
    shutil.move(str(tmp), str(xlsx_path))

# ── 색상 팔레트 ──────────────────────────────────────────────
C_DARK_BLUE  = "#1F4E79"
C_MID_BLUE   = "#2E75B6"
C_LIGHT_BLUE = "#EBF3FB"
C_LINK       = "#0563C1"
C_WHITE      = "#FFFFFF"
C_BLACK      = "#000000"
C_SUMMARY_BG = "#D6E4F0"

COLS = [
    # (헤더,            필드 키,           너비)
    ("번호",            "idx",             5),
    ("숙소명",          "title",           42),
    ("숙소 유형",       "room_type",       13),
    ("건물 유형",       "property_type",   13),
    ("침실",            "bedrooms",        7),
    ("침대",            "beds",            7),
    ("욕실",            "bathrooms",       7),
    ("평점",            "rating",          8),
    ("1박 가격 (₩)",    "price_per_night", 16),
    ("총 가격 (₩)",     "total_price",     16),
    ("링크",            "url",             55),
    ("위도",            "latitude",        12),
    ("경도",            "longitude",       12),
    ("지역 쿼리",       "region_query",    12),
]


def build_excel(
    listings: list[dict],
    query: str,
    checkin: str,
    checkout: str,
    out_path: Path,
) -> None:
    nights = (date.fromisoformat(checkout) - date.fromisoformat(checkin)).days

    wb = xlsxwriter.Workbook(str(out_path))
    ws = wb.add_worksheet("숙소 목록")

    # ── 공통 포맷 ────────────────────────────────────────────
    fmt_banner = wb.add_format({
        "bold": True, "font_size": 12,
        "font_color": C_WHITE, "bg_color": C_DARK_BLUE,
        "align": "center", "valign": "vcenter",
        "border": 0,
    })
    fmt_header = wb.add_format({
        "bold": True, "font_size": 10,
        "font_color": C_WHITE, "bg_color": C_MID_BLUE,
        "align": "center", "valign": "vcenter",
        "border": 1, "border_color": C_DARK_BLUE,
        "text_wrap": True,
    })
    fmt_num = wb.add_format({
        "num_format": "#,##0", "valign": "vcenter",
        "border": 1, "border_color": "#BDD7EE",
    })
    fmt_num_alt = wb.add_format({
        "num_format": "#,##0", "valign": "vcenter",
        "bg_color": C_LIGHT_BLUE,
        "border": 1, "border_color": "#BDD7EE",
    })
    fmt_dec = wb.add_format({
        "num_format": "0.00", "valign": "vcenter",
        "border": 1, "border_color": "#BDD7EE",
    })
    fmt_dec_alt = wb.add_format({
        "num_format": "0.00", "valign": "vcenter",
        "bg_color": C_LIGHT_BLUE,
        "border": 1, "border_color": "#BDD7EE",
    })
    fmt_coord = wb.add_format({
        "num_format": "0.0000", "valign": "vcenter",
        "border": 1, "border_color": "#BDD7EE",
    })
    fmt_coord_alt = wb.add_format({
        "num_format": "0.0000", "valign": "vcenter",
        "bg_color": C_LIGHT_BLUE,
        "border": 1, "border_color": "#BDD7EE",
    })
    fmt_text = wb.add_format({
        "valign": "vcenter",
        "border": 1, "border_color": "#BDD7EE",
    })
    fmt_text_alt = wb.add_format({
        "valign": "vcenter", "bg_color": C_LIGHT_BLUE,
        "border": 1, "border_color": "#BDD7EE",
    })
    fmt_wrap = wb.add_format({
        "valign": "vcenter", "text_wrap": True,
        "border": 1, "border_color": "#BDD7EE",
    })
    fmt_wrap_alt = wb.add_format({
        "valign": "vcenter", "text_wrap": True, "bg_color": C_LIGHT_BLUE,
        "border": 1, "border_color": "#BDD7EE",
    })
    fmt_link = wb.add_format({
        "font_color": C_LINK, "underline": True, "valign": "vcenter",
        "border": 1, "border_color": "#BDD7EE",
    })
    fmt_link_alt = wb.add_format({
        "font_color": C_LINK, "underline": True, "valign": "vcenter",
        "bg_color": C_LIGHT_BLUE,
        "border": 1, "border_color": "#BDD7EE",
    })

    # 필드별 포맷 매핑 (일반, alt)
    FMT: dict[str, tuple] = {
        "idx":             (fmt_num,    fmt_num_alt),
        "title":           (fmt_wrap,   fmt_wrap_alt),
        "room_type":       (fmt_text,   fmt_text_alt),
        "property_type":   (fmt_text,   fmt_text_alt),
        "bedrooms":        (fmt_num,    fmt_num_alt),
        "beds":            (fmt_num,    fmt_num_alt),
        "bathrooms":       (fmt_dec,    fmt_dec_alt),
        "rating":          (fmt_dec,    fmt_dec_alt),
        "price_per_night": (fmt_num,    fmt_num_alt),
        "total_price":     (fmt_num,    fmt_num_alt),
        "url":             (fmt_link,   fmt_link_alt),
        "latitude":        (fmt_coord,  fmt_coord_alt),
        "longitude":       (fmt_coord,  fmt_coord_alt),
        "region_query":    (fmt_text,   fmt_text_alt),
    }

    # ── 열 너비 설정 ─────────────────────────────────────────
    for ci, (_, _, width) in enumerate(COLS):
        ws.set_column(ci, ci, width)

    # ── 1행: 배너 ────────────────────────────────────────────
    banner_txt = (
        f"여행지: {query}   |   체크인: {checkin}   |   "
        f"체크아웃: {checkout}   |   숙박: {nights}박   |   "
        f"수집 숙소: {len(listings)}개"
    )
    ws.merge_range(0, 0, 0, len(COLS) - 1, banner_txt, fmt_banner)
    ws.set_row(0, 24)

    # ── 2행: 컬럼 헤더 ──────────────────────────────────────
    for ci, (header, _, _) in enumerate(COLS):
        ws.write(1, ci, header, fmt_header)
    ws.set_row(1, 22)

    # ── 데이터 행 ───────────────────────────────────────────
    for ri, lst in enumerate(listings):
        row   = 2 + ri
        alt   = ri % 2 == 1
        total = lst["price_per_night"] * nights

        field_map: dict[str, object] = {
            "idx":             ri + 1,
            "title":           _strip_emoji(lst["title"]),
            "room_type":       _strip_emoji(lst["room_type"]),
            "property_type":   _strip_emoji(lst["property_type"]),
            "bedrooms":        lst["bedrooms"] if lst["bedrooms"] else 0,
            "beds":            lst["beds"] if lst["beds"] else 0,
            "bathrooms":       lst["bathrooms"] if lst["bathrooms"] else 0.0,
            "rating":          lst["rating"] if lst["rating"] else 0.0,
            "price_per_night": lst["price_per_night"],
            "total_price":     total,
            "url":             lst["url"],
            "latitude":        lst["latitude"],
            "longitude":       lst["longitude"],
            "region_query":    _strip_emoji(lst["region_query"]),
        }

        for ci, (_, field, _) in enumerate(COLS):
            fmt = FMT[field][1 if alt else 0]
            val = field_map[field]
            if field == "url":
                ws.write_url(row, ci, str(val), fmt, str(val))
            else:
                ws.write(row, ci, val, fmt)

        ws.set_row(row, 18)

    # 헤더 고정 + 자동필터
    ws.freeze_panes(2, 1)
    ws.autofilter(1, 0, 1 + len(listings), len(COLS) - 1)

    # ── 요약 시트 ───────────────────────────────────────────
    ws2    = wb.add_worksheet("요약")
    prices = [l["price_per_night"] for l in listings]
    totals = [p * nights for p in prices]
    ratings = [l["rating"] for l in listings if l["rating"]]

    fmt_title2 = wb.add_format({
        "bold": True, "font_size": 11,
        "font_color": C_WHITE, "bg_color": C_DARK_BLUE,
        "align": "center", "valign": "vcenter",
    })
    fmt_key = wb.add_format({
        "bold": True, "font_size": 10,
        "align": "right", "valign": "vcenter",
    })
    fmt_val = wb.add_format({
        "font_size": 10, "align": "left", "valign": "vcenter",
    })
    fmt_key_hl = wb.add_format({
        "bold": True, "font_size": 10,
        "align": "right", "valign": "vcenter",
        "bg_color": C_SUMMARY_BG,
    })
    fmt_val_hl = wb.add_format({
        "font_size": 10, "align": "left", "valign": "vcenter",
        "bg_color": C_SUMMARY_BG,
    })

    ws2.merge_range(0, 0, 0, 1, f"검색 요약 — {query}  {checkin} ~ {checkout}", fmt_title2)
    ws2.set_row(0, 22)

    rows2 = [
        ("검색 지역",      query,                                    True),
        ("체크인",         checkin,                                   False),
        ("체크아웃",       checkout,                                  False),
        ("숙박 박수",      f"{nights}박",                             False),
        ("수집 숙소 수",   f"{len(listings)}개",                      True),
        ("",               "",                                        False),
        ("최저 1박가",     f"₩{min(prices):,}",                      False),
        ("최고 1박가",     f"₩{max(prices):,}",                      False),
        ("평균 1박가",     f"₩{round(sum(prices)/len(prices)):,}",    False),
        ("중간값 1박가",   f"₩{sorted(prices)[len(prices)//2]:,}",    False),
        ("",               "",                                        False),
        (f"최저 {nights}박 총액", f"₩{min(totals):,}",               False),
        (f"최고 {nights}박 총액", f"₩{max(totals):,}",               False),
        (f"평균 {nights}박 총액", f"₩{round(sum(totals)/len(totals)):,}", False),
        ("",               "",                                        False),
        ("평균 평점",      f"{sum(ratings)/len(ratings):.2f}" if ratings else "-", False),
        ("평점 있는 숙소", f"{len(ratings)}개",                       False),
    ]
    for r2, (k, v, hl) in enumerate(rows2, start=1):
        ws2.write(r2, 0, k, fmt_key_hl if hl else fmt_key)
        ws2.write(r2, 1, v, fmt_val_hl if hl else fmt_val)
        ws2.set_row(r2, 17)

    ws2.set_column(0, 0, 20)
    ws2.set_column(1, 1, 24)

    wb.close()
    _fix_colors(out_path)
    print(f"✅ 저장 완료: {out_path}")


def run(query: str, checkin: str, checkout: str) -> None:
    print(f"[1/3] 지오코딩: {query}")
    geo = geocode_region(query)
    print(f"      → {geo}")

    print(f"[2/3] Airbnb 크롤링: {query} ({checkin} ~ {checkout})")
    listings = crawl_airbnb(query, checkin, checkout, geo=geo, max_results=80)
    print(f"      → {len(listings)}개 수집")

    if not listings:
        print("❌ 수집된 숙소가 없습니다.")
        return

    nights   = (date.fromisoformat(checkout) - date.fromisoformat(checkin)).days
    ci_tag    = checkin.replace("-", "")[2:]    # 2026-06-22 → 260622
    co_tag    = checkout.replace("-", "")[2:]
    hhmm      = datetime.now().strftime("%H%M")
    folder_ts = datetime.now().strftime("%y%m%d_%H%M")
    out_dir   = Path(__file__).parent / "output" / f"{folder_ts}_{query}"
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path  = out_dir / f"{query}_기본_{ci_tag}-{co_tag}_{hhmm}.xlsx"

    print(f"[3/3] Excel 생성: {out_path.name}")
    build_excel(listings, query, checkin, checkout, out_path)

    print()
    print(f"{'번호':>3}  {'숙소명':<35}  {'유형':<8}  {'1박가':>10}  {f'{nights}박 총액':>12}  평점")
    print("-" * 85)
    for i, l in enumerate(listings, 1):
        total = l["price_per_night"] * nights
        name  = l["title"][:33]
        print(f"{i:>3}  {name:<35}  {l['room_type']:<8}  ₩{l['price_per_night']:>8,}  ₩{total:>10,}  {l['rating']}")


def main() -> None:
    query    = sys.argv[1] if len(sys.argv) > 1 else "충무로"
    checkin  = sys.argv[2] if len(sys.argv) > 2 else "2026-09-08"
    checkout = sys.argv[3] if len(sys.argv) > 3 else "2026-09-10"
    run(query, checkin, checkout)


if __name__ == "__main__":
    main()
