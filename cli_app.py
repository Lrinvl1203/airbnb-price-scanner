"""
Airbnb 시장 분석 — 콘솔 인터페이스
사용: python cli_app.py  (또는 airbnb_CLI.exe 더블클릭)
"""
from __future__ import annotations

import sys

# ── Windows 콘솔 UTF-8 설정 (배포 exe 에서 한글 깨짐 방지) ────────────
if sys.platform == "win32":
    import ctypes
    ctypes.windll.kernel32.SetConsoleOutputCP(65001)
    ctypes.windll.kernel32.SetConsoleCP(65001)
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")

from datetime import date, timedelta
from pathlib import Path

# ── 경로 설정 (PyInstaller 대응) ────────────────────────────────────────
BASE_DIR = (
    Path(sys.executable).parent if getattr(sys, "frozen", False)
    else Path(__file__).parent
)
sys.path.insert(0, str(BASE_DIR))


def _ask(prompt: str, default: str = "") -> str:
    hint = f" [{default}]" if default else ""
    val = input(f"  {prompt}{hint}: ").strip()
    return val if val else default


def _ask_date(prompt: str, default_offset: int) -> str:
    default = str(date.today() + timedelta(days=default_offset))
    while True:
        val = _ask(prompt, default)
        try:
            date.fromisoformat(val)
            return val
        except ValueError:
            print("    ⚠  YYYY-MM-DD 형식으로 입력해주세요.")


def _hr():
    print("─" * 52)


def main() -> None:
    print()
    _hr()
    print("  Airbnb 시장 분석 도구")
    _hr()
    print("  A  기본 수집  — 가격·위치·평점 Excel (~1분)")
    print("  B  상세 수집  — 소개글·편의시설 포함 (~5분)")
    print("  M  시장 분석  — 평일/주말 비교 리포트 (~15분)")
    _hr()

    mode = ""
    while mode not in ("A", "B", "M"):
        mode = _ask("모드 선택 (A/B/M)").upper()
        if mode not in ("A", "B", "M"):
            print("    ⚠  A, B, M 중 하나를 입력해주세요.")

    region = ""
    while not region:
        region = _ask("지역명 (예: 홍대, 해운대)")

    print()

    if mode == "A":
        ci = _ask_date("체크인  (YYYY-MM-DD)", 7)
        co = _ask_date("체크아웃 (YYYY-MM-DD)", 8)
        _hr()
        print(f"  실행: A 모드 | {region} | {ci} ~ {co}")
        _hr()
        from export_excel import run
        run(region, ci, co)

    elif mode == "B":
        ci = _ask_date("체크인  (YYYY-MM-DD)", 7)
        co = _ask_date("체크아웃 (YYYY-MM-DD)", 8)
        _hr()
        print(f"  실행: B 모드 | {region} | {ci} ~ {co}")
        _hr()
        from export_excel_detail import run
        run(region, ci, co)

    elif mode == "M":
        beds_str  = _ask("침실 수 (없으면 Enter)", "")
        baths_str = _ask("욕실 수 (없으면 Enter)", "")
        checkin   = _ask("기준 날짜 (없으면 3주 후 자동, YYYY-MM-DD)", "")
        cleaning  = _ask("청소비 (원, 기본 80000)", "80000")
        monthly   = _ask("월 고정비 (원, 기본 0)", "0")

        beds  = int(beds_str)  if beds_str.isdigit()  else None
        baths = int(baths_str) if baths_str.isdigit() else None
        checkin_date = None
        if checkin:
            try:
                from datetime import date as _d
                checkin_date = _d.fromisoformat(checkin)
            except ValueError:
                print("    ⚠  날짜 형식 오류 — 자동 날짜 사용")

        _hr()
        print(f"  실행: M 모드 | {region}"
              + (f" | 침실 {beds}" if beds else "")
              + (f" 욕실 {baths}" if baths else "")
              + (f" | 기준 {checkin}" if checkin else " | 날짜 자동"))
        _hr()
        from market_report import run as m_run
        m_run(
            query        = region,
            beds         = beds,
            baths        = baths,
            target_date  = checkin_date,
            cleaning_fee = int(cleaning) if cleaning.isdigit() else 80_000,
            monthly_cost = int(monthly)  if monthly.isdigit()  else 0,
        )

    print()
    input("  완료. Enter 키를 누르면 종료합니다...")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n  취소되었습니다.")
    except Exception as exc:
        print(f"\n  ❌ 오류: {exc}")
        import traceback
        traceback.print_exc()
        input("\n  Enter 키를 누르면 종료합니다...")
        sys.exit(1)
