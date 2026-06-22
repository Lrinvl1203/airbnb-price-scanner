"""
Airbnb 시장 분석 GUI
실행: python gui_app.py
"""
from __future__ import annotations

import importlib
import io
import os
import queue
import sys
import threading
import tkinter as tk
from datetime import date, timedelta
from pathlib import Path
from tkinter import messagebox, ttk

from tkcalendar import DateEntry

PROJECT_DIR = Path(__file__).parent


# ── stdout 리다이렉터 ──────────────────────────────────────────────────
class _QueueWriter(io.TextIOBase):
    def __init__(self, q: "queue.Queue[str | tuple]") -> None:
        self._q = q

    def write(self, s: str) -> int:
        if s:
            self._q.put(s)           # 일반 문자열로 넣음
        return len(s)

    def flush(self) -> None:
        pass


# ── 메인 윈도우 ───────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Airbnb 시장 분석 도구")
        self.resizable(True, True)
        self.minsize(820, 720)

        self._log_q: queue.Queue[str | tuple] = queue.Queue()
        self._running = False
        self._last_out_dir: Path | None = None

        self._build_ui()
        self._poll_log()

    # ── UI 구성 ────────────────────────────────────────────────────────
    def _build_ui(self) -> None:
        pad = {"padx": 10, "pady": 4}

        # ── 기본 설정 ──
        frm_basic = ttk.LabelFrame(self, text="기본 설정")
        frm_basic.pack(fill="x", padx=12, pady=(10, 4))

        ttk.Label(frm_basic, text="지역 (예: 홍대, 해운대, 충신동)").grid(
            row=0, column=0, sticky="w", **pad)
        self.var_region = tk.StringVar(value="홍대")
        ttk.Entry(frm_basic, textvariable=self.var_region, width=22).grid(
            row=0, column=1, columnspan=3, sticky="w", **pad)

        ttk.Label(frm_basic, text="체크인").grid(row=1, column=0, sticky="w", **pad)
        ci_default = date.today() + timedelta(days=7)
        self.ent_checkin = DateEntry(
            frm_basic, date_pattern="yyyy-mm-dd", width=13,
            year=ci_default.year, month=ci_default.month, day=ci_default.day,
            showweeknumbers=False, firstweekday="sunday",
        )
        self.ent_checkin.grid(row=1, column=1, sticky="w", **pad)
        ttk.Label(frm_basic, text="체크아웃").grid(row=1, column=2, sticky="w", padx=(16,4))
        co_default = date.today() + timedelta(days=8)
        self.ent_checkout = DateEntry(
            frm_basic, date_pattern="yyyy-mm-dd", width=13,
            year=co_default.year, month=co_default.month, day=co_default.day,
            showweeknumbers=False, firstweekday="sunday",
        )
        self.ent_checkout.grid(row=1, column=3, sticky="w", **pad)
        ttk.Label(frm_basic, text="  ← 달력 아이콘 클릭 또는 ▲▼키").grid(
            row=1, column=4, sticky="w", padx=2)

        # 모드
        ttk.Label(frm_basic, text="모드").grid(row=2, column=0, sticky="nw", **pad)
        self.var_mode = tk.StringVar(value="A")
        frm_radio = ttk.Frame(frm_basic)
        frm_radio.grid(row=2, column=1, columnspan=4, sticky="w", **pad)
        for label, val in [
            ("A — 기본  (가격·평점 Excel, ~1분)", "A"),
            ("B — 상세  (편의시설·소개글 추가, ~5분)", "B"),
            ("M — 시장 분석 리포트  (손님용+내부용+HTML, ~20분)", "M"),
        ]:
            ttk.Radiobutton(frm_radio, text=label, variable=self.var_mode,
                            value=val, command=self._on_mode_change).pack(anchor="w", pady=2)

        # 모드 설명 인포 박스
        _IB = "#f0f4fa"
        self._frm_info = tk.Frame(frm_radio, bg=_IB, relief="groove", bd=1)
        self._frm_info.pack(fill="x", padx=(22, 4), pady=(6, 4))

        _rows = ["수집 방식", "입     력", "출     력", "소요 시간"]
        self._info_lbls: dict[str, tk.Label] = {}
        for i, key in enumerate(_rows):
            tk.Label(self._frm_info, text=key, bg=_IB,
                     font=("", 9, "bold"), fg="#444444",
                     anchor="e", width=9).grid(
                row=i, column=0,
                padx=(10, 4), pady=(6 if i == 0 else 3, 6 if i == len(_rows)-1 else 3),
                sticky="e")
            tk.Label(self._frm_info, text="│", bg=_IB, fg="#bbbbbb",
                     font=("", 9)).grid(row=i, column=1, sticky="ns")
            lbl = tk.Label(self._frm_info, text="", bg=_IB,
                           font=("", 9), fg="#333333",
                           anchor="w", justify="left", wraplength=450)
            lbl.grid(row=i, column=2,
                     padx=(8, 10),
                     pady=(6 if i == 0 else 3, 6 if i == len(_rows)-1 else 3),
                     sticky="w")
            self._info_lbls[key] = lbl

        # ── M 모드 옵션 (초기에는 숨김) ──
        self.frm_m = ttk.LabelFrame(self, text="M 모드 옵션 (시장 분석)")
        # 내부 위젯만 구성, pack은 _on_mode_change에서

        _DG = "#777777"   # description gray

        def _spin_row(row, label, var, from_, to, inc, fmt="%.0f", desc=""):
            ttk.Label(self.frm_m, text=label, width=16, anchor="w").grid(
                row=row, column=0, sticky="w", padx=(10, 4), pady=4)
            ttk.Spinbox(self.frm_m, from_=from_, to=to, increment=inc,
                        textvariable=var, width=12, format=fmt).grid(
                row=row, column=1, sticky="w", padx=4, pady=4)
            if desc:
                ttk.Label(self.frm_m, text=desc, foreground=_DG,
                          wraplength=310, justify="left").grid(
                    row=row, column=2, sticky="w", padx=(8, 10), pady=4)

        self.var_beds       = tk.StringVar(value="0")
        self.var_baths      = tk.StringVar(value="0")
        self.var_cleaning   = tk.StringVar(value="80000")
        self.var_monthly    = tk.StringVar(value="0")
        self.var_occ_low    = tk.StringVar(value="0.40")
        self.var_occ_base   = tk.StringVar(value="0.60")
        self.var_occ_high   = tk.StringVar(value="0.70")
        self.var_avg_nights = tk.StringVar(value="2.0")

        _spin_row(0, "침실 수",         self.var_beds,       0, 20,      1,
                  desc="0 = 침실 수 무관 전체 수집\n특정 규모만 비교하려면 숫자 지정")
        _spin_row(1, "욕실 수",         self.var_baths,      0, 10,      1,
                  desc="0 = 욕실 수 무관 전체 수집")
        _spin_row(2, "청소비  (원/회)", self.var_cleaning,   0, 500000, 10000,
                  desc="손님에게 부과하는 1회 청소비\n→ 수익 시뮬레이터 계산에 반영됨")
        _spin_row(3, "월 고정비  (원)", self.var_monthly,    0, 5000000, 50000,
                  desc="임대료·관리비 등 매달 고정 지출\n0이면 비용 없음으로 처리")
        _spin_row(4, "평균 숙박일",     self.var_avg_nights, 1, 30, 0.5, "%.1f",
                  desc="예약 1건당 평균 숙박일\n청소비 발생 빈도 계산에 사용")

        ttk.Label(self.frm_m, text="예약률  하/기/상", width=16, anchor="w").grid(
            row=5, column=0, sticky="w", padx=(10, 4), pady=4)
        frm_occ = ttk.Frame(self.frm_m)
        frm_occ.grid(row=5, column=1, sticky="w", padx=4, pady=4)
        for var in (self.var_occ_low, self.var_occ_base, self.var_occ_high):
            ttk.Spinbox(frm_occ, from_=0.0, to=1.0, increment=0.05,
                        textvariable=var, width=7, format="%.2f").pack(side="left", padx=(0, 4))
        ttk.Label(self.frm_m, text="연간 예약률 시나리오 3가지 (보수적 / 기준 / 공격적)\n수익 시뮬레이터에서 3개 열로 나란히 비교됨",
                  foreground=_DG, wraplength=310, justify="left").grid(
            row=5, column=2, sticky="w", padx=(8, 10), pady=4)

        ttk.Label(self.frm_m, text="출력 형식", width=16, anchor="w").grid(
            row=6, column=0, sticky="w", padx=(10, 4), pady=4)
        self.var_output_mode = tk.StringVar(value="both")
        frm_om = ttk.Frame(self.frm_m)
        frm_om.grid(row=6, column=1, sticky="w", padx=4, pady=4)
        for txt, val in [("손님+내부+HTML", "both"), ("손님용만", "client"), ("내부용만", "internal")]:
            ttk.Radiobutton(frm_om, text=txt, variable=self.var_output_mode,
                            value=val).pack(side="left", padx=(0, 8))
        ttk.Label(self.frm_m,
                  text="both: 손님용·내부용 Excel + HTML 모두 생성\nclient: 손님 전달용 Excel만  /  internal: 내부 분석용 Excel만",
                  foreground=_DG, wraplength=310, justify="left").grid(
            row=6, column=2, sticky="w", padx=(8, 10), pady=4)

        # ── 수집 매물 수 ──
        self.frm_adv = ttk.LabelFrame(self, text="수집 매물 수")
        self.frm_adv.pack(fill="x", padx=12, pady=4)

        ttk.Label(self.frm_adv, text="최대 매물", font=("", 10, "bold")).grid(
            row=0, column=0, sticky="w", **pad)
        self.var_max = tk.StringVar(value="200")
        ttk.Spinbox(self.frm_adv, from_=10, to=9999, increment=10,
                    textvariable=self.var_max, width=7).grid(
            row=0, column=1, sticky="w", **pad)

        ttk.Label(self.frm_adv, text="개     최대 페이지").grid(
            row=0, column=2, sticky="w")
        self.var_pages = tk.StringVar(value="20")
        ttk.Spinbox(self.frm_adv, from_=1, to=50, increment=1,
                    textvariable=self.var_pages, width=5).grid(
            row=0, column=3, sticky="w", **pad)
        ttk.Label(self.frm_adv,
                  text="페이지   (페이지당 ~20개 · Airbnb 실제 한계 ~200개)").grid(
            row=0, column=4, sticky="w", padx=4)

        # ── 버튼 ──
        frm_btn = ttk.Frame(self)
        frm_btn.pack(pady=8)
        self.btn_run = ttk.Button(frm_btn, text="▶  실행",
                                  command=self._run, width=14)
        self.btn_run.pack(side="left", padx=6)
        self.btn_open = ttk.Button(frm_btn, text="📂  결과 폴더 열기",
                                   command=self._open_output, width=18, state="disabled")
        self.btn_open.pack(side="left", padx=6)
        ttk.Button(frm_btn, text="🗑  로그 지우기",
                   command=self._clear_log, width=14).pack(side="left", padx=6)

        # ── 로그 창 ──
        frm_log = ttk.LabelFrame(self, text="실행 로그")
        frm_log.pack(fill="both", expand=True, padx=12, pady=(0, 10))

        self.log_text = tk.Text(
            frm_log, state="disabled", wrap="word",
            bg="#1e1e1e", fg="#d4d4d4", font=("Consolas", 9),
            relief="flat", borderwidth=0, selectbackground="#264f78",
        )
        self.log_text.tag_configure("hdr",  foreground="#4ec9b0", font=("Consolas", 9, "bold"))
        self.log_text.tag_configure("file", foreground="#9cdcfe")
        self.log_text.tag_configure("dir",  foreground="#dcdcaa")
        self.log_text.tag_configure("err",  foreground="#f44747")
        self.log_text.tag_configure("ok",   foreground="#4ec9b0")

        sb = ttk.Scrollbar(frm_log, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.log_text.pack(fill="both", expand=True, padx=2, pady=2)

        # 초기 상태: 기본 모드 설명 표시
        self._on_mode_change()

    # ── 모드 설명 데이터 ────────────────────────────────────────────────
    _MODE_INFO: dict[str, dict[str, str]] = {
        "A": {
            "수집 방식": "가격·위치·평점·숙소유형 등  Airbnb 목록 페이지 데이터만 빠르게 수집",
            "입     력": "지역명  /  체크인·체크아웃 날짜  /  최대 매물 수",
            "출     력": "output/raw/airbnb_{지역}_{날짜}.xlsx   (14컬럼)",
            "소요 시간": "약 30~60초",
        },
        "B": {
            "수집 방식": "A 모드 수집 후  각 숙소 상세 페이지를 직접 방문 →  소개글·편의시설·호스트 정보 추가",
            "입     력": "지역명  /  체크인·체크아웃 날짜  /  최대 매물 수",
            "출     력": "output/raw/airbnb_detail_{지역}_{날짜}.xlsx   (22컬럼)",
            "소요 시간": "숙소당 약 2초 × 최대 80개  ≈  3~5분",
        },
        "M": {
            "수집 방식": "체크인이 속한 주의  평일(월→화) · 주말(금→토)  2개 날짜창 자동 수집 →  통계 분석",
            "입     력": "지역명  /  체크인(해당 주 기준 자동 분리)  /  아래 M 모드 옵션 전체",
            "출     력": "output/market/   손님용.xlsx  +  내부용.xlsx  +  .html",
            "소요 시간": "약 10~20분",
        },
    }

    # ── 모드 토글 ──────────────────────────────────────────────────────
    def _on_mode_change(self) -> None:
        mode = self.var_mode.get()
        info = self._MODE_INFO.get(mode, {})
        for key, lbl in self._info_lbls.items():
            lbl.configure(text=info.get(key, ""))
        if mode == "M":
            self.frm_m.pack(fill="x", padx=12, pady=4, before=self.frm_adv)
        else:
            self.frm_m.pack_forget()

    # ── 로그 폴링 ──────────────────────────────────────────────────────
    def _poll_log(self) -> None:
        try:
            while True:
                item = self._log_q.get_nowait()
                if isinstance(item, tuple):
                    text, tag = item
                else:
                    text, tag = item, None
                self.log_text.configure(state="normal")
                if tag:
                    self.log_text.insert("end", text, tag)
                else:
                    self.log_text.insert("end", text)
                self.log_text.see("end")
                self.log_text.configure(state="disabled")
        except queue.Empty:
            pass
        self.after(80, self._poll_log)

    def _log(self, msg: str, tag: str | None = None) -> None:
        self._log_q.put((msg + "\n", tag))

    def _clear_log(self) -> None:
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    # ── 입력 검증 + 값 수집 (메인 스레드에서만 호출) ──────────────────
    def _collect_inputs(self) -> dict | None:
        """UI 값을 메인 스레드에서 읽어 dict로 반환. 검증 실패 시 None."""
        region = self.var_region.get().strip()
        if not region:
            messagebox.showerror("입력 오류", "지역을 입력하세요.")
            return None
        try:
            ci: date = self.ent_checkin.get_date()
            co: date = self.ent_checkout.get_date()
        except Exception:
            messagebox.showerror("입력 오류", "날짜 형식을 확인하세요.")
            return None
        if co <= ci:
            messagebox.showerror("입력 오류", "체크아웃은 체크인 이후여야 합니다.")
            return None

        try:
            max_res = int(self.var_max.get() or 200)
            max_pgs = int(self.var_pages.get() or 20)
        except ValueError:
            messagebox.showerror("입력 오류", "매물 수 / 페이지 수는 정수여야 합니다.")
            return None
        if max_res == 0:
            max_res = 9999

        beds_raw  = self.var_beds.get().strip()
        baths_raw = self.var_baths.get().strip()

        return {
            "mode":        self.var_mode.get(),
            "region":      region,
            "checkin":     ci.strftime("%Y-%m-%d"),
            "checkout":    co.strftime("%Y-%m-%d"),
            "max_res":     max_res,
            "max_pgs":     max_pgs,
            # M 옵션
            "beds":        int(beds_raw)  if beds_raw  and beds_raw  != "0" else None,
            "baths":       int(baths_raw) if baths_raw and baths_raw != "0" else None,
            "cleaning":    int(float(self.var_cleaning.get()   or 80000)),
            "monthly":     int(float(self.var_monthly.get()    or 0)),
            "occ_low":     float(self.var_occ_low.get()        or 0.40),
            "occ_base":    float(self.var_occ_base.get()       or 0.60),
            "occ_high":    float(self.var_occ_high.get()       or 0.70),
            "avg_nights":  float(self.var_avg_nights.get()     or 2.0),
            "output_mode": self.var_output_mode.get(),
        }

    # ── 실행 ──────────────────────────────────────────────────────────
    def _run(self) -> None:
        if self._running:
            return
        inputs = self._collect_inputs()   # 메인 스레드에서 값 수집
        if inputs is None:
            return

        self._running = True
        self._last_out_dir = None
        self.btn_run.configure(state="disabled", text="⏳ 실행 중...")
        self.btn_open.configure(state="disabled")
        self._clear_log()

        threading.Thread(target=self._worker, args=(inputs,), daemon=True).start()

    # ── 워커 (별도 스레드 — tkinter 위젯 접근 금지) ────────────────────
    def _worker(self, inp: dict) -> None:
        old_out, old_err = sys.stdout, sys.stderr
        sys.stdout = sys.stderr = _QueueWriter(self._log_q)

        sys.path.insert(0, str(PROJECT_DIR))

        # airbnb_fetch 모듈 로드 + 원본 crawl 저장
        import airbnb_fetch as _af
        _orig_crawl = _af.AirbnbClient.crawl

        try:
            mode    = inp["mode"]
            region  = inp["region"]
            checkin = inp["checkin"]
            checkout= inp["checkout"]
            max_res = inp["max_res"]
            max_pgs = inp["max_pgs"]

            self._log(f"=== 실행 시작: 모드={mode}, 지역={region} ===")
            self._log(f"    체크인={checkin}, 체크아웃={checkout}")
            self._log(f"    최대 매물={max_res}개, 최대 페이지={max_pgs}페이지")
            self._log("")

            # GUI 설정값으로 crawl 강제 적용
            def _patched_crawl(self_c, query, ci, co, geo=None,
                               max_results=60, max_pages=4):
                return _orig_crawl(self_c, query, ci, co,
                                   geo=geo, max_results=max_res, max_pages=max_pgs)
            _af.AirbnbClient.crawl = _patched_crawl

            if mode == "A":
                import export_excel as _m
                importlib.reload(_m)
                _m.run(region, checkin, checkout)

            elif mode == "B":
                import export_excel_detail as _m
                importlib.reload(_m)
                _m.run(region, checkin, checkout)

            elif mode == "M":
                import market_report as _m
                importlib.reload(_m)
                kwargs: dict = {
                    "checkin":      inp["checkin"],
                    "cleaning_fee": inp["cleaning"],
                    "monthly_cost": inp["monthly"],
                    "occ_low":      inp["occ_low"],
                    "occ_base":     inp["occ_base"],
                    "occ_high":     inp["occ_high"],
                    "avg_nights":   inp["avg_nights"],
                    "output_mode":  inp["output_mode"],
                }
                if inp["beds"]  is not None: kwargs["beds"]  = inp["beds"]
                if inp["baths"] is not None: kwargs["baths"] = inp["baths"]
                _m.run(region, **kwargs)

            # ── 완료 후 파일 요약 ────────────────────────────────────
            today_str = date.today().strftime("%Y%m%d")
            out_dir   = PROJECT_DIR / "output" / f"{today_str}_{region}"
            self._last_out_dir = out_dir

            self._log("")
            self._log("═" * 56, "hdr")
            self._log("  생성된 파일 목록", "hdr")
            self._log("═" * 56, "hdr")
            self._log(f"  📁 output/{today_str}_{region}/", "dir")
            self._log("─" * 56, "hdr")

            if out_dir.exists():
                files = sorted(out_dir.iterdir())
                if files:
                    for f in files:
                        size_kb = f.stat().st_size / 1024
                        icon = "🌐" if f.suffix == ".html" else "📊"
                        self._log(f"  {icon} {f.name:<44}  {size_kb:>6.0f} KB", "file")
                else:
                    self._log("  (생성된 파일 없음)", "err")
            else:
                self._log("  (출력 폴더가 없습니다 — 수집 결과 0개일 수 있음)", "err")

            self._log("═" * 56, "hdr")
            self._log("")
            self._log("✅ 완료", "ok")
            self.after(0, lambda: self.btn_open.configure(state="normal"))

        except Exception as exc:
            import traceback
            if isinstance(exc, PermissionError) or "Permission denied" in str(exc):
                fname = str(exc).split("'")[-2] if "'" in str(exc) else "파일"
                import os as _os
                fname = _os.path.basename(fname)
                self._log(f"\n❌ 파일 저장 실패: {fname}", "err")
                self._log("   → Excel에서 해당 파일이 열려 있습니다.", "err")
                self._log("   → 파일을 닫은 후 다시 실행하세요.", "err")
            else:
                self._log(f"\n❌ 오류: {exc}", "err")
                self._log(traceback.format_exc(), "err")

        finally:
            sys.stdout, sys.stderr = old_out, old_err
            _af.AirbnbClient.crawl = _orig_crawl
            self._running = False
            self.after(0, lambda: self.btn_run.configure(
                state="normal", text="▶  실행"))

    # ── 결과 폴더 열기 ──────────────────────────────────────────────────
    def _open_output(self) -> None:
        folder = self._last_out_dir or (PROJECT_DIR / "output")
        folder.mkdir(parents=True, exist_ok=True)
        os.startfile(folder)


# ── 진입점 ────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = App()
    app.mainloop()
