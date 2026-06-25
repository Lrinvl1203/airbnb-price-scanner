"""
Airbnb 시장 분석 GUI
실행: python gui_app.py
"""
from __future__ import annotations

import importlib
import io
import os
import queue
import re
import sys
import threading
import tkinter as tk
import tkinter.font as tkfont
from datetime import date, datetime, timedelta
from pathlib import Path
from tkinter import messagebox

import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.widgets import Floodgauge

from tkcalendar import DateEntry

PROJECT_DIR = (
    Path(sys.executable).parent if getattr(sys, "frozen", False)
    else Path(__file__).parent
)

# ── 로그 필터 ──────────────────────────────────────────────────────────
_STEP_RE    = re.compile(r'\[(\d+)/(\d+)\]\s*(.*)')
_VERBOSE_RE = re.compile(r'^\d+\s{2,}')

def _is_verbose(s: str) -> bool:
    t = s.strip()
    if not t:
        return True
    if _VERBOSE_RE.match(t):
        return True
    if re.match(r'^[-─━=]{4,}$', t):
        return True
    if '번호' in t and '숙소명' in t:
        return True
    return False


# ── stdout 리다이렉터 ──────────────────────────────────────────────────
class _QueueWriter(io.TextIOBase):
    def __init__(self, q: "queue.Queue[str | tuple]") -> None:
        self._q = q

    def write(self, s: str) -> int:
        if s:
            m = _STEP_RE.search(s)
            if m:
                n, total = int(m.group(1)), int(m.group(2))
                desc = m.group(3).strip().split("(")[0].strip()
                self._q.put(("PROGRESS", n, total, desc))
            if not _is_verbose(s):
                self._q.put(s)
        return len(s)

    def flush(self) -> None:
        pass


# ── DateEntry 내비게이션 버그 픽스 ────────────────────────────────────
class _FixedDateEntry(DateEntry):
    """월·년 이동 화살표 클릭 시 달력이 닫히는 tkcalendar 버그 수정."""

    def _on_focus_out_cal(self, event) -> None:
        try:
            fw = self.focus_get()
            if fw is not None and fw is not self:
                top = str(self._top_cal)
                fw_path = str(fw)
                if fw_path == top or fw_path.startswith(top + "."):
                    return
        except Exception:
            pass
        super()._on_focus_out_cal(event)


# ── 메인 윈도우 ───────────────────────────────────────────────────────
class App(ttk.Window):
    # ── 테마 상수 ──────────────────────────────────────────────────────
    _FONT_FAMILY = "맑은 고딕"
    _FONT_SIZE   = 10
    _FONT_SMALL  = 9
    _FONT_XSMALL = 9

    _LOG_BG      = "#141821"
    _LOG_FG      = "#e6edf3"
    _INFO_BG     = "#252b38"
    _INFO_FG     = "#b4bdcf"
    _INFO_KEY_FG = "#95a0b5"
    _INFO_BAR    = "#3b465a"
    _MUTED_FG    = "#8b94a8"
    _DATE_BG     = "#f3f6fa"
    _DATE_FG     = "#151922"

    def __init__(self) -> None:
        super().__init__(themename="darkly")
        self._configure_fonts()
        self.title("Airbnb 시장 분석 도구")
        self.resizable(True, True)
        self.minsize(900, 720)

        self._log_q: queue.Queue[str | tuple] = queue.Queue()
        self._running = False
        self._last_out_dir: Path | None = None

        self._build_ui()
        self._poll_log()

        # 화면 중앙 배치
        self.update_idletasks()
        w, h = 980, 880
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}")

    def _configure_fonts(self) -> None:
        """Tk/ttk 기본 폰트를 한 단계 올려 가독성을 맞춘다."""
        font_specs = {
            "TkDefaultFont":      (self._FONT_SIZE, "normal"),
            "TkTextFont":         (self._FONT_SIZE, "normal"),
            "TkMenuFont":         (self._FONT_SIZE, "normal"),
            "TkHeadingFont":      (self._FONT_SIZE + 1, "bold"),
            "TkCaptionFont":      (self._FONT_SMALL, "normal"),
            "TkSmallCaptionFont": (self._FONT_SMALL, "normal"),
            "TkIconFont":         (self._FONT_SIZE, "normal"),
        }
        for name, (size, weight) in font_specs.items():
            try:
                tkfont.nametofont(name).configure(
                    family=self._FONT_FAMILY,
                    size=size,
                    weight=weight,
                )
            except tk.TclError:
                pass

        style = ttk.Style()
        base = (self._FONT_FAMILY, self._FONT_SIZE)
        small = (self._FONT_FAMILY, self._FONT_SMALL)
        style.configure("TLabel", font=base)
        style.configure("TButton", font=base)
        style.configure("TEntry", font=base)
        style.configure("TSpinbox", font=base)
        style.configure("TRadiobutton", font=base)
        style.configure("TLabelframe.Label", font=(self._FONT_FAMILY, self._FONT_SIZE, "bold"))
        style.configure("TCheckbutton", font=base)
        style.configure("Toolbutton", font=small)

    # ── UI 구성 ────────────────────────────────────────────────────────
    def _build_ui(self) -> None:
        # 메인 grid — row 3(로그)만 늘어남
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        # ── 0: 기본 설정 ────────────────────────────────────────────────
        frm_basic = ttk.LabelFrame(self, text="기본 설정")
        frm_basic.grid(row=0, column=0, sticky="ew", padx=14, pady=(12, 4))
        frm_basic.columnconfigure(1, weight=1)

        # 지역
        ttk.Label(frm_basic, text="지역").grid(
            row=0, column=0, sticky="w", padx=(0, 10), pady=5)
        self.var_region = tk.StringVar(value="홍대")
        ttk.Entry(frm_basic, textvariable=self.var_region, width=24).grid(
            row=0, column=1, sticky="w", pady=5)
        ttk.Label(frm_basic, text="예: 홍대, 해운대, 충신동",
                  foreground=self._MUTED_FG,
                  font=(self._FONT_FAMILY, self._FONT_XSMALL)).grid(
            row=0, column=2, sticky="w", padx=(10, 0), pady=5)

        # 날짜 — 체크인 ~ 체크아웃 한 줄
        ttk.Label(frm_basic, text="날짜").grid(
            row=1, column=0, sticky="w", padx=(0, 10), pady=5)
        frm_dates = ttk.Frame(frm_basic)
        frm_dates.grid(row=1, column=1, columnspan=2, sticky="w", pady=5)

        ci_default = date.today() + timedelta(days=7)
        self.ent_checkin = _FixedDateEntry(
            frm_dates, date_pattern="yyyy-mm-dd", width=13,
            year=ci_default.year, month=ci_default.month, day=ci_default.day,
            showweeknumbers=False, firstweekday="sunday",
        )
        self.ent_checkin.pack(side="left")

        ttk.Label(frm_dates, text="  →  ").pack(side="left")

        co_default = date.today() + timedelta(days=8)
        self.ent_checkout = _FixedDateEntry(
            frm_dates, date_pattern="yyyy-mm-dd", width=13,
            year=co_default.year, month=co_default.month, day=co_default.day,
            showweeknumbers=False, firstweekday="sunday",
        )
        self.ent_checkout.pack(side="left")

        # DateEntry는 TCombobox 기반 'DateEntry' 스타일 사용 — 위젯 생성 후 덮어씀
        _s = ttk.Style()
        _s.configure("DateEntry",
                      fieldbackground=self._DATE_BG,
                      foreground=self._DATE_FG,
                      insertcolor=self._DATE_FG,
                      arrowcolor=self._DATE_FG,
                      selectforeground="#ffffff",
                      selectbackground="#2563eb",
                      font=(self._FONT_FAMILY, self._FONT_SIZE))
        _s.map("DateEntry",
               fieldbackground=[("readonly", self._DATE_BG), ("disabled", "#d8dde6")],
               foreground=[("readonly", self._DATE_FG), ("disabled", "#7c8493")],
               arrowcolor=[("disabled", "#888888")])

        # 모드 선택
        ttk.Label(frm_basic, text="모드").grid(
            row=2, column=0, sticky="nw", padx=(0, 10), pady=(8, 5))
        self.var_mode = tk.StringVar(value="A")
        frm_radio = ttk.Frame(frm_basic)
        frm_radio.grid(row=2, column=1, columnspan=2, sticky="ew", pady=(8, 0))

        for label, val in [
            ("A — 기본  (가격·평점 Excel, ~1분)", "A"),
            ("B — 상세  (편의시설·소개글 추가, ~5분)", "B"),
            ("M — 시장 분석 리포트  (손님용+내부용+HTML, ~30분)", "M"),
        ]:
            ttk.Radiobutton(
                frm_radio, text=label, variable=self.var_mode,
                value=val, command=self._on_mode_change,
            ).pack(anchor="w", pady=2)

        # 모드 인포박스
        self._frm_info = tk.Frame(
            frm_radio, bg=self._INFO_BG,
            highlightbackground=self._INFO_BAR, highlightthickness=2,
        )
        self._frm_info.pack(fill="x", padx=(22, 0), pady=(8, 4))

        _rows = ["수집 방식", "입     력", "출     력", "소요 시간"]
        self._info_lbls: dict[str, tk.Label] = {}
        for i, key in enumerate(_rows):
            tk.Label(self._frm_info, text=key, bg=self._INFO_BG,
                     font=(self._FONT_FAMILY, self._FONT_SMALL, "bold"),
                     fg=self._INFO_KEY_FG,
                     anchor="e", width=9).grid(
                row=i, column=0,
                padx=(10, 4),
                pady=(7 if i == 0 else 3, 7 if i == len(_rows)-1 else 3),
                sticky="e")
            tk.Label(self._frm_info, text="│", bg=self._INFO_BG,
                     fg="#596579",
                     font=(self._FONT_FAMILY, self._FONT_SMALL)).grid(row=i, column=1, sticky="ns")
            lbl = tk.Label(self._frm_info, text="", bg=self._INFO_BG,
                           font=(self._FONT_FAMILY, self._FONT_SMALL),
                           fg=self._INFO_FG,
                           anchor="w", justify="left", wraplength=500)
            lbl.grid(row=i, column=2,
                     padx=(8, 12),
                     pady=(7 if i == 0 else 3, 7 if i == len(_rows)-1 else 3),
                     sticky="w")
            self._info_lbls[key] = lbl

        # ── 1: M 모드 옵션 ────────────────────────────────────────────
        self.frm_m = ttk.LabelFrame(self, text="M 모드 옵션 (시장 분석)")
        self.frm_m.grid(row=1, column=0, sticky="ew", padx=14, pady=4)
        self.frm_m.grid_remove()   # 초기 숨김 — grid_remove는 공간 예약 없음

        _DG = self._MUTED_FG

        def _spin_row(row, label, var, from_, to, inc, fmt="%.0f", desc=""):
            ttk.Label(self.frm_m, text=label, width=16, anchor="w").grid(
                row=row, column=0, sticky="w", padx=(0, 8), pady=4)
            ttk.Spinbox(self.frm_m, from_=from_, to=to, increment=inc,
                        textvariable=var, width=12, format=fmt).grid(
                row=row, column=1, sticky="w", pady=4)
            if desc:
                ttk.Label(self.frm_m, text=desc, foreground=_DG,
                          wraplength=340, justify="left").grid(
                    row=row, column=2, sticky="w", padx=(16, 0), pady=4)

        self.var_beds       = tk.StringVar(value="0")
        self.var_baths      = tk.StringVar(value="0")
        self.var_cleaning   = tk.StringVar(value="80000")
        self.var_monthly    = tk.StringVar(value="0")
        self.var_occ_low    = tk.StringVar(value="0.40")
        self.var_occ_base   = tk.StringVar(value="0.60")
        self.var_occ_high   = tk.StringVar(value="0.70")
        self.var_avg_nights = tk.StringVar(value="2.0")

        _spin_row(0, "침실 수",         self.var_beds,       0, 20,      1,
                  desc="0 = 전체 수집 / 숫자 지정 시 해당 규모만 비교")
        _spin_row(1, "욕실 수",         self.var_baths,      0, 10,      1,
                  desc="0 = 전체 수집")
        _spin_row(2, "청소비  (원/회)", self.var_cleaning,   0, 500000, 10000,
                  desc="손님에게 부과하는 1회 청소비 → 수익 시뮬레이터에 반영")
        _spin_row(3, "월 고정비  (원)", self.var_monthly,    0, 5000000, 50000,
                  desc="임대료·관리비 등 매달 고정 지출 / 0이면 비용 없음")
        _spin_row(4, "평균 숙박일",     self.var_avg_nights, 1, 30, 0.5, "%.1f",
                  desc="예약 1건당 평균 숙박일 (청소비 발생 빈도 계산에 사용)")

        ttk.Label(self.frm_m, text="예약률  하/기/상", width=16, anchor="w").grid(
            row=5, column=0, sticky="w", padx=(0, 8), pady=4)
        frm_occ = ttk.Frame(self.frm_m)
        frm_occ.grid(row=5, column=1, sticky="w", pady=4)
        for var in (self.var_occ_low, self.var_occ_base, self.var_occ_high):
            ttk.Spinbox(frm_occ, from_=0.0, to=1.0, increment=0.05,
                        textvariable=var, width=7, format="%.2f").pack(
                side="left", padx=(0, 6))
        ttk.Label(self.frm_m,
                  text="보수적 / 기준 / 공격적 시나리오 3가지로 수익 비교",
                  foreground=_DG, wraplength=340, justify="left").grid(
            row=5, column=2, sticky="w", padx=(16, 0), pady=4)

        ttk.Label(self.frm_m, text="출력 형식", width=16, anchor="w").grid(
            row=6, column=0, sticky="w", padx=(0, 8), pady=4)
        self.var_output_mode = tk.StringVar(value="both")
        frm_om = ttk.Frame(self.frm_m)
        frm_om.grid(row=6, column=1, columnspan=2, sticky="w", pady=4)
        for txt, val in [
            ("손님+내부+HTML", "both"),
            ("손님용만", "client"),
            ("내부용만", "internal"),
        ]:
            ttk.Radiobutton(frm_om, text=txt, variable=self.var_output_mode,
                            value=val).pack(side="left", padx=(0, 16))

        # ── 2: 수집 설정 ─────────────────────────────────────────────
        frm_adv = ttk.LabelFrame(self, text="수집 설정")
        frm_adv.grid(row=2, column=0, sticky="ew", padx=14, pady=4)

        ttk.Label(frm_adv, text="최대 매물").pack(side="left")
        self.var_max = tk.StringVar(value="200")
        ttk.Spinbox(frm_adv, from_=10, to=9999, increment=10,
                    textvariable=self.var_max, width=7).pack(side="left", padx=(6, 4))
        ttk.Label(frm_adv, text="개").pack(side="left")

        ttk.Separator(frm_adv, orient="vertical").pack(side="left", fill="y", padx=16)

        ttk.Label(frm_adv, text="최대 페이지").pack(side="left")
        self.var_pages = tk.StringVar(value="20")
        ttk.Spinbox(frm_adv, from_=1, to=50, increment=1,
                    textvariable=self.var_pages, width=5).pack(side="left", padx=(6, 4))
        ttk.Label(frm_adv, text="페이지").pack(side="left")
        ttk.Label(frm_adv, text="  (페이지당 ~20개 · Airbnb 한계 ~200개)",
                  foreground=self._MUTED_FG,
                  font=(self._FONT_FAMILY, self._FONT_XSMALL)).pack(side="left")

        # ── 3: 로그창 (weight=1 로 늘어남) ──────────────────────────
        frm_log = ttk.Frame(self)
        frm_log.grid(row=3, column=0, sticky="nsew", padx=14, pady=(4, 0))
        frm_log.grid_rowconfigure(1, weight=1)
        frm_log.grid_columnconfigure(0, weight=1)

        # 툴바: 버튼 왼쪽 + 진행바 중앙 + 유틸 오른쪽
        frm_toolbar = ttk.Frame(frm_log)
        frm_toolbar.grid(row=0, column=0, sticky="ew", pady=(6, 4))
        frm_toolbar.columnconfigure(1, weight=1)

        self.btn_run = ttk.Button(
            frm_toolbar, text="▶  실행",
            bootstyle="success", width=14,
            command=self._run,
        )
        self.btn_run.grid(row=0, column=0, sticky="w")

        # 진행바 (중앙, 늘어남)
        frm_prog = ttk.Frame(frm_toolbar)
        frm_prog.grid(row=0, column=1, sticky="ew", padx=12)
        frm_prog.columnconfigure(0, weight=1)

        self.prog_bar = Floodgauge(
            frm_prog,
            bootstyle="success",
            font=(self._FONT_FAMILY, self._FONT_SMALL),
            mask="{}%",
            maximum=100,
            value=0,
        )
        self.prog_bar.grid(row=0, column=0, sticky="ew")
        self.lbl_step = ttk.Label(frm_prog, text="", foreground=self._MUTED_FG,
                                  font=(self._FONT_FAMILY, self._FONT_XSMALL), anchor="w")
        self.lbl_step.grid(row=1, column=0, sticky="ew")

        # 우측 유틸 버튼
        frm_util = ttk.Frame(frm_toolbar)
        frm_util.grid(row=0, column=2, sticky="e")

        self.btn_open = ttk.Button(
            frm_util, text="📂  결과 폴더",
            bootstyle="info-outline", width=16,
            command=self._open_output, state="disabled",
        )
        self.btn_open.pack(side="left", padx=(0, 6))
        ttk.Button(
            frm_util, text="📋  로그 복사",
            bootstyle="secondary-outline", width=14,
            command=self._copy_log,
        ).pack(side="left", padx=(0, 6))
        ttk.Button(
            frm_util, text="🗑  로그 지우기",
            bootstyle="secondary-outline", width=14,
            command=self._clear_log,
        ).pack(side="left")

        # 로그 텍스트
        frm_logbody = ttk.LabelFrame(frm_log, text="실행 로그")
        frm_logbody.grid(row=1, column=0, sticky="nsew", pady=(0, 10))
        frm_logbody.grid_rowconfigure(0, weight=1)
        frm_logbody.grid_columnconfigure(0, weight=1)

        self.log_text = tk.Text(
            frm_logbody, state="disabled", wrap="word",
            bg=self._LOG_BG, fg=self._LOG_FG,
            font=("Consolas", 10),
            relief="flat", borderwidth=0,
            insertbackground=self._LOG_FG,
            selectbackground="#2f5f98",
        )
        self.log_text.tag_configure("hdr",  foreground="#20d3a2", font=("Consolas", 10, "bold"))
        self.log_text.tag_configure("file", foreground="#58a6ff")
        self.log_text.tag_configure("dir",  foreground="#f2b84b")
        self.log_text.tag_configure("err",  foreground="#ff6b6b")
        self.log_text.tag_configure("ok",   foreground="#20d3a2")

        sb = ttk.Scrollbar(frm_logbody, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=sb.set)
        sb.grid(row=0, column=1, sticky="ns")
        self.log_text.grid(row=0, column=0, sticky="nsew")

        # 초기 상태
        self._on_mode_change()

    # ── 모드 설명 데이터 ────────────────────────────────────────────────
    _MODE_INFO: dict[str, dict[str, str]] = {
        "A": {
            "수집 방식": "가격·위치·평점·숙소유형 등  Airbnb 목록 페이지 데이터만 빠르게 수집",
            "입     력": "지역명  /  체크인·체크아웃 날짜  /  최대 매물 수",
            "출     력": "output/{날짜}_{지역}/  {지역}_기본_{날짜}.xlsx   (14컬럼)",
            "소요 시간": "약 30~60초",
        },
        "B": {
            "수집 방식": "A 모드 수집 후  각 숙소 상세 페이지를 직접 방문 →  소개글·편의시설·호스트 정보 추가",
            "입     력": "지역명  /  체크인·체크아웃 날짜  /  최대 매물 수",
            "출     력": "output/{날짜}_{지역}/  {지역}_상세_{날짜}.xlsx   (22컬럼)",
            "소요 시간": "숙소당 약 2초 × 최대 80개  ≈  3~5분",
        },
        "M": {
            "수집 방식": "체크인 기준 주의 평일(월→화) · 금→토 · 토→일  3개 날짜창 자동 수집 → 통계 분석",
            "입     력": "지역명  /  체크인(기준 날짜 — 체크아웃은 무시됨)  /  아래 M 모드 옵션 전체",
            "출     력": "output/{날짜}_{지역}/  손님용.xlsx  +  내부용.xlsx  +  .html",
            "소요 시간": "약 20~30분",
        },
    }

    # ── 모드 토글 ──────────────────────────────────────────────────────
    def _on_mode_change(self) -> None:
        mode = self.var_mode.get()
        info = self._MODE_INFO.get(mode, {})
        for key, lbl in self._info_lbls.items():
            lbl.configure(text=info.get(key, ""))
        if mode == "M":
            self.frm_m.grid()       # 공간 복원
        else:
            self.frm_m.grid_remove()   # 공간 제거 (창 높이 변동 없이 숨김)

    # ── 로그 폴링 ──────────────────────────────────────────────────────
    def _poll_log(self) -> None:
        try:
            while True:
                item = self._log_q.get_nowait()
                if isinstance(item, tuple) and item[0] == "PROGRESS":
                    _, n, total, desc = item
                    pct = int(n / total * 100)
                    self.prog_bar.configure(value=pct)
                    self.lbl_step.configure(text=f"[{n}/{total}]  {desc}")
                elif isinstance(item, tuple):
                    text, tag = item
                    self.log_text.configure(state="normal")
                    self.log_text.insert("end", text, tag)
                    self.log_text.see("end")
                    self.log_text.configure(state="disabled")
                else:
                    self.log_text.configure(state="normal")
                    self.log_text.insert("end", item)
                    self.log_text.see("end")
                    self.log_text.configure(state="disabled")
        except queue.Empty:
            pass
        self.after(80, self._poll_log)

    def _log(self, msg: str, tag: str | None = None) -> None:
        self._log_q.put((msg + "\n", tag))

    def _copy_log(self) -> None:
        content = self.log_text.get("1.0", "end").strip()
        if content:
            self.clipboard_clear()
            self.clipboard_append(content)

    def _clear_log(self) -> None:
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    # ── 입력 검증 ──────────────────────────────────────────────────────
    def _collect_inputs(self) -> dict | None:
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
        inputs = self._collect_inputs()
        if inputs is None:
            return

        self._running = True
        self._last_out_dir = None
        self.btn_run.configure(state="disabled", text="⏳ 실행 중...", bootstyle="secondary")
        self.btn_open.configure(state="disabled")
        self._clear_log()
        self.prog_bar.configure(value=0)
        self.lbl_step.configure(text="준비 중...")

        threading.Thread(target=self._worker, args=(inputs,), daemon=True).start()

    # ── 워커 ──────────────────────────────────────────────────────────
    def _worker(self, inp: dict) -> None:
        old_out, old_err = sys.stdout, sys.stderr
        sys.stdout = sys.stderr = _QueueWriter(self._log_q)

        sys.path.insert(0, str(PROJECT_DIR))

        import airbnb_fetch as _af
        _orig_crawl = _af.AirbnbClient.crawl

        try:
            mode     = inp["mode"]
            region   = inp["region"]
            checkin  = inp["checkin"]
            checkout = inp["checkout"]
            max_res  = inp["max_res"]
            max_pgs  = inp["max_pgs"]

            self._log(f"=== 실행 시작: 모드={mode}, 지역={region} ===")
            self._log(f"    체크인={checkin}, 체크아웃={checkout}")
            self._log(f"    최대 매물={max_res}개, 최대 페이지={max_pgs}페이지")
            self._log("")

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

            # ── 완료 후 파일 목록 ────────────────────────────────────
            folder_ts = datetime.now().strftime("%y%m%d_%H%M")
            out_dir   = PROJECT_DIR / "output" / f"{folder_ts}_{region}"
            self._last_out_dir = out_dir

            self._log("")
            self._log("═" * 56, "hdr")
            self._log("  생성된 파일 목록", "hdr")
            self._log("═" * 56, "hdr")
            self._log(f"  📁 output/{folder_ts}_{region}/", "dir")
            self._log("─" * 56, "hdr")

            if out_dir.exists():
                import time as _time
                now    = _time.time()
                recent = [f for f in sorted(out_dir.iterdir())
                          if f.is_file() and now - f.stat().st_mtime < 30]
                files  = recent if recent else sorted(
                    [f for f in out_dir.iterdir() if f.is_file()])
                if files:
                    for f in files:
                        size_kb = f.stat().st_size / 1024
                        icon = "🌐" if f.suffix == ".html" else "📊"
                        self._log(f"  {icon} {f.name:<46}  {size_kb:>6.0f} KB", "file")
                else:
                    self._log("  (생성된 파일 없음)", "err")
            else:
                self._log("  (출력 폴더 없음 — 수집 결과 0개일 수 있음)", "err")

            self._log("═" * 56, "hdr")
            self._log("")
            self._log("✅ 완료", "ok")
            self.after(0, lambda: self.prog_bar.configure(value=100))
            self.after(0, lambda: self.lbl_step.configure(text="완료"))
            self.after(0, lambda: self.btn_open.configure(state="normal"))
            if out_dir.exists():
                self.after(600, lambda d=out_dir: os.startfile(str(d)))

        except Exception as exc:
            import traceback
            if isinstance(exc, PermissionError) or "Permission denied" in str(exc):
                import os as _os
                fname = _os.path.basename(str(exc).split("'")[-2] if "'" in str(exc) else "파일")
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
                state="normal", text="▶  실행", bootstyle="success"))

    # ── 결과 폴더 열기 ──────────────────────────────────────────────────
    def _open_output(self) -> None:
        folder = self._last_out_dir or (PROJECT_DIR / "output")
        folder.mkdir(parents=True, exist_ok=True)
        os.startfile(folder)


# ── 진입점 ────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = App()
    app.mainloop()
