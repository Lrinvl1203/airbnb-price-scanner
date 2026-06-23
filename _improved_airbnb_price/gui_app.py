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
from tkcalendar import DateEntry

PROJECT_DIR = (
    Path(sys.executable).parent if getattr(sys, "frozen", False)
    else Path(__file__).parent
)

_STEP_RE = re.compile(r"\[\s*(\d+)/(\d+)\]\s*(.*)")
_VERBOSE_RE = re.compile(r"^\d+\s{2,}")


def _is_verbose(s: str) -> bool:
    t = s.strip()
    if not t:
        return True
    if _VERBOSE_RE.match(t):
        return True
    if re.match(r"^[-=]{4,}$", t):
        return True
    if "번호" in t and "숙소명" in t:
        return True
    return False


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


class _FixedDateEntry(DateEntry):
    """월/년 이동 화살표 클릭 시 달력이 닫히는 tkcalendar 버그 보정."""

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


class App(ttk.Window):
    FONT = "맑은 고딕"
    BG = "#f7f7f8"
    SIDEBAR = "#f1f1f3"
    CARD = "#ffffff"
    CARD_ALT = "#fafafa"
    TEXT = "#111827"
    MUTED = "#6b7280"
    FAINT = "#9ca3af"
    BORDER = "#e5e7eb"
    BORDER_DARK = "#d1d5db"
    ACCENT = "#10a37f"
    ACCENT_DARK = "#0d8f70"
    LOG_BG = "#0f172a"
    LOG_FG = "#dbeafe"
    LOG_MUTED = "#93a4b8"
    ERROR = "#dc2626"

    MODE_INFO: dict[str, dict[str, str]] = {
        "A": {
            "title": "빠른 엑셀",
            "subtitle": "가격, 위치, 평점 중심",
            "runtime": "약 30~60초",
            "output": "기본 Excel",
            "body": "목록 페이지 데이터만 빠르게 수집합니다. 시세를 빠르게 훑어볼 때 사용합니다.",
        },
        "B": {
            "title": "상세 엑셀",
            "subtitle": "소개글, 편의시설 추가",
            "runtime": "약 3~5분",
            "output": "상세 Excel",
            "body": "각 숙소 상세 페이지까지 방문합니다. 숙소별 설명과 편의시설 비교가 필요할 때 사용합니다.",
        },
        "M": {
            "title": "시장분석 리포트",
            "subtitle": "추천가, 수익, HTML 리포트",
            "runtime": "약 20~30분",
            "output": "손님용/내부용/HTML",
            "body": "기준 날짜 주간의 평일, 금-토, 토-일을 비교해 시장성과 추천가를 계산합니다.",
        },
    }

    def __init__(self) -> None:
        super().__init__(themename="litera")
        self.title("Airbnb Market Studio")
        self.resizable(True, True)
        self.minsize(1080, 650)
        self.configure(bg=self.BG)

        self._log_q: queue.Queue[str | tuple] = queue.Queue()
        self._running = False
        self._last_out_dir: Path | None = None
        self._run_started_at: datetime | None = None
        self._mode_cards: dict[str, tk.Frame] = {}
        self._mode_title_labels: dict[str, tk.Label] = {}
        self._mode_subtitle_labels: dict[str, tk.Label] = {}
        self._checkout_visible = True

        self._configure_fonts()
        self._build_ui()
        self._poll_log()

        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w = min(1240, max(1080, sw - 80))
        h = min(820, max(650, sh - 150))
        self.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}")

    def _configure_fonts(self) -> None:
        for name in (
            "TkDefaultFont",
            "TkTextFont",
            "TkMenuFont",
            "TkCaptionFont",
            "TkSmallCaptionFont",
            "TkIconFont",
        ):
            try:
                tkfont.nametofont(name).configure(family=self.FONT, size=10)
            except tk.TclError:
                pass
        try:
            tkfont.nametofont("TkHeadingFont").configure(family=self.FONT, size=11, weight="bold")
        except tk.TclError:
            pass

        style = ttk.Style()
        style.configure("TLabel", font=(self.FONT, 10))
        style.configure("TEntry", font=(self.FONT, 11), padding=(8, 6))
        style.configure("TSpinbox", font=(self.FONT, 10), padding=(8, 4))
        style.configure("TRadiobutton", font=(self.FONT, 10))
        style.configure("Studio.TButton", font=(self.FONT, 10, "bold"), padding=(14, 9))
        style.configure("Soft.TButton", font=(self.FONT, 10), padding=(12, 8))
        style.configure("Studio.Horizontal.TProgressbar", thickness=7)
        style.configure(
            "DateEntry",
            fieldbackground="#ffffff",
            foreground=self.TEXT,
            insertcolor=self.TEXT,
            arrowcolor=self.TEXT,
            selectforeground="#ffffff",
            selectbackground=self.ACCENT,
            font=(self.FONT, 10),
        )

    def _build_ui(self) -> None:
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self._build_sidebar()
        self._build_workspace()
        self._on_mode_change()

    def _build_sidebar(self) -> None:
        sidebar = tk.Frame(self, bg=self.SIDEBAR, width=284, highlightthickness=0)
        sidebar.grid(row=0, column=0, sticky="ns")
        sidebar.grid_propagate(False)
        sidebar.grid_rowconfigure(4, weight=1)

        tk.Label(
            sidebar,
            text="Airbnb\nMarket Studio",
            bg=self.SIDEBAR,
            fg=self.TEXT,
            font=(self.FONT, 18, "bold"),
            justify="left",
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=24, pady=(20, 4))
        tk.Label(
            sidebar,
            text="수집부터 가격 추천까지 한 번에 정리합니다.",
            bg=self.SIDEBAR,
            fg=self.MUTED,
            font=(self.FONT, 9),
            justify="left",
            wraplength=220,
            anchor="w",
        ).grid(row=1, column=0, sticky="ew", padx=24, pady=(0, 16))

        self.var_mode = tk.StringVar(value="A")
        for i, mode in enumerate(("A", "B", "M"), start=2):
            self._create_mode_card(sidebar, mode).grid(
                row=i,
                column=0,
                sticky="ew",
                padx=16,
                pady=(0, 8),
            )

        footer = tk.Frame(sidebar, bg=self.SIDEBAR)
        footer.grid(row=5, column=0, sticky="ew", padx=20, pady=16)
        tk.Label(
            footer,
            text="상태",
            bg=self.SIDEBAR,
            fg=self.FAINT,
            font=(self.FONT, 9, "bold"),
            anchor="w",
        ).pack(fill="x")
        self.lbl_sidebar_status = tk.Label(
            footer,
            text="실행 대기 중",
            bg=self.SIDEBAR,
            fg=self.MUTED,
            font=(self.FONT, 10),
            anchor="w",
        )
        self.lbl_sidebar_status.pack(fill="x", pady=(4, 12))
        ttk.Button(
            footer,
            text="결과 폴더 열기",
            command=self._open_output,
            style="Soft.TButton",
            bootstyle="secondary-outline",
        ).pack(fill="x")

    def _create_mode_card(self, parent: tk.Widget, mode: str) -> tk.Frame:
        info = self.MODE_INFO[mode]
        card = tk.Frame(
            parent,
            bg=self.CARD,
            highlightthickness=1,
            highlightbackground=self.BORDER,
            cursor="hand2",
        )
        card.columnconfigure(1, weight=1)

        badge = tk.Label(
            card,
            text=mode,
            bg="#eef2ff",
            fg="#374151",
            font=(self.FONT, 10, "bold"),
            width=3,
        )
        badge.grid(row=0, column=0, rowspan=2, sticky="n", padx=(14, 10), pady=14)
        title = tk.Label(
            card,
            text=info["title"],
            bg=self.CARD,
            fg=self.TEXT,
            font=(self.FONT, 11, "bold"),
            anchor="w",
        )
        title.grid(row=0, column=1, sticky="ew", padx=(0, 14), pady=(12, 0))
        subtitle = tk.Label(
            card,
            text=info["subtitle"],
            bg=self.CARD,
            fg=self.MUTED,
            font=(self.FONT, 9),
            anchor="w",
        )
        subtitle.grid(row=1, column=1, sticky="ew", padx=(0, 14), pady=(0, 12))

        for widget in (card, badge, title, subtitle):
            widget.bind("<Button-1>", lambda _e, m=mode: self._select_mode(m))

        self._mode_cards[mode] = card
        self._mode_title_labels[mode] = title
        self._mode_subtitle_labels[mode] = subtitle
        return card

    def _build_workspace(self) -> None:
        workspace = tk.Frame(self, bg=self.BG)
        workspace.grid(row=0, column=1, sticky="nsew")
        workspace.grid_columnconfigure(0, weight=1)
        workspace.grid_rowconfigure(4, weight=1)

        self._build_header(workspace)
        self._build_primary_cards(workspace)
        self._build_market_options(workspace)
        self._build_run_panel(workspace)
        self._build_log_panel(workspace)

    def _build_header(self, parent: tk.Widget) -> None:
        header = tk.Frame(parent, bg=self.BG)
        header.grid(row=0, column=0, sticky="ew", padx=24, pady=(12, 6))
        header.columnconfigure(0, weight=1)

        tk.Label(
            header,
            text="분석 작업 만들기",
            bg=self.BG,
            fg=self.TEXT,
            font=(self.FONT, 20, "bold"),
            anchor="w",
        ).grid(row=0, column=0, sticky="w")
        tk.Label(
            header,
            text="지역과 날짜를 입력하고 실행하면 Excel 또는 리포트를 생성합니다.",
            bg=self.BG,
            fg=self.MUTED,
            font=(self.FONT, 10),
            anchor="w",
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))

        self.lbl_mode_pill = tk.Label(
            header,
            text="빠른 엑셀",
            bg="#e8f8f3",
            fg=self.ACCENT_DARK,
            font=(self.FONT, 10, "bold"),
            padx=14,
            pady=7,
        )
        self.lbl_mode_pill.grid(row=0, column=1, rowspan=2, sticky="e")

    def _build_primary_cards(self, parent: tk.Widget) -> None:
        row = tk.Frame(parent, bg=self.BG)
        row.grid(row=1, column=0, sticky="ew", padx=24)
        row.grid_columnconfigure(0, weight=3)
        row.grid_columnconfigure(1, weight=2)

        request = self._card(row)
        request.grid(row=0, column=0, sticky="nsew", padx=(0, 14))
        request.grid_columnconfigure(1, weight=1)

        self._card_title(request, "분석 요청", "지역과 날짜를 지정합니다.").grid(
            row=0, column=0, columnspan=3, sticky="ew", padx=20, pady=(12, 8)
        )

        self.var_region = tk.StringVar(value="홍대")
        self._field_label(request, "지역").grid(row=1, column=0, sticky="w", padx=(22, 14), pady=8)
        ttk.Entry(request, textvariable=self.var_region).grid(
            row=1, column=1, columnspan=2, sticky="ew", padx=(0, 20), pady=6
        )

        self._field_label(request, "날짜").grid(row=2, column=0, sticky="w", padx=(20, 14), pady=6)
        dates = tk.Frame(request, bg=self.CARD)
        dates.grid(row=2, column=1, columnspan=2, sticky="w", padx=(0, 20), pady=6)
        ci_default = date.today() + timedelta(days=7)
        co_default = date.today() + timedelta(days=8)
        self.ent_checkin = _FixedDateEntry(
            dates,
            date_pattern="yyyy-mm-dd",
            width=13,
            year=ci_default.year,
            month=ci_default.month,
            day=ci_default.day,
            showweeknumbers=False,
            firstweekday="sunday",
        )
        self.ent_checkin.pack(side="left")
        self.lbl_date_between = tk.Label(
            dates,
            text="to",
            bg=self.CARD,
            fg=self.FAINT,
            font=(self.FONT, 9, "bold"),
            padx=10,
        )
        self.lbl_date_between.pack(side="left")
        self.ent_checkout = _FixedDateEntry(
            dates,
            date_pattern="yyyy-mm-dd",
            width=13,
            year=co_default.year,
            month=co_default.month,
            day=co_default.day,
            showweeknumbers=False,
            firstweekday="sunday",
        )
        self.ent_checkout.pack(side="left")
        self.lbl_date_hint = tk.Label(
            request,
            text="M 모드는 체크인 날짜만 기준일로 사용합니다.",
            bg=self.CARD,
            fg=self.FAINT,
            font=(self.FONT, 9),
            anchor="w",
        )
        self.lbl_date_hint.grid(row=3, column=1, columnspan=2, sticky="ew", padx=(0, 20), pady=(0, 12))

        scope = self._card(row)
        scope.grid(row=0, column=1, sticky="nsew")
        scope.grid_columnconfigure(1, weight=1)

        self._card_title(scope, "수집 범위", "분석에 사용할 숙소 수만 정하면 됩니다.").grid(
            row=0, column=0, columnspan=2, sticky="ew", padx=20, pady=(12, 8)
        )

        self.var_max = tk.StringVar(value="200")
        self._field_label(scope, "분석 매물 수").grid(row=1, column=0, sticky="w", padx=(20, 14), pady=6)
        ttk.Spinbox(scope, from_=1, to=9999, increment=10, textvariable=self.var_max, width=10).grid(
            row=1, column=1, sticky="ew", padx=(0, 20), pady=6
        )
        tk.Label(
            scope,
            text="입력한 개수만큼 모이면 자동으로 멈춥니다.",
            bg=self.CARD,
            fg=self.FAINT,
            font=(self.FONT, 9),
            anchor="w",
        ).grid(
            row=2, column=0, columnspan=2, sticky="ew", padx=20, pady=(0, 6)
        )

        self.var_output_mode = tk.StringVar(value="both")
        self.output_frame = tk.Frame(scope, bg=self.CARD)
        self.output_frame.grid(row=3, column=0, columnspan=2, sticky="ew", padx=20, pady=(6, 14))
        tk.Label(
            self.output_frame,
            text="출력",
            bg=self.CARD,
            fg=self.MUTED,
            font=(self.FONT, 9, "bold"),
            anchor="w",
        ).pack(side="left", padx=(0, 10))
        for txt, val in [
            ("손님+내부+HTML", "both"),
            ("손님용만", "client"),
            ("내부용만", "internal"),
        ]:
            ttk.Radiobutton(
                self.output_frame,
                text=txt,
                variable=self.var_output_mode,
                value=val,
            ).pack(side="left", padx=(0, 10))

    def _build_market_options(self, parent: tk.Widget) -> None:
        self.frm_m = self._card(parent)
        self.frm_m.grid(row=2, column=0, sticky="ew", padx=24, pady=(10, 0))
        for i in range(8):
            self.frm_m.grid_columnconfigure(i, weight=1)

        self._card_title(
            self.frm_m,
            "시장분석 옵션",
            "추천가와 수익 시뮬레이션에 들어가는 가정값입니다.",
        ).grid(row=0, column=0, columnspan=8, sticky="ew", padx=22, pady=(14, 10))

        self.var_beds = tk.StringVar(value="0")
        self.var_baths = tk.StringVar(value="0")
        self.var_cleaning = tk.StringVar(value="80000")
        self.var_monthly = tk.StringVar(value="0")
        self.var_occ_low = tk.StringVar(value="0.40")
        self.var_occ_base = tk.StringVar(value="0.60")
        self.var_occ_high = tk.StringVar(value="0.70")
        self.var_avg_nights = tk.StringVar(value="2.0")

        fields = [
            ("침실", self.var_beds, 0, 20, 1, "%.0f"),
            ("욕실", self.var_baths, 0, 10, 1, "%.0f"),
            ("청소비", self.var_cleaning, 0, 500000, 10000, "%.0f"),
            ("월 고정비", self.var_monthly, 0, 5000000, 50000, "%.0f"),
            ("평균 숙박일", self.var_avg_nights, 1, 30, 0.5, "%.1f"),
            ("예약률 하", self.var_occ_low, 0.0, 1.0, 0.05, "%.2f"),
            ("예약률 기준", self.var_occ_base, 0.0, 1.0, 0.05, "%.2f"),
            ("예약률 상", self.var_occ_high, 0.0, 1.0, 0.05, "%.2f"),
        ]
        for i, (label, var, from_, to, inc, fmt) in enumerate(fields):
            cell = tk.Frame(self.frm_m, bg=self.CARD)
            cell.grid(row=1, column=i, sticky="ew", padx=(22 if i == 0 else 4, 22 if i == len(fields) - 1 else 4), pady=(0, 14))
            tk.Label(
                cell,
                text=label,
                bg=self.CARD,
                fg=self.MUTED,
                font=(self.FONT, 9, "bold"),
                anchor="w",
            ).pack(fill="x", pady=(0, 5))
            ttk.Spinbox(
                cell,
                from_=from_,
                to=to,
                increment=inc,
                textvariable=var,
                width=8,
                format=fmt,
            ).pack(fill="x")

    def _build_run_panel(self, parent: tk.Widget) -> None:
        panel = self._card(parent)
        panel.grid(row=3, column=0, sticky="ew", padx=24, pady=(10, 0))
        panel.grid_columnconfigure(0, weight=1)

        self.lbl_status = tk.Label(
            panel,
            text="실행 대기 중",
            bg=self.CARD,
            fg=self.TEXT,
            font=(self.FONT, 11, "bold"),
            anchor="w",
        )
        self.lbl_status.grid(row=0, column=0, sticky="ew", padx=22, pady=(14, 4))
        self.lbl_step = tk.Label(
            panel,
            text="설정을 확인한 뒤 분석 실행을 누르세요.",
            bg=self.CARD,
            fg=self.MUTED,
            font=(self.FONT, 9),
            anchor="w",
        )
        self.lbl_step.grid(row=1, column=0, sticky="ew", padx=22, pady=(0, 12))

        self.prog_bar = ttk.Progressbar(
            panel,
            mode="determinate",
            maximum=100,
            value=0,
            style="Studio.Horizontal.TProgressbar",
            bootstyle="success",
        )
        self.prog_bar.grid(row=2, column=0, sticky="ew", padx=22, pady=(0, 14))

        actions = tk.Frame(panel, bg=self.CARD)
        actions.grid(row=0, column=1, rowspan=3, sticky="e", padx=22, pady=14)
        self.btn_run = ttk.Button(
            actions,
            text="분석 실행",
            command=self._run,
            style="Studio.TButton",
            bootstyle="success",
            width=14,
        )
        self.btn_run.pack(side="left", padx=(0, 8))
        self.btn_open = ttk.Button(
            actions,
            text="결과 폴더",
            command=self._open_output,
            style="Soft.TButton",
            bootstyle="secondary-outline",
            state="disabled",
            width=12,
        )
        self.btn_open.pack(side="left", padx=(0, 8))
        ttk.Button(
            actions,
            text="로그 복사",
            command=self._copy_log,
            style="Soft.TButton",
            bootstyle="secondary-outline",
            width=11,
        ).pack(side="left")

    def _build_log_panel(self, parent: tk.Widget) -> None:
        card = self._card(parent)
        card.grid(row=4, column=0, sticky="nsew", padx=24, pady=(10, 18))
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(1, weight=1)

        header = tk.Frame(card, bg=self.CARD)
        header.grid(row=0, column=0, sticky="ew", padx=22, pady=(12, 8))
        header.columnconfigure(0, weight=1)
        tk.Label(
            header,
            text="실행 로그",
            bg=self.CARD,
            fg=self.TEXT,
            font=(self.FONT, 12, "bold"),
            anchor="w",
        ).grid(row=0, column=0, sticky="w")
        ttk.Button(
            header,
            text="지우기",
            command=self._clear_log,
            style="Soft.TButton",
            bootstyle="secondary-outline",
            width=8,
        ).grid(row=0, column=1, sticky="e")

        log_wrap = tk.Frame(card, bg=self.LOG_BG, highlightthickness=0)
        log_wrap.grid(row=1, column=0, sticky="nsew", padx=22, pady=(0, 14))
        log_wrap.grid_columnconfigure(0, weight=1)
        log_wrap.grid_rowconfigure(0, weight=1)

        self.log_text = tk.Text(
            log_wrap,
            state="disabled",
            wrap="word",
            height=7,
            bg=self.LOG_BG,
            fg=self.LOG_FG,
            font=("Consolas", 10),
            relief="flat",
            borderwidth=0,
            padx=16,
            pady=10,
            insertbackground=self.LOG_FG,
            selectbackground="#334155",
        )
        self.log_text.tag_configure("hdr", foreground="#34d399", font=("Consolas", 10, "bold"))
        self.log_text.tag_configure("file", foreground="#93c5fd")
        self.log_text.tag_configure("dir", foreground="#fbbf24")
        self.log_text.tag_configure("err", foreground="#f87171")
        self.log_text.tag_configure("ok", foreground="#34d399")
        self.log_text.grid(row=0, column=0, sticky="nsew")

        scroll = ttk.Scrollbar(log_wrap, command=self.log_text.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scroll.set)

    def _card(self, parent: tk.Widget) -> tk.Frame:
        return tk.Frame(
            parent,
            bg=self.CARD,
            highlightthickness=1,
            highlightbackground=self.BORDER,
            highlightcolor=self.BORDER,
        )

    def _card_title(self, parent: tk.Widget, title: str, subtitle: str) -> tk.Frame:
        frame = tk.Frame(parent, bg=self.CARD)
        tk.Label(
            frame,
            text=title,
            bg=self.CARD,
            fg=self.TEXT,
            font=(self.FONT, 13, "bold"),
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            frame,
            text=subtitle,
            bg=self.CARD,
            fg=self.MUTED,
            font=(self.FONT, 9),
            anchor="w",
        ).pack(fill="x", pady=(3, 0))
        return frame

    def _field_label(self, parent: tk.Widget, text: str) -> tk.Label:
        return tk.Label(
            parent,
            text=text,
            bg=self.CARD,
            fg=self.MUTED,
            font=(self.FONT, 9, "bold"),
            anchor="w",
        )

    def _select_mode(self, mode: str) -> None:
        if self._running:
            return
        self.var_mode.set(mode)
        self._on_mode_change()

    def _on_mode_change(self) -> None:
        mode = self.var_mode.get()
        info = self.MODE_INFO[mode]
        self.lbl_mode_pill.configure(text=info["title"])

        for key, card in self._mode_cards.items():
            selected = key == mode
            bg = "#ffffff" if selected else self.SIDEBAR
            border = self.ACCENT if selected else self.BORDER
            fg_title = self.TEXT if selected else "#374151"
            fg_sub = self.MUTED if selected else self.FAINT
            card.configure(bg=bg, highlightbackground=border)
            self._mode_title_labels[key].configure(bg=bg, fg=fg_title)
            self._mode_subtitle_labels[key].configure(bg=bg, fg=fg_sub)
            for child in card.winfo_children():
                if isinstance(child, tk.Label) and child not in (
                    self._mode_title_labels[key],
                    self._mode_subtitle_labels[key],
                ):
                    child.configure(bg="#e8f8f3" if selected else "#eef2ff", fg=self.ACCENT_DARK if selected else "#374151")

        self.lbl_status.configure(text=f"{info['title']} 준비")
        self.lbl_step.configure(text=f"{info['body']} 예상 소요 시간: {info['runtime']}.")

        if mode == "M":
            self.frm_m.grid()
            self.output_frame.grid()
            self._hide_checkout()
            self.lbl_date_hint.configure(text="시장분석 리포트는 체크인 날짜가 속한 주를 기준으로 계산합니다.")
        else:
            self.frm_m.grid_remove()
            self.output_frame.grid_remove()
            self._show_checkout()
            self.lbl_date_hint.configure(text="체크인과 체크아웃 사이의 숙박 요금을 수집합니다.")

    def _hide_checkout(self) -> None:
        if not self._checkout_visible:
            return
        self.lbl_date_between.pack_forget()
        self.ent_checkout.pack_forget()
        self._checkout_visible = False

    def _show_checkout(self) -> None:
        if self._checkout_visible:
            return
        self.lbl_date_between.pack(side="left")
        self.ent_checkout.pack(side="left")
        self.ent_checkout.configure(state="normal")
        self._checkout_visible = True

    def _poll_log(self) -> None:
        try:
            while True:
                item = self._log_q.get_nowait()
                if isinstance(item, tuple) and item[0] == "PROGRESS":
                    _, n, total, desc = item
                    pct = int(n / total * 100)
                    self.prog_bar.configure(value=pct)
                    self.lbl_status.configure(text=f"진행 중 {pct}%")
                    self.lbl_sidebar_status.configure(text=f"진행 중 {pct}%")
                    self.lbl_step.configure(text=f"[{n}/{total}] {desc}")
                elif isinstance(item, tuple):
                    text, tag = item
                    self._write_log(text, tag)
                else:
                    self._write_log(item, None)
        except queue.Empty:
            pass
        self.after(80, self._poll_log)

    def _write_log(self, text: str, tag: str | None = None) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert("end", text, tag)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

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

    def _auto_page_limit(self, max_results: int) -> int:
        return min(50, max(1, (max_results + 9) // 10 + 2))

    def _collect_inputs(self) -> dict | None:
        region = self.var_region.get().strip()
        if not region:
            messagebox.showerror("입력 오류", "지역을 입력하세요.")
            return None
        mode = self.var_mode.get()
        try:
            ci: date = self.ent_checkin.get_date()
            co: date = self.ent_checkout.get_date() if mode != "M" else ci + timedelta(days=1)
        except Exception:
            messagebox.showerror("입력 오류", "날짜 형식을 확인하세요.")
            return None

        if mode != "M" and co <= ci:
            messagebox.showerror("입력 오류", "체크아웃은 체크인 이후여야 합니다.")
            return None

        try:
            max_res = int(self.var_max.get() or 200)
        except ValueError:
            messagebox.showerror("입력 오류", "분석 매물 수는 정수여야 합니다.")
            return None
        if max_res < 1:
            messagebox.showerror("입력 오류", "분석 매물 수는 1개 이상이어야 합니다.")
            return None
        max_pgs = self._auto_page_limit(max_res)

        try:
            beds_raw = self.var_beds.get().strip()
            baths_raw = self.var_baths.get().strip()
            beds = int(beds_raw) if beds_raw and beds_raw != "0" else None
            baths = float(baths_raw) if baths_raw and baths_raw != "0" else None
            cleaning = int(float(self.var_cleaning.get() or 80000))
            monthly = int(float(self.var_monthly.get() or 0))
            occ_low = float(self.var_occ_low.get() or 0.40)
            occ_base = float(self.var_occ_base.get() or 0.60)
            occ_high = float(self.var_occ_high.get() or 0.70)
            avg_nights = float(self.var_avg_nights.get() or 2.0)
        except ValueError:
            messagebox.showerror("입력 오류", "시장분석 옵션에는 숫자만 입력하세요.")
            return None
        if not (0 <= occ_low <= occ_base <= occ_high <= 1):
            messagebox.showerror("입력 오류", "예약률은 0~1 사이이며 하/기준/상 순서로 커야 합니다.")
            return None
        if avg_nights <= 0:
            messagebox.showerror("입력 오류", "평균 숙박일은 0보다 커야 합니다.")
            return None

        return {
            "mode": mode,
            "region": region,
            "checkin": ci.strftime("%Y-%m-%d"),
            "checkout": co.strftime("%Y-%m-%d"),
            "max_res": max_res,
            "max_pgs": max_pgs,
            "beds": beds,
            "baths": baths,
            "cleaning": cleaning,
            "monthly": monthly,
            "occ_low": occ_low,
            "occ_base": occ_base,
            "occ_high": occ_high,
            "avg_nights": avg_nights,
            "output_mode": self.var_output_mode.get(),
        }

    def _run(self) -> None:
        if self._running:
            return
        inputs = self._collect_inputs()
        if inputs is None:
            return

        self._running = True
        self._last_out_dir = None
        self._run_started_at = datetime.now()
        self.btn_run.configure(state="disabled", text="실행 중")
        self.btn_open.configure(state="disabled")
        self.prog_bar.configure(value=0)
        self.lbl_status.configure(text="실행 준비 중")
        self.lbl_sidebar_status.configure(text="실행 준비 중")
        self.lbl_step.configure(text="Airbnb 데이터를 요청할 준비를 하고 있습니다.")
        self._clear_log()
        threading.Thread(target=self._worker, args=(inputs,), daemon=True).start()

    def _worker(self, inp: dict) -> None:
        old_out, old_err = sys.stdout, sys.stderr
        sys.stdout = sys.stderr = _QueueWriter(self._log_q)
        sys.path.insert(0, str(PROJECT_DIR))

        import airbnb_fetch as _af

        _orig_crawl = _af.AirbnbClient.crawl
        out_dir: Path | None = None

        try:
            mode = inp["mode"]
            region = inp["region"]
            checkin = inp["checkin"]
            checkout = inp["checkout"]
            max_res = inp["max_res"]
            max_pgs = inp["max_pgs"]

            self._log(f"Airbnb Market Studio")
            self._log(f"mode={mode}, region={region}")
            self._log(f"checkin={checkin}, checkout={checkout}")
            self._log(f"target_listings={max_res}")
            self._log("")

            def _patched_crawl(self_c, query, ci, co, geo=None, max_results=60, max_pages=4):
                return _orig_crawl(
                    self_c,
                    query,
                    ci,
                    co,
                    geo=geo,
                    max_results=max_res,
                    max_pages=max_pgs,
                )

            _af.AirbnbClient.crawl = _patched_crawl

            if mode == "A":
                import export_excel as module
                importlib.reload(module)
                out_dir = module.run(region, checkin, checkout, max_res, max_pgs)
            elif mode == "B":
                import export_excel_detail as module
                importlib.reload(module)
                out_dir = module.run(region, checkin, checkout, max_res, max_pgs)
            elif mode == "M":
                import market_report as module
                importlib.reload(module)
                kwargs: dict = {
                    "checkin": inp["checkin"],
                    "cleaning_fee": inp["cleaning"],
                    "monthly_cost": inp["monthly"],
                    "occ_low": inp["occ_low"],
                    "occ_base": inp["occ_base"],
                    "occ_high": inp["occ_high"],
                    "avg_nights": inp["avg_nights"],
                    "output_mode": inp["output_mode"],
                    "max_results": max_res,
                    "max_pages": max_pgs,
                }
                if inp["beds"] is not None:
                    kwargs["beds"] = inp["beds"]
                if inp["baths"] is not None:
                    kwargs["baths"] = inp["baths"]
                out_dir = module.run(region, **kwargs)

            self._finish_success(out_dir)

        except Exception as exc:
            import traceback
            if isinstance(exc, PermissionError) or "Permission denied" in str(exc):
                name = os.path.basename(str(exc).split("'")[-2] if "'" in str(exc) else "파일")
                self._log(f"\n파일 저장 실패: {name}", "err")
                self._log("Excel에서 해당 파일이 열려 있습니다. 닫은 뒤 다시 실행하세요.", "err")
            else:
                self._log(f"\n오류: {exc}", "err")
                self._log(traceback.format_exc(), "err")
            self.after(0, lambda: self.lbl_status.configure(text="실행 실패"))
            self.after(0, lambda: self.lbl_sidebar_status.configure(text="실행 실패"))
        finally:
            sys.stdout, sys.stderr = old_out, old_err
            _af.AirbnbClient.crawl = _orig_crawl
            self._running = False
            self.after(0, lambda: self.btn_run.configure(state="normal", text="분석 실행"))

    def _finish_success(self, out_dir: Path | None) -> None:
        self._last_out_dir = out_dir if out_dir else None
        self._log("")
        self._log("=" * 64, "hdr")
        self._log("생성된 파일", "hdr")
        self._log("=" * 64, "hdr")
        if out_dir:
            self._log(f"output/{out_dir.name}/", "dir")

        if out_dir and out_dir.exists():
            import time as _time
            now = _time.time()
            recent = [f for f in sorted(out_dir.iterdir()) if f.is_file() and now - f.stat().st_mtime < 45]
            files = recent if recent else sorted([f for f in out_dir.iterdir() if f.is_file()])
            if files:
                for f in files:
                    kind = "HTML" if f.suffix == ".html" else "XLSX"
                    size_kb = f.stat().st_size / 1024
                    self._log(f"{kind:<4} {f.name:<48} {size_kb:>7.0f} KB", "file")
            else:
                self._log("생성된 파일이 없습니다.", "err")
        else:
            self._log("출력 폴더가 없습니다. 수집 결과가 0개일 수 있습니다.", "err")
        self._log("=" * 64, "hdr")

        elapsed_text = ""
        if self._run_started_at:
            elapsed = datetime.now() - self._run_started_at
            elapsed_text = f" - {elapsed.seconds // 60}분 {elapsed.seconds % 60}초"
        self._log(f"완료{elapsed_text}", "ok")

        self.after(0, lambda: self.prog_bar.configure(value=100))
        self.after(0, lambda: self.lbl_status.configure(text="완료"))
        self.after(0, lambda: self.lbl_sidebar_status.configure(text="완료"))
        self.after(0, lambda: self.lbl_step.configure(text=f"분석이 완료되었습니다{elapsed_text}."))
        self.after(0, lambda: self.btn_open.configure(state="normal"))
        if out_dir and out_dir.exists():
            self.after(600, lambda d=out_dir: os.startfile(str(d)))

    def _open_output(self) -> None:
        folder = self._last_out_dir or (PROJECT_DIR / "output")
        folder.mkdir(parents=True, exist_ok=True)
        os.startfile(folder)


if __name__ == "__main__":
    app = App()
    app.mainloop()
