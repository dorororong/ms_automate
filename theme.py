"""공통 디자인 토큰과 ttk 스타일.

CATMOA의 단순한 입력 흐름은 유지하고, Microsoft 앱처럼 차분한 파란색과
작은 여백을 사용해 화면의 시선을 줄인다.
"""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk

# --- 색 ---------------------------------------------------------------
BG = "#f5f8fc"
SURFACE = "#ffffff"
BORDER = "#d9e2ec"
TEXT = "#1f2937"
MUTED = "#60758a"
ACCENT = "#0f6cbd"       # Microsoft blue
ACCENT_HOVER = "#115ea3"
ACCENT_SOFT = "#eaf3fb"
OK = "#107c10"
FAIL = "#c42b1c"
WARN = "#8a6116"

SCOPE_COLORS = {"담임": "#0f6cbd", "업무": "#2b579a", "참조": "#7a8ca5"}
SCOPE_TINTS = {"담임": "#eaf3fb", "업무": "#eef4fb", "참조": "#f1f5f9"}

# --- 글꼴 -------------------------------------------------------------
FAMILY = "Malgun Gothic"
F_TITLE = (FAMILY, 13, "bold")
F_HEAD = (FAMILY, 10, "bold")
F_BODY = (FAMILY, 9)
F_SMALL = (FAMILY, 8)
F_HOTKEY = ("Consolas", 14, "bold")

PAD_X = 14


def apply_theme(root: tk.Misc) -> ttk.Style:
    """창 하나에 공통 스타일을 적용한다."""

    style = ttk.Style(root)
    try:
        style.theme_use("clam")  # 색 지정이 실제로 먹는 테마
    except tk.TclError:
        pass

    root.configure(bg=BG)
    style.configure(".", background=BG, foreground=TEXT, font=F_BODY)
    style.configure("TFrame", background=BG)
    style.configure("TLabel", background=BG, foreground=TEXT, font=F_BODY)
    style.configure("Title.TLabel", font=F_TITLE, foreground=TEXT)
    style.configure("Sub.TLabel", font=F_SMALL, foreground=MUTED)
    style.configure("Head.TLabel", font=F_HEAD, foreground=TEXT)
    style.configure("Muted.TLabel", font=F_SMALL, foreground=MUTED)
    style.configure("Hotkey.TLabel", font=F_HOTKEY, foreground=ACCENT)
    style.configure("Card.TFrame", background=SURFACE)
    style.configure("Card.TLabel", background=SURFACE, foreground=TEXT, font=F_BODY)
    style.configure("CardMuted.TLabel", background=SURFACE, foreground=MUTED, font=F_SMALL)
    style.configure("CardHead.TLabel", background=SURFACE, foreground=TEXT, font=F_HEAD)
    # 체크박스는 tk.Checkbutton 을 쓴다 (theme.checkbox 참고). clam 의 ttk
    # 체크 글리프는 'X' 로 그려져 '거부'처럼 읽히고, indicatorcolor 를 덮으면
    # 글리프가 통째로 사라진다.
    for name, bg in (("Card.TRadiobutton", SURFACE), ("TRadiobutton", BG)):
        style.configure(name, background=bg, font=F_BODY, indicatormargin=3)
        style.map(name,
                  background=[("active", bg)],
                  indicatorcolor=[("selected", ACCENT), ("!selected", "#ffffff")],
                  foreground=[("selected", ACCENT)])

    style.configure(
        "TButton", font=F_BODY, padding=(10, 5),
        background="#e7eef7", foreground=TEXT, borderwidth=0,
    )
    style.map("TButton",
              background=[("active", "#d6e6f5"), ("disabled", "#edf1f5")],
              foreground=[("disabled", "#a3a9b8")])
    style.configure("Accent.TButton", background=ACCENT, foreground="#ffffff", padding=(14, 6))
    style.map("Accent.TButton",
              background=[("active", ACCENT_HOVER), ("disabled", "#a9c9e4")],
              foreground=[("disabled", "#f5f9fd")])

    style.configure("TLabelframe", background=BG, borderwidth=0, relief="flat")
    style.configure("TLabelframe.Label", background=BG, foreground=MUTED, font=F_SMALL)
    style.configure("TEntry", fieldbackground=SURFACE, bordercolor=BORDER, padding=4)
    style.configure("TCombobox", fieldbackground=SURFACE, padding=3)
    style.configure("Status.TLabel", background=ACCENT_SOFT, foreground=TEXT, padding=(8, 5))
    style.configure("Treeview", background=SURFACE, fieldbackground=SURFACE, rowheight=24)
    style.configure("Treeview.Heading", font=F_SMALL, background="#e7eef7")
    return style


def header(parent: tk.Misc, title: str, subtitle: str = "") -> ttk.Frame:
    """CATMOA식 헤더 — 굵은 제목 아래 설명 한 줄."""

    frame = ttk.Frame(parent, padding=(PAD_X, 10, PAD_X, 7))
    frame.columnconfigure(0, weight=1)
    ttk.Label(frame, text=title, style="Title.TLabel").grid(row=0, column=0, sticky="w")
    if subtitle:
        ttk.Label(frame, text=subtitle, style="Sub.TLabel").grid(
            row=1, column=0, sticky="w", pady=(4, 0))
    return frame


def checkbox(parent: tk.Misc, text: str, variable: tk.BooleanVar,
             background: str = SURFACE) -> tk.Checkbutton:
    """네이티브 체크 표시가 나오는 체크박스."""

    return tk.Checkbutton(
        parent, text=text, variable=variable, font=F_BODY,
        background=background, activebackground=background,
        foreground=TEXT, activeforeground=ACCENT, selectcolor="#ffffff",
        highlightthickness=0, borderwidth=0, anchor="w", padx=0,
    )


def divider(parent: tk.Misc) -> tk.Frame:
    return tk.Frame(parent, height=1, bg=BORDER)


def card(parent: tk.Misc, accent: str) -> tuple[tk.Frame, ttk.Frame]:
    """왼쪽에 분류 색 띠가 있는 흰 카드. (바깥틀, 내용틀) 을 돌려준다."""

    outer = tk.Frame(parent, bg=BORDER, highlightthickness=0)
    outer.columnconfigure(1, weight=1)
    strip = tk.Frame(outer, bg=accent, width=4)
    strip.grid(row=0, column=0, sticky="ns")
    strip.grid_propagate(False)
    body = ttk.Frame(outer, style="Card.TFrame", padding=(10, 8))
    body.grid(row=0, column=1, sticky="nsew", padx=(0, 1), pady=1)
    body.columnconfigure(1, weight=1)
    return outer, body
