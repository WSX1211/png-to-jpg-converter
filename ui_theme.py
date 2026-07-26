#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""共享的现代商务 Tkinter 主题，不依赖第三方 UI 库。"""

from __future__ import annotations

import sys
import tkinter as tk
from tkinter import font as tkfont
from tkinter import ttk


COLORS = {
    "navy": "#101828",
    "navy_hover": "#1D2939",
    "blue": "#2563EB",
    "blue_hover": "#1D4ED8",
    "blue_soft": "#EFF6FF",
    "background": "#F4F7FB",
    "card": "#FFFFFF",
    "surface": "#F8FAFC",
    "text": "#172033",
    "muted": "#667085",
    "border": "#D9E1EC",
    "success": "#16803C",
    "danger": "#B42318",
    "disabled": "#98A2B3",
}

_DPI_ENABLED = False


def enable_windows_dpi_awareness() -> None:
    """在创建第一个 Tk 根窗口前启用 Windows 高 DPI。"""
    global _DPI_ENABLED
    if _DPI_ENABLED or sys.platform != "win32":
        return
    try:
        import ctypes

        try:
            ctypes.windll.user32.SetProcessDpiAwarenessContext(-4)
        except (AttributeError, OSError):
            ctypes.windll.user32.SetProcessDPIAware()
    except (ImportError, AttributeError, OSError):
        pass
    _DPI_ENABLED = True


def _preferred_font(root: tk.Misc) -> str:
    available = set(tkfont.families(root))
    for family in ("Microsoft YaHei UI", "Microsoft YaHei", "Segoe UI", "PingFang SC", "Noto Sans CJK SC"):
        if family in available:
            return family
    return tkfont.nametofont("TkDefaultFont", root=root).cget("family")


def configure_theme(root: tk.Misc) -> str:
    """配置统一主题并返回实际字体名称。"""
    family = _preferred_font(root)
    for name, size, weight in (
        ("TkDefaultFont", 10, "normal"),
        ("TkTextFont", 10, "normal"),
        ("TkHeadingFont", 10, "bold"),
        ("TkMenuFont", 10, "normal"),
    ):
        try:
            tkfont.nametofont(name, root=root).configure(family=family, size=size, weight=weight)
        except tk.TclError:
            pass

    root.configure(background=COLORS["background"])
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass

    style.configure(".", font=(family, 10), background=COLORS["background"], foreground=COLORS["text"])
    style.configure("App.TFrame", background=COLORS["background"])
    style.configure("Card.TFrame", background=COLORS["card"])
    style.configure("Header.TFrame", background=COLORS["navy"])
    style.configure("Title.TLabel", background=COLORS["background"], foreground=COLORS["text"], font=(family, 22, "bold"))
    style.configure("Subtitle.TLabel", background=COLORS["background"], foreground=COLORS["muted"], font=(family, 10))
    style.configure("HeaderTitle.TLabel", background=COLORS["navy"], foreground="#FFFFFF", font=(family, 21, "bold"))
    style.configure("HeaderSubtitle.TLabel", background=COLORS["navy"], foreground="#CBD5E1", font=(family, 10))
    style.configure("CardTitle.TLabel", background=COLORS["card"], foreground=COLORS["text"], font=(family, 11, "bold"))
    style.configure("CardText.TLabel", background=COLORS["card"], foreground=COLORS["muted"], font=(family, 10))
    style.configure("Muted.TLabel", background=COLORS["background"], foreground=COLORS["muted"])
    style.configure("Status.TLabel", background=COLORS["card"], foreground=COLORS["blue"], font=(family, 10, "bold"))

    style.configure(
        "Card.TLabelframe", background=COLORS["card"], bordercolor=COLORS["border"],
        lightcolor=COLORS["border"], darkcolor=COLORS["border"], borderwidth=1, relief="solid",
    )
    style.configure(
        "Card.TLabelframe.Label", background=COLORS["card"], foreground=COLORS["text"],
        font=(family, 11, "bold"), padding=(2, 4),
    )
    style.configure(
        "Drop.TLabelframe", background=COLORS["card"], bordercolor="#BFDBFE",
        lightcolor="#BFDBFE", darkcolor="#BFDBFE", borderwidth=1, relief="solid",
    )
    style.configure("Drop.TLabelframe.Label", background=COLORS["card"], foreground=COLORS["blue"], font=(family, 11, "bold"))
    style.configure(
        "DropActive.TLabelframe", background=COLORS["blue_soft"], bordercolor=COLORS["blue"],
        lightcolor=COLORS["blue"], darkcolor=COLORS["blue"], borderwidth=2, relief="solid",
    )
    style.configure("DropActive.TLabelframe.Label", background=COLORS["blue_soft"], foreground=COLORS["blue"], font=(family, 11, "bold"))
    style.configure("DropTitle.TLabel", background=COLORS["card"], foreground=COLORS["text"], font=(family, 15, "bold"))
    style.configure("DropActiveTitle.TLabel", background=COLORS["blue_soft"], foreground=COLORS["blue"], font=(family, 15, "bold"))

    _configure_buttons(style, family)
    _configure_inputs(style)
    return family


def _configure_buttons(style: ttk.Style, family: str) -> None:
    style.configure("Primary.TButton", background=COLORS["blue"], foreground="#FFFFFF", borderwidth=0, padding=(16, 10), font=(family, 10, "bold"))
    style.map("Primary.TButton", background=[("active", COLORS["blue_hover"]), ("pressed", COLORS["blue_hover"]), ("disabled", "#AFC6F9")])
    style.configure("Dark.TButton", background=COLORS["navy"], foreground="#FFFFFF", borderwidth=0, padding=(16, 10), font=(family, 10, "bold"))
    style.map("Dark.TButton", background=[("active", COLORS["navy_hover"]), ("pressed", COLORS["navy_hover"])])
    style.configure("Secondary.TButton", background=COLORS["card"], foreground=COLORS["text"], bordercolor=COLORS["border"], borderwidth=1, padding=(14, 9), font=(family, 10, "bold"))
    style.map("Secondary.TButton", background=[("active", COLORS["surface"]), ("pressed", COLORS["blue_soft"])], bordercolor=[("active", COLORS["blue"])])
    style.configure("Ghost.TButton", background=COLORS["background"], foreground=COLORS["muted"], borderwidth=0, padding=(10, 7))
    style.map("Ghost.TButton", foreground=[("active", COLORS["blue"])], background=[("active", COLORS["blue_soft"])])
    style.configure("Header.Primary.TButton", background="#FFFFFF", foreground=COLORS["navy"], borderwidth=0, padding=(18, 11), font=(family, 10, "bold"))
    style.map("Header.Primary.TButton", background=[("active", "#E8EEF8"), ("disabled", "#D0D5DD")])


def _configure_inputs(style: ttk.Style) -> None:
    style.configure("Modern.TEntry", fieldbackground=COLORS["surface"], foreground=COLORS["text"], bordercolor=COLORS["border"], padding=(10, 8), insertcolor=COLORS["text"])
    style.map("Modern.TEntry", bordercolor=[("focus", COLORS["blue"])], lightcolor=[("focus", COLORS["blue"])], darkcolor=[("focus", COLORS["blue"])])
    style.configure("Business.Horizontal.TProgressbar", troughcolor="#E5EAF2", background=COLORS["blue"], borderwidth=0, thickness=8)
    style.configure("Modern.Vertical.TScrollbar", background="#C8D1DF", troughcolor=COLORS["surface"], borderwidth=0, arrowcolor=COLORS["muted"])


def style_text_widget(widget: tk.Text) -> None:
    """统一原生 Text 控件视觉，不改变其 state 或内容行为。"""
    widget.configure(
        background=COLORS["surface"], foreground=COLORS["text"], insertbackground=COLORS["text"],
        selectbackground="#BFDBFE", selectforeground=COLORS["text"], relief=tk.FLAT,
        borderwidth=0, highlightthickness=1, highlightbackground=COLORS["border"],
        highlightcolor=COLORS["blue"], padx=10, pady=8,
    )
