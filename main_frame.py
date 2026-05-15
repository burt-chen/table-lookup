"""嵌入式包裝 — 讓 資料對照彙整工具 的 MainWindow 跑在 Launcher 的分頁裡。

實作 create_frame(parent) -> ttk.Frame，由 Launcher 動態載入。
"""
from __future__ import annotations

import importlib.util
import sys
from pathlib import Path
import tkinter as tk
from tkinter import ttk

_TOOL_ROOT = Path(__file__).parent  # 安裝後指向工具目錄，無需硬編路徑


def _load_tool_main():
    """用 importlib 直接從絕對路徑載入，完全不動 Launcher 的 app.* 名稱空間。"""

    # 1. 載入 core.py → 唯一名稱 _table_lookup.core
    core_spec = importlib.util.spec_from_file_location(
        "_table_lookup.core", _TOOL_ROOT / "app" / "core.py"
    )
    core_mod = importlib.util.module_from_spec(core_spec)
    sys.modules["_table_lookup.core"] = core_mod

    # 2. 暫時讓 app.core 指向它（main.py 內部 from app.core import 需要）
    _prev_app_core = sys.modules.get("app.core")
    sys.modules["app.core"] = core_mod
    try:
        core_spec.loader.exec_module(core_mod)

        # 3. 載入 main.py → 唯一名稱 _table_lookup.main
        main_spec = importlib.util.spec_from_file_location(
            "_table_lookup.main", _TOOL_ROOT / "app" / "main.py"
        )
        main_mod = importlib.util.module_from_spec(main_spec)
        sys.modules["_table_lookup.main"] = main_mod
        main_spec.loader.exec_module(main_mod)
    finally:
        # 4. 不管成功或失敗，一定還原 app.core
        if _prev_app_core is not None:
            sys.modules["app.core"] = _prev_app_core
        else:
            sys.modules.pop("app.core", None)

    return main_mod


_tool_main = _load_tool_main()
_MainWindow = _tool_main.MainWindow
_configure_global_fonts = _tool_main._configure_global_fonts


class _EmbeddedWindow(_MainWindow):
    """把 MainWindow 嵌進任意 Tkinter widget（不需要 tk.Tk）。"""

    def __init__(self, parent: tk.Widget) -> None:
        self.root = parent
        _configure_global_fonts()

        self._df1 = None
        self._df2 = None
        self._cols1: list[str] = []
        self._cols2: list[str] = []
        self._result_df = None
        self._preview_ready = False
        self._load_gen = {1: 0, 2: 0}
        self._merge_running = False

        self._build_menu()
        self._build_ui()
        self._load_settings_silent()
        self._invalidate_preview()

    def _build_menu(self) -> None:
        """嵌入模式：不建立 menubar，直接略過。"""


def create_frame(parent: tk.Widget) -> ttk.Frame:
    frame = ttk.Frame(parent)
    _EmbeddedWindow(frame)
    return frame
