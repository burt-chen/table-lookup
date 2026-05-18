"""Tkinter implementation of the Excel merger tool.

Reuses ``app.core`` for all data processing. Threading is plain
``threading.Thread`` + ``root.after`` to marshal results back to the UI thread.
"""
from __future__ import annotations

import json
import sys
import threading
import traceback
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

import tkinter as tk
from tkinter import ttk

from app.core import (
    MergeConfig,
    MergeResult,
    OutputColumn,
    Table,
    export,
    list_sheets,
    read_table,
    run_merge,
)


def _is_empty(v) -> bool:
    """Empty-value check that doesn't need pandas."""
    if v is None:
        return True
    if isinstance(v, float) and v != v:  # NaN
        return True
    return False


APP_TITLE = "資料對照彙整工具"
SETTINGS_PATH = Path.home() / ".table_lookup_settings.json"
# Source previews now load every row via batched insertion (see PreviewTable)


MODE_LABELS = [("來源資料", "copy"), ("自訂", "template")]
MODE_LABEL_OF = {v: k for k, v in MODE_LABELS}
MODE_VALUE_OF = {k: v for k, v in MODE_LABELS}


def _ref_of(src: int, col: str) -> str:
    return f"{src}.{col}"


def _parse_ref(value: str, cols1: list[str], cols2: list[str]):
    v = (value or "").strip()
    if v.startswith("1."):
        return 1, v[2:]
    if v.startswith("2."):
        return 2, v[2:]
    if v in cols1:
        return 1, v
    if v in cols2:
        return 2, v
    return None


# ---- widgets ----------------------------------------------------------------

class FilePickerRow(ttk.Frame):
    """One row: label + path + browse + sheet picker + status."""

    def __init__(self, parent, label: str, on_pick):
        super().__init__(parent)
        self._on_pick = on_pick
        ttk.Label(self, text=label, width=20).pack(side="left", padx=(0, 6))
        self.path_var = tk.StringVar()
        self.path_entry = ttk.Entry(self, textvariable=self.path_var, state="readonly")
        self.path_entry.pack(side="left", fill="x", expand=True)
        self.browse_btn = ttk.Button(self, text="選檔…", command=self._browse)
        self.browse_btn.pack(side="left", padx=4)
        ttk.Label(self, text="工作表").pack(side="left", padx=(8, 2))
        self.sheet_var = tk.StringVar()
        self.sheet_cb = ttk.Combobox(
            self, textvariable=self.sheet_var, width=22, state="readonly"
        )
        self.sheet_cb.pack(side="left")
        self.sheet_cb.bind("<<ComboboxSelected>>", lambda _e: self._fire(sheet_only=True))
        self.status_var = tk.StringVar(value="")
        self.status_lbl = ttk.Label(self, textvariable=self.status_var, foreground="#1976d2", width=14)
        self.status_lbl.pack(side="left", padx=(8, 0))
        self._sheet_was_enabled = False
        self.last_error: str | None = None

    def _browse(self):
        try:
            start = str(Path(self.path_var.get()).parent) if self.path_var.get() else ""
            path = filedialog.askopenfilename(
                title="選擇檔案",
                initialdir=start,
                filetypes=[
                    ("資料檔", "*.xlsx *.xlsm *.xls *.csv *.json"),
                    ("所有檔案", "*.*"),
                ],
            )
            if path:
                self.set_path(path)
        except Exception as e:
            messagebox.showerror("選檔失敗", f"{type(e).__name__}: {e}\n\n{traceback.format_exc()}")

    def set_path(self, path: str):
        self.path_var.set(path)
        ext = Path(path).suffix.lower() if path else ""
        sheets: list[str] = []
        sheet_err: str | None = None
        if path and ext in (".xlsx", ".xlsm", ".xls"):
            try:
                sheets = list_sheets(path)
            except Exception as e:
                sheet_err = f"{type(e).__name__}: {e}\n{traceback.format_exc()}"
        if sheets:
            self.sheet_cb.configure(values=sheets, state="readonly")
            self.sheet_cb.set(sheets[0])
        else:
            placeholder = "(JSON)" if ext == ".json" else "(CSV)" if ext == ".csv" else ""
            self.sheet_cb.configure(values=[placeholder] if placeholder else [], state="disabled")
            self.sheet_var.set(placeholder)
        self.last_error = sheet_err  # 給父層讀取
        if sheet_err is not None:
            # 寫到一個固定路徑的 debug 檔,絕對看得到
            try:
                debug_log = Path.home() / "table_lookup_debug.log"
                with open(debug_log, "a", encoding="utf-8") as f:
                    from datetime import datetime as _dt
                    f.write(f"\n=== {_dt.now().isoformat(timespec='seconds')} list_sheets failed ===\n")
                    f.write(f"path = {path}\n")
                    f.write(f"sys.path:\n")
                    for p in sys.path:
                        f.write(f"  {p}\n")
                    f.write(f"\nerror:\n{sheet_err}\n")
            except Exception:
                pass
            self.status_var.set("讀取失敗(看下方錯誤)")
        self._fire(sheet_only=False)

    def _fire(self, sheet_only: bool):
        self._on_pick(self, sheet_only)

    def path(self) -> str:
        return self.path_var.get()

    def sheet(self) -> str:
        if str(self.sheet_cb["state"]) == "disabled":
            return ""
        return self.sheet_var.get()

    def set_sheet(self, name: str):
        if name and name in self.sheet_cb["values"]:
            self.sheet_var.set(name)

    def set_loaded(self, n_rows: int):
        self.status_var.set(f"✓ {n_rows:,} 列")

    def set_loading(self, busy: bool, text: str = "讀取中…"):
        self.status_var.set(text if busy else "")
        if busy:
            self._sheet_was_enabled = (str(self.sheet_cb["state"]) == "readonly")
            self.browse_btn.configure(state="disabled")
            self.sheet_cb.configure(state="disabled")
        else:
            self.browse_btn.configure(state="normal")
            if self._sheet_was_enabled:
                self.sheet_cb.configure(state="readonly")


class MappingRow(tk.Frame):
    """One editable row: 輸出欄位名稱 | 模式 | 來源/模板.

    Single-click anywhere on the row makes it the current row; operations
    (delete / move / duplicate) act on the current row, matching PySide6.
    """

    NORMAL_BG = "SystemButtonFace"
    SELECTED_BG = "#cfe2ff"   # soft blue fill
    BORDER_NORMAL = "#cccccc" # subtle separator when not selected
    BORDER_SELECTED = "#1976d2"  # strong blue border for selection
    BAR_SELECTED = "#1976d2"  # left indicator bar

    def __init__(self, parent, table, name="", mode="copy", value=""):
        super().__init__(
            parent,
            bg=self.NORMAL_BG,
            padx=3, pady=2,
            highlightthickness=2,
            highlightbackground=self.NORMAL_BG,
            highlightcolor=self.NORMAL_BG,
        )
        self.table = table
        self._pending_value = ""
        self._is_current = False

        # Left indicator bar (col 0) — narrow strip that lights up when selected.
        self.indicator = tk.Frame(self, width=5, bg=self.NORMAL_BG)
        self.indicator.grid(row=0, column=0, sticky="ns", padx=(0, 4))

        self.name_var = tk.StringVar(value=name)
        self.mode_var = tk.StringVar(value=MODE_LABEL_OF.get(mode, "來源資料"))

        # name (col 1)
        self.name_entry = ttk.Entry(self, textvariable=self.name_var, width=22)
        self.name_entry.grid(row=0, column=1, padx=(0, 4), sticky="ew")
        self.name_var.trace_add("write", lambda *_: table._on_changed())

        # mode (col 2)
        self.mode_cb = ttk.Combobox(
            self, textvariable=self.mode_var,
            values=[label for label, _ in MODE_LABELS],
            state="readonly", width=10,
        )
        self.mode_cb.grid(row=0, column=2, padx=(0, 4))
        self.mode_cb.bind("<<ComboboxSelected>>", self._on_mode_changed)

        # value holder (col 3) — use tk.Frame so its bg tints when selected
        self.value_holder = tk.Frame(self, bg=self.NORMAL_BG)
        self.value_holder.grid(row=0, column=3, sticky="ew")
        self.columnconfigure(3, weight=1)
        self.value_holder.columnconfigure(1, weight=1)

        self._build_value_widget(value)
        self._attach_click_handlers()

    # ---- selection ----

    def set_current(self, current: bool) -> None:
        self._is_current = current
        bg = self.SELECTED_BG if current else self.NORMAL_BG
        border = self.BORDER_SELECTED if current else self.NORMAL_BG
        bar = self.BAR_SELECTED if current else self.NORMAL_BG
        try:
            self.configure(
                bg=bg,
                highlightbackground=border,
                highlightcolor=border,
            )
            self.indicator.configure(bg=bar)
            self.value_holder.configure(bg=bg)
        except tk.TclError:
            pass

    def _attach_click_handlers(self, root_widget=None) -> None:
        """Make clicks on the row (or any inner widget) select it.

        ``add="+"`` so the inner widget's own click handlers (focus, dropdown,
        cursor placement) still run. Pass ``root_widget`` to scope the walk
        to a newly-built subtree (avoids duplicate bindings on rebuild).
        """
        target = root_widget if root_widget is not None else self
        def stamp(widget):
            try:
                widget.bind("<Button-1>", self._on_click, add="+")
            except tk.TclError:
                pass
            for child in widget.winfo_children():
                stamp(child)
        stamp(target)

    def _on_click(self, _e=None) -> None:
        self.table._select_row(self)

    # ---- mode ----

    def mode_internal(self) -> str:
        return MODE_VALUE_OF.get(self.mode_var.get(), "copy")

    def _on_mode_changed(self, _e=None):
        current = self.value()
        self._build_value_widget(current)
        self.table._on_changed()

    # ---- value widgets ----

    def _build_value_widget(self, value: str):
        for w in self.value_holder.winfo_children():
            w.destroy()
        if self.mode_internal() == "copy":
            self._build_copy(value)
        else:
            self._build_template(value)
        # Only stamp the freshly-built subtree; the row + mode combo already have bindings.
        self._attach_click_handlers(self.value_holder)

    def _build_copy(self, value: str):
        self.src_var = tk.StringVar()
        self.col_var = tk.StringVar()
        self.src_cb = ttk.Combobox(
            self.value_holder, textvariable=self.src_var,
            values=["來源 1", "來源 2"], state="readonly", width=10,
        )
        self.src_cb.grid(row=0, column=0, padx=(0, 4))
        self.col_cb = ttk.Combobox(
            self.value_holder, textvariable=self.col_var,
            values=[], state="readonly",
        )
        self.col_cb.grid(row=0, column=1, sticky="ew")
        self.src_cb.bind("<<ComboboxSelected>>", self._on_src_changed)
        self.col_cb.bind("<<ComboboxSelected>>", lambda _e: self._on_col_changed())
        self._apply_copy_value(value)

    def _build_template(self, value: str):
        self.tpl_var = tk.StringVar(value=value)
        self.tpl_entry = ttk.Entry(self.value_holder, textvariable=self.tpl_var)
        self.tpl_entry.grid(row=0, column=0, columnspan=2, sticky="ew")
        self.tpl_var.trace_add("write", lambda *_: self.table._on_changed())
        self.insert_btn = ttk.Menubutton(self.value_holder, text="插入欄位 ▾")
        self.insert_btn.grid(row=0, column=2, padx=(4, 0))
        self._rebuild_insert_menu()

    def _rebuild_insert_menu(self):
        if self.mode_internal() != "template":
            return
        menu_font = ("TkDefaultFont", UI_FONT_SIZE)
        menu = tk.Menu(self.insert_btn, tearoff=False, font=menu_font)
        for src, cols in ((1, self.table._cols1), (2, self.table._cols2)):
            sub = tk.Menu(menu, tearoff=False, font=menu_font)
            if cols:
                for c in cols:
                    sub.add_command(
                        label=c,
                        command=lambda s=src, name=c: self._insert_token(s, name),
                    )
            else:
                sub.add_command(label="(尚未載入)", state="disabled")
            menu.add_cascade(label=f"來源 {src}", menu=sub)
        self.insert_btn.configure(menu=menu)

    def _insert_token(self, src: int, col: str):
        token = f"{{{src}.{col}}}"
        pos = self.tpl_entry.index("insert")
        self.tpl_var.set(self.tpl_var.get()[:pos] + token + self.tpl_var.get()[pos:])
        self.tpl_entry.icursor(pos + len(token))
        self.tpl_entry.focus_set()

    # ---- copy mode helpers ----

    def _apply_copy_value(self, v: str):
        self._pending_value = (v or "").strip()
        parsed = _parse_ref(v, self.table._cols1, self.table._cols2)
        if parsed is not None:
            src, col = parsed
            self.src_var.set(f"來源 {src}")
            self._refill_col_cb()
            if col in self.col_cb["values"]:
                self.col_var.set(col)
                self._pending_value = ""
        else:
            if not self.src_var.get():
                self.src_var.set("來源 1")
            self._refill_col_cb()

    def _refill_col_cb(self):
        src = 1 if self.src_var.get() == "來源 1" else 2
        cols = self.table._cols1 if src == 1 else self.table._cols2
        self.col_cb.configure(values=[""] + list(cols))

    def _on_src_changed(self, _e=None):
        self._pending_value = ""
        self._refill_col_cb()
        self.col_var.set("")
        self.table._on_changed()

    def _on_col_changed(self):
        if self.col_var.get():
            self._pending_value = ""
        self.table._on_changed()

    # ---- public ----

    def name(self) -> str:
        return self.name_var.get().strip()

    def value(self) -> str:
        if self.mode_internal() == "copy":
            col = getattr(self, "col_var", tk.StringVar()).get()
            if col:
                src = 1 if self.src_var.get() == "來源 1" else 2
                return _ref_of(src, col)
            return self._pending_value
        return getattr(self, "tpl_var", tk.StringVar()).get()

    def set_value(self, v: str):
        if self.mode_internal() == "copy":
            self._apply_copy_value(v)
        else:
            self.tpl_var.set(v or "")

    def refresh_source_columns(self):
        if self.mode_internal() == "copy":
            pending = self._pending_value or self.value()
            self._refill_col_cb()
            if pending:
                self._apply_copy_value(pending)
        else:
            self._rebuild_insert_menu()


class ScrollableRowFrame(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.inner = ttk.Frame(self.canvas)
        self.inner_window = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        self.inner.bind("<Configure>", lambda _e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(self.inner_window, width=e.width))
        # mouse wheel scroll
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel, add="+")

    def _on_mousewheel(self, event):
        # Only scroll if pointer is over this canvas
        x, y = event.x_root, event.y_root
        try:
            wx = self.canvas.winfo_rootx()
            wy = self.canvas.winfo_rooty()
            ww = self.canvas.winfo_width()
            wh = self.canvas.winfo_height()
            if wx <= x <= wx + ww and wy <= y <= wy + wh:
                first, last = self.canvas.yview()
                # 內容比畫面短(全部已可見)時不滾動,避免頂端被推出空白
                if last - first >= 1.0:
                    return
                step = int(-event.delta / 120)
                # 已到頂/到底就不要再往該方向滾過頭
                if (step < 0 and first <= 0.0) or (step > 0 and last >= 1.0):
                    return
                self.canvas.yview_scroll(step, "units")
        except tk.TclError:
            pass


class MappingTable(ttk.Frame):
    """Manages the output column rows + toolbar.

    Single-click on a row makes it the "current" row; toolbar actions
    (刪除 / 上移 / 下移 / 複製列) operate on that row.
    """

    HEADERS = [("輸出欄位名稱", 22), ("模式", 12), ("來源欄位 或 模板", 50)]

    def __init__(self, parent, on_changed):
        super().__init__(parent)
        self._cols1: list[str] = []
        self._cols2: list[str] = []
        self._rows: list[MappingRow] = []
        self._current: MappingRow | None = None
        self._on_changed_cb = on_changed
        self._emit_changes = True

        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        # Toolbar (one row, at top)
        bar = ttk.Frame(self)
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 4))
        # Row-editing buttons on the left
        left_actions = [
            ("新增", self.add_row),
            ("刪除", self.remove_current),
            ("上移", lambda: self._move(-1)),
            ("下移", lambda: self._move(+1)),
            ("複製列", self.duplicate_current),
            ("從來源 1 帶入", lambda: self._autofill(1)),
            ("從來源 2 帶入", lambda: self._autofill(2)),
            ("清空", self.clear_all),
        ]
        for text, fn in left_actions:
            ttk.Button(bar, text=text, command=fn).pack(side="left", padx=(0, 4))

        # 欄位設定 import/export pinned to the right
        # Pack rightmost element first so the visual order reads "欄位設定 匯出 匯入".
        ttk.Button(bar, text="匯入", command=self.import_json).pack(side="right", padx=(0, 4))
        ttk.Button(bar, text="匯出", command=self.export_json).pack(side="right", padx=(0, 4))
        ttk.Label(bar, text="欄位設定").pack(side="right", padx=(8, 6))

        # Column header strip — col 0 spacer matches the row indicator bar.
        headers = ttk.Frame(self)
        headers.grid(row=1, column=0, sticky="ew", padx=5)  # 5px to clear row's highlightthickness=2 + padx=3
        headers.columnconfigure(3, weight=1)
        tk.Frame(headers, width=9).grid(row=0, column=0)  # spacer for row indicator
        col_widths = [self.HEADERS[0][1], self.HEADERS[1][1], None]
        for col, ((text, _w), width) in enumerate(zip(self.HEADERS, col_widths), start=1):
            kw = dict(width=width) if width else dict()
            tk.Label(
                headers, text=text, anchor="w",
                font=("Microsoft JhengHei UI", UI_FONT_SIZE, "bold"),
                background="#e8edf3", padx=6, pady=4, **kw,
            ).grid(row=0, column=col, sticky="ew", padx=(0, 2))

        # Scrollable row container
        self.scroll = ScrollableRowFrame(self)
        self.scroll.grid(row=2, column=0, sticky="nsew")

    # ---- public ----

    def set_source_columns(self, cols1: list[str], cols2: list[str]):
        self._cols1 = list(cols1)
        self._cols2 = list(cols2)
        for row in self._rows:
            row.refresh_source_columns()

    def get_columns(self) -> list[OutputColumn]:
        out = []
        for row in self._rows:
            n = row.name()
            if not n:
                continue
            out.append(OutputColumn(name=n, mode=row.mode_internal(), value=row.value()))
        return out

    def set_columns(self, cols: list[OutputColumn]):
        self._emit_changes = False
        try:
            for row in list(self._rows):
                row.destroy()
            self._rows.clear()
            self._current = None
            for c in cols:
                self._append_row(c.name, c.mode, c.value)
        finally:
            self._emit_changes = True
        self._on_changed()

    # ---- row ops ----

    def add_row(self):
        self._append_row("", "copy", "")
        self._on_changed()

    def _append_row(self, name: str, mode: str, value: str):
        row = MappingRow(self.scroll.inner, self, name=name, mode=mode, value=value)
        row.pack(fill="x", padx=2, pady=1)
        self._rows.append(row)

    def _select_row(self, row: MappingRow) -> None:
        """Mark ``row`` as the current row (single selection)."""
        if self._current is row:
            return
        if self._current is not None:
            try:
                self._current.set_current(False)
            except tk.TclError:
                pass
        self._current = row
        if row is not None:
            row.set_current(True)

    def _require_current(self, action: str) -> MappingRow | None:
        if self._current is None or self._current not in self._rows:
            messagebox.showinfo("提示", f"請先點選要{action}的列")
            return None
        return self._current

    def remove_current(self):
        row = self._require_current("刪除")
        if row is None:
            return
        idx = self._rows.index(row)
        self._rows.remove(row)
        row.destroy()
        # Auto-select neighbour so the user can chain ops
        if self._rows:
            self._select_row(self._rows[min(idx, len(self._rows) - 1)])
        else:
            self._current = None
        self._on_changed()

    def duplicate_current(self):
        row = self._require_current("複製")
        if row is None:
            return
        idx = self._rows.index(row)
        new_row = MappingRow(
            self.scroll.inner, self,
            name=row.name(), mode=row.mode_internal(), value=row.value(),
        )
        self._rows.insert(idx + 1, new_row)
        self._repack()
        self._select_row(new_row)
        self._on_changed()

    def _move(self, delta: int):
        row = self._require_current("移動")
        if row is None:
            return
        i = self._rows.index(row)
        j = i + delta
        if not (0 <= j < len(self._rows)):
            return  # silently no-op at the boundaries
        self._rows[i], self._rows[j] = self._rows[j], self._rows[i]
        self._repack()
        self._select_row(row)
        self._on_changed()

    def _repack(self):
        for r in self._rows:
            r.pack_forget()
        for r in self._rows:
            r.pack(fill="x", padx=2, pady=1)

    def _autofill(self, src: int):
        cols = self._cols1 if src == 1 else self._cols2
        if not cols:
            messagebox.showinfo("提示", f"來源 {src} 尚未載入")
            return
        existing = {r.name() for r in self._rows}
        for c in cols:
            if c in existing:
                continue
            self._append_row(c, "copy", _ref_of(src, c))
        self._on_changed()

    def clear_all(self):
        if not self._rows:
            return
        if not messagebox.askyesno("確認", "清空所有輸出欄位設定?"):
            return
        for r in list(self._rows):
            r.destroy()
        self._rows.clear()
        self._current = None
        self._on_changed()

    # ---- JSON IO ----

    def import_json(self):
        path = filedialog.askopenfilename(
            title="匯入輸出欄位 JSON", filetypes=[("JSON", "*.json"), ("所有檔案", "*.*")]
        )
        if not path:
            return
        try:
            data = json.loads(Path(path).read_text(encoding="utf-8"))
        except Exception as e:
            messagebox.showwarning("匯入失敗", str(e))
            return
        items = data.get("outputs", data) if isinstance(data, dict) else data
        if not isinstance(items, list):
            messagebox.showwarning("匯入失敗", "JSON 結構不正確,應為 outputs 陣列")
            return
        cols = [OutputColumn.from_dict(x) for x in items if isinstance(x, dict)]
        if self._rows:
            ans = messagebox.askyesnocancel(
                "匯入方式",
                f"目前已有 {len(self._rows)} 列。\n是 = 取代,否 = 附加在尾端,取消 = 中止",
            )
            if ans is None:
                return
            if ans:
                self.set_columns(cols)
                return
            for c in cols:
                self._append_row(c.name, c.mode, c.value)
            self._on_changed()
        else:
            self.set_columns(cols)

    def export_json(self):
        cols = self.get_columns()
        if not cols:
            messagebox.showinfo("提示", "目前沒有可匯出的欄位")
            return
        path = filedialog.asksaveasfilename(
            title="匯出輸出欄位 JSON",
            defaultextension=".json",
            initialfile="outputs.json",
            filetypes=[("JSON", "*.json")],
        )
        if not path:
            return
        data = {"outputs": [c.to_dict() for c in cols]}
        try:
            Path(path).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
            messagebox.showinfo("完成", f"已匯出至:\n{path}")
        except Exception as e:
            messagebox.showwarning("存檔失敗", str(e))

    # ---- internal ----

    def _on_changed(self):
        if self._emit_changes and self._on_changed_cb:
            self._on_changed_cb()


class PreviewTable(ttk.Frame):
    """Read-only data preview using ttk.Treeview.

    Inserts rows in batches via ``after()`` so the UI stays responsive on
    large tables (Tk's Treeview isn't virtualised, every insert is a Tcl
    call). The generation counter cancels stale in-flight loads when a new
    show_df arrives.
    """

    BATCH_SIZE = 500       # rows per scheduled chunk
    BATCH_INTERVAL = 1     # ms between chunks (yields to the event loop)

    def __init__(self, parent):
        super().__init__(parent)
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        self.tree = ttk.Treeview(self, show="headings")
        vbar = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        hbar = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vbar.set, xscrollcommand=hbar.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vbar.grid(row=0, column=1, sticky="ns")
        hbar.grid(row=1, column=0, sticky="ew")
        self.info = ttk.Label(self, text="", anchor="w")
        self.info.grid(row=2, column=0, columnspan=2, sticky="ew", padx=4, pady=(2, 0))
        self._load_gen = 0  # increment to cancel in-flight loads

    def show_df(self, df: Table, max_rows: int | None = None) -> None:
        # Cancel any prior batch by bumping the generation.
        self._load_gen += 1
        gen = self._load_gen

        self.tree.delete(*self.tree.get_children())
        cols = [str(c) for c in df.columns]
        self.tree["columns"] = cols
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=120, anchor="w", stretch=False)

        rows = df.rows if max_rows is None else df.head(max_rows).rows
        total = len(rows)
        capped = max_rows is not None and len(df) > max_rows
        suffix = (f"(僅顯示前 {max_rows:,})" if capped else "")

        if total == 0:
            self.info.configure(text=f"共 0 列、{len(cols)} 欄")
            return

        def insert_batch(start: int) -> None:
            if gen != self._load_gen:
                return  # superseded by a newer show_df
            end = min(start + self.BATCH_SIZE, total)
            for r in rows[start:end]:
                self.tree.insert(
                    "", "end",
                    values=["" if _is_empty(v) else str(v) for v in r],
                )
            if end < total:
                self.info.configure(
                    text=f"載入中 {end:,}/{total:,} 列、{len(cols)} 欄 {suffix}".rstrip()
                )
                self.after(self.BATCH_INTERVAL, lambda: insert_batch(end))
            else:
                self.info.configure(
                    text=f"共 {len(df):,} 列、{len(cols)} 欄 {suffix}".rstrip()
                )

        self.info.configure(text=f"準備載入 {total:,} 列…")
        insert_batch(0)


# ---- main window ------------------------------------------------------------

UI_FONT_SIZE = 12


def _configure_global_fonts(size: int = UI_FONT_SIZE) -> None:
    """Resize every named Tk/ttk font + apply matching size to ttk styles.

    Tk widgets read from the standard named fonts (``TkDefaultFont`` etc);
    ttk widgets pick up font via the Style system. We poke both so the
    whole window scales consistently.
    """
    import tkinter.font as tkfont
    for name in (
        "TkDefaultFont", "TkTextFont", "TkFixedFont", "TkMenuFont",
        "TkHeadingFont", "TkCaptionFont", "TkSmallCaptionFont",
        "TkIconFont", "TkTooltipFont",
    ):
        try:
            tkfont.nametofont(name).configure(size=size)
        except tk.TclError:
            pass
    style = ttk.Style()
    for st in (
        "TButton", "TLabel", "TEntry", "TCombobox", "TCheckbutton",
        "TRadiobutton", "TMenubutton", "TNotebook", "TNotebook.Tab",
        "TLabelframe", "TLabelframe.Label", "Treeview", "Treeview.Heading",
        "TProgressbar",
    ):
        try:
            style.configure(st, font=("TkDefaultFont", size))
        except tk.TclError:
            pass
    # Treeview row height needs to grow with the font; rough heuristic:
    try:
        style.configure("Treeview", rowheight=int(size * 2.0))
    except tk.TclError:
        pass


class MainWindow:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title(APP_TITLE)
        _configure_global_fonts()
        try:
            root.state("zoomed")
        except tk.TclError:
            root.geometry("1280x800")

        # State
        self._df1: Table | None = None
        self._df2: Table | None = None
        self._cols1: list[str] = []
        self._cols2: list[str] = []
        self._result_df: Table | None = None
        self._preview_ready = False
        self._load_gen = {1: 0, 2: 0}
        self._merge_running = False

        # 防止滑鼠滾輪在 Combobox 上停留時誤改選取值;
        # 下拉清單本身是另一個 widget class,捲動不受影響。
        root.bind_class("TCombobox", "<MouseWheel>", lambda _e: "break")

        self._build_menu()
        self._build_ui()
        self._load_settings_silent()
        self._invalidate_preview()

    # ---- menu ----

    def _build_menu(self):
        menu_font = ("TkDefaultFont", UI_FONT_SIZE)
        menubar = tk.Menu(self.root, font=menu_font)
        m = tk.Menu(menubar, tearoff=False, font=menu_font)
        m.add_command(label="儲存設定", accelerator="Ctrl+S", command=lambda: self._save_settings(SETTINGS_PATH))
        m.add_command(label="載入設定", accelerator="Ctrl+O", command=lambda: self._load_settings(SETTINGS_PATH))
        m.add_separator()
        m.add_command(label="另存設定…", command=self._save_settings_as)
        m.add_command(label="從檔案載入…", command=self._load_settings_from)
        menubar.add_cascade(label="設定", menu=m)
        self.root.configure(menu=menubar)
        self.root.bind("<Control-s>", lambda _e: self._save_settings(SETTINGS_PATH))
        self.root.bind("<Control-o>", lambda _e: self._load_settings(SETTINGS_PATH))

    # ---- ui ----

    def _build_ui(self):
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=8, pady=8)
        self.notebook.add(self._build_files_tab(self.notebook), text="1. 檔案")
        self.notebook.add(self._build_mapping_tab(self.notebook), text="2. 比對與欄位")
        self.notebook.add(self._build_preview_tab(self.notebook), text="3. 預覽 / 執行")
        self.notebook.add(self._build_log_tab(self.notebook), text="4. 日誌")
        # Tab 2 / 3 require both sources; gate them until they're loaded.
        self._refresh_tab_gating()
        self._refresh_tab_gating()

    def _refresh_tab_gating(self):
        """Tab 2 (比對與欄位) requires both source files loaded; Tab 3 too."""
        ready = self._df1 is not None and self._df2 is not None
        try:
            self.notebook.tab(1, state=("normal" if ready else "disabled"))
            self.notebook.tab(2, state=("normal" if ready else "disabled"))
        except tk.TclError:
            pass

    def _build_files_tab(self, parent):
        page = ttk.Frame(parent)
        page.columnconfigure(0, weight=1)

        intro_text = (
            "角色說明\n"
            "  • 來源 1:要彙整的資料(欄位較多的母檔,輸出欄位的內容都來自這裡)\n"
            "  • 來源 2:篩選條件(只挑出 key 在此清單中的列)\n"
            "處理邏輯:用「來源 2」某欄的值,去「來源 1」找對應列,組出新表。"
        )
        tk.Label(
            page, text=intro_text, justify="left", anchor="w",
            background="#f5f8ff", foreground="#000000",
            relief="solid", borderwidth=1, padx=10, pady=8,
        ).grid(row=0, column=0, sticky="ew", pady=(0, 8))

        files_box = ttk.LabelFrame(page, text="檔案來源")
        files_box.grid(row=1, column=0, sticky="ew")
        files_box.columnconfigure(0, weight=1)
        self.picker1 = FilePickerRow(files_box, "來源 1 (要彙整)", on_pick=self._on_picker_change)
        self.picker1.grid(row=0, column=0, sticky="ew", padx=6, pady=4)
        self.picker2 = FilePickerRow(files_box, "來源 2 (篩選條件)", on_pick=self._on_picker_change)
        self.picker2.grid(row=1, column=0, sticky="ew", padx=6, pady=(0, 6))

        tip_text = (
            "操作流程\n"
            "1. 選兩個來源檔案(xlsx 可選工作表,亦支援 csv/json)\n"
            "2. 到「比對與欄位」分頁設定 key 與輸出欄位\n"
            "3. 到「預覽 / 執行」分頁挑輸出格式與路徑,先預覽再執行\n"
            "4. 設定可從上方「設定」選單存檔,下次自動帶入"
        )
        tk.Label(page, text=tip_text, justify="left", anchor="w",
                 foreground="#000000", padx=4, pady=6).grid(row=2, column=0, sticky="ew", pady=(8, 0))
        page.rowconfigure(3, weight=1)
        return page

    def _build_mapping_tab(self, parent):
        page = ttk.Frame(parent)
        page.columnconfigure(0, weight=1)
        page.rowconfigure(1, weight=1)

        match_box = ttk.LabelFrame(page, text="比對設定")
        match_box.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        match_box.columnconfigure(1, weight=1)
        ttk.Label(match_box, text="來源 1 比對欄位:").grid(row=0, column=0, sticky="w", padx=6, pady=4)
        self.match_col1_var = tk.StringVar()
        self.match_col1_cb = ttk.Combobox(match_box, textvariable=self.match_col1_var, state="readonly", values=[])
        self.match_col1_cb.grid(row=0, column=1, sticky="ew", padx=6, pady=4)
        ttk.Label(match_box, text="來源 2 提供 key 的欄位:").grid(row=1, column=0, sticky="w", padx=6, pady=4)
        self.filter_col2_var = tk.StringVar()
        self.filter_col2_cb = ttk.Combobox(match_box, textvariable=self.filter_col2_var, state="readonly", values=[])
        self.filter_col2_cb.grid(row=1, column=1, sticky="ew", padx=6, pady=(0, 6))
        for var in (self.match_col1_var, self.filter_col2_var):
            var.trace_add("write", lambda *_: self._invalidate_preview())

        mapping_box = ttk.LabelFrame(
            page,
            text="輸出欄位 (「來源資料」直接複製欄位、「自訂」用 {欄名} 組合;前綴 1./2. 指定來源)",
        )
        mapping_box.grid(row=1, column=0, sticky="nsew")
        mapping_box.columnconfigure(0, weight=1)
        mapping_box.rowconfigure(0, weight=1)
        self.mapping = MappingTable(mapping_box, on_changed=self._invalidate_preview)
        self.mapping.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)
        return page

    def _build_preview_tab(self, parent):
        page = ttk.Frame(parent)
        page.columnconfigure(0, weight=1)
        page.rowconfigure(1, weight=1)

        out_box = ttk.LabelFrame(page, text="輸出檔")
        out_box.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        out_box.columnconfigure(2, weight=1)
        ttk.Label(out_box, text="格式:").grid(row=0, column=0, padx=(8, 4), pady=6)
        self.format_var = tk.StringVar(value="Excel (.xlsx)")
        self.format_cb = ttk.Combobox(
            out_box, textvariable=self.format_var,
            values=["Excel (.xlsx)", "CSV (.csv)", "JSON (.json)"],
            state="readonly", width=14,
        )
        self.format_cb.grid(row=0, column=1, padx=(0, 4), pady=6)
        self.format_cb.bind("<<ComboboxSelected>>", lambda _e: self._on_format_changed())
        self.out_path_var = tk.StringVar()
        ttk.Entry(out_box, textvariable=self.out_path_var).grid(
            row=0, column=2, sticky="ew", padx=(0, 4), pady=6
        )
        ttk.Button(out_box, text="另存為…", command=self._pick_output).grid(
            row=0, column=3, padx=(0, 6), pady=6
        )

        inner = ttk.Notebook(page)
        inner.grid(row=1, column=0, sticky="nsew")
        self.preview_src1 = PreviewTable(inner)
        self.preview_src2 = PreviewTable(inner)
        self.preview_out = PreviewTable(inner)
        inner.add(self.preview_src1, text="來源 1")
        inner.add(self.preview_src2, text="來源 2")
        inner.add(self.preview_out, text="輸出結果")
        self.inner_preview = inner

        action_box = ttk.LabelFrame(page, text="執行")
        action_box.grid(row=2, column=0, sticky="ew", pady=(6, 0))
        action_box.columnconfigure(3, weight=1)
        self.preview_btn = ttk.Button(action_box, text="預覽", command=self._do_preview)
        self.preview_btn.grid(row=0, column=0, padx=6, pady=6)
        self.run_btn = ttk.Button(action_box, text="執行並輸出", state="disabled", command=self._do_run)
        self.run_btn.grid(row=0, column=1, padx=(0, 16), pady=6)
        ttk.Label(action_box, text="進度:").grid(row=0, column=2, padx=(0, 4))
        self.progress = ttk.Progressbar(action_box, mode="determinate", maximum=100)
        self.progress.grid(row=0, column=3, sticky="ew", padx=(0, 8))
        self.status_var = tk.StringVar(value="待命")
        ttk.Label(action_box, textvariable=self.status_var, width=28).grid(
            row=0, column=4, padx=(0, 6)
        )
        return page

    def _build_log_tab(self, parent):
        page = ttk.Frame(parent)
        page.rowconfigure(0, weight=1)
        page.columnconfigure(0, weight=1)
        self.log_text = ScrolledText(page, wrap="word", state="disabled", height=20)
        self.log_text.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)
        btn_row = ttk.Frame(page)
        btn_row.grid(row=1, column=0, sticky="e", padx=4, pady=(0, 4))
        ttk.Button(btn_row, text="清除日誌", command=self._clear_log).pack(side="right")
        return page

    # ---- log ----

    def _log(self, msg: str):
        t = datetime.now().strftime("%H:%M:%S")
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"[{t}] {msg}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    # ---- file loading (async) ----

    def _on_picker_change(self, picker: FilePickerRow, sheet_only: bool):
        idx = 1 if picker is self.picker1 else 2
        path = picker.path()
        if not path or not Path(path).exists():
            return
        # 若 list_sheets 同步失敗,先把錯誤寫進日誌讓使用者看得到
        if picker.last_error:
            self._log(f"【來源 {idx} list_sheets 失敗】")
            for line in picker.last_error.splitlines():
                self._log(f"  {line}")
            self._log(f"完整錯誤也寫入: {Path.home() / 'table_lookup_debug.log'}")
            # 自動切到日誌分頁讓使用者看到
            try:
                self.notebook.select(3)
            except Exception:
                pass
            return
        self._invalidate_preview()
        self._load_gen[idx] += 1
        gen = self._load_gen[idx]
        picker.set_loading(True, "讀取中…")
        self.status_var.set(f"讀取來源 {idx}…")
        self.progress.configure(mode="indeterminate")
        self.progress.start(10)
        self._log(f"開始讀取來源 {idx}:{Path(path).name}")
        sheet = picker.sheet()

        def worker():
            try:
                df = read_table(path, sheet or None)
                self.root.after(0, lambda: self._on_load_finished(idx, gen, df))
            except Exception as e:
                msg = f"{e}\n\n{traceback.format_exc()}"
                self.root.after(0, lambda: self._on_load_failed(idx, gen, msg))

        threading.Thread(target=worker, daemon=True).start()

    def _on_load_finished(self, idx: int, gen: int, df: Table):
        if gen != self._load_gen.get(idx):
            return
        picker = self.picker1 if idx == 1 else self.picker2
        picker.set_loading(False)
        picker.set_loaded(len(df))
        cols = [str(c) for c in df.columns]
        if idx == 1:
            self._df1 = df
            self._cols1 = cols
            self.match_col1_cb.configure(values=cols)
            if self.match_col1_var.get() not in cols:
                self.match_col1_var.set(cols[0] if cols else "")
            self.preview_src1.show_df(df)
        else:
            self._df2 = df
            self._cols2 = cols
            self.filter_col2_cb.configure(values=cols)
            if self.filter_col2_var.get() not in cols:
                self.filter_col2_var.set(cols[0] if cols else "")
            self.preview_src2.show_df(df)
        self.mapping.set_source_columns(self._cols1, self._cols2)
        self.progress.stop()
        self.progress.configure(mode="determinate", value=0)
        self.status_var.set("待命")
        self._log(f"來源 {idx} 載入完成:{len(df)} 列、{len(cols)} 欄")
        self._refresh_tab_gating()
        if self._df1 is not None and self._df2 is not None:
            try:
                self.notebook.select(1)
            except Exception:
                pass

    def _on_load_failed(self, idx: int, gen: int, msg: str):
        if gen != self._load_gen.get(idx):
            return
        picker = self.picker1 if idx == 1 else self.picker2
        picker.set_loading(False)
        self.progress.stop()
        self.progress.configure(mode="determinate", value=0)
        self.status_var.set("讀取失敗")
        if idx == 1:
            self._df1 = None
        else:
            self._df2 = None
        self._refresh_tab_gating()
        self._log(f"來源 {idx} 讀取失敗:{msg.splitlines()[0]}")
        messagebox.showerror(f"來源 {idx} 讀檔失敗", msg)

    # ---- output settings ----

    def _on_format_changed(self):
        ext = self._format_ext()
        path = self.out_path_var.get().strip()
        if path and Path(path).suffix.lower() != ext:
            self.out_path_var.set(str(Path(path).with_suffix(ext)))

    def _format_ext(self) -> str:
        m = {"Excel (.xlsx)": ".xlsx", "CSV (.csv)": ".csv", "JSON (.json)": ".json"}
        return m.get(self.format_var.get(), ".xlsx")

    def _pick_output(self):
        ext = self._format_ext()
        labels = {".xlsx": "Excel", ".csv": "CSV", ".json": "JSON"}
        path = filedialog.asksaveasfilename(
            title="輸出至",
            defaultextension=ext,
            filetypes=[(labels[ext], f"*{ext}")]
                      + [(labels[e], f"*{e}") for e in labels if e != ext],
        )
        if not path:
            return
        p = Path(path)
        if p.suffix == "":
            path = str(p.with_suffix(ext))
        else:
            for label, e in (("Excel (.xlsx)", ".xlsx"), ("CSV (.csv)", ".csv"), ("JSON (.json)", ".json")):
                if e == p.suffix.lower():
                    self.format_var.set(label)
                    break
        self.out_path_var.set(path)

    # ---- preview gating ----

    def _invalidate_preview(self, *_a):
        self._preview_ready = False
        if hasattr(self, "run_btn"):
            self.run_btn.configure(state="disabled")

    def _mark_preview_ready(self):
        self._preview_ready = True
        self.run_btn.configure(state="normal")

    # ---- run / preview ----

    def _collect_config(self) -> MergeConfig | None:
        if self._df1 is None:
            messagebox.showwarning("缺少資料", "請先選擇來源 1")
            return None
        if self._df2 is None:
            messagebox.showwarning("缺少資料", "請先選擇來源 2")
            return None
        cfg = MergeConfig(
            file1=self.picker1.path(), sheet1=self.picker1.sheet(),
            file2=self.picker2.path(), sheet2=self.picker2.sheet(),
            match_col1=self.match_col1_var.get(),
            filter_col2=self.filter_col2_var.get(),
            drop_duplicates=True, keep_unmatched=False,
            outputs=self.mapping.get_columns(),
        )
        if not cfg.match_col1:
            messagebox.showwarning("缺少設定", "請選擇來源 1 的比對欄位")
            return None
        if not cfg.filter_col2:
            messagebox.showwarning("缺少設定", "請選擇來源 2 的 key 欄位")
            return None
        if not cfg.outputs:
            messagebox.showwarning("缺少設定", "請至少新增一個輸出欄位")
            return None
        return cfg

    def _do_preview(self):
        if self._merge_running:
            messagebox.showinfo("處理中", "目前已有作業正在執行")
            return
        cfg = self._collect_config()
        if cfg is None:
            return
        self._merge_running = True
        self.preview_btn.configure(state="disabled")
        self.run_btn.configure(state="disabled")
        self.progress.configure(mode="determinate", value=0, maximum=100)
        self.status_var.set("開始處理…")
        self._log("開始處理…")

        df1, df2 = self._df1, self._df2

        def progress_cb(i, n, msg):
            self.root.after(0, lambda: self._on_progress(i, n, msg))

        def worker():
            try:
                result = run_merge(cfg, df1, df2, progress=progress_cb)
                self.root.after(0, lambda: self._on_merge_done(result))
            except Exception as e:
                msg = f"{e}\n\n{traceback.format_exc()}"
                self.root.after(0, lambda: self._on_merge_failed(msg))

        threading.Thread(target=worker, daemon=True).start()

    def _on_progress(self, i, n, msg):
        if n > 0:
            self.progress.configure(value=int(100 * i / n))
        if msg:
            self.status_var.set(msg)
            self._log(msg)

    def _on_merge_done(self, result: MergeResult):
        self._merge_running = False
        self.preview_btn.configure(state="normal")
        self._log(
            f"完成:輸出 {len(result.df)} 列,來源 2 共 {result.total_keys} 個 key,"
            f"未匹配 {len(result.missing_keys)} 個"
        )
        if result.missing_keys:
            head = ", ".join(result.missing_keys[:10])
            more = f" …(其餘 {len(result.missing_keys) - 10})" if len(result.missing_keys) > 10 else ""
            self._log(f"未匹配範例:{head}{more}")
        self._result_df = result.df
        self.preview_out.show_df(result.df)
        self.inner_preview.select(2)
        self.notebook.select(2)
        self.status_var.set(f"預覽 {len(result.df)} 列 — 確認後可執行")
        self._mark_preview_ready()

    def _on_merge_failed(self, msg: str):
        self._merge_running = False
        self.preview_btn.configure(state="normal")
        self.progress.configure(value=0)
        self.status_var.set("處理失敗")
        self._log(f"處理失敗:{msg.splitlines()[0]}")
        messagebox.showerror("處理失敗", msg)

    def _do_run(self):
        if not self._preview_ready or self._result_df is None:
            messagebox.showinfo("請先預覽", "請先按「預覽」確認結果後再執行輸出")
            return
        out_path = self.out_path_var.get().strip()
        if not out_path:
            messagebox.showwarning("缺少輸出", "請選擇輸出檔位置")
            return
        ext = self._format_ext()
        if Path(out_path).suffix.lower() != ext:
            out_path = str(Path(out_path).with_suffix(ext))
            self.out_path_var.set(out_path)
        try:
            export(self._result_df, out_path)
            self._log(f"已輸出 {out_path}")
            messagebox.showinfo("完成", f"已輸出 {len(self._result_df)} 列至:\n{out_path}")
        except Exception as e:
            messagebox.showerror("輸出失敗", str(e))
            self._log(f"輸出失敗:{e}")

    # ---- settings IO ----

    def _collect_settings_dict(self) -> dict:
        cfg = MergeConfig(
            file1=self.picker1.path(), sheet1=self.picker1.sheet(),
            file2=self.picker2.path(), sheet2=self.picker2.sheet(),
            match_col1=self.match_col1_var.get(),
            filter_col2=self.filter_col2_var.get(),
            drop_duplicates=True, keep_unmatched=False,
            outputs=self.mapping.get_columns(),
        )
        data = cfg.to_dict()
        data["output_path"] = self.out_path_var.get()
        data["output_format"] = self.format_var.get()
        return data

    def _save_settings(self, path: Path):
        try:
            path.write_text(
                json.dumps(self._collect_settings_dict(), ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            self._log(f"設定已存至 {path}")
        except Exception as e:
            messagebox.showerror("存檔失敗", str(e))

    def _load_settings_silent(self):
        if SETTINGS_PATH.exists():
            try:
                self._load_settings(SETTINGS_PATH)
            except Exception:
                pass

    def _load_settings(self, path: Path):
        if not path.exists():
            return
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except Exception as e:
            messagebox.showerror("讀設定失敗", str(e))
            return
        cfg = MergeConfig.from_dict(data)
        if cfg.file1 and Path(cfg.file1).exists():
            self.picker1.set_path(cfg.file1)
            self.picker1.set_sheet(cfg.sheet1)
        if cfg.file2 and Path(cfg.file2).exists():
            self.picker2.set_path(cfg.file2)
            self.picker2.set_sheet(cfg.sheet2)
        if cfg.match_col1:
            self.match_col1_var.set(cfg.match_col1)
        if cfg.filter_col2:
            self.filter_col2_var.set(cfg.filter_col2)
        self.mapping.set_columns(cfg.outputs)
        if data.get("output_path"):
            self.out_path_var.set(data["output_path"])
        if data.get("output_format") in ("Excel (.xlsx)", "CSV (.csv)", "JSON (.json)"):
            self.format_var.set(data["output_format"])
        self._log(f"已載入設定 {path}")

    def _save_settings_as(self):
        path = filedialog.asksaveasfilename(
            title="另存設定", defaultextension=".json", filetypes=[("JSON", "*.json")],
        )
        if path:
            self._save_settings(Path(path))

    def _load_settings_from(self):
        path = filedialog.askopenfilename(title="載入設定", filetypes=[("JSON", "*.json")])
        if path:
            self._load_settings(Path(path))


def main() -> int:
    root = tk.Tk()
    try:
        ttk.Style().theme_use("vista")
    except tk.TclError:
        pass
    MainWindow(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
