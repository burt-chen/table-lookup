"""Full end-to-end smoke test for the Tkinter app.

Drives the app via root.after() scheduled steps with a real mainloop so the
worker thread's root.after() calls succeed.
"""
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import tkinter as tk
from app.main import MainWindow
from app.core import OutputColumn


class TestRunner:
    POLL_MS = 100

    def __init__(self, root, w):
        self.root = root
        self.w = w
        self.error: str | None = None
        self.tmp = tempfile.TemporaryDirectory()
        self.export_idx = 0
        self.export_formats = [
            (".xlsx", "Excel (.xlsx)"),
            (".csv", "CSV (.csv)"),
            (".json", "JSON (.json)"),
        ]

    def start(self):
        self.root.after(100, self.step_load_file1)

    def fail(self, msg):
        self.error = msg
        self.root.quit()

    def finish(self):
        self.tmp.cleanup()
        self.root.quit()

    # ---- steps ----

    def step_load_file1(self):
        file1 = str(ROOT / "Frame_5K_Name_1100823行政界_1110617彙整.xlsx")
        print("Loading file1…")
        self.w.picker1.set_path(file1)
        self._deadline = 30000  # ms remaining
        self.root.after(self.POLL_MS, self._wait_file1)

    def _wait_file1(self):
        if self.w._df1 is not None:
            print(f"  ok: {len(self.w._df1)} rows, {len(self.w._cols1)} cols")
            self.step_load_file2()
            return
        self._deadline -= self.POLL_MS
        if self._deadline <= 0:
            return self.fail("file1 load timed out")
        self.root.after(self.POLL_MS, self._wait_file1)

    def step_load_file2(self):
        file2 = str(ROOT / "115國土利用調查第2階段幅幅清冊.xlsx")
        print("Loading file2…")
        self.w.picker2.set_path(file2)
        self._deadline = 30000
        self.root.after(self.POLL_MS, self._wait_file2)

    def _wait_file2(self):
        if self.w._df2 is not None:
            print(f"  ok: {len(self.w._df2)} rows, {len(self.w._cols2)} cols")
            self.step_configure()
            return
        self._deadline -= self.POLL_MS
        if self._deadline <= 0:
            return self.fail("file2 load timed out")
        self.root.after(self.POLL_MS, self._wait_file2)

    def step_configure(self):
        print("Configuring mapping…")
        self.w.match_col1_var.set("圖幅_5K")
        self.w.filter_col2_var.set("圖號")
        self.w.mapping.set_columns([
            OutputColumn(name="識別碼", mode="template", value="{2.圖號}_115LU"),
            OutputColumn(name="圖號", mode="copy", value="2.圖號"),
            OutputColumn(name="Town", mode="copy", value="1.Town"),
            OutputColumn(name="圖名", mode="copy", value="1.圖名"),
        ])
        if str(self.w.run_btn.cget("state")) != "disabled":
            return self.fail("run_btn should be disabled before preview")
        self.step_preview()

    def step_preview(self):
        print("Running preview…")
        self.w._do_preview()
        self._deadline = 60000
        self.root.after(self.POLL_MS, self._wait_preview)

    def _wait_preview(self):
        if self.w._preview_ready:
            df = self.w._result_df
            print(f"  ok: result {len(df)} rows")
            if str(self.w.run_btn.cget("state")) != "normal":
                return self.fail("run_btn should be normal after preview")
            tn_idx = df.column_index("圖號")
            hits = [i for i, r in enumerate(df.rows) if str(r[tn_idx]) == "97211001"]
            if not hits:
                return self.fail("97211001 not in result")
            row = df.row_dict(hits[0])
            print(f"  spot 97211001: 識別碼={row['識別碼']!r} 圖名={row['圖名']!r}")
            if row["識別碼"] != "97211001_115LU":
                return self.fail(f"unexpected 識別碼 {row['識別碼']!r}")
            if row["圖名"] != "南澳(四)":
                return self.fail(f"unexpected 圖名 {row['圖名']!r}")
            self.step_export_next()
            return
        self._deadline -= self.POLL_MS
        if self._deadline <= 0:
            return self.fail("preview timed out")
        self.root.after(self.POLL_MS, self._wait_preview)

    def step_export_next(self):
        if self.export_idx >= len(self.export_formats):
            self.step_invalidate()
            return
        ext, fmt_label = self.export_formats[self.export_idx]
        out = str(Path(self.tmp.name) / f"out{ext}")
        self.w.out_path_var.set(out)
        self.w.format_var.set(fmt_label)
        self.w._do_run()
        size = Path(out).stat().st_size if Path(out).exists() else 0
        if size <= 0:
            return self.fail(f"export {ext} produced empty file")
        print(f"  exported {ext}: {size} bytes")
        self.export_idx += 1
        self.root.after(100, self.step_export_next)

    def step_invalidate(self):
        print("Verifying invalidate-on-change…")
        self.w.mapping.add_row()
        if str(self.w.run_btn.cget("state")) != "disabled":
            return self.fail("config change must disable run_btn")
        print("  ok")
        print("ALL OK")
        self.finish()


def main():
    root = tk.Tk()
    root.withdraw()
    w = MainWindow(root)
    # _do_run uses messagebox.showinfo which is blocking — patch it out for the test
    import tkinter.messagebox as mb
    mb.showinfo = lambda *a, **k: None
    mb.showwarning = lambda *a, **k: None
    mb.showerror = lambda *a, **k: print("  [showerror]", a, k)

    runner = TestRunner(root, w)
    runner.start()
    root.mainloop()
    root.destroy()
    return 0 if runner.error is None else 1


if __name__ == "__main__":
    raise SystemExit(main())
