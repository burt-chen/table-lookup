from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from manual_video_common import build_tk_video, parse_args


PROJECT = Path(__file__).resolve().parent
TITLE = "資料對照彙整工具"
PURPOSE = "兩份資料表 key 比對與自訂欄位輸出"
SCENE_PLAN = [
    ("01", "title", 4.0, None),
    ("02", "files", 6.0, 0),
    ("03", "mapping", 6.0, 1),
    ("04", "missing", 6.5, 1),
    ("05", "preview", 6.0, 2),
    ("06", "log", 5.0, 3),
]


def build_samples(project_dir: Path) -> None:
    from openpyxl import Workbook

    sample_dir = project_dir / "範例"
    sample_dir.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "來源1"
    ws.append(["客戶代號", "客戶名稱", "區域", "負責人"])
    ws.append(["C001", "晨星商行", "北區", "王小明"])
    ws.append(["C002", "青禾工程", "中區", "林雅婷"])
    ws.append(["C003", "遠山食品", "南區", "陳志豪"])
    wb.save(sample_dir / "來源1_客戶母檔.xlsx")

    wb = Workbook()
    ws = wb.active
    ws.title = "來源2"
    ws.append(["客戶代號"])
    ws.append(["C001"])
    ws.append(["C003"])
    wb.save(sample_dir / "來源2_篩選清單.xlsx")


def setup_scene(root, frame, app, scene_id: str, scene_name: str) -> None:
    wrapper = sys.modules[app.__class__.__module__]
    mod = wrapper._tool_main
    sample_dir = PROJECT / "範例"
    out_dir = sample_dir / "輸出"
    out_dir.mkdir(parents=True, exist_ok=True)
    src1 = sample_dir / "來源1_客戶母檔.xlsx"
    src2 = sample_dir / "來源2_篩選清單.xlsx"
    out_path = out_dir / "客戶對照輸出.xlsx"

    if app._df1 is None or app._df2 is None:
        app.picker1.path_var.set(str(src1))
        app.picker1.sheet_var.set("來源1")
        app.picker1.status_var.set("已載入 3 列")
        app.picker2.path_var.set(str(src2))
        app.picker2.sheet_var.set("來源2")
        app.picker2.status_var.set("已載入 2 列")
        app._load_gen[1] += 1
        app._load_gen[2] += 1
        app._on_load_finished(1, app._load_gen[1], mod.read_table(str(src1), "來源1"))
        app._on_load_finished(2, app._load_gen[2], mod.read_table(str(src2), "來源2"))
        app.match_col1_var.set("客戶代號")
        app.filter_col2_var.set("客戶代號")
        app.only_missing_var.set(False)
        app._on_scope_changed()
        app.mapping.set_columns([
            mod.OutputColumn("客戶名稱", "copy", "1.客戶名稱"),
            mod.OutputColumn("區域", "copy", "1.區域"),
            mod.OutputColumn("負責人", "copy", "1.負責人"),
        ])
        app.out_path_var.set(str(out_path))
        app._log("demo 來源檔已載入，並設定客戶代號為 key")

    if scene_id == "04":
        app.only_missing_var.set(True)
        app._on_scope_changed()
        app.mapping.set_columns([
            mod.OutputColumn("客戶代號", "copy", "1.客戶代號"),
            mod.OutputColumn("未對應客戶", "copy", "1.客戶名稱"),
            mod.OutputColumn("區域", "copy", "1.區域"),
            mod.OutputColumn("負責人", "copy", "1.負責人"),
        ])
        app._log("demo 已勾選「篩選沒有對應資料」，輸出欄位鎖定來源 1")
        app._result_df = None

    if scene_id in {"05", "06"} and app._result_df is None:
        cfg = app._collect_config()
        if cfg:
            result = mod.run_merge(cfg, app._df1, app._df2)
            app._on_merge_done(result)
            mod.export(app._result_df, str(out_path))
            app._log(f"demo 輸出完成：{out_path}")
            app.status_var.set(f"已輸出 demo 結果：{out_path.name}")


if __name__ == "__main__":
    args = parse_args()
    build_tk_video(
        PROJECT,
        TITLE,
        PURPOSE,
        SCENE_PLAN,
        sample_builder=build_samples,
        setup_scene=setup_scene,
        silent=args.silent,
        keep_temp=args.keep_temp,
    )
