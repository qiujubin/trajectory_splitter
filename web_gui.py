"""pywebview 启动器 — 原生窗口跑 HTML/CSS/JS 前端。

启动:  python -m trajectory_splitter.web_gui

前端文件位于 web/ 下, 默认从 web/index.html 加载。
JS 通过 window.pywebview.api.* 调下面的 Api 类方法。
"""

from __future__ import annotations

import sys
from pathlib import Path

import webview

from trajectory_splitter import (
    default_end_poi,
    default_start_poi,
    find_main_date,
    split_at_poi,
)
from trajectory_splitter import logging_setup
from trajectory_splitter.merger import merge_two
from trajectory_splitter.modifier import ModifyOptions, modify_xls


def _resolve_web_dir(fallback: Path) -> Path:
    """PyInstaller onefile 启动时会解压到 sys._MEIPASS, web/ 资源跟着过去。
    没打包时 (开发模式) 走 __file__ 同级目录。"""
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        return Path(meipass) / "web"
    return fallback


WEB_DIR = _resolve_web_dir(Path(__file__).parent / "web")


class Api:
    """pywebview 暴露给 JS 的 API。

    pywebview 把方法名 camelCase 之后注入 window.pywebview.api。
    这里的方法名按 JS 习惯用 snake_case 也行, JS 端按 api.method_name() 调用即可。
    """

    def pick_xls_file(self) -> str | None:
        """弹原生文件选择对话框, 过滤 .xls。返回路径字符串或 None。"""
        window = webview.windows[0]
        result = window.create_file_dialog(
            webview.OPEN_DIALOG,
            file_types=("XLS files (*.xls)", "All files (*.*)"),
        )
        if not result:
            return None
        # create_file_dialog 可能返回 list 或 tuple, 取第一个
        path = result[0] if isinstance(result, (list, tuple)) else result
        return str(path) if path else None

    def pick_directory(self) -> str | None:
        """弹原生目录选择对话框。返回路径字符串或 None。"""
        window = webview.windows[0]
        result = window.create_file_dialog(webview.FOLDER_DIALOG)
        if not result:
            return None
        path = result[0] if isinstance(result, (list, tuple)) else result
        return str(path) if path else None

    def preview_main_date(self, src: str) -> dict:
        """返回 {ok, main_date, count} 或 {ok:false, error}。"""
        try:
            main_date, count = find_main_date(src)
            return {"ok": True, "main_date": main_date, "count": count}
        except Exception as e:  # noqa: BLE001
            return {"ok": False, "error": str(e)}

    def default_split_poi(self, direction: str) -> dict:
        """按 direction 返回默认 start/end POI。"""
        return {
            "ok": True,
            "start_poi": default_start_poi(direction),
            "end_poi": default_end_poi(direction),
        }

    def run_split(self, params: dict) -> dict:
        """运行切片。

        params keys: src, direction, out_dir, start_poi, end_poi, noon,
                     clamp, round_to_minute, tpl
        返回 {ok, out_path} 或 {ok:false, error}。
        """
        try:
            out_path = split_at_poi(
                params["src"],
                params["direction"],
                params["out_dir"],
                start_poi=(params.get("start_poi") or None),
                end_poi=(params.get("end_poi") or None),
                noon_cutoff=(params.get("noon") or None),
                clamp_temp_to_0to4=bool(params.get("clamp", True)),
                round_to_minute=bool(params.get("round_to_minute", False)),
                filename_template=(params.get("tpl") or "1{direction}.xls"),
            )
            return {"ok": True, "out_path": str(out_path)}
        except Exception as e:  # noqa: BLE001
            return {"ok": False, "error": str(e)}

    def run_merge(self, params: dict) -> dict:
        """运行双文件拼接。

        params keys: first, second, out_dir, gap_minutes, output_plate,
                     clamp, round_to_minute, tpl
        """
        try:
            gap_str = (params.get("gap_minutes") or "").strip() or "1"
            gap = int(gap_str)
            out_plate = (params.get("output_plate") or "").strip() or None
            out_path = merge_two(
                params["first"],
                params["second"],
                params["out_dir"],
                time_gap_minutes=gap,
                clamp_temp_to_0to4=bool(params.get("clamp", True)),
                filename_template=(params.get("tpl") or "1日前.xls"),
                output_plate=out_plate,
                round_to_minute=bool(params.get("round_to_minute", False)),
            )
            return {"ok": True, "out_path": str(out_path)}
        except Exception as e:  # noqa: BLE001
            return {"ok": False, "error": str(e)}

    def run_modify(self, params: dict) -> dict:
        """运行数据微改。

        params keys: src, out_dir, output_plate, output_main_date,
                     time_offset_minutes, time_jitter_seconds, temp_jitter,
                     clamp, round_to_minute, seed, tpl
        """
        try:
            offset = int((params.get("time_offset_minutes") or "0").strip() or "0")
            jitter_s = float((params.get("time_jitter_seconds") or "8").strip() or "8")
            jitter_t = float((params.get("temp_jitter") or "0.3").strip() or "0.3")
        except ValueError:
            return {"ok": False, "error": "时间偏移/微抖/温度微抖都必须是数字"}

        out_plate = (params.get("output_plate") or "").strip()
        out_main_date = (params.get("output_main_date") or "").strip()

        seed_raw = (params.get("seed") or "").strip()
        try:
            seed = int(seed_raw) if seed_raw else None
        except ValueError:
            return {"ok": False, "error": "随机种子必须是整数"}

        opts = ModifyOptions(
            time_offset_minutes=offset,
            time_jitter_seconds=jitter_s,
            temp_jitter=jitter_t,
            output_main_date=out_main_date,
            output_plate=out_plate,
            round_to_minute=bool(params.get("round_to_minute", False)),
            clamp_temp_to_0to4=bool(params.get("clamp", True)),
            filename_template=(params.get("tpl") or "1modified.xls"),
            seed=seed,
        )
        try:
            out_path = modify_xls(params["src"], params["out_dir"], opts)
            return {"ok": True, "out_path": str(out_path)}
        except Exception as e:  # noqa: BLE001
            return {"ok": False, "error": str(e)}


def _resolve_index_url() -> str:
    """把 web/index.html 转成 file:// URL。pywebview 直接吃 path 也行,
    但用 as_uri() 更显式且对子资源 (style.css/app.js) 解析更稳。
    """
    return (WEB_DIR / "index.html").resolve().as_uri()


def main() -> int:
    logging_setup.configure()
    api = Api()
    window = webview.create_window(
        title="trajectory_splitter",
        url=_resolve_index_url(),
        js_api=api,
        width=1040,
        height=860,
        resizable=True,
        min_size=(900, 700),
    )
    # Linux 需要显式指定 gui; Windows / macOS 由 pywebview 自动选择 (默认 Edge WebView2 / WKWebView)。
    webview.start()
    return 0


if __name__ == "__main__":
    sys.exit(main())
