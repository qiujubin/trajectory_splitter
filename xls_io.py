"""xls 读写工具 (win32com / Excel COM 自动化, 常驻单实例)。

读写都走 Excel 本体 (机器上必须装 Excel; WPS 表格注册了兼容 ProgID 也能用):
  - 进程内共享一个隐藏 Excel 实例, 由专用工作线程持有, 首次调用时启动
    (~1s), 之后每次读写 0.1~0.6s — 避免旧版每次操作都新起/退出一个
    Excel 进程带来的 3~5s 开销
  - 读: Workbooks.Open(ReadOnly=True), UsedRange 一次性取值
  - 写: 新建工作簿, 整块赋值, SaveAs FileFormat=56 (xlExcel8 = .xls 97-2003)
  - Excel 实例意外崩溃/被手动结束时 (com_error), 自动重启实例并重试一次
  - 解释器退出时经 atexit 统一 Quit, 不留僵尸 EXCEL.EXE

线程安全: 所有 COM 调用都封送到唯一的 Excel 工作线程串行执行
(pywebview 的 js_api 回调来自任意线程, Tk 版后台线程也直接用),
并发调用在队列里排队, 不会打架。
xlsx 不在本项目范围。

日期单元格 COM 读出来是 datetime, 统一转回 "YYYY-MM-DD HH:MM:SS"
字符串, 对齐旧版 (xlrd 时代) 文本单元格的行为, 上游 strptime 不用改。
"""

from __future__ import annotations

import atexit
import queue
import threading
from concurrent.futures import Future
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Iterable

import pywintypes

from trajectory_splitter.logging_setup import get_logger


logger = get_logger(__name__)


# COM ProgID 依次尝试: MS Excel 优先, 后两个是 WPS 表格的兼容注册名
_PROGIDS = ("Excel.Application", "ket.Application", "et.Application")

# Excel 单元格错误值 (COM 读出来是这些负整数), 转回错误串
_EXCEL_ERRORS = {
    -2146826281: "#DIV/0!",
    -2146826246: "#N/A",
    -2146826259: "#NAME?",
    -2146826288: "#NULL!",
    -2146826252: "#NUM!",
    -2146826265: "#REF!",
    -2146826273: "#VALUE!",
}

# SaveAs FileFormat 56 = xlExcel8 (.xls 97-2003)
_XLS_FILE_FORMAT = 56


# ── 常驻 Excel 工作线程 ─────────────────────────────────────────
# 所有 Excel COM 对象只在 _excel_worker_loop 线程里创建和访问 (COM 单元
# 线程亲和), 其它线程通过 _submit() 排队投递任务并阻塞等结果。
_queue: "queue.Queue[tuple[Callable[[], Any] | None, Future]]" = queue.Queue()
_worker_thread: threading.Thread | None = None
_worker_lock = threading.Lock()
_app = None  # Excel Application 代理, 仅工作线程内访问


def _start_excel():
    """新起一个独立 Excel 进程 (DispatchEx), 不碰用户自己开着的 Excel。"""
    import win32com.client as win32

    last_err: Exception | None = None
    for progid in _PROGIDS:
        try:
            app = win32.DispatchEx(progid)
            app.Visible = False
            app.DisplayAlerts = False
            app.ScreenUpdating = False
            logger.info("xls_io: Excel COM 实例已启动 (%s, 常驻复用)", progid)
            return app
        except Exception as e:
            last_err = e
    raise RuntimeError(
        f"无法启动 Excel COM (机器上没装 Excel / WPS 表格?): {last_err}"
    ) from last_err


def _quit_app() -> None:
    """关闭并遗忘常驻实例 (仅工作线程内调用)。"""
    global _app
    app = _app
    _app = None
    if app is None:
        return
    try:
        app.Quit()
    except Exception as e:
        logger.warning("xls_io: Excel Quit 失败 (进程可能已被手动结束): %s", e)


def _excel_worker_loop() -> None:
    import pythoncom

    pythoncom.CoInitialize()
    try:
        while True:
            fn, future = _queue.get()
            if fn is None:  # 关机药丸: 退出 Excel, 确认后收工
                _quit_app()
                future.set_result(True)
                return
            if not future.set_running_or_notify_cancel():
                continue
            try:
                future.set_result(fn())
            except BaseException as e:  # 原样传回调用线程
                future.set_exception(e)
    finally:
        pythoncom.CoUninitialize()


def _ensure_worker() -> None:
    global _worker_thread
    with _worker_lock:
        if _worker_thread is not None and _worker_thread.is_alive():
            return
        _worker_thread = threading.Thread(
            target=_excel_worker_loop, name="excel-com-worker", daemon=True,
        )
        _worker_thread.start()


def _submit(fn: Callable[[], Any]) -> Any:
    """把 fn 投递到 Excel 工作线程执行并阻塞等结果, 异常原样抛回。"""
    _ensure_worker()
    future: Future = Future()
    _queue.put((fn, future))
    return future.result()


def _with_retry(fn: Callable[[], Any]) -> Any:
    """执行 fn; com_error (Excel 崩溃/被杀) 时重启常驻实例再试一次。"""
    try:
        return _submit(fn)
    except pywintypes.com_error as e:
        logger.warning("xls_io: COM 调用失败 (%s), 重启 Excel 实例后重试一次", e)
        try:
            _submit(_quit_app)
        except Exception:
            pass
        return _submit(fn)


def _shutdown_worker() -> None:
    """解释器退出时关闭常驻 Excel (atexit 注册, 只在主线程跑一次)。"""
    t = _worker_thread
    if t is None or not t.is_alive():
        return
    ack: Future = Future()
    _queue.put((None, ack))
    try:
        ack.result(timeout=10)
    except Exception:
        logger.warning("xls_io: 退出时 Excel 工作线程未在 10s 内确认关闭")


atexit.register(_shutdown_worker)


# ── 数值/表单转换 ───────────────────────────────────────────────
def _from_com(v: Any) -> Any:
    if v is None:
        return ""  # 对齐 xlrd: 空单元格给 ''
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, int) and not isinstance(v, bool) and v in _EXCEL_ERRORS:
        return _EXCEL_ERRORS[v]
    return v


def _read_sheet(ws) -> list[list[Any]]:
    """把一个 worksheet 从 A1 读到 UsedRange 末尾, 归一化成 list[list]。"""
    ur = ws.UsedRange
    last_row = ur.Row + ur.Rows.Count - 1
    last_col = ur.Column + ur.Columns.Count - 1
    if last_row < 1 or last_col < 1:
        return []
    vals = ws.Range(ws.Cells(1, 1), ws.Cells(last_row, last_col)).Value
    if vals is None:
        return []

    # Value 的形状: 单格=标量 / 单行=1D tuple / 单列=tuple of 1-tuple / 其它=2D
    if not isinstance(vals, tuple):
        rows: list[list[Any]] = [[vals]]
    elif last_row == 1:
        rows = [list(vals)]
    elif last_col == 1:
        rows = [[v[0] if isinstance(v, tuple) else v for v in vals]]
    else:
        rows = [list(r) for r in vals]

    # 掐掉末尾整行空白 (UsedRange 可能带格式撑出来的空行; xlrd 不会给)
    while rows and all(v in ("", None) for v in rows[-1]):
        rows.pop()
    return [[_from_com(v) for v in row] for row in rows]


# ── 公开 API (签名与旧版一致) ───────────────────────────────────
def read_xls(path: str | Path) -> dict[str, list[list[Any]]]:
    """读 .xls 返回 {sheet_name: [[row], ...]} (含表头)。

    sheet 按工作簿顺序返回; 调用方仍应按内容而非名字定位 sheet
    (本项目其它模块用 `_find_header_col` 按表头内容识别)。
    """
    path_str = str(Path(path).resolve())
    return _with_retry(lambda: _read_xls_impl(path_str))


def _read_xls_impl(path_str: str) -> dict[str, list[list[Any]]]:
    global _app
    if _app is None:
        _app = _start_excel()
    wb = _app.Workbooks.Open(path_str, UpdateLinks=0, ReadOnly=True)
    try:
        out: dict[str, list[list[Any]]] = {}
        for i in range(1, wb.Worksheets.Count + 1):
            ws = wb.Worksheets(i)
            out[ws.Name] = _read_sheet(ws)
        return out
    finally:
        wb.Close(SaveChanges=False)
        wb = None


def write_xls(
    path: str | Path,
    sheet_name: str,
    rows: Iterable[Iterable[Any]],
    *,
    bold_header: bool = True,
) -> None:
    """写单 sheet .xls。第一行可选加粗表头。

    文本列 (数据里出现非数字值) 会把整列设成文本格式 "@", 避免
    "2026-08-25 12:00:00" 被 Excel 吃成日期、车牌前导零丢失; 纯数字
    列 (序号/温度) 保持常规格式写成数值。
    """
    path_str = str(Path(path).resolve())
    rows_list = [list(r) for r in rows]
    _with_retry(lambda: _write_xls_impl(path_str, sheet_name, rows_list, bold_header=bold_header))


def _write_xls_impl(
    path_str: str,
    sheet_name: str,
    rows_list: list[list[Any]],
    *,
    bold_header: bool,
) -> None:
    global _app
    if _app is None:
        _app = _start_excel()
    wb = _app.Workbooks.Add()
    ws = None
    try:
        # 多余 sheet 删掉, 只留第一个 (对齐 xlwt 单 sheet 行为)
        for i in range(wb.Worksheets.Count, 1, -1):
            wb.Worksheets(i).Delete()
        ws = wb.Worksheets(1)
        try:
            ws.Name = sheet_name
        except Exception:
            logger.warning("write_xls: sheet 重命名失败, 保留默认名 (%r)", sheet_name)

        if rows_list:
            width = max(len(r) for r in rows_list)
            height = len(rows_list)

            def _is_num(v: Any) -> bool:
                return isinstance(v, (int, float)) and not isinstance(v, bool)

            for c in range(1, width + 1):
                col_vals = [
                    r[c - 1] for r in rows_list[1:]
                    if c <= len(r) and r[c - 1] not in ("", None)
                ]
                if col_vals and not all(_is_num(v) for v in col_vals):
                    ws.Range(ws.Cells(1, c), ws.Cells(height, c)).NumberFormat = "@"

            data = tuple(
                tuple((r[c] if c < len(r) else "") for c in range(width))
                for r in rows_list
            )
            ws.Range(ws.Cells(1, 1), ws.Cells(height, width)).Value = data
            if bold_header:
                ws.Range(ws.Cells(1, 1), ws.Cells(1, width)).Font.Bold = True

        wb.SaveAs(path_str, FileFormat=_XLS_FILE_FORMAT)
    finally:
        ws = None
        wb.Close(SaveChanges=False)
        wb = None
