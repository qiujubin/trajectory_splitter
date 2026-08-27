"""xls 读写工具。

读用 xlrd (注意: xlrd 2.x 还能读 .xls, 但文档已 deprecated), 写用 xlwt。
xlsx 不在本项目范围。
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Iterable

import xlrd
import xlwt


def read_xls(path: str | Path) -> dict[str, list[list[Any]]]:
    """读 .xls 返回 {sheet_name: [[row], ...]} (含表头)。

    xlrd 2.x 在 Windows + 中文 sheet 名下可能拿到乱码 sheet 名, 调用方
    应该按内容而非名字定位 sheet (本项目其它模块用 `_find_header_col`
    按表头内容识别)。
    """
    wb = xlrd.open_workbook(str(path))
    out: dict[str, list[list[Any]]] = {}
    for s in wb.sheets():
        rows: list[list[Any]] = []
        for r in range(s.nrows):
            rows.append([s.cell_value(r, c) for c in range(s.ncols)])
        out[s.name] = rows
    return out


def write_xls(
    path: str | Path,
    sheet_name: str,
    rows: Iterable[Iterable[Any]],
    *,
    bold_header: bool = True,
) -> None:
    """写单 sheet .xls。第一行可选加粗表头。"""
    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet(sheet_name)
    rows_list = list(rows)
    if not rows_list:
        wb.save(str(path))
        return

    header_style = xlwt.easyxf("font: bold on") if bold_header else xlwt.easyxf()
    body_style = xlwt.easyxf()
    for r_idx, row in enumerate(rows_list):
        for c_idx, val in enumerate(row):
            ws.write(r_idx, c_idx, val, header_style if r_idx == 0 else body_style)
    wb.save(str(path))
