"""双文件拼接: 把两份 .xls 按 [日前 + 后半] 顺序拼成一个, 时间接续, 文件名按参考文件夹命名。

时间接续规则:
  - 锚点 = 日前最后一条时间 + gap_minutes
  - 日后第一行新时间 = 锚点
  - 日后整体时间偏移 offset = 锚点 - 日后第一行原时间
  - 通常 offset < 0 (日后第一行原时间 > 锚点, 即日后数据时间更晚, 平移让它"提前"接续)
  - 如果日后第一行原时间 < 日前最后一行原时间, 说明切片顺序错了, 抛 ValueError
  - 跨天 (e.g. 前半 23:55, 后半原时间 00:05 第二天) 也走同一套逻辑:
    timedelta 自动跨天, 排序仍然按 datetime 比较。

温度: 超 [0, 4] 范围按传感器故障模式分段线性映射 (不用随机数, 保持温度曲线自然),
      输出统一保留 1 位小数。详见 splitter._clamp_temp。

输出文件名模板: 默认 `1日前.xls` (沿用参考文件夹 `N+方向.xls` 命名), 支持 {plate}。
"""

from __future__ import annotations

from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

from trajectory_splitter.splitter import (
    COL_LOC,
    COL_PLATE,
    COL_TEMP1,
    COL_TEMP2,
    COL_TIME,
    OUTPUT_COLUMNS,
    _find_header_col,
    _project_rows,
    _read_main_sheet,
    pick_temperature_column,
)
from trajectory_splitter.logging_setup import get_logger
from trajectory_splitter.xls_io import write_xls


logger = get_logger(__name__)


def _safe_for_filename(s: str) -> str:
    return "".join(c for c in str(s) if c.isalnum() or c in "_-")


def _parse_dt(s: Any) -> datetime | None:
    if not s:
        return None
    try:
        return datetime.strptime(str(s)[:19], "%Y-%m-%d %H:%M:%S")
    except ValueError:
        return None


def merge_two(
    first_path: str | Path,
    second_path: str | Path,
    out_dir: str | Path,
    *,
    time_gap_minutes: int = 1,
    clamp_temp_to_0to4: bool = True,
    filename_template: str = "1日前.xls",
    dry_run: bool = False,
) -> Path:
    """拼接两个 .xls (日前 + 后半), 时间接续, 温度夹紧。

    Args:
        first_path: 日前 .xls (车牌号从该文件抽取)。
        second_path: 后半 .xls。
        out_dir: 输出目录。
        time_gap_minutes: 日后第一行新时间相对日前最后一条的间隔 (分钟)。
        clamp_temp_to_0to4: 温度超 [0,4] 是否夹紧。
        filename_template: 输出文件名模板, 支持 {plate} (日前车牌)。
    """
    first_path = Path(first_path)
    second_path = Path(second_path)
    out_dir = Path(out_dir)
    if not dry_run:
        out_dir.mkdir(parents=True, exist_ok=True)

    f_hdr, f_data = _read_main_sheet(first_path)
    s_hdr, s_data = _read_main_sheet(second_path)

    f_cols = {
        "plate": _find_header_col(f_hdr, COL_PLATE),
        "time": _find_header_col(f_hdr, COL_TIME, "时间"),
        "temp1": _find_header_col(f_hdr, COL_TEMP1, "温度1"),
        "temp2": _find_header_col(f_hdr, COL_TEMP2, "温度2"),
        "loc": _find_header_col(f_hdr, COL_LOC),
    }
    s_cols = {
        "plate": _find_header_col(s_hdr, COL_PLATE),
        "time": _find_header_col(s_hdr, COL_TIME, "时间"),
        "temp1": _find_header_col(s_hdr, COL_TEMP1, "温度1"),
        "temp2": _find_header_col(s_hdr, COL_TEMP2, "温度2"),
        "loc": _find_header_col(s_hdr, COL_LOC),
    }

    if f_cols["plate"] is None:
        raise ValueError("前半文件找不到车牌号列")
    if not f_data:
        raise ValueError("前半文件无数据")

    plate = f_data[0][f_cols["plate"]]
    safe_plate = _safe_for_filename(plate) or "unknown"

    # 把后半所有行的车牌号统一改成前半的车牌 (跨车拼接场景)
    if s_cols["plate"] is not None:
        s_data = [
            [plate if i == s_cols["plate"] else v for i, v in enumerate(row)]
            for row in s_data
        ]
    s_cols["plate"] = f_cols["plate"]

    f_temp_col = pick_temperature_column(f_data, f_cols["temp1"], f_cols["temp2"])
    s_temp_col = pick_temperature_column(s_data, s_cols["temp1"], s_cols["temp2"])

    # 时间接续
    f_last_dt: datetime | None = None
    if f_cols["time"] is not None:
        for row in reversed(f_data):
            t = _parse_dt(row[f_cols["time"]])
            if t is not None:
                f_last_dt = t
                break
    s_first_dt: datetime | None = None
    if s_cols["time"] is not None:
        for row in s_data:
            t = _parse_dt(row[s_cols["time"]])
            if t is not None:
                s_first_dt = t
                break

    if f_last_dt is None or s_first_dt is None:
        raise ValueError("两份文件里至少一份没有可解析的定位时间")

    anchor = f_last_dt + timedelta(minutes=time_gap_minutes)
    offset = anchor - s_first_dt

    # 验证: 日后第一行原时间不能早于日前最后一行原时间 (否则切片顺序错)
    if s_first_dt < f_last_dt:
        raise ValueError(
            f"日后第一行时间 {s_first_dt} 早于日前最后一行时间 {f_last_dt}, "
            "说明切片顺序错了 (应该先日前再后半)"
        )

    # 平移 second 时间
    s_shifted: list[list[Any]] = []
    if s_cols["time"] is not None and offset != timedelta(0):
        for row in s_data:
            new_row = list(row)
            t = _parse_dt(new_row[s_cols["time"]])
            if t is not None:
                new_t = (t + offset).strftime("%Y-%m-%d %H:%M:%S")
                new_row[s_cols["time"]] = new_t
            s_shifted.append(new_row)
    else:
        s_shifted = [list(r) for r in s_data]

    out_rows: list[list[Any]] = [OUTPUT_COLUMNS]
    f_out = _project_rows(
        f_data, f_cols, f_temp_col,
        clamp_temp_to_0to4=clamp_temp_to_0to4, start_seq=1,
    )
    out_rows.extend(f_out)
    s_out = _project_rows(
        s_shifted, s_cols, s_temp_col,
        clamp_temp_to_0to4=clamp_temp_to_0to4, start_seq=len(f_out) + 1,
    )
    out_rows.extend(s_out)

    out_name = filename_template.format(plate=safe_plate)
    out_path = out_dir / out_name
    logger.info(
        "merge_two: 锚点=%s offset=%s clamp=%s 输出 %s 共 %d 行%s",
        anchor, offset, clamp_temp_to_0to4, out_path, len(out_rows),
        " (DRY-RUN)" if dry_run else "",
    )
    if not dry_run:
        write_xls(out_path, "轨迹", out_rows)
    return out_path
