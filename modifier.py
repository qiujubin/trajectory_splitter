"""数据微改: 输入 .xls → 输出 车牌替换 + 时间偏移/微抖 + 温度微变, 严格递增。

不改业务列 (位置不动, 序号重排)。

时间微抖: 整体偏移 (±N 分钟) + 每行正态微抖 (μ=0, σ=time_jitter_seconds),
         然后强制严格递增 (如果新时间 ≤ 上一行, 用上一行 + 1秒)。
温度微变: 每行均匀随机 ±temp_jitter 度, 输出走 _clamp_temp 映射, 仍在 [0,4]。
车牌替换: 纯字符串替换, 列里每一格的 cell 整串替换为新值。

设置面板留默认值 (时间 ±5-30 秒, 温度 ±0.2-0.5 度)。
"""

from __future__ import annotations

import random
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

from trajectory_splitter.logging_setup import get_logger
from trajectory_splitter.splitter import (
    COL_LOC,
    COL_PLATE,
    COL_TEMP1,
    COL_TEMP2,
    COL_TIME,
    OUTPUT_COLUMNS,
    _DATE_FMT_RE,
    _clamp_temp,
    _find_header_col,
    _format_filename,
    _read_main_sheet,
    find_main_date,
    pick_temperature_column,
)
from trajectory_splitter.xls_io import read_xls, write_xls


logger = get_logger(__name__)


@dataclass
class ModifyOptions:
    """微改参数。"""
    time_offset_minutes: int = 0       # 整体偏移 (分钟, 可正可负)
    time_jitter_seconds: float = 8.0   # 每行正态微抖 σ (秒), μ=0
    temp_jitter: float = 0.3           # 每行温度均匀随机 ±temp_jitter (度)
    plate_old: str = ""                # 要替换的旧车牌 (空 = 不替换)
    plate_new: str = ""                # 替换成的新车牌 (空 = 不替换)
    filename_template: str = "1modified.xls"
    clamp_temp_to_0to4: bool = True    # 微变后再夹紧到 [0,4] 映射规则
    seed: int | None = None            # 随机种子 (None = 真随机)
    dry_run: bool = False


def modify_xls(
    src_path: str | Path,
    out_dir: str | Path,
    options: ModifyOptions | None = None,
) -> Path:
    """微改单文件: 车牌替换 + 时间偏移/微抖 + 温度微变。

    Args:
        src_path: 输入 .xls 路径。
        out_dir: 输出目录。
        options: ModifyOptions; None 用默认 (只整体偏移 0, 不微抖)。

    Returns:
        输出 .xls 路径 (dry_run 时也不算 makedirs / 写文件)。
    """
    src_path = Path(src_path)
    out_dir = Path(out_dir)
    opts = options or ModifyOptions()

    if not opts.dry_run:
        out_dir.mkdir(parents=True, exist_ok=True)

    if opts.seed is not None:
        rng_normal = random.Random(opts.seed)
        rng_uniform = random.Random(opts.seed + 1)  # 独立序列
    else:
        rng_normal = random.Random()
        rng_uniform = random.Random()

    headers, data_rows = _read_main_sheet(src_path)
    cols = {
        "plate": _find_header_col(headers, COL_PLATE),
        "time": _find_header_col(headers, COL_TIME, "时间"),
        "temp1": _find_header_col(headers, COL_TEMP1, "温度1"),
        "temp2": _find_header_col(headers, COL_TEMP2, "温度2"),
        "loc": _find_header_col(headers, COL_LOC),
    }
    if cols["plate"] is None:
        raise ValueError("找不到车牌号列")
    if cols["time"] is None:
        raise ValueError("找不到定位时间列")
    if cols["loc"] is None:
        raise ValueError("找不到位置列")

    temp_col = pick_temperature_column(data_rows, cols["temp1"], cols["temp2"])
    time_col = cols["time"]
    plate_col = cols["plate"]

    # 1) 车牌替换
    plate_changes = 0
    if opts.plate_old and opts.plate_new:
        for row in data_rows:
            if plate_col < len(row) and str(row[plate_col]) == opts.plate_old:
                row[plate_col] = opts.plate_new
                plate_changes += 1

    # 2) 时间: 整体偏移 + 每行正态微抖 + 严格递增
    overall = timedelta(minutes=opts.time_offset_minutes)
    prev_dt: datetime | None = None
    parsed: list[tuple[datetime, float]] = []  # (原dt, 抖动秒数)
    for row in data_rows:
        if time_col >= len(row):
            parsed.append((datetime.min, 0.0))
            continue
        raw = row[time_col]
        s = str(raw)[:19]
        try:
            dt = datetime.strptime(s, "%Y-%m-%d %H:%M:%S")
        except ValueError:
            parsed.append((datetime.min, 0.0))
            continue
        # 正态微抖 (gauss σ)
        jitter = rng_normal.gauss(0.0, opts.time_jitter_seconds)
        new_dt = dt + overall + timedelta(seconds=jitter)
        # 强制严格递增 (微抖后 + 整体偏移可能让个别行 < 上一行)
        if prev_dt is not None and new_dt <= prev_dt:
            new_dt = prev_dt + timedelta(seconds=1)
        parsed.append((new_dt, jitter))
        prev_dt = new_dt

    # 3) 温度微变 (均匀随机 ±Y)
    out_rows: list[list[Any]] = [OUTPUT_COLUMNS]
    for i, row in enumerate(data_rows):
        plate = row[plate_col] if plate_col < len(row) else ""
        new_dt, _jitter = parsed[i] if i < len(parsed) else (datetime.min, 0.0)
        if new_dt == datetime.min:
            time_str = ""
        else:
            time_str = new_dt.strftime("%Y-%m-%d %H:%M:%S")
        loc = row[cols["loc"]] if cols["loc"] < len(row) else ""
        loc = "" if loc is None else loc

        # 温度微变
        temp_val: Any = ""
        if temp_col is not None and temp_col < len(row):
            v = row[temp_col]
            if v not in ("", None):
                base = _clamp_temp(v) if opts.clamp_temp_to_0to4 else v
                try:
                    f = float(base)
                    delta = rng_uniform.uniform(-opts.temp_jitter, opts.temp_jitter)
                    new_f = f + delta
                    temp_val = _clamp_temp(new_f) if opts.clamp_temp_to_0to4 else new_f
                except (TypeError, ValueError):
                    temp_val = base

        out_rows.append([i + 1, plate, time_str, temp_val, loc])

    # 文件名: 车牌 (替换后的第一个) 走 {plate}
    first_plate = out_rows[1][1] if len(out_rows) > 1 else "unknown"
    safe_plate = "".join(c for c in str(first_plate) if c.isalnum() or c in "_-") or "unknown"
    out_name = _format_filename(
        opts.filename_template,
        plate=safe_plate,
        date=find_main_date(src_path)[0] if _DATE_FMT_RE.search(opts.filename_template) else None,
    )
    # modify 模板里没有 {direction} 语义
    if "{direction}" in out_name or "{direction:" in out_name:
        raise ValueError(
            f"modify_xls 的 filename_template 不支持 {{direction}}: {opts.filename_template!r}"
        )
    out_path = out_dir / out_name

    logger.info(
        "modify_xls: 车牌替换=%d 行 整体偏移=%+d 分钟 微抖σ=%.1f秒 温度±%.2f度 输出=%s 行=%d%s",
        plate_changes, opts.time_offset_minutes, opts.time_jitter_seconds,
        opts.temp_jitter, out_path, len(out_rows),
        " (DRY-RUN)" if opts.dry_run else "",
    )

    if not opts.dry_run:
        write_xls(out_path, "轨迹", out_rows)
    return out_path


def build_default_options() -> ModifyOptions:
    """默认参数 (时间微抖 ±5-30 秒 σ=8, 温度 ±0.3 度)。"""
    return ModifyOptions()