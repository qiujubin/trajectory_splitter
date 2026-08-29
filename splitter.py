"""单文件切片: 输入完整 .xls → 按 (start_poi, end_poi) 切成日前段 / 日后段。

业务场景 (春梅里朴朴配送):
  - 文件头 → ... → 【麻涌卸货地址】 → ... → 【广州朴朴】
  - 日前 (前半段) = 文件头 → 第一次出现 end_poi (含)
    (start_poi 留空, 表示从文件头开始)
  - 日后 (后半段) = 最后一次出现 start_poi → 第一次出现 end_poi (含)
    (后半段可加 noon_cutoff: 只取 noon_cutoff 之前的最后一次 start_poi,
     业务上"中午前最后一次到卸货点")

POI 字符串可配置 (settings.yaml / 函数参数), 默认值是春梅里业务场景。

温度: 超 [0, 4] 范围按传感器故障模式分段线性映射 (不用随机数, 保持温度曲线自然),
      输出统一保留 1 位小数。详见 _clamp_temp 注释。
"""

from __future__ import annotations

from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Any, Literal

from trajectory_splitter.logging_setup import get_logger
from trajectory_splitter.xls_io import read_xls, write_xls


logger = get_logger(__name__)


import re as _re
from datetime import datetime as _datetime


# 文件名模板占位符 (按用户口述):
#   {direction}     切片方向 (日前/后半)
#   {plate}         第一行的车牌号 (sanitize 后的安全字符串)
#   {date}          主要日期 (YYYY-MM-DD), 可用 {date:%m%d} 改格式 (strftime)
#   其它 key 会抛 KeyError (模板写错时立刻报错)
#
# {date:fmt} 走 strptime("YYYY-MM-DD") + strftime(fmt); fmt 为空 = 默认 YYYY-MM-DD
_DATE_FMT_RE = _re.compile(r"\{date(?::([^}]*))?\}")


def _format_filename(
    template: str,
    *,
    direction: str | None = None,
    plate: str | None = None,
    date: str | None = None,
) -> str:
    # 先处理 {date[:fmt]} 占位符 (str 类型的 format 不支持 strftime)
    def _date_sub(m: _re.Match[str]) -> str:
        if date is None:
            raise KeyError("date")
        fmt = m.group(1) or "%Y-%m-%d"
        try:
            return _datetime.strptime(date, "%Y-%m-%d").strftime(fmt)
        except ValueError as e:
            raise ValueError(f"模板里的 {m.group(0)!r} 无法解析 date={date!r}: {e}") from None

    out = _DATE_FMT_RE.sub(_date_sub, template)

    fmt_args: dict[str, str] = {}
    if direction is not None:
        fmt_args["direction"] = direction
    if plate is not None:
        fmt_args["plate"] = plate
    # 不再让 str.format 处理 {date}, 已 sub 过了
    return out.format(**fmt_args)


COL_PLATE = "车牌号"
COL_TIME = "定位时间"
COL_TEMP1 = "温度1(℃)"
COL_TEMP2 = "温度2(℃)"
COL_LOC = "位置"

OUTPUT_COLUMNS = ["序号", COL_PLATE, COL_TIME, "温度2(℃)", COL_LOC]

# "日前" = 前半段, "后半" = 后半段 (沿用历史命名)
Direction = Literal["日前", "后半"]


def default_start_poi(direction: Direction) -> str:
    """方向对应的默认 start_poi。日前留空 = 从文件头切。"""
    if direction == "日前":
        return ""
    return "【麻涌卸货地址】"


def default_end_poi(direction: Direction) -> str:
    """方向对应的默认 end_poi。"""
    if direction == "日前":
        return "【麻涌卸货地址】"
    return "【广州朴朴】"


# 保留兼容旧 import 名 (CLI / settings.yaml 可能还在引用)
DEFAULT_POI_START = "【麻涌卸货地址】"
DEFAULT_POI_END = "【广州朴朴】"


def _extract_date(time_str: Any) -> str:
    if not time_str:
        return ""
    return str(time_str)[:10]


def _extract_time_of_day(time_str: Any) -> str:
    """'2026-08-24 17:09:44' -> '17:09:44'。"""
    s = str(time_str) if time_str else ""
    if len(s) < 19:
        return ""
    return s[11:19]


def _find_header_col(headers: list[Any], *candidates: str) -> int | None:
    for i, h in enumerate(headers):
        text = str(h)
        for cand in candidates:
            if cand in text:
                return i
    return None


def _read_main_sheet(path: Path) -> tuple[list[Any], list[list[Any]]]:
    sheets = read_xls(path)
    for rows in sheets.values():
        if rows:
            return rows[0], rows[1:]
    raise ValueError(f"输入文件无任何 sheet 数据: {path}")


def _try_float(v: Any) -> float | None:
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def find_main_date(path: str | Path) -> tuple[str, int]:
    """返回 (主要日期, 记录数)。主要日期 = 文件里记录数最多的那一天。"""
    headers, data_rows = _read_main_sheet(Path(path))
    time_col = _find_header_col(headers, COL_TIME, "时间")
    if time_col is None:
        raise ValueError("找不到定位时间列")
    counter = Counter(_extract_date(r[time_col]) for r in data_rows if r[time_col])
    if not counter:
        raise ValueError("输入文件无定位时间数据")
    date, count = counter.most_common(1)[0]
    return date, count


def pick_temperature_column(
    rows: list[list[Any]],
    temp1_col: int | None,
    temp2_col: int | None,
) -> int | None:
    """决定使用哪一列温度。

    - 只有 temp1 有数据 -> temp1_col
    - 只有 temp2 有数据 -> temp2_col
    - 都有 -> 选 0-4 度范围内出现次数较多的那列
    - 都没数据 -> temp1_col (fallback, 输出时会留空)
    """
    t1_in = t1_total = 0
    t2_in = t2_total = 0
    for row in rows:
        if temp1_col is not None and temp1_col < len(row):
            f = _try_float(row[temp1_col])
            if f is not None:
                t1_total += 1
                if 0 <= f <= 4:
                    t1_in += 1
        if temp2_col is not None and temp2_col < len(row):
            f = _try_float(row[temp2_col])
            if f is not None:
                t2_total += 1
                if 0 <= f <= 4:
                    t2_in += 1

    if t1_total == 0 and t2_total == 0:
        return temp1_col if temp1_col is not None else temp2_col
    if t1_total == 0:
        return temp2_col
    if t2_total == 0:
        return temp1_col
    return temp1_col if t1_in >= t2_in else temp2_col


def _clamp_temp(v: Any, low: float = 0.0, high: float = 4.0) -> Any:
    """按传感器故障模式把 v 映射回 [low, high] 范围, 非数值返回原值。

    映射规则 (按你的口述):
      - [0, 4]       原值 (传感器正常)
      - (4, 12]      反向映射回 (4, 0]   (偏高但不至于彻底坏)
      - (12, 28)     正向映射回 [0, 4)   (偏高厉害)
      - [28, +inf)   固定 4              (彻底坏)
      - [-8, 0)      反向映射回 (0, 4]   (偏低但不至于彻底坏)
      - (-inf, -8)   固定 0              (彻底坏)

    输出统一保留 1 位小数 (跟你选的精度)。
    """
    f = _try_float(v)
    if f is None:
        return v
    if low <= f <= high:
        out = f
    elif high < f <= 12:
        # (4, 12] -> (4, 0]: f(v) = 6 - 0.5*v   (4->4, 12->0)
        out = 6.0 - 0.5 * f
    elif 12 < f < 28:
        # (12, 28) -> [0, 4): f(v) = 0.25*v - 3
        out = 0.25 * f - 3.0
    elif f >= 28:
        out = high
    elif -8 <= f < low:
        # [-8, 0) -> (0, 4]: f(v) = -0.5*v
        out = -0.5 * f
    else:  # f < -8
        out = low
    return round(out, 1)


def _project_rows(
    rows: list[list[Any]],
    cols: dict[str, int | None],
    temp_col: int | None,
    *,
    clamp_temp_to_0to4: bool = False,
    start_seq: int = 1,
) -> list[list[Any]]:
    """把任意输入行映射到固定 5 列输出, 重编号, 可选夹紧温度到 [0,4]。"""

    def _get(key: str, row: list[Any]) -> Any:
        i = cols.get(key)
        if i is None or i >= len(row):
            return ""
        v = row[i]
        return "" if v is None else v

    out: list[list[Any]] = []
    seq = start_seq
    for row in rows:
        plate = _get("plate", row)
        time_v = _get("time", row)
        loc = _get("loc", row)
        temp: Any = ""
        if temp_col is not None and temp_col < len(row):
            v = row[temp_col]
            if v in ("", None):
                temp = ""
            elif clamp_temp_to_0to4:
                temp = _clamp_temp(v)
            else:
                temp = v
        out.append([seq, plate, time_v, temp, loc])
        seq += 1
    return out


def _loc_at(row: list[Any], loc_col: int | None) -> str:
    if loc_col is None or loc_col >= len(row):
        return ""
    return str(row[loc_col]) if row[loc_col] is not None else ""


def _time_at(row: list[Any], time_col: int | None) -> str:
    if time_col is None or time_col >= len(row):
        return ""
    return str(row[time_col]) if row[time_col] is not None else ""


def round_to_minute_drop_collisions(out_rows: list[list[Any]]) -> list[list[Any]]:
    """整点化去重: 把 out_rows[1:] 时间列秒清零, 同分钟内后续重复行丢弃, 序号重排。

    - 输入约定: out_rows[0] = 表头 (OUTPUT_COLUMNS), 之后是数据行
    - 数据行时间列在 index 2 (= COL_TIME)
    - 整点化前先按时间排序避免乱序丢行? 不: 调用方负责传入已排好序的行, 本函数不做排序
    - 时间不可解析的行原样保留 (不参与整点化, 也不丢)
    - 返回新 list (不修改入参)
    """
    if len(out_rows) <= 1:
        return list(out_rows)

    header = out_rows[0]
    TIME_COL = 2
    seen_minutes: set[tuple[int, int, int, int, int]] = set()
    kept: list[list[Any]] = []
    for row in out_rows[1:]:
        if TIME_COL >= len(row):
            kept.append(row)
            continue
        t = str(row[TIME_COL])[:19]
        try:
            dt = datetime.strptime(t, "%Y-%m-%d %H:%M:%S")
        except ValueError:
            kept.append(row)
            continue
        key = (dt.year, dt.month, dt.day, dt.hour, dt.minute)
        if key in seen_minutes:
            logger.info(
                "round_to_minute: 丢弃 %s (与已保留的同分钟行重复)",
                t,
            )
            continue
        seen_minutes.add(key)
        new_row = list(row)
        new_row[TIME_COL] = dt.strftime("%Y-%m-%d %H:%M:00")
        kept.append(new_row)

    # 序号重排
    for i, row in enumerate(kept, start=1):
        if row:
            row[0] = i

    logger.info(
        "round_to_minute: 输入 %d 行, 去重保留 %d 行",
        len(out_rows) - 1, len(kept),
    )
    return [header] + kept


def split_at_poi(
    src_path: str | Path,
    direction: Direction,
    out_dir: str | Path,
    *,
    start_poi: str | None = None,
    end_poi: str | None = None,
    noon_cutoff: str | None = "12:00:00",
    clamp_temp_to_0to4: bool = True,
    filename_template: str = "1{direction}.xls",
    round_to_minute: bool = False,
    dry_run: bool = False,
) -> Path:
    """按 (start_poi, end_poi) 把 src_path 切成两段。

    direction="日前":
        start_poi 为空 / None  -> 从文件头切到第一次出现 end_poi (含)
        start_poi 非空         -> 从第一次出现 start_poi (含) 切到第一次出现 end_poi (含)
        (注意: 日前场景下 start_poi 通常留空, 表示"从头")

    direction="后半":
        从最后一次出现 start_poi (时间 < noon_cutoff) 切到第一次出现 end_poi (含)
        noon_cutoff="" 或 None  -> 不限时间, 取最后一次 start_poi

    默认 POI 取自 default_start_poi / default_end_poi (按 direction)。
    round_to_minute=True 时: 时间列秒清零, 同分钟内后续重复行直接丢弃, 序号重排。
    dry_run=True 时不算 makedirs、不写文件, 只返回将输出的路径。
    """
    src_path = Path(src_path)
    out_dir = Path(out_dir)
    if not dry_run:
        out_dir.mkdir(parents=True, exist_ok=True)

    if start_poi is None:
        start_poi = default_start_poi(direction)
    if end_poi is None:
        end_poi = default_end_poi(direction)
    if end_poi == "":
        raise ValueError(f"end_poi 不能为空, direction={direction}")
    if direction == "后半" and not start_poi:
        raise ValueError("direction='后半' 必须指定 start_poi (后半段起点)")

    logger.info(
        "split_at_poi: direction=%s start_poi=%r end_poi=%r noon_cutoff=%r "
        "clamp_temp_to_0to4=%s src=%s out_dir=%s",
        direction, start_poi, end_poi, noon_cutoff, clamp_temp_to_0to4, src_path, out_dir,
    )

    headers, data_rows = _read_main_sheet(src_path)
    cols = {
        "plate": _find_header_col(headers, COL_PLATE),
        "time": _find_header_col(headers, COL_TIME, "时间"),
        "temp1": _find_header_col(headers, COL_TEMP1, "温度1"),
        "temp2": _find_header_col(headers, COL_TEMP2, "温度2"),
        "loc": _find_header_col(headers, COL_LOC),
    }
    if cols["time"] is None:
        raise ValueError("找不到定位时间列")
    if cols["loc"] is None:
        raise ValueError("找不到位置列")

    temp_col = pick_temperature_column(data_rows, cols["temp1"], cols["temp2"])
    loc_col = cols["loc"]
    time_col = cols["time"]

    # 找 end_poi 第一次出现
    first_end_idx: int | None = None
    for i, row in enumerate(data_rows):
        loc = _loc_at(row, loc_col)
        if end_poi in loc:
            first_end_idx = i
            break
    if first_end_idx is None:
        raise ValueError(f"输入文件里找不到 POI 终点 {end_poi!r}")

    if direction == "日前":
        if start_poi:
            # 找 [0, first_end_idx] 内第一次 start_poi
            first_start_idx: int | None = None
            for i in range(first_end_idx + 1):
                if start_poi in _loc_at(data_rows[i], loc_col):
                    first_start_idx = i
                    break
            if first_start_idx is None:
                raise ValueError(
                    f"输入文件 end_poi {end_poi!r} 之前找不到 start_poi {start_poi!r}"
                )
            if first_start_idx > first_end_idx:
                raise ValueError(
                    f"start_poi {start_poi!r} (idx={first_start_idx}) 在 "
                    f"end_poi {end_poi!r} (idx={first_end_idx}) 之后, 不合理"
                )
            selected = data_rows[first_start_idx: first_end_idx + 1]
        else:
            # start_poi 为空 -> 从文件头切到 first_end_idx (含)
            selected = data_rows[: first_end_idx + 1]
    elif direction == "后半":
        # 找 [0, first_end_idx) 范围内最后一次 start_poi
        # 若 noon_cutoff 非空, 再加时间约束 (时间 < noon_cutoff)
        last_start_idx: int | None = None
        for j in range(first_end_idx):
            row = data_rows[j]
            loc = _loc_at(row, loc_col)
            if start_poi not in loc:
                continue
            if noon_cutoff:
                t = _extract_time_of_day(_time_at(row, time_col))
                if t and t >= noon_cutoff:
                    continue
            last_start_idx = j
        if last_start_idx is None:
            raise ValueError(
                f"输入文件里找不到 start_poi {start_poi!r} "
                f"(noon_cutoff={noon_cutoff!r})"
            )
        selected = data_rows[last_start_idx: first_end_idx + 1]
    else:
        raise ValueError(f"direction 必须是 '日前' 或 '后半', 得到: {direction!r}")

    out_rows: list[list[Any]] = [OUTPUT_COLUMNS]
    out_rows.extend(_project_rows(selected, cols, temp_col, clamp_temp_to_0to4=clamp_temp_to_0to4))

    if round_to_minute:
        out_rows = round_to_minute_drop_collisions(out_rows)

    # 单文件切片 = 一辆车, 用首行车牌给 {plate} 占位符 (跟 merge/modifier 同套 sanitization 逻辑)
    plate_for_name = "unknown"
    if cols["plate"] is not None and selected:
        raw = selected[0][cols["plate"]] if cols["plate"] < len(selected[0]) else ""
        if raw not in (None, ""):
            plate_for_name = "".join(c for c in str(raw) if c.isalnum() or c in "_-") or "unknown"

    out_name = _format_filename(
        filename_template,
        direction=direction,
        plate=plate_for_name,
        date=find_main_date(src_path)[0] if _DATE_FMT_RE.search(filename_template) else None,
    )
    out_path = out_dir / out_name
    logger.info(
        "split_at_poi: 输出 %s 共 %d 行%s",
        out_path, len(out_rows), " (DRY-RUN)" if dry_run else "",
    )
    if not dry_run:
        write_xls(out_path, "轨迹", out_rows)
    return out_path