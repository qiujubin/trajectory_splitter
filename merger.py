"""双文件拼接: 把两份 .xls 按 [日前 + 后半] 顺序拼成一个, 时间接续, 文件名按参考文件夹命名。

时间接续规则 (两种模式):
  默认:
    - 锚点 = 日前最后一条时间 + gap_minutes
    - 日后第一行新时间 = 锚点
    - 日后整体时间偏移 offset = 锚点 - 日后第一行原时间
    - 通常 offset < 0 (日后第一行原时间 > 锚点, 即日后数据时间更晚, 平移让它"提前"接续)
    - 跨天 (e.g. 前半 23:55, 后半原时间 00:05 第二天) 也走同一套逻辑:
      timedelta 自动跨天, 排序仍然按 datetime 比较。
    - 跨日期/不同车的两段: offset 自动取正, 把后半段整个平移到前半末条之后;
      仅打 logger.warning 提示用户关注, 不阻止 (因为操作本来就合法)。

  preserve_endpoints=True (保留两端 + 中间压接):
    - A 的开头保留 (本来就保留); B 的结尾保留 (必要时把 B 整段日期前移,
      时分秒不变, 匹配「以第一个的开头为开头, 第二个的结尾为结尾」)
    - A 的最后 K 行 + B 的最先 K 行 (共 2K 行) 时间戳重新分配:
      在 A[N_A-K-1] 时间 与 B[K] 时间 (可能已 +delta_days) 之间均匀插值
    - K = min(squeeze_rows, N_A - 1, N_B - 1)
    - 温度、地点列原值不动, 只改时间列
    - 当 R_raw <= L (B 的保留首行时间 <= A 的保留末行时间),
      把 B 整段日期前移 delta_days = ceil((L - R_raw) / 1 day) 天,
      保证输出单调

温度: 超 [0, 4] 范围按传感器故障模式分段线性映射 (不用随机数, 保持温度曲线自然),
      输出统一保留 1 位小数。详见 splitter._clamp_temp。

输出文件名模板: 默认 `1日前.xls` (沿用参考文件夹 `N+方向.xls` 命名), 支持 {plate}。
"""

from __future__ import annotations

import math
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
    round_to_minute_drop_collisions,
)
from trajectory_splitter.logging_setup import get_logger
from trajectory_splitter.splitter import _DATE_FMT_RE, _format_filename, find_main_date
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


def _linear_squeeze(l_dt: datetime, r_dt: datetime, n: int) -> list[datetime]:
    """返回 n 个均匀间隔的时间戳, 落在 (l_dt, r_dt) 之间。

    间距 = (r - l) / (n + 1), i-th (i=1..n) -> l + i * 间距。
    n=0 返回空列表; n<0 抛 ValueError。
    """
    if n < 0:
        raise ValueError(f"_linear_squeeze: n 必须 >= 0, 得到 {n}")
    if n == 0:
        return []
    span = (r_dt - l_dt).total_seconds()
    if span <= 0:
        raise ValueError(f"_linear_squeeze: 区间非正 (l={l_dt}, r={r_dt})")
    step = span / (n + 1)
    return [l_dt + timedelta(seconds=(i + 1) * step) for i in range(n)]


def _transform_preserve_endpoints(
    f_data: list[list[Any]],
    s_data: list[list[Any]],
    f_cols: dict[str, int | None],
    s_cols: dict[str, int | None],
    squeeze_rows: int,
) -> bool:
    """preserve_endpoints 模式的时间变换。原地修改 f_data / s_data。

    成功应用返回 True; 数据太少 / 锚点缺失 / 区间非正 时返回 False, 此时
    f_data / s_data 可能已部分改动 (保险起见, 调用方应在失败时退回默认模式)。

    算法:
      1. effective_K = min(squeeze_rows, N_A - 1, N_B - 1)
      2. 左锚 L = A[N_A-K-1] 时间; 右锚 R_raw = B[K] 时间
      3. 若 R_raw <= L, 把 B 整段日期前移 delta_days = ceil((L - R_raw)/1day) 天
      4. R = R_raw + delta_days; span = (R - L).total_seconds()
      5. _linear_squeeze(L, R, 2K) 得到 2K 个时间戳, 前 K 覆盖 A 末 K 行, 后 K 覆盖 B 首 K 行
      6. B 整段 (包括保留段 K..N_B-1) +delta_days
    """
    n_a = len(f_data)
    n_b = len(s_data)
    f_time_col = f_cols.get("time")
    s_time_col = s_cols.get("time")

    if n_a <= 1 or n_b <= 1 or f_time_col is None or s_time_col is None:
        logger.warning(
            "_transform_preserve_endpoints: 数据/时间列不足 (A=%d, B=%d, "
            "f_time_col=%s, s_time_col=%s), 跳过压接",
            n_a, n_b, f_time_col, s_time_col,
        )
        return False

    effective_K = min(squeeze_rows, n_a - 1, n_b - 1)
    if effective_K <= 0:
        logger.warning(
            "_transform_preserve_endpoints: effective_K=%d, 跳过压接",
            effective_K,
        )
        return False

    # 找左锚 (A[N_A-K-1] 时间, 往前找最近的能解析的时间行)
    l_dt: datetime | None = None
    for i in range(n_a - effective_K - 1, -1, -1):
        if f_time_col < len(f_data[i]):
            l_dt = _parse_dt(f_data[i][f_time_col])
            if l_dt is not None:
                break
    if l_dt is None:
        logger.warning(
            "_transform_preserve_endpoints: 找不到左锚时间 (A[N_A-K-1..0]), 跳过压接",
        )
        return False

    # 找右锚原值 (B[K] 时间, 往后找最近的能解析的时间行)
    r_raw_dt: datetime | None = None
    for i in range(effective_K, n_b):
        if s_time_col < len(s_data[i]):
            r_raw_dt = _parse_dt(s_data[i][s_time_col])
            if r_raw_dt is not None:
                break
    if r_raw_dt is None:
        logger.warning(
            "_transform_preserve_endpoints: 找不到右锚时间 (B[K..N_B-1]), 跳过压接",
        )
        return False

    # delta_days: 让 R > L
    delta_days = 0
    if r_raw_dt <= l_dt:
        delta_days = math.ceil((l_dt - r_raw_dt).total_seconds() / 86400)
    r_dt = r_raw_dt + timedelta(days=delta_days)

    span_s = (r_dt - l_dt).total_seconds()
    if span_s <= 0:
        logger.warning(
            "_transform_preserve_endpoints: 区间 = 0 (L=%s, R=%s), 跳过压接",
            l_dt, r_dt,
        )
        return False

    # 行就地改成 list (xls_io 已经返回 list, 这里只是保险)
    for i in range(n_a):
        if not isinstance(f_data[i], list):
            f_data[i] = list(f_data[i])
    for i in range(n_b):
        if not isinstance(s_data[i], list):
            s_data[i] = list(s_data[i])

    # 2K 个均匀时间戳
    try:
        squeeze_times = _linear_squeeze(l_dt, r_dt, 2 * effective_K)
    except ValueError as e:
        logger.warning("_transform_preserve_endpoints: %s, 跳过压接", e)
        return False
    a_squeeze = squeeze_times[:effective_K]
    b_squeeze = squeeze_times[effective_K:]

    # A 末 K 行应用压接
    for i, new_t in enumerate(a_squeeze):
        row_idx = n_a - effective_K + i
        if f_time_col < len(f_data[row_idx]):
            f_data[row_idx][f_time_col] = new_t.strftime("%Y-%m-%d %H:%M:%S")

    # B 整段日期前移
    if delta_days > 0:
        delta_td = timedelta(days=delta_days)
        for i in range(n_b):
            if s_time_col < len(s_data[i]):
                t = _parse_dt(s_data[i][s_time_col])
                if t is not None:
                    s_data[i][s_time_col] = (t + delta_td).strftime("%Y-%m-%d %H:%M:%S")

    # B 首 K 行覆盖压接时间
    for i, new_t in enumerate(b_squeeze):
        if s_time_col < len(s_data[i]):
            s_data[i][s_time_col] = new_t.strftime("%Y-%m-%d %H:%M:%S")

    logger.info(
        "merge_two preserve_endpoints: squeeze_rows=%d effective_K=%d delta_days=%d "
        "L=%s R=%s step=%.3fs",
        squeeze_rows, effective_K, delta_days, l_dt, r_dt,
        span_s / (2 * effective_K + 1),
    )
    return True


def merge_two(
    first_path: str | Path,
    second_path: str | Path,
    out_dir: str | Path,
    *,
    time_gap_minutes: int = 1,
    clamp_temp_to_0to4: bool = True,
    filename_template: str = "1日前.xls",
    output_plate: str | None = None,
    round_to_minute: bool = False,
    preserve_endpoints: bool = False,
    squeeze_rows: int = 10,
    dry_run: bool = False,
) -> Path:
    """拼接两个 .xls (日前 + 后半), 时间接续, 温度夹紧。

    Args:
        first_path: 日前 .xls (默认从该文件抽取车牌号)。
        second_path: 后半 .xls。
        out_dir: 输出目录。
        time_gap_minutes: 日后第一行新时间相对日前最后一条的间隔 (分钟)。
                          仅默认模式用, preserve_endpoints 模式忽略。
        clamp_temp_to_0to4: 温度超 [0,4] 是否夹紧。
        filename_template: 输出文件名模板, 支持 {plate}。
        output_plate: 强制覆盖输出车牌 (覆盖前后两份所有行 + {plate} 占位符)。
                       None = 沿用日前首行的车牌作为输出车牌 (向后兼容)。
        round_to_minute: 时间列秒清零, 同分钟内重复行丢弃 (最后一步)。
        preserve_endpoints: True = 保留 A 开头 + B 结尾, 中间按 squeeze_rows 压接。
                            False (默认) = 走原 B 整体偏移模式。
        squeeze_rows: preserve_endpoints=True 时, A 末尾和 B 开头各取多少行压接。
                      默认 10。实际 K = min(squeeze_rows, N_A - 1, N_B - 1)。
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

    # 决定输出车牌: output_plate 优先, 否则用日前首行的车牌。
    #   有 output_plate -> 覆盖前后两份所有行 (跨车拼接场景里这是用户希望的)
    #   没有          -> 沿用日前首行的车牌, 把后半的车牌列也改成它 (旧行为, 向后兼容)
    if output_plate:
        plate = output_plate
    else:
        plate = f_data[0][f_cols["plate"]]
    safe_plate = _safe_for_filename(plate) or "unknown"

    # 把所有行 (前后两份) 的车牌号统一改成最终输出车牌
    for cols, data in ((f_cols, f_data), (s_cols, s_data)):
        if cols["plate"] is None:
            continue
        for row in data:
            if cols["plate"] < len(row):
                row[cols["plate"]] = plate
        cols["plate"] = f_cols["plate"]

    f_temp_col = pick_temperature_column(f_data, f_cols["temp1"], f_cols["temp2"])
    s_temp_col = pick_temperature_column(s_data, s_cols["temp1"], s_cols["temp2"])

    # 时间接续: 两条路径
    #   1) preserve_endpoints=True -> _transform_preserve_endpoints 原地改 f_data/s_data
    #   2) 默认 -> 找 A 末 / B 首, 算 offset, 整体平移 B
    applied_squeeze = False
    if preserve_endpoints:
        applied_squeeze = _transform_preserve_endpoints(
            f_data, s_data, f_cols, s_cols, squeeze_rows,
        )
        if not applied_squeeze:
            logger.warning(
                "merge_two: preserve_endpoints=True 但压接未应用, 退回默认模式",
            )

    if not applied_squeeze:
        # 默认: 整体偏移 B
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

        # 跨日期 / 不同车 拼接时 s_first_dt < f_last_dt 是合法的: offset 自动取正,
        # 后半段整体平移到前半末条之后。仅打 warning 提示, 不阻止。
        if s_first_dt < f_last_dt:
            gap_days = (f_last_dt - s_first_dt).total_seconds() / 86400
            logger.warning(
                "merge_two: 后半首条 (%s) 早于前半末条 (%s), 间隔 %.1f 天 — 走正向偏移自动接续; "
                "如非预期, 检查是否把切片顺序搞反了",
                s_first_dt, f_last_dt, gap_days,
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
        # 用 s_shifted 替换 s_data 走下面的 _project_rows
        s_data = s_shifted

    out_rows: list[list[Any]] = [OUTPUT_COLUMNS]
    f_out = _project_rows(
        f_data, f_cols, f_temp_col,
        clamp_temp_to_0to4=clamp_temp_to_0to4, start_seq=1,
    )
    out_rows.extend(f_out)
    s_out = _project_rows(
        s_data, s_cols, s_temp_col,
        clamp_temp_to_0to4=clamp_temp_to_0to4, start_seq=len(f_out) + 1,
    )
    out_rows.extend(s_out)

    if round_to_minute:
        out_rows = round_to_minute_drop_collisions(out_rows)

    out_name = _format_filename(
        filename_template,
        plate=safe_plate,
        date=find_main_date(first_path)[0] if _DATE_FMT_RE.search(filename_template) else None,
    )
    # merge 模板里没有 {direction} 语义, 拼错立刻报错
    if "{direction}" in out_name or "{direction:" in out_name:
        raise ValueError(
            f"merge_two 的 filename_template 不支持 {{direction}}: {filename_template!r}"
        )
    out_path = out_dir / out_name
    if applied_squeeze:
        logger.info(
            "merge_two: preserve_endpoints=True clamp=%s 输出 %s 共 %d 行%s",
            clamp_temp_to_0to4, out_path, len(out_rows),
            " (DRY-RUN)" if dry_run else "",
        )
    else:
        logger.info(
            "merge_two: 锚点=%s offset=%s clamp=%s 输出 %s 共 %d 行%s",
            anchor, offset, clamp_temp_to_0to4, out_path, len(out_rows),
            " (DRY-RUN)" if dry_run else "",
        )
    if not dry_run:
        write_xls(out_path, "轨迹", out_rows)
    return out_path
