r"""冒烟测试: 用真实输入文件跑一遍 split_at_poi + merge_two, 验证输出。

跳过条件: 真实输入文件不存在时自动 skip。

注意: 测试用 【麻涌卸货地址】 -> 【东莞鸣鲜】 这对 POI。
这份 8月24-25日输入里行程顺序是 泉州烹趣(8/24 晚) 在前、
麻涌卸货地址(8/25 上午) 在后, 所以"后半"段的终点 POI 取麻涌之后
第一次出现的 【东莞鸣鲜】 (用 【泉州烹趣】 会因它在麻涌之前而报
找不到 start_poi)。日前段不指定 start_poi (留空 = 从文件头切到 end_poi)。
"""

from __future__ import annotations

import shutil
from pathlib import Path

import pytest

from trajectory_splitter import find_main_date, merge_two, read_xls, split_at_poi


SRC = Path(r"D:\downloads2\15239468_20260824170652至20260825192937_轨迹数据.xls")
TMP = Path(r"C:\Temp\trajectory_splitter_test")

# 测试用 POI (这份数据里实际有的; 麻涌 8/25 上午 08:49-10:24, 东莞鸣鲜 10:29)
POI_MID = "【麻涌卸货地址】"
POI_END = "【东莞鸣鲜】"


@pytest.fixture(scope="module", autouse=True)
def _setup_tmp():
    if TMP.exists():
        shutil.rmtree(TMP)
    TMP.mkdir(parents=True)
    yield
    if TMP.exists():
        shutil.rmtree(TMP)


def _skip_if_no_src():
    if not SRC.exists():
        pytest.skip(f"真实输入文件不存在: {SRC}")


def test_find_main_date():
    _skip_if_no_src()
    date, count = find_main_date(SRC)
    assert date == "2026-08-25"
    assert count == 233


def test_split_at_poi_first():
    """日前段: start_poi 空 (从文件头) -> end_poi=麻涌。"""
    _skip_if_no_src()
    out = split_at_poi(
        SRC, "日前", TMP,
        start_poi="",
        end_poi=POI_MID,
        clamp_temp_to_0to4=False,
    )
    assert out.exists()
    assert out.name == "1日前.xls"
    sheets = read_xls(out)
    rows = next(iter(sheets.values()))
    assert rows[0] == ["序号", "车牌号", "定位时间", "温度2(℃)", "位置"]
    assert len(rows) > 1


def test_split_at_poi_second():
    """后半段: start_poi=麻涌 -> end_poi=烹趣。"""
    _skip_if_no_src()
    out = split_at_poi(
        SRC, "后半", TMP,
        start_poi=POI_MID,
        end_poi=POI_END,
        clamp_temp_to_0to4=False,
    )
    assert out.exists()
    assert out.name == "1后半.xls"


def test_split_at_poi_with_clamp():
    """温度夹紧: 所有温度必须在 [0, 4] 范围内。"""
    _skip_if_no_src()
    out = split_at_poi(
        SRC, "日前", TMP,
        start_poi="",
        end_poi=POI_MID,
        clamp_temp_to_0to4=True,
        filename_template="1clamp.xls",
    )
    sheets = read_xls(out)
    rows = next(iter(sheets.values()))
    for row in rows[1:]:
        temp = row[3]
        if temp in ("", None):
            continue
        try:
            f = float(temp)
        except (TypeError, ValueError):
            continue
        assert 0.0 <= f <= 4.0, f"温度 {f} 超出 [0, 4] 范围"


def test_merge():
    _skip_if_no_src()
    a = split_at_poi(
        SRC, "日前", TMP,
        start_poi="",
        end_poi=POI_MID,
        clamp_temp_to_0to4=False,
    )
    b = split_at_poi(
        SRC, "后半", TMP,
        start_poi=POI_MID,
        end_poi=POI_END,
        clamp_temp_to_0to4=False,
    )
    out = merge_two(a, b, TMP, time_gap_minutes=1, clamp_temp_to_0to4=False)
    assert out.exists()
    sheets = read_xls(out)
    rows = next(iter(sheets.values()))
    assert rows[0] == ["序号", "车牌号", "定位时间", "温度2(℃)", "位置"]
    # 序号应连续 1..N
    seqs = [r[0] for r in rows[1:]]
    assert seqs == list(range(1, len(rows)))


def test_merge_preserve_endpoints_squeeze():
    """preserve_endpoints=True: A 开头 + B 结尾保留, 中间 2K 行压接。"""
    _skip_if_no_src()
    a = split_at_poi(
        SRC, "日前", TMP,
        start_poi="",
        end_poi=POI_MID,
        clamp_temp_to_0to4=False,
    )
    b = split_at_poi(
        SRC, "后半", TMP,
        start_poi=POI_MID,
        end_poi=POI_END,
        clamp_temp_to_0to4=False,
    )

    # 取 A 首行 / B 末行 原时间作为锚点
    a_rows = read_xls(a)[next(iter(read_xls(a)))][1:]
    b_rows = read_xls(b)[next(iter(read_xls(b)))][1:]
    a_first_time = a_rows[0][2]
    b_last_time = b_rows[-1][2]

    out = merge_two(
        a, b, TMP,
        time_gap_minutes=1,
        clamp_temp_to_0to4=False,
        filename_template="preserve.xls",
        preserve_endpoints=True,
        squeeze_rows=10,
    )
    assert out.exists()
    sheets = read_xls(out)
    rows = next(iter(sheets.values()))
    assert rows[0] == ["序号", "车牌号", "定位时间", "温度2(℃)", "位置"]
    data = rows[1:]
    seqs = [r[0] for r in data]
    assert seqs == list(range(1, len(rows)))

    # 首行时间 == A 首行原时间 (preserved)
    assert data[0][2] == a_first_time, (
        f"preserve_endpoints: 首行时间应 == A 首行原时间\n"
        f"  got      {data[0][2]}\n  expected {a_first_time}"
    )
    # 末行时间 == B 末行原时间 (preserved, 可能因日期前移而不同 — 但时分秒相同)
    out_last_time = data[-1][2]
    out_last_hms = out_last_time[11:19]
    b_last_hms = b_last_time[11:19]
    assert out_last_hms == b_last_hms, (
        f"preserve_endpoints: 末行时分秒应 == B 末行原时分秒\n"
        f"  got      {out_last_time}\n  expected {b_last_time}"
    )

    # 时间戳严格单调递增 (因为 delta_days 已经处理过 R>L)
    from datetime import datetime
    prev = None
    for r in data:
        t = datetime.strptime(r[2], "%Y-%m-%d %H:%M:%S")
        if prev is not None:
            assert t > prev, f"时间戳倒退: {prev} -> {t}"
        prev = t

    # 压接段验证: 倒数 K+1 行 (含) 之前的时间应保持原始 A 的非压接段单调
    # 简单校验: 至少前几行的时间是 A 原时间 (没动)
    K = min(10, len(a_rows) - 1, len(b_rows) - 1)
    assert K > 0, f"需要 K>0 才能验证压接, a={len(a_rows)} b={len(b_rows)}"
    # A 的非压接段最后一行 (a_rows[len(a_rows)-K-1]) 应出现在输出里
    expected_a_last_preserved = a_rows[len(a_rows) - K - 1][2]
    found_idx = None
    for i, r in enumerate(data):
        if r[2] == expected_a_last_preserved:
            found_idx = i
            break
    assert found_idx is not None, (
        f"压接模式下, A 的非压接段最后一行 ({expected_a_last_preserved}) "
        f"应在输出中找到, 但未匹配"
    )


def test_merge_preserve_endpoints_too_small():
    """preserve_endpoints=True 但 K=0: 应该退回默认模式 (不报错)。"""
    _skip_if_no_src()
    a = split_at_poi(
        SRC, "日前", TMP,
        start_poi="",
        end_poi=POI_MID,
        clamp_temp_to_0to4=False,
    )
    b = split_at_poi(
        SRC, "后半", TMP,
        start_poi=POI_MID,
        end_poi=POI_END,
        clamp_temp_to_0to4=False,
    )
    # squeeze_rows=0 -> effective_K=0 -> 退回默认
    out = merge_two(
        a, b, TMP,
        time_gap_minutes=1,
        clamp_temp_to_0to4=False,
        filename_template="fallback.xls",
        preserve_endpoints=True,
        squeeze_rows=0,
    )
    assert out.exists()
    sheets = read_xls(out)
    rows = next(iter(sheets.values()))
    seqs = [r[0] for r in rows[1:]]
    assert seqs == list(range(1, len(rows)))

def _make_xls(path, rows):
    """写一个最小可用的 .xls 给 merge_two 测, 不依赖真实源文件。
    rows = list of [plate, time_str, temp, loc] (4 列, 不含 序号, 序号自动写)。"""
    from trajectory_splitter.xls_io import write_xls

    # 列顺序: 序号 / 车牌号 / 定位时间 / 温度2(℃) / 位置 (5 列)
    out_rows = [["序号", "车牌号", "定位时间", "温度2(℃)", "位置"]]
    for r_idx, (plate, t_str, temp, loc) in enumerate(rows, start=1):
        out_rows.append([r_idx, plate, t_str, temp, loc])
    write_xls(path, "轨迹", out_rows)


def test_merge_preserve_endpoints_with_mock_files(tmp_path=None):
    """不依赖真实 .xls: 用临时 .xls 跑 merge_two preserve_endpoints 全流程。
    A: 12 行 12:00-12:11; B: 12 行 13:00-13:11; K=3。"""
    from datetime import datetime, timedelta
    import tempfile, shutil
    from trajectory_splitter import merge_two, read_xls

    tmp = Path(tempfile.mkdtemp(prefix="merge_mock_"))
    try:
        a_path = tmp / "a.xls"
        b_path = tmp / "b.xls"
        a_rows = [
            ["粤BNX001", (datetime(2026, 8, 25, 12, 0, 0) + timedelta(minutes=i)).strftime("%Y-%m-%d %H:%M:%S"), "1.0", f"POI_{i}"]
            for i in range(12)
        ]
        b_rows = [
            ["粤BNX002", (datetime(2026, 8, 25, 13, 0, 0) + timedelta(minutes=i)).strftime("%Y-%m-%d %H:%M:%S"), "2.0", f"POI_{i+100}"]
            for i in range(12)
        ]
        _make_xls(a_path, a_rows)
        _make_xls(b_path, b_rows)

        out = merge_two(
            a_path, b_path, tmp,
            time_gap_minutes=1,
            clamp_temp_to_0to4=False,
            filename_template="merged.xls",
            preserve_endpoints=True,
            squeeze_rows=3,
        )
        assert out.exists()
        sheets = read_xls(out)
        rows = next(iter(sheets.values()))[1:]
        assert len(rows) == 24

        # 首行时间 == A[0] 原时间
        assert rows[0][2] == "2026-08-25 12:00:00", f"got {rows[0][2]}"
        # 末行时间 == B[11] 原时间 (时分秒)
        assert rows[-1][2][11:] == "13:11:00", f"got {rows[-1][2]}"

        # 时间戳严格单调
        prev = None
        for r in rows:
            t = datetime.strptime(r[2], "%Y-%m-%d %H:%M:%S")
            assert prev is None or t > prev, f"非单调: {prev} -> {t}"
            prev = t

        # A 非压接段最后一行 (a_rows[8] = 12:08) 应原样出现在输出里
        # 序号在输出里是连续的, 但内容上应该能找到 12:08 这一行
        found = any(r[2] == "2026-08-25 12:08:00" for r in rows)
        assert found, "A 非压接段最后一行 12:08 应原样保留"
        # B 非压接段第一行 (b_rows[3] = 13:03) 应原样出现
        found = any(r[2] == "2026-08-25 13:03:00" for r in rows)
        assert found, "B 非压接段第一行 13:03 应原样保留"
    finally:
        shutil.rmtree(tmp)


def test_merge_preserve_endpoints_cross_day_with_mock():
    """跨天场景: A 在 8月25, B 在 8月24。delta_days 应自动 +1。"""
    from datetime import datetime, timedelta
    import tempfile, shutil
    from trajectory_splitter import merge_two, read_xls

    tmp = Path(tempfile.mkdtemp(prefix="merge_crossday_"))
    try:
        a_path = tmp / "a.xls"
        b_path = tmp / "b.xls"
        a_rows = [
            ["粤BNX001", (datetime(2026, 8, 25, 12, 0, 0) + timedelta(minutes=i)).strftime("%Y-%m-%d %H:%M:%S"), "1.0", f"POI_{i}"]
            for i in range(12)
        ]
        b_rows = [
            ["粤BNX002", (datetime(2026, 8, 24, 13, 0, 0) + timedelta(minutes=i)).strftime("%Y-%m-%d %H:%M:%S"), "2.0", f"POI_{i+100}"]
            for i in range(12)
        ]
        _make_xls(a_path, a_rows)
        _make_xls(b_path, b_rows)

        out = merge_two(
            a_path, b_path, tmp,
            time_gap_minutes=1,
            clamp_temp_to_0to4=False,
            filename_template="merged.xls",
            preserve_endpoints=True,
            squeeze_rows=3,
        )
        sheets = read_xls(out)
        rows = next(iter(sheets.values()))[1:]
        # B 末行应 = B 原末行时分秒 + 8月25 (而非 8月24)
        assert rows[-1][2].startswith("2026-08-25 "), f"got {rows[-1][2]}"
        assert rows[-1][2][11:] == "13:11:00", f"got {rows[-1][2]}"

        # 时间戳严格单调
        prev = None
        for r in rows:
            t = datetime.strptime(r[2], "%Y-%m-%d %H:%M:%S")
            assert prev is None or t > prev, f"非单调: {prev} -> {t}"
            prev = t
    finally:
        shutil.rmtree(tmp)


def test_merge_default_unchanged():
    """回归: preserve_endpoints=False 应该跟原行为一致。"""
    from datetime import datetime, timedelta
    import tempfile, shutil
    from trajectory_splitter import merge_two, read_xls

    tmp = Path(tempfile.mkdtemp(prefix="merge_default_"))
    try:
        a_path = tmp / "a.xls"
        b_path = tmp / "b.xls"
        a_rows = [
            ["粤BNX001", (datetime(2026, 8, 25, 12, 0, 0) + timedelta(minutes=i)).strftime("%Y-%m-%d %H:%M:%S"), "1.0", f"POI_{i}"]
            for i in range(12)
        ]
        b_rows = [
            ["粤BNX002", (datetime(2026, 8, 25, 13, 0, 0) + timedelta(minutes=i)).strftime("%Y-%m-%d %H:%M:%S"), "2.0", f"POI_{i+100}"]
            for i in range(12)
        ]
        _make_xls(a_path, a_rows)
        _make_xls(b_path, b_rows)

        out = merge_two(
            a_path, b_path, tmp,
            time_gap_minutes=1,
            clamp_temp_to_0to4=False,
            filename_template="default.xls",
            preserve_endpoints=False,  # 默认模式
        )
        sheets = read_xls(out)
        rows = next(iter(sheets.values()))[1:]
        # 默认模式: B 整体偏移, B 末行时间应是 B 原末行 + offset
        # anchor = A 末 (12:11) + 1min = 12:12
        # offset = 12:12 - B首(13:00) = -48min
        # B 末行新时间 = 13:11 - 48min = 12:23
        assert rows[-1][2] == "2026-08-25 12:23:00", f"got {rows[-1][2]}"

        # 首行 = A 首行 原样
        assert rows[0][2] == "2026-08-25 12:00:00"
    finally:
        shutil.rmtree(tmp)
