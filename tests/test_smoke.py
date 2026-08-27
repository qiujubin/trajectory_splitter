r"""冒烟测试: 用真实输入文件跑一遍 split_at_poi + merge_two, 验证输出。

跳过条件: 真实输入文件不存在时自动 skip。

注意: 测试用 【麻涌卸货地址】 -> 【泉州烹趣】 这对 POI
(那份 8月24-25日输入里没有 【广州朴朴】)。
日前段不指定 start_poi (留空 = 从文件头切到 end_poi)。
"""

from __future__ import annotations

import shutil
from pathlib import Path

import pytest

from trajectory_splitter import find_main_date, merge_two, read_xls, split_at_poi


SRC = Path(r"D:\downloads2\15239468_20260824170652至20260825192937_轨迹数据.xls")
TMP = Path(r"C:\Temp\trajectory_splitter_test")

# 测试用 POI (这份数据里实际有的)
POI_MID = "【麻涌卸货地址】"
POI_END = "【泉州烹趣】"


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