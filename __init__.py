"""trajectory_splitter — 单/双 .xls 切片与拼接工具。

两步用法:
  1) 单文件切片: 输入完整 .xls, 按 (start_poi, end_poi) 切日前/后半段,
     温度夹紧到 [0,4]
  2) 双文件拼接: 选两份 .xls, 无缝拼接, 时间接续, 文件名按参考文件夹命名

启动 GUI:  python -m trajectory_splitter.gui
命令行:    python -m trajectory_splitter.cli split|merge|main-date ...
"""

from trajectory_splitter.config import Settings, get_settings, reset_settings_cache
from trajectory_splitter.logging_setup import configure as configure_logging, get_logger
from trajectory_splitter.merger import merge_two
from trajectory_splitter.modifier import ModifyOptions, build_default_options, modify_xls
from trajectory_splitter.splitter import (
    COL_LOC,
    COL_PLATE,
    COL_TEMP1,
    COL_TEMP2,
    COL_TIME,
    DEFAULT_POI_END,
    DEFAULT_POI_START,
    OUTPUT_COLUMNS,
    Direction,
    default_end_poi,
    default_start_poi,
    find_main_date,
    pick_temperature_column,
    split_at_poi,
)
from trajectory_splitter.xls_io import read_xls, write_xls

__all__ = [
    "Settings",
    "get_settings",
    "reset_settings_cache",
    "configure_logging",
    "get_logger",
    "find_main_date",
    "split_at_poi",
    "merge_two",
    "modify_xls",
    "ModifyOptions",
    "build_default_options",
    "pick_temperature_column",
    "read_xls",
    "write_xls",
    "OUTPUT_COLUMNS",
    "COL_PLATE",
    "COL_TIME",
    "COL_TEMP1",
    "COL_TEMP2",
    "COL_LOC",
    "Direction",
    "DEFAULT_POI_START",
    "DEFAULT_POI_END",
    "default_start_poi",
    "default_end_poi",
]