"""配置加载 (settings.yaml)。非敏感可调参数统一放这里, 不写死在代码里。"""

from __future__ import annotations

from pathlib import Path
from typing import Optional

import yaml


PROJECT_ROOT = Path(__file__).resolve().parent.parent
SETTINGS_PATH = PROJECT_ROOT / "settings.yaml"


class SplitHalfCfg:
    __slots__ = ("start_poi", "end_poi")

    def __init__(self, start_poi: str, end_poi: str) -> None:
        self.start_poi = start_poi
        self.end_poi = end_poi

    def __repr__(self) -> str:
        return f"SplitHalfCfg(start_poi={self.start_poi!r}, end_poi={self.end_poi!r})"


class SplitCfg:
    __slots__ = ("first_half", "second_half", "noon_cutoff")

    def __init__(self, first_half: SplitHalfCfg, second_half: SplitHalfCfg, noon_cutoff: str) -> None:
        self.first_half = first_half
        self.second_half = second_half
        self.noon_cutoff = noon_cutoff

    def __repr__(self) -> str:
        return (
            f"SplitCfg(first_half={self.first_half!r}, "
            f"second_half={self.second_half!r}, noon_cutoff={self.noon_cutoff!r})"
        )


class MergeCfg:
    __slots__ = ("time_gap_minutes", "filename_template")

    def __init__(self, time_gap_minutes: int, filename_template: str) -> None:
        self.time_gap_minutes = time_gap_minutes
        self.filename_template = filename_template

    def __repr__(self) -> str:
        return (
            f"MergeCfg(time_gap_minutes={self.time_gap_minutes}, "
            f"filename_template={self.filename_template!r})"
        )


class TemperatureCfg:
    __slots__ = ("clamp_to_0to4",)

    def __init__(self, clamp_to_0to4: bool) -> None:
        self.clamp_to_0to4 = clamp_to_0to4

    def __repr__(self) -> str:
        return f"TemperatureCfg(clamp_to_0to4={self.clamp_to_0to4})"


class Settings:
    """扁平配置对象。故意不引入 pydantic, 减少依赖。"""

    def __init__(self, data: dict) -> None:
        split = data.get("split") or {}
        merge = data.get("merge") or {}
        temp = data.get("temperature") or {}

        fh = split.get("first_half") or {}
        sh = split.get("second_half") or {}
        self.split = SplitCfg(
            first_half=SplitHalfCfg(
                start_poi=fh.get("start_poi", ""),
                end_poi=fh.get("end_poi", "【麻涌卸货地址】"),
            ),
            second_half=SplitHalfCfg(
                start_poi=sh.get("start_poi", "【麻涌卸货地址】"),
                end_poi=sh.get("end_poi", "【广州朴朴】"),
            ),
            noon_cutoff=split.get("noon_cutoff", "12:00:00"),
        )

        self.merge = MergeCfg(
            time_gap_minutes=int(merge.get("time_gap_minutes", 1)),
            filename_template=merge.get("filename_template", "1日前.xls"),
        )

        self.temperature = TemperatureCfg(
            clamp_to_0to4=bool(temp.get("clamp_to_0to4", True)),
        )

        log_cfg = data.get("logging") or {}
        self.log_level: str = log_cfg.get("level", "INFO")

    def __repr__(self) -> str:
        return f"Settings(split={self.split!r}, merge={self.merge!r}, temperature={self.temperature!r})"


_settings_cache: Optional[Settings] = None


def get_settings() -> Settings:
    global _settings_cache
    if _settings_cache is not None:
        return _settings_cache
    if SETTINGS_PATH.exists():
        with SETTINGS_PATH.open("r", encoding="utf-8") as f:
            data = yaml.safe_load(f) or {}
    else:
        data = {}
    _settings_cache = Settings(data)
    return _settings_cache


def reset_settings_cache() -> None:
    """测试时清缓存, 让 settings.yaml 改动生效。"""
    global _settings_cache
    _settings_cache = None