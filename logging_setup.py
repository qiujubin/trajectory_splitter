"""集中 logging 配置。

默认行为是不配置 (避免 import 时副作用), 用 `configure()` 在 CLI / GUI 入口按
settings.yaml 的 logging.level 设置级别。
"""

from __future__ import annotations

import logging
import sys

from trajectory_splitter.config import get_settings


_LEVELS = {
    "DEBUG": logging.DEBUG,
    "INFO": logging.INFO,
    "WARNING": logging.WARNING,
    "ERROR": logging.ERROR,
    "CRITICAL": logging.CRITICAL,
}


def configure() -> None:
    """根据 settings.yaml 的 logging.level 设置根 logger 级别。"""
    level_name = get_settings().log_level.upper()
    level = _LEVELS.get(level_name, logging.INFO)
    root = logging.getLogger()
    if not root.handlers:
        handler = logging.StreamHandler(sys.stderr)
        handler.setFormatter(
            logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s"),
        )
        root.addHandler(handler)
    root.setLevel(level)


def get_logger(name: str) -> logging.Logger:
    """模块级 logger。"""
    return logging.getLogger(name)