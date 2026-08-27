"""GUI 持久化状态: 加载/保存到独立 JSON 文件, 不污染 settings.yaml。

路径: <project_root>/gui_state.json
- 启动 App 时 load_state() 拿 defaults
- 关窗 / 退出前 save_state(state) 写回

state 结构: {split: {...}, merge: {...}}
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from trajectory_splitter.config import PROJECT_ROOT


STATE_PATH = PROJECT_ROOT / "gui_state.json"


def _read(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {}
    try:
        with path.open("r", encoding="utf-8") as f:
            return json.load(f) or {}
    except (OSError, json.JSONDecodeError):
        return {}


def _write(path: Path, data: dict[str, Any]) -> None:
    try:
        with path.open("w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except OSError:
        pass


def load_state() -> dict[str, dict[str, Any]]:
    """读取持久化 GUI 状态, 失败返回空 dict。"""
    return _read(STATE_PATH)


def save_state(state: dict[str, dict[str, Any]]) -> None:
    """写回持久化 GUI 状态, 失败静默 (UI 关闭时不应抛错)。"""
    _write(STATE_PATH, state)