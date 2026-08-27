"""包入口: 默认启 GUI (跟 `python -m trajectory_splitter` 等价)。

显式启 CLI/GUI 仍可用 `python -m trajectory_splitter.cli` / `gui`。
"""

from __future__ import annotations

import sys


def main() -> int:
    from trajectory_splitter import logging_setup
    logging_setup.configure()
    from trajectory_splitter.gui import main as gui_main
    return gui_main()


if __name__ == "__main__":
    sys.exit(main())