from __future__ import annotations

import argparse
import runpy
import sys
from pathlib import Path

sys.dont_write_bytecode = True


SCRIPT_DIR = Path(__file__).resolve().parent
COMMANDS = {
    "audit": "审计参考文献.py",
    "apply": "应用参考文献.py",
    "extract": "抽取文档引用.py",
    "overview": "生成参考文献概览.py",
    "reorder": "重排序号.py",
    "search": "检索正式文献.py",
    "split": "拆分引用组.py",
}


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "ASCII dispatcher for GB/T 7714 helper scripts. "
            "Use this when a terminal has trouble typing Chinese script filenames."
        )
    )
    parser.add_argument("command", choices=sorted(COMMANDS), help="script command to run")
    parser.add_argument("args", nargs=argparse.REMAINDER, help="arguments passed to the target script")
    ns = parser.parse_args()

    target = SCRIPT_DIR / COMMANDS[ns.command]
    if not target.exists():
        raise SystemExit(f"Target script is missing: {target}")
    sys.argv = [str(target), *ns.args]
    runpy.run_path(str(target), run_name="__main__")


if __name__ == "__main__":
    main()
