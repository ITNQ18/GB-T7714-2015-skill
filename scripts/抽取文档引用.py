from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.dont_write_bytecode = True

from 公共库 import dump_json, load_docx, package_summary


def main() -> None:
    parser = argparse.ArgumentParser(description="Extract citation and reference-section information from DOCX.")
    parser.add_argument("docx", help="DOCX file to inspect")
    parser.add_argument("-o", "--output", help="optional JSON output path")
    args = parser.parse_args()

    package = load_docx(args.docx)
    data = package_summary(package)
    content = dump_json(data)
    if args.output:
        Path(args.output).write_text(content + "\n", encoding="utf-8")
    else:
        print(content)


if __name__ == "__main__":
    main()
