from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.dont_write_bytecode = True

from 公共库 import audit_document, dump_json, load_docx, render_markdown_report


def main() -> None:
    parser = argparse.ArgumentParser(description="Audit a DOCX against strict GB/T 7714-2015 reference rules.")
    parser.add_argument("docx", help="DOCX file to audit")
    parser.add_argument("-o", "--output", help="optional report output path")
    parser.add_argument("--json", action="store_true", help="emit JSON instead of Markdown")
    parser.add_argument("--fail-on-issues", action="store_true", help="exit with code 1 when issues are found")
    args = parser.parse_args()

    package = load_docx(args.docx)
    issues = audit_document(package)
    content = dump_json({"path": str(Path(args.docx)), "issue_count": len(issues), "issues": issues})
    if not args.json:
        content = render_markdown_report(args.docx, issues)

    if args.output:
        Path(args.output).write_text(content + "\n", encoding="utf-8")
    else:
        print(content)

    if args.fail_on_issues and issues:
        sys.exit(1)


if __name__ == "__main__":
    main()
