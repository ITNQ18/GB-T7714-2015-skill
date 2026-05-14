from __future__ import annotations

import argparse
import json
import sys
from datetime import date
from pathlib import Path
from typing import Any

sys.dont_write_bytecode = True


def normalize_records(data: Any) -> list[dict[str, Any]]:
    if isinstance(data, list):
        return data
    if isinstance(data, dict):
        for key in ("记录", "records", "references", "items"):
            value = data.get(key)
            if isinstance(value, list):
                return value
    raise ValueError("输入必须是 JSON 列表或包含 记录/references/items 的对象。")


def record_status(record: dict[str, Any]) -> str:
    if record.get("跳过原因"):
        return "skipped"
    if record.get("已写入文档") is True:
        return "accepted"
    return str(record.get("状态", "accepted")).lower()


def render_record(record: dict[str, Any], index: int, 核验日期: str) -> list[str]:
    status = record_status(record)
    lines = [
        f"### {index}. {record.get('标识') or record.get('key') or record.get('id') or '未命名文献'}",
        "",
        f"- 候选编号：{record.get('候选编号', index)}",
        f"- 正文引用位置：{record.get('正文引用位置', '')}",
        f"- 支撑内容：{record.get('支撑内容', '')}",
        f"- 文献类型：{record.get('文献类型', record.get('type', record.get('kind', '')))}",
        f"- GB/T 7714 著录条目：{record.get('GB/T 7714 著录条目', record.get('gbt7714', record.get('entry', '')))}",
        f"- 核验来源 URL/DOI：{record.get('核验来源 URL/DOI', record.get('source', record.get('url', record.get('doi', ''))))}",
        f"- 核验日期：{record.get('核验日期', 核验日期)}",
        f"- 已写入文档：{'是' if status in {'accepted', 'include', 'included'} else '否'}",
        f"- 跳过原因：{record.get('跳过原因', '')}",
        "",
    ]
    return lines


def render_overview(records: list[dict[str, Any]], batch_size: int = 3, 核验日期: str | None = None) -> str:
    if 核验日期 is None:
        核验日期 = date.today().isoformat()
    lines = [
        "# 参考文献概览",
        "",
        f"生成日期：{核验日期}",
        "",
        "本文件记录候选文献、核验来源、正文支撑位置、是否写入正文以及跳过原因。真实性优先于数量，无法核实的文献不得进入最终参考文献表。",
        "",
    ]
    for batch_start in range(0, len(records), batch_size):
        batch = records[batch_start : batch_start + batch_size]
        batch_no = batch_start // batch_size + 1
        lines.extend([f"## 批次 {batch_no}", ""])
        for offset, record in enumerate(batch, start=1):
            lines.extend(render_record(record, batch_start + offset, 核验日期))
    return "\n".join(lines).rstrip() + "\n"


def build_task_state(
    目标文档: str = "",
    所需数量: int | None = None,
    核心文献: list[str] | None = None,
    引用样式: str = "顺序编码制",
    状态复用条件: bool = True,
    核验日期: str | None = None,
) -> dict[str, Any]:
    if 核验日期 is None:
        核验日期 = date.today().isoformat()
    return {
        "目标文档": 目标文档,
        "所需数量": 所需数量,
        "核心文献": (核心文献 or [])[:3],
        "引用样式": 引用样式,
        "正文字体": "Times New Roman",
        "正文字号_中文": "小四",
        "正文字号_磅值": 12,
        "正文上标": True,
        "交叉引用类型": "编号项/段落编号",
        "参考文献_首行缩进": 0,
        "参考文献_左缩进": 0,
        "参考文献_悬挂缩进": 0,
        "参考文献_对齐方式": "双端对齐",
        "最近核验日期": 核验日期,
        "最近完成阶段": "",
        "状态复用条件": 状态复用条件,
    }


def write_task_state(output: str | Path, state: dict[str, Any]) -> None:
    path = Path(output)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(state, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(description="生成或更新 参考文献/参考文献概览.md 和 参考文献/任务状态.json。")
    parser.add_argument("references_json", help="JSON 列表或对象，包含参考文献记录")
    parser.add_argument("--project-root", default=".", help="项目根目录，参考文献/ 将创建在此")
    parser.add_argument("--output", help="显式指定 Markdown 输出路径")
    parser.add_argument("--batch-size", type=int, default=3, help="每个批次的记录数")
    parser.add_argument("--核验日期", default=date.today().isoformat(), help="写入记录的日期")
    parser.add_argument("--目标文档", default="", help="记录到任务状态中的目标 DOCX")
    parser.add_argument("--所需数量", type=int, help="记录到任务状态中的参考文献数量")
    parser.add_argument("--核心文献", action="append", default=[], help="核心文献路径/标题，最多三次")
    parser.add_argument("--引用样式", default="顺序编码制", help="记录到任务状态中的引用样式")
    parser.add_argument("--no-reuse-state", action="store_true", help="记录后续运行不应自动复用状态")
    args = parser.parse_args()

    data = json.loads(Path(args.references_json).read_text(encoding="utf-8"))
    records = normalize_records(data)
    content = render_overview(records, batch_size=args.batch_size, 核验日期=args.核验日期)

    if args.output:
        output = Path(args.output)
    else:
        output = Path(args.project_root) / "参考文献" / "参考文献概览.md"
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(content, encoding="utf-8")
    if args.目标文档 or args.所需数量 is not None or args.核心文献:
        state = build_task_state(
            目标文档=args.目标文档,
            所需数量=args.所需数量,
            核心文献=args.核心文献[:3],
            引用样式=args.引用样式,
            状态复用条件=not args.no_reuse_state,
            核验日期=args.核验日期,
        )
        write_task_state(Path(args.project_root) / "参考文献" / "任务状态.json", state)
    print(str(output))


if __name__ == "__main__":
    main()
