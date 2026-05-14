from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any

sys.dont_write_bytecode = True

from 公共库 import extract_citations, load_docx, parse_citation_numbers


def find_oversized_groups(citations: list[dict], threshold: int = 3) -> list[dict]:
    """找出超过阈值的引用组。"""
    groups: list[dict] = []
    for citation in citations:
        numbers = parse_citation_numbers(citation["text"])
        if len(numbers) > threshold:
            groups.append({
                "段落索引": citation["paragraph_index"],
                "运行索引": citation["run_index"],
                "引用文本": citation["text"],
                "编号数量": len(numbers),
                "编号列表": sorted(numbers),
                "上下文": citation["context"],
            })
    return groups


def suggest_split_points(group: dict) -> list[dict]:
    """为单个引用组提出拆分建议。

    基于上下文中的逗号、分号、模型名、方法名等语义边界提出建议。
    不改写事实，只建议引用放置位置。
    """
    context = group["上下文"]
    numbers = group["编号列表"]
    suggestions: list[dict] = []

    sentences = re.split(r"[。；;]", context)
    current_num_idx = 0
    for sent_idx, sentence in enumerate(sentences):
        sent = sentence.strip()
        if not sent:
            continue
        sub_parts = re.split(r"[，,]", sent)
        for part_idx, part in enumerate(sub_parts):
            part = part.strip()
            if not part:
                continue
            if current_num_idx < len(numbers):
                suggestions.append({
                    "建议短语": part[:50],
                    "分配编号": [numbers[current_num_idx]],
                    "原句片段": sent[:80],
                })
                current_num_idx += 1

    remaining = numbers[current_num_idx:]
    if remaining:
        if suggestions:
            suggestions[-1]["分配编号"].extend(remaining)
        else:
            suggestions.append({
                "建议短语": context[:50],
                "分配编号": remaining,
                "原句片段": context[:80],
            })

    merged: list[dict] = []
    for s in suggestions:
        if merged and len(merged[-1]["分配编号"]) + len(s["分配编号"]) <= 3:
            merged[-1]["分配编号"].extend(s["分配编号"])
            merged[-1]["建议短语"] += "、" + s["建议短语"]
        else:
            merged.append(s)

    return merged


def analyze_docx(docx_path: str | Path, threshold: int = 3) -> dict[str, Any]:
    """分析 DOCX 中超过阈值的引用组并提出拆分建议。"""
    package = load_docx(docx_path)
    citations = extract_citations(package)
    oversized = find_oversized_groups(citations, threshold)

    results: list[dict] = []
    for group in oversized:
        suggestions = suggest_split_points(group)
        results.append({
            "段落索引": group["段落索引"],
            "运行索引": group["运行索引"],
            "引用文本": group["引用文本"],
            "编号数量": group["编号数量"],
            "编号列表": group["编号列表"],
            "上下文": group["上下文"][:200],
            "拆分建议": suggestions,
        })

    return {
        "文件": str(docx_path),
        "阈值": threshold,
        "总引用数": len(citations),
        "超限引用组数": len(results),
        "超限引用组": results,
    }


def render_markdown_report(data: dict[str, Any]) -> str:
    """将分析结果渲染为 Markdown。"""
    lines = [
        "# 引用拆分分析报告",
        "",
        f"文件：`{data['文件']}`",
        f"阈值：每组最多 {data['阈值']} 条",
        f"总引用数：{data['总引用数']}",
        f"超限引用组数：{data['超限引用组数']}",
        "",
    ]

    if not data["超限引用组"]:
        lines.append("未发现超过阈值的引用组。")
        return "\n".join(lines)

    for idx, group in enumerate(data["超限引用组"], start=1):
        lines.extend([
            f"## {idx}. 段落 {group['段落索引']}，运行 {group['运行索引']}",
            "",
            f"- 引用文本：`{group['引用文本']}`",
            f"- 编号数量：{group['编号数量']}",
            f"- 编号列表：{group['编号列表']}",
            f"- 上下文：{group['上下文']}",
            "",
            "### 拆分建议",
            "",
        ])
        for s_idx, sug in enumerate(group["拆分建议"], start=1):
            lines.extend([
                f"{s_idx}. **{sug['建议短语']}** → `{sug['分配编号']}`",
                "",
            ])

    return "\n".join(lines)


def main() -> None:
    parser = argparse.ArgumentParser(description="找出超过阈值的引用组并提出拆分建议。")
    parser.add_argument("docx", help="DOCX 文件路径")
    parser.add_argument("--threshold", type=int, default=3, help="每个引用点最大文献数，默认 3")
    parser.add_argument("--json", action="store_true", help="以 JSON 格式输出")
    parser.add_argument("-o", "--output", help="输出文件路径")
    args = parser.parse_args()

    data = analyze_docx(args.docx, args.threshold)

    if args.json:
        content = json.dumps(data, ensure_ascii=False, indent=2)
    else:
        content = render_markdown_report(data)

    if args.output:
        Path(args.output).write_text(content + "\n", encoding="utf-8")
    else:
        print(content)


if __name__ == "__main__":
    main()
