from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any

sys.dont_write_bytecode = True

from 公共库 import extract_citations, load_docx, parse_citation_numbers


KEY_PHRASE_PATTERN = re.compile(
    r"[A-Za-z][A-Za-z0-9+\-]*(?:\s+[A-Za-z][A-Za-z0-9+\-]*){0,4}"
    r"|[\u4e00-\u9fffA-Za-z0-9+\-]{2,}(?:模型|方法|算法|框架|工具|系统|机制|流程|策略|结构|编码器|解码器|数据集|损失|训练|微调|推理|搜索|编译|渲染|向量化|图形|代码|表示)"
)


def candidate_anchor(part: str) -> str:
    """选择尽量短的引用落点，优先模型/方法/工具等短语。"""
    part = part.strip()
    matches = [m.group(0).strip() for m in KEY_PHRASE_PATTERN.finditer(part) if m.group(0).strip()]
    if matches:
        # Prefer the most specific short anchor over a long clause.
        def score(item: str) -> tuple[int, bool, int]:
            has_domain_suffix = bool(re.search(r"(模型|方法|算法|框架|工具|系统|机制|流程|策略|结构|编码器|解码器|数据集|损失|训练|微调|推理|搜索|编译|渲染|向量化|图形|代码|表示)$", item))
            has_model_like_case = bool(re.search(r"[A-Z][A-Za-z0-9+\-]+(?:\s+[A-Z][A-Za-z0-9+\-]+)+|[A-Z]{2,}", item))
            starts_with_model_name = bool(re.match(r"^[A-Z][A-Za-z0-9+\-]+(?:\s+[A-Z][A-Za-z0-9+\-]+)*$", item))
            rank = 0 if starts_with_model_name else 1 if has_model_like_case else 2 if has_domain_suffix else 3
            return (rank, len(item) > 40, len(item))

        matches.sort(key=score)
        return matches[0][:40]
    return part[:40]


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

    基于上下文中的模型名、方法名、工具名和名词短语提出建议。
    尽量给出短语级引用落点，避免把引用默认放在子句或句末。
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
                anchor = candidate_anchor(part)
                suggestions.append({
                    "建议短语": anchor,
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
                "建议短语": candidate_anchor(context),
                "分配编号": remaining,
                "原句片段": context[:80],
            })

    merged: list[dict] = []
    for s in suggestions:
        if merged and len(merged[-1]["分配编号"]) + len(s["分配编号"]) <= 3:
            merged[-1]["分配编号"].extend(s["分配编号"])
            if s["建议短语"] not in merged[-1]["建议短语"]:
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
