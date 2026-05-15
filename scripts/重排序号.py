from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from copy import deepcopy
from pathlib import Path
from xml.etree import ElementTree as ET

sys.dont_write_bytecode = True

from 公共库 import (
    NS,
    XML_SPACE,
    extract_citations,
    load_docx,
    parse_citation_numbers,
    qn,
)

for prefix, uri in NS.items():
    if prefix not in {"rel", "ct"}:
        ET.register_namespace(prefix, uri)
ET.register_namespace("", NS["ct"])


def collect_first_appearance_order(citations: list[dict]) -> list[int]:
    """按正文首次出现顺序收集参考文献编号。"""
    seen: set[int] = set()
    order: list[int] = []
    for citation in citations:
        for number in parse_citation_numbers(citation["text"]):
            if number not in seen:
                seen.add(number)
                order.append(number)
    return order


def build_renumber_map(old_order: list[int]) -> dict[int, int]:
    """构建旧编号到新编号的映射。"""
    return {old: new for new, old in enumerate(old_order, start=1)}


def renumber_citation_text(text: str, renumber_map: dict[int, int]) -> str:
    """重写引用文本中的编号。"""
    inner = text.strip()[1:-1]
    parts = re.split(r"([,，])", inner)
    new_parts: list[str] = []
    for part in parts:
        stripped = part.strip()
        if not stripped or stripped in {",", "，"}:
            new_parts.append(part)
            continue
        range_match = re.match(r"^(\d+)\s*[-–—]\s*(\d+)$", stripped)
        if range_match:
            old_start = int(range_match.group(1))
            old_end = int(range_match.group(2))
            new_start = renumber_map.get(old_start, old_start)
            new_end = renumber_map.get(old_end, old_end)
            new_parts.append(f"{new_start}-{new_end}")
            continue
        if stripped.isdigit():
            old_num = int(stripped)
            new_num = renumber_map.get(old_num, old_num)
            new_parts.append(str(new_num))
            continue
        new_parts.append(stripped)
    return "[" + ",".join(new_parts) + "]"


def renumber_reference_entries(body: ET.Element, title_text: str, renumber_map: dict[int, int], total_old: int) -> None:
    """重排参考文献条目段落顺序。"""
    title_para = None
    for para in body.findall(qn("w:p")):
        text = "".join(node.text or "" for node in para.iter(qn("w:t"))).strip()
        if text == title_text:
            title_para = para
            break
    if title_para is None:
        return

    children = list(body)
    title_index = children.index(title_para)
    entry_paras: list[tuple[int, ET.Element]] = []
    for idx, child in enumerate(children[title_index + 1:], start=title_index + 1):
        if child.tag != qn("w:p"):
            break
        text = "".join(node.text or "" for node in child.iter(qn("w:t"))).strip()
        if not text:
            break
        entry_paras.append((idx, child))

    if len(entry_paras) < total_old:
        total_old = len(entry_paras)

    reordered: list[ET.Element] = [None] * total_old
    for old_num in range(1, total_old + 1):
        new_num = renumber_map.get(old_num, old_num)
        if old_num <= len(entry_paras):
            reordered[new_num - 1] = entry_paras[old_num - 1][1]

    reordered = [p for p in reordered if p is not None]

    for idx, _ in entry_paras:
        body.remove(children[idx])
    for offset, para in enumerate(reordered):
        body.insert(title_index + 1 + offset, para)


def renumber_docx(source: Path, output: Path) -> dict:
    """读取 DOCX，按正文首次出现顺序重排编号，输出新 DOCX。"""
    package = load_docx(source)
    citations = extract_citations(package)
    old_order = collect_first_appearance_order(citations)
    if not old_order:
        return {"changed": False, "reason": "未检测到正文引用。"}

    total_old = max(old_order)
    renumber_map = build_renumber_map(old_order)
    identity = all(old == new for old, new in renumber_map.items())
    if identity:
        return {"changed": False, "reason": "编号已经是正文首次出现顺序。"}

    body = package.document_root.find(qn("w:body"))
    if body is None:
        raise ValueError("word/document.xml 缺少 w:body")

    for para in body.iter(qn("w:p")):
        full_text = "".join(node.text or "" for node in para.iter(qn("w:t")))
        if not re.search(r"\[\d+", full_text):
            continue
        for run in para.findall(qn("w:r")):
            for t_node in run.findall(qn("w:t")):
                old_text = t_node.text or ""
                if re.search(r"\[\d+", old_text):
                    t_node.text = renumber_citation_text(old_text, renumber_map)

    renumber_reference_entries(body, "参考文献", renumber_map, total_old)

    files = dict(package.files)
    files["word/document.xml"] = ET.tostring(package.document_root, encoding="utf-8", xml_declaration=True)
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)

    return {
        "changed": True,
        "mapping": {str(k): v for k, v in renumber_map.items()},
        "output": str(output),
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="按正文首次出现顺序重排参考文献编号。")
    parser.add_argument("source_docx", help="输入 DOCX")
    parser.add_argument("output_docx", help="输出 DOCX")
    parser.add_argument("--json", action="store_true", help="以 JSON 格式输出结果")
    args = parser.parse_args()

    result = renumber_docx(Path(args.source_docx), Path(args.output_docx))
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        if result["changed"]:
            print(f"已重排编号，输出: {result['output']}")
            for old, new in result.get("mapping", {}).items():
                print(f"  [{old}] -> [{new}]")
        else:
            print(result["reason"])


if __name__ == "__main__":
    main()
