from __future__ import annotations

import argparse
import json
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "v": "urn:schemas-microsoft-com:vml",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
}

for prefix, uri in NS.items():
    if prefix not in {"rel", "ct"}:
        ET.register_namespace(prefix, uri)
ET.register_namespace("", NS["ct"])

TITLE_TEXT = "参考文献"
SIMSUN = "宋体"
HEITI = "黑体"
TNR = "Times New Roman"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
CITATION_PATTERN = re.compile(r"\[(?:\d+(?:\s*[-,，]\s*\d+)*)\]")
MANUAL_NUMBER_PATTERN = re.compile(r"^\s*\[\d+\]")
FULLWIDTH_PUNCTUATION_PATTERN = re.compile(r"[，：；（）]")
ACCESS_DATE_PATTERN = re.compile(r"\[\d{4}-\d{2}-\d{2}\]")
URL_PATTERN = re.compile(r"https?://\S+")
ARXIV_PATTERN = re.compile(r"\barxiv\b|arxiv:\s*\d{4}\.\d{4,5}", re.IGNORECASE)
PREPRINT_PATTERN = re.compile(r"\bpreprint\b", re.IGNORECASE)
ELECTRONIC_TRACE_PATTERN = re.compile(
    r"\b(openreview|github|ctan|hugging\s*face|web\s?page|website|official\s+docs?|documentation|accessed|retrieved)\b"
    r"|访问日期|引用日期|网页|项目主页|官方文档|电子资源|预印本",
    re.IGNORECASE,
)
PAGE_RANGE_PATTERN = r"\d+[A-Za-z0-9:.\-–—]*\s*[-–—]\s*\d+[A-Za-z0-9:.\-–—]*"
JOURNAL_REFERENCE_PATTERN = re.compile(
    rf"^.+?\. .+?\[J\]\. .+?, \d{{4}}, [^:]+?\([^)]+\): {PAGE_RANGE_PATTERN}\.$"
)
BOOK_REFERENCE_PATTERN = re.compile(
    rf"^.+?\. .+?\[M\]\. [^:]+: [^,]+, \d{{4}}: {PAGE_RANGE_PATTERN}\.$"
)
IN_CONFERENCE_REFERENCE_PATTERN = re.compile(
    rf"^.+?\. .+?\. In.+?, .+?, .+?, .+?\. .+?, Eds?\., [^:]+: ?[^,]+, \d{{4}}: {PAGE_RANGE_PATTERN}\.$"
)
EXTRACTED_CONFERENCE_REFERENCE_PATTERN = re.compile(
    rf"^.+?\. .+?\[C\]\. //.+?\. .+?\. [^:]+: [^,]+, \d{{4}}: {PAGE_RANGE_PATTERN}\.$"
)
DISSERTATION_REFERENCE_PATTERN = re.compile(r"^.+?\. .+?\[D\]\. [^:]+: [^,]+, \d{4}\.$")
PATENT_REFERENCE_PATTERN = re.compile(r"^.+?\. .+?\[P\]\. [^,]+, .+\.$")
ENGLISH_CHAPTER_HEADING_PATTERN = re.compile(
    r"^(Appendix|Acknowledg(?:e)?ments|Conclusion|References)\b",
    re.IGNORECASE,
)


def qn(name: str) -> str:
    prefix, local = name.split(":", 1)
    return f"{{{NS[prefix]}}}{local}"


def get_w_val(element: ET.Element | None, default: str | None = None) -> str | None:
    if element is None:
        return default
    return element.get(qn("w:val"), default)


@dataclass
class DocxPackage:
    path: Path
    files: dict[str, bytes]
    document_root: ET.Element
    numbering_root: ET.Element | None


def load_docx(path: str | Path) -> DocxPackage:
    docx_path = Path(path)
    with zipfile.ZipFile(docx_path) as docx:
        files = {name: docx.read(name) for name in docx.namelist()}
    if "word/document.xml" not in files:
        raise ValueError(f"{docx_path} is missing word/document.xml")
    document_root = ET.fromstring(files["word/document.xml"])
    numbering_root = None
    if "word/numbering.xml" in files:
        numbering_root = ET.fromstring(files["word/numbering.xml"])
    return DocxPackage(docx_path, files, document_root, numbering_root)


def document_body(package: DocxPackage) -> ET.Element:
    body = package.document_root.find(qn("w:body"))
    if body is None:
        raise ValueError("word/document.xml is missing w:body")
    return body


def paragraphs(package: DocxPackage) -> list[ET.Element]:
    return [child for child in document_body(package) if child.tag == qn("w:p")]


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.iter(qn("w:t")))


def run_text(run: ET.Element) -> str:
    return "".join(node.text or "" for node in run.iter(qn("w:t")))


def child(parent: ET.Element | None, name: str) -> ET.Element | None:
    if parent is None:
        return None
    return parent.find(qn(name))


def paragraph_properties(paragraph: ET.Element) -> dict[str, Any]:
    ppr = child(paragraph, "w:pPr")
    spacing = child(ppr, "w:spacing")
    ind = child(ppr, "w:ind")
    jc = child(ppr, "w:jc")
    p_style = child(ppr, "w:pStyle")
    outline = child(ppr, "w:outlineLvl")
    num_pr = child(ppr, "w:numPr")
    ilvl = child(num_pr, "w:ilvl")
    num_id = child(num_pr, "w:numId")
    return {
        "style": get_w_val(p_style),
        "outline_level": get_w_val(outline),
        "alignment": get_w_val(jc),
        "spacing_before": spacing.get(qn("w:before")) if spacing is not None else None,
        "spacing_after": spacing.get(qn("w:after")) if spacing is not None else None,
        "spacing_line": spacing.get(qn("w:line")) if spacing is not None else None,
        "spacing_line_rule": spacing.get(qn("w:lineRule")) if spacing is not None else None,
        "hanging": ind.get(qn("w:hanging")) if ind is not None else None,
        "hanging_chars": ind.get(qn("w:hangingChars")) if ind is not None else None,
        "first_line": ind.get(qn("w:firstLine")) if ind is not None else None,
        "first_line_chars": ind.get(qn("w:firstLineChars")) if ind is not None else None,
        "left": ind.get(qn("w:left")) if ind is not None else None,
        "left_chars": ind.get(qn("w:leftChars")) if ind is not None else None,
        "num_id": get_w_val(num_id),
        "ilvl": get_w_val(ilvl, "0"),
    }


def run_properties(run: ET.Element | None) -> dict[str, Any]:
    if run is None:
        return {
            "font_east_asia": None,
            "font_ascii": None,
            "font_hansi": None,
            "font_size_pt": None,
            "vert_align": None,
            "is_superscript": False,
        }
    rpr = child(run, "w:rPr")
    fonts = child(rpr, "w:rFonts")
    sz = child(rpr, "w:sz")
    vert_align = child(rpr, "w:vertAlign")
    size_pt = None
    size_val = get_w_val(sz)
    if size_val and size_val.isdigit():
        size_pt = int(size_val) / 2
    vert = get_w_val(vert_align)
    return {
        "font_east_asia": fonts.get(qn("w:eastAsia")) if fonts is not None else None,
        "font_ascii": fonts.get(qn("w:ascii")) if fonts is not None else None,
        "font_hansi": fonts.get(qn("w:hAnsi")) if fonts is not None else None,
        "font_size_pt": size_pt,
        "vert_align": vert,
        "is_superscript": vert == "superscript",
    }


def first_run(paragraph: ET.Element) -> ET.Element | None:
    return paragraph.find(qn("w:r"))


def run_infos(paragraph: ET.Element) -> list[dict[str, Any]]:
    infos: list[dict[str, Any]] = []
    for index, run in enumerate(paragraph.findall(qn("w:r"))):
        text = run_text(run)
        if not text:
            continue
        infos.append({"run_index": index, "text": text, **run_properties(run)})
    return infos


def paragraph_run_fragments(paragraph: ET.Element) -> tuple[str, list[dict[str, Any]]]:
    text_parts: list[str] = []
    fragments: list[dict[str, Any]] = []
    offset = 0
    for index, run in enumerate(paragraph.findall(qn("w:r"))):
        text = run_text(run)
        if not text:
            continue
        start = offset
        end = start + len(text)
        text_parts.append(text)
        fragments.append({"run_index": index, "start": start, "end": end, "text": text, **run_properties(run)})
        offset = end
    return "".join(text_parts), fragments


def aggregate_fragment_value(fragments: list[dict[str, Any]], key: str) -> Any:
    values = [fragment.get(key) for fragment in fragments if fragment.get(key) is not None]
    if not values:
        return None
    first = values[0]
    if all(value == first for value in values):
        return first
    return "__mixed__"


def paragraph_field_instructions(paragraph: ET.Element) -> list[str]:
    return [node.text or "" for node in paragraph.iter(qn("w:instrText"))]


def paragraph_has_ref_field(paragraph: ET.Element) -> bool:
    joined = " ".join(paragraph_field_instructions(paragraph))
    return bool(re.search(r"\bREF\b", joined))


def paragraph_has_bookmark(paragraph: ET.Element) -> bool:
    return paragraph.find(qn("w:bookmarkStart")) is not None or paragraph.find(qn("w:bookmarkEnd")) is not None


def is_word_generated_ref_bookmark(name: str) -> bool:
    return bool(re.match(r"^_Ref\d+$", name))


def paragraph_bookmark_names(paragraph: ET.Element) -> set[str]:
    names = set()
    for node in paragraph.iter(qn("w:bookmarkStart")):
        name = node.get(qn("w:name"))
        if name:
            names.add(name)
    return names


def paragraph_has_fallback_bookmark(paragraph: ET.Element) -> bool:
    return any(not is_word_generated_ref_bookmark(name) for name in paragraph_bookmark_names(paragraph))


def resolve_numbering_formats(package: DocxPackage) -> dict[tuple[str, str], str]:
    if package.numbering_root is None:
        return {}
    abstract_map: dict[str, dict[str, str]] = {}
    for abstract in package.numbering_root.findall(qn("w:abstractNum")):
        abstract_id = abstract.get(qn("w:abstractNumId"))
        if abstract_id is None:
            continue
        levels: dict[str, str] = {}
        for level in abstract.findall(qn("w:lvl")):
            ilvl = level.get(qn("w:ilvl"), "0")
            lvl_text = child(level, "w:lvlText")
            value = get_w_val(lvl_text)
            if value:
                levels[ilvl] = value
        abstract_map[abstract_id] = levels

    resolved: dict[tuple[str, str], str] = {}
    for num in package.numbering_root.findall(qn("w:num")):
        num_id = num.get(qn("w:numId"))
        abstract_id_node = child(num, "w:abstractNumId")
        abstract_id = get_w_val(abstract_id_node)
        if num_id is None or abstract_id is None:
            continue
        for ilvl, fmt in abstract_map.get(abstract_id, {}).items():
            resolved[(num_id, ilvl)] = fmt
    return resolved


def reference_title_index(package: DocxPackage) -> int | None:
    for index, para in enumerate(paragraphs(package)):
        if paragraph_text(para).strip() == TITLE_TEXT:
            return index
    return None


def looks_like_chapter_heading(paragraph: ET.Element) -> bool:
    text = paragraph_text(paragraph).strip()
    if not text:
        return False
    compact = re.sub(r"\s+", "", text)
    if re.match(r"^第[0-9一二三四五六七八九十百千万]+[章节篇部分]", compact):
        return True
    if re.match(r"^(附录[0-9一二三四五六七八九十百千万]*|致谢|结论|摘要)", compact):
        return True
    if ENGLISH_CHAPTER_HEADING_PATTERN.match(text):
        return True

    props = paragraph_properties(paragraph)
    if props.get("outline_level") == "0":
        return True
    style = str(props.get("style") or "").lower()
    return style in {"heading1", "title"} or style.startswith("heading1")


def paragraph_info(package: DocxPackage, paragraph: ET.Element, index: int) -> dict[str, Any]:
    props = paragraph_properties(paragraph)
    first_props = run_properties(first_run(paragraph))
    numbering_formats = resolve_numbering_formats(package)
    numbering_format = None
    if props["num_id"]:
        numbering_format = numbering_formats.get((props["num_id"], props["ilvl"] or "0"))
    text = paragraph_text(paragraph).strip()
    return {
        "index": index,
        "text": text,
        "has_numbering": bool(props["num_id"]),
        "numbering_format": numbering_format,
        "starts_with_manual_number": bool(MANUAL_NUMBER_PATTERN.search(text)),
        "ends_with_period": text.endswith("."),
        "runs": run_infos(paragraph),
        **props,
        **first_props,
    }


def extract_reference_section(package: DocxPackage) -> dict[str, Any]:
    paras = paragraphs(package)
    title_index = reference_title_index(package)
    if title_index is None:
        return {"title": None, "entries": []}
    title = paragraph_info(package, paras[title_index], title_index)
    entries: list[dict[str, Any]] = []
    for index, para in enumerate(paras[title_index + 1 :], start=title_index + 1):
        text = paragraph_text(para).strip()
        if not text:
            continue
        if looks_like_chapter_heading(para):
            break
        entries.append(paragraph_info(package, para, index))
    return {"title": title, "entries": entries}


def extract_citations(package: DocxPackage) -> list[dict[str, Any]]:
    title_index = reference_title_index(package)
    target_paras = paragraphs(package)
    if title_index is not None:
        target_paras = target_paras[:title_index]

    citations: list[dict[str, Any]] = []
    for para_index, para in enumerate(target_paras):
        para_text = paragraph_text(para)
        has_ref = paragraph_has_ref_field(para)
        visible_text, fragments = paragraph_run_fragments(para)
        for match in CITATION_PATTERN.finditer(visible_text):
            covered = [fragment for fragment in fragments if fragment["start"] < match.end() and fragment["end"] > match.start()]
            if not covered:
                continue
            citations.append(
                {
                    "paragraph_index": para_index,
                    "run_index": covered[0]["run_index"],
                    "text": match.group(0),
                    "context": para_text,
                    "has_ref_field": has_ref,
                    "font_east_asia": aggregate_fragment_value(covered, "font_east_asia"),
                    "font_ascii": aggregate_fragment_value(covered, "font_ascii"),
                    "font_hansi": aggregate_fragment_value(covered, "font_hansi"),
                    "font_size_pt": aggregate_fragment_value(covered, "font_size_pt"),
                    "vert_align": aggregate_fragment_value(covered, "vert_align"),
                    "is_superscript": all(fragment.get("is_superscript") for fragment in covered),
                }
            )
    return citations


def bookmark_names(package: DocxPackage) -> set[str]:
    names = set()
    for node in package.document_root.iter(qn("w:bookmarkStart")):
        name = node.get(qn("w:name"))
        if name:
            names.add(name)
    return names


def ref_field_targets(package: DocxPackage) -> list[str]:
    targets: list[str] = []
    for instruction in package.document_root.iter(qn("w:instrText")):
        text = instruction.text or ""
        match = re.search(r"\bREF\s+([A-Za-z_][A-Za-z0-9_]*)", text)
        if match:
            targets.append(match.group(1))
    return targets


def fallback_ref_field_targets(package: DocxPackage) -> list[str]:
    return [target for target in ref_field_targets(package) if not is_word_generated_ref_bookmark(target)]


def paragraph_contains_locked_or_noneditable_construct(paragraph: ET.Element) -> bool:
    blocked_tags = {
        qn("w:sdt"),
        qn("w:txbxContent"),
        qn("w:object"),
        f"{{{NS['v']}}}shape",
    }
    return any(node.tag in blocked_tags for node in paragraph.iter())


def has_punctuation_spacing_issue(text: str) -> bool:
    for index, char in enumerate(text[:-1]):
        if char not in {",", ":"}:
            continue
        next_char = text[index + 1]
        if next_char.isspace():
            continue
        if char == ":":
            prev_char = text[index - 1] if index else ""
            if prev_char.isdigit() and next_char.isdigit():
                continue
            prefix = text[max(0, index - 5) : index].lower()
            if prefix.endswith("http") or prefix.endswith("https") or next_char == "/":
                continue
        return True
    return False


def has_default_reference_indent(entry: dict[str, Any]) -> bool:
    no_first_line = entry.get("first_line") in {None, "0"} and entry.get("first_line_chars") in {None, "0"}
    no_left = entry.get("left") in {None, "0"} and entry.get("left_chars") in {None, "0"}
    has_two_char_hanging = entry.get("hanging_chars") == "200" or entry.get("hanging") == "420"
    return no_first_line and no_left and has_two_char_hanging


def normalize_reference_text(text: str) -> str:
    return re.sub(r"^\s*\[\d+\]\s*", "", text.strip())


def _format_details(reference: str, pattern: re.Pattern[str], expected_fields: list[tuple[str, str]]) -> str:
    if pattern.match(reference):
        return ""
    missing = [label for label, detector in expected_fields if not re.search(detector, reference)]
    if missing:
        return "缺失或异常字段: " + "、".join(missing)
    return "字段顺序、标点或空格不符合模板"


def _reference_format_error(code: str, template: str, details: str = "") -> dict[str, str]:
    message = f"参考文献格式必须符合模板：{template}"
    if details:
        message += f"；{details}"
    return {"code": code, "message": message}


def validate_reference_text(text: str) -> list[dict[str, str]]:
    reference = normalize_reference_text(text)
    errors: list[dict[str, str]] = []

    def add(code: str, message: str) -> None:
        errors.append({"code": code, "message": message})

    def add_format(code: str, template: str, pattern: re.Pattern[str], fields: list[tuple[str, str]]) -> None:
        details = _format_details(reference, pattern, fields)
        if details:
            errors.append(_reference_format_error(code, template, details))

    if "[EB/OL]" in reference:
        add("reference.eb_ol_disallowed", "默认禁止使用第 6 条电子文献格式 [EB/OL]；请替换为正式出版文献，或向用户说明无法按当前规则写入。")
    if URL_PATTERN.search(reference):
        add("reference.url_disallowed", "最终参考文献条目不得残留 URL；电子来源只能作为检索线索或核验入口。")
    if ARXIV_PATTERN.search(reference) or PREPRINT_PATTERN.search(reference):
        add("reference.preprint_disallowed", "不得把 arXiv、preprint 或预印本伪装成正式参考文献；请核验正式出版版本或向用户说明。")
    if ACCESS_DATE_PATTERN.search(reference):
        add("reference.access_date_disallowed", "默认禁用电子文献格式，最终条目不得残留 [YYYY-MM-DD] 访问日期。")
    if ELECTRONIC_TRACE_PATTERN.search(reference):
        add("reference.electronic_trace_disallowed", "最终参考文献不得保留电子页面、官方文档、项目主页、访问说明或类似网页型来源痕迹；请核验正式出版记录。")

    if "[J]" in reference:
        add_format(
            "reference.journal_format",
            "作者. 题名[J]. 刊名, 年, 卷(期): 起-止页码.",
            JOURNAL_REFERENCE_PATTERN,
            [
                ("作者与题名", r"^.+?\. .+?\[J\]\."),
                ("刊名", r"\[J\]\. .+?,"),
                ("年份", r", \d{4},"),
                ("卷(期)", r", [^:]+?\([^)]+\):"),
                ("起-止页码", PAGE_RANGE_PATTERN + r"\.$"),
            ],
        )
    elif "[M]" in reference:
        add_format(
            "reference.book_format",
            "著者. 书名[M]. 出版地: 出版者, 出版年: 起-止页码.",
            BOOK_REFERENCE_PATTERN,
            [
                ("著者与书名", r"^.+?\. .+?\[M\]\."),
                ("出版地", r"\[M\]\. [^:]+:"),
                ("出版者", r": [^,]+,"),
                ("出版年", r", \d{4}:"),
                ("起-止页码", PAGE_RANGE_PATTERN + r"\.$"),
            ],
        )
    elif "[C]" in reference:
        add_format(
            "reference.extracted_conference_format",
            "作者. 题名[C]. //编者. 文集名. 出版地: 出版者, 出版年: 起-止页码.",
            EXTRACTED_CONFERENCE_REFERENCE_PATTERN,
            [
                ("作者与题名", r"^.+?\. .+?\[C\]"),
                ("[C]. //类型标识", r"\[C\]\. //"),
                ("编者", r"//.+?\. "),
                ("文集名", r"//.+?\. .+?\. "),
                ("出版地", r"\. [^:]+:"),
                ("出版者", r": [^,]+,"),
                ("出版年", r", \d{4}:"),
                ("起-止页码", PAGE_RANGE_PATTERN + r"\.$"),
            ],
        )
    elif "[D]" in reference:
        add_format(
            "reference.dissertation_format",
            "作者. 题名[D]. 授予学位地: 授予学位单位, 出版年.",
            DISSERTATION_REFERENCE_PATTERN,
            [
                ("作者与题名", r"^.+?\. .+?\[D\]\."),
                ("授予学位地", r"\[D\]\. [^:]+:"),
                ("授予学位单位", r": [^,]+,"),
                ("出版年", r", \d{4}\.$"),
            ],
        )
    elif "[P]" in reference:
        add_format(
            "reference.patent_format",
            "著者. 专利题名[P]. 专利号, 公告日期或公开日期.",
            PATENT_REFERENCE_PATTERN,
            [
                ("著者与专利题名", r"^.+?\. .+?\[P\]\."),
                ("专利号", r"\[P\]\. [^,]+,"),
                ("公告日期或公开日期", r", .+\.$"),
            ],
        )
    elif "[EB/OL]" in reference:
        # Already reported as disallowed. Do not add type_missing.
        pass
    elif ". In" in reference:
        add_format(
            "reference.in_conference_format",
            "著者. 题名. In文集名, 会议名, 会址, 开会时间. 编者, Eds., 出版地:出版者, 出版年: 页码范围.",
            IN_CONFERENCE_REFERENCE_PATTERN,
            [
                ("著者与题名", r"^.+?\. .+?\. In"),
                ("文集名/会议名/会址/开会时间", r"\. In.+?, .+?, .+?, .+?\."),
                ("编者", r"\. .+?, Eds?\."),
                ("出版地", r"Eds?\., [^:]+:"),
                ("出版者", r": ?[^,]+,"),
                ("出版年", r", \d{4}:"),
                ("页码范围", PAGE_RANGE_PATTERN + r"\.$"),
            ],
        )
    else:
        add("reference.type_missing", "参考文献条目缺少可识别类型，或未匹配期刊、专著、会议论文、论文集析出、学位论文、专利格式。")

    return errors


def parse_citation_numbers(citation_text: str) -> set[int]:
    inner = citation_text.strip()[1:-1]
    numbers: set[int] = set()
    for part in re.split(r"[,，]", inner):
        part = part.strip()
        if not part:
            continue
        range_match = re.match(r"^(\d+)\s*[-–—]\s*(\d+)$", part)
        if range_match:
            start, end = int(range_match.group(1)), int(range_match.group(2))
            if start <= end:
                numbers.update(range(start, end + 1))
            else:
                numbers.update(range(end, start + 1))
            continue
        if part.isdigit():
            numbers.add(int(part))
    return numbers


def audit_document(package: DocxPackage) -> list[dict[str, Any]]:
    issues: list[dict[str, Any]] = []

    def issue(code: str, message: str, location: str, severity: str = "error", text: str = "") -> None:
        issues.append(
            {
                "severity": severity,
                "code": code,
                "location": location,
                "message": message,
                "text": text,
            }
        )

    section = extract_reference_section(package)
    title = section["title"]
    if title is None:
        issue("reference.title_missing", "未找到“参考文献”标题。", "document")
    else:
        if title["alignment"] != "center":
            issue("reference.title_alignment", "“参考文献”标题未直接设置为居中。", f"p{title['index']}", "warning", title["text"])
        if title["spacing_before"] != "480":
            issue("reference.title_spacing_before", "“参考文献”标题段前不是 24 磅。", f"p{title['index']}", "warning", title["text"])
        if title["spacing_after"] != "240":
            issue("reference.title_spacing_after", "“参考文献”标题段后不是 12 磅。", f"p{title['index']}", "warning", title["text"])
        if title["font_east_asia"] and title["font_east_asia"] != HEITI:
            issue("reference.title_font", "“参考文献”标题字体不是黑体。", f"p{title['index']}", "warning", title["text"])
        if title["font_size_pt"] and title["font_size_pt"] != 16:
            issue("reference.title_size", "“参考文献”标题字号不是三号 16 磅。", f"p{title['index']}", "warning", title["text"])

    entries = section["entries"]
    if title is not None and not entries:
        issue("reference.entries_missing", "参考文献节没有条目。", "reference-section")

    for entry in entries:
        location = f"p{entry['index']}"
        text = entry["text"]
        if entry["starts_with_manual_number"] and not entry["has_numbering"]:
            issue("reference.manual_numbering", "条目疑似使用手打 [n]，不是 Word 真实编号项。", location, "error", text)
        if entry["starts_with_manual_number"] and entry["has_numbering"]:
            issue("reference.manual_number_in_numbered_entry", "条目已使用 Word 编号项，但正文又手打了 [n]，会形成双编号。", location, "error", text)
        if not entry["has_numbering"]:
            issue("reference.missing_numbering", "参考文献条目未使用 Word 真实编号项。", location, "error", text)
        elif entry["numbering_format"] != "[%1]":
            issue("reference.numbering_format", "参考文献编号格式不是 [%1]。", location, "error", text)
        if FULLWIDTH_PUNCTUATION_PATTERN.search(text):
            issue("reference.fullwidth_punctuation", "条目中存在全角逗号、冒号、括号或分号。", location, "error", text)
        if has_punctuation_spacing_issue(text):
            issue("reference.punctuation_spacing", "半角逗号或冒号后缺少空格。", location, "error", text)
        if not entry["ends_with_period"]:
            issue("reference.period_missing", "参考文献条目未以英文句点结尾。", location, "error", text)
        if entry["alignment"] not in {None, "both"}:
            issue("reference.alignment", "参考文献条目不是双端对齐。", location, "warning", text)
        if entry["spacing_line"] != "300":
            issue("reference.line_spacing", "参考文献条目未直接设置为 1.25 倍行距。", location, "warning", text)
        if not has_default_reference_indent(entry):
            issue("reference.indent", "参考文献条目应无首行缩进、无左缩进，并设置悬挂缩进两个字符。", location, "warning", text)
        if entry["font_east_asia"] and entry["font_east_asia"] != SIMSUN:
            issue("reference.cjk_font", "参考文献中文字体不是宋体。", location, "warning", text)
        if entry["font_size_pt"] and entry["font_size_pt"] != 10.5:
            issue("reference.font_size", "参考文献字号不是五号 10.5 磅。", location, "warning", text)
        for run in entry.get("runs", []):
            run_location = f"{location}:r{run['run_index']}"
            run_text_value = run["text"]
            if run["font_east_asia"] and run["font_east_asia"] != SIMSUN:
                issue("reference.run_cjk_font", "参考文献条目中存在非宋体中文字体运行。", run_location, "warning", run_text_value)
            if run["font_ascii"] and run["font_ascii"] != TNR:
                issue("reference.run_ascii_font", "参考文献条目中存在非 Times New Roman 英文字体运行。", run_location, "warning", run_text_value)
            if run["font_size_pt"] and run["font_size_pt"] != 10.5:
                issue("reference.run_font_size", "参考文献条目中存在非五号 10.5 磅运行。", run_location, "warning", run_text_value)
        for validation_error in validate_reference_text(text):
            issue(validation_error["code"], validation_error["message"], location, "error", text)
        paragraph = paragraphs(package)[entry["index"]]
        if paragraph_has_fallback_bookmark(paragraph):
            issue("reference.bookmark_present", "参考文献条目包含书签；这应仅作为编号项交叉引用失败后的 fallback 使用。", location, "warning", text)
        if paragraph_contains_locked_or_noneditable_construct(paragraph):
            issue("reference.noneditable_construct", "参考文献条目包含文本框、内容控件或对象等不利于 Word 正常编辑的结构。", location, "error", text)

    citations = extract_citations(package)
    used_reference_numbers: set[int] = set()
    for citation in citations:
        location = f"p{citation['paragraph_index']}:r{citation['run_index']}"
        citation_numbers = parse_citation_numbers(citation["text"])
        for number in citation_numbers:
            if number < 1 or number > len(entries):
                issue("citation.number_out_of_range", f"正文引用编号 {number} 超出参考文献条目数量 {len(entries)}。", location, "error", citation["text"])
            else:
                used_reference_numbers.add(number)
        if not citation["is_superscript"]:
            issue("citation.not_superscript", "正文引用不是上标。", location, "error", citation["text"])
        raw_fonts = {citation.get("font_ascii"), citation.get("font_hansi"), citation.get("font_east_asia")}
        if "__mixed__" in raw_fonts:
            issue("citation.font", "正文引用字体不一致；应统一为 Times New Roman。", location, "warning", citation["text"])
        citation_fonts = raw_fonts - {None, "__mixed__"}
        if citation_fonts and citation_fonts != {TNR}:
            issue("citation.font", "正文引用字体不是 Times New Roman。", location, "warning", citation["text"])
        if citation["font_size_pt"] and citation["font_size_pt"] != 12:
            issue("citation.font_size", "正文引用字号不是小四 12 磅。", location, "warning", citation["text"])
        if not citation["has_ref_field"]:
            issue("citation.not_cross_reference", "正文引用未检测到可机器识别的编号项交叉引用字段。", location, "error", citation["text"])

    if citations:
        for number, entry in enumerate(entries, start=1):
            if number not in used_reference_numbers:
                issue("reference.unused_entry", f"参考文献条目 {number} 未在正文可检测引用中出现。", f"p{entry['index']}", "warning", entry["text"])

    targets = fallback_ref_field_targets(package)
    if targets:
        issue("cross_reference.bookmark_ref_present", "检测到 REF 书签交叉引用；应确认这是编号项交叉引用失败后的 fallback，而不是首选方案。", "document", "warning", ", ".join(sorted(set(targets))))

    return issues


def render_markdown_report(path: str | Path, issues: list[dict[str, Any]]) -> str:
    lines = [f"# GB/T 7714 DOCX 审计报告", "", f"文件: `{Path(path)}`", ""]
    if not issues:
        lines.extend(["未发现机器可检测的问题。", ""])
        return "\n".join(lines)
    lines.append(f"共发现 {len(issues)} 个问题。")
    lines.append("")
    for index, item in enumerate(issues, start=1):
        lines.append(f"## {index}. {item['code']} ({item['severity']})")
        lines.append("")
        lines.append(f"- 位置: `{item['location']}`")
        lines.append(f"- 说明: {item['message']}")
        if item.get("text"):
            lines.append(f"- 文本: {item['text']}")
        lines.append("")
    return "\n".join(lines)


def package_summary(package: DocxPackage) -> dict[str, Any]:
    section = extract_reference_section(package)
    return {
        "path": str(package.path),
        "citations": extract_citations(package),
        "reference_section": section,
        "ref_field_targets": ref_field_targets(package),
        "fallback_ref_field_targets": fallback_ref_field_targets(package),
        "bookmarks": sorted(bookmark_names(package)),
    }


def dump_json(data: Any) -> str:
    return json.dumps(data, ensure_ascii=False, indent=2)


def main() -> None:
    parser = argparse.ArgumentParser(description="Inspect GB/T 7714 structures in a DOCX file.")
    parser.add_argument("docx", help="DOCX file to inspect")
    parser.add_argument("--audit", action="store_true", help="emit audit issues instead of an extraction summary")
    args = parser.parse_args()

    package = load_docx(args.docx)
    if args.audit:
        print(dump_json(audit_document(package)))
    else:
        print(dump_json(package_summary(package)))


if __name__ == "__main__":
    main()
