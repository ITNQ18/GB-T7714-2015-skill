from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from copy import deepcopy
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

sys.dont_write_bytecode = True

from 公共库 import CITATION_PATTERN, HEITI, NS, SIMSUN, TITLE_TEXT, TNR, XML_SPACE, looks_like_chapter_heading, qn, run_text, validate_reference_text


PLACEHOLDER_PATTERN = re.compile(r"\[\[CITE:([A-Za-z0-9_,.\- ]+)\]\]")
FORMAL_SOURCE_KEYS = (
    "verification_source",
    "verification_url",
    "source",
    "source_url",
    "publisher_url",
    "database_record",
    "formal_source",
    "doi",
    "DOI",
    "isbn",
    "ISBN",
    "issn",
    "ISSN",
    "核验来源",
    "核验来源 URL/DOI",
    "正式来源",
    "来源",
)
ELECTRONIC_ONLY_SOURCE_PATTERN = re.compile(
    r"\b(arxiv|preprint|openreview|github|ctan|hugging\s*face|official\s+docs?|documentation|project\s+page)\b"
    r"|预印本|官方文档|项目主页|网页",
    re.IGNORECASE,
)
NUMBERING_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
SETTINGS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"
NUMBERING_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
SETTINGS_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
MC_IGNORABLE = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable"


class NumberedItemCrossReferenceError(RuntimeError):
    pass


def safe_bookmark_name(key: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9_]", "_", key).strip("_")
    if not cleaned or cleaned[0].isdigit():
        cleaned = f"r_{cleaned}"
    return f"ref_{cleaned}"[:40]


def read_manifest(path: str | Path) -> list[dict[str, Any]]:
    data = json.loads(Path(path).read_text(encoding="utf-8-sig"))
    if isinstance(data, list):
        return data
    if isinstance(data, dict):
        for key in ("references", "records", "items"):
            value = data.get(key)
            if isinstance(value, list):
                return value
    raise ValueError("Manifest must be a JSON list or object containing references/records/items.")


def make_rpr(font_size_half_points: str, east_asia: str = SIMSUN, ascii_font: str = TNR, superscript: bool = False) -> ET.Element:
    rpr = ET.Element(qn("w:rPr"))
    fonts = ET.SubElement(rpr, qn("w:rFonts"))
    fonts.set(qn("w:eastAsia"), east_asia)
    fonts.set(qn("w:ascii"), ascii_font)
    fonts.set(qn("w:hAnsi"), ascii_font)
    sz = ET.SubElement(rpr, qn("w:sz"))
    sz.set(qn("w:val"), font_size_half_points)
    sz_cs = ET.SubElement(rpr, qn("w:szCs"))
    sz_cs.set(qn("w:val"), font_size_half_points)
    if superscript:
        vert = ET.SubElement(rpr, qn("w:vertAlign"))
        vert.set(qn("w:val"), "superscript")
    return rpr


def make_text_run(text: str, rpr: ET.Element | None = None) -> ET.Element:
    run = ET.Element(qn("w:r"))
    if rpr is not None:
        run.append(deepcopy(rpr))
    text_node = ET.SubElement(run, qn("w:t"))
    if text[:1].isspace() or text[-1:].isspace():
        text_node.set(XML_SPACE, "preserve")
    text_node.text = text
    return run


def make_field_char_run(kind: str, rpr: ET.Element) -> ET.Element:
    run = ET.Element(qn("w:r"))
    run.append(deepcopy(rpr))
    fld = ET.SubElement(run, qn("w:fldChar"))
    fld.set(qn("w:fldCharType"), kind)
    return run


def make_instr_run(instruction: str, rpr: ET.Element) -> ET.Element:
    run = ET.Element(qn("w:r"))
    run.append(deepcopy(rpr))
    instr = ET.SubElement(run, qn("w:instrText"))
    instr.set(XML_SPACE, "preserve")
    instr.text = instruction
    return run


def make_ref_field_runs(bookmark: str, display_text: str) -> list[ET.Element]:
    rpr = make_rpr("24", TNR, TNR, superscript=True)
    return [
        make_field_char_run("begin", rpr),
        make_instr_run(f" REF {bookmark} \\r \\h ", rpr),
        make_field_char_run("separate", rpr),
        make_text_run(display_text, rpr),
        make_field_char_run("end", rpr),
    ]


def make_citation_replacement_runs(source_text: str, references: list[dict[str, Any]]) -> list[ET.Element]:
    key_to_index = {str(item["key"]): index for index, item in enumerate(references, start=1)}
    key_to_bookmark = {key: safe_bookmark_name(key) for key in key_to_index}
    new_runs: list[ET.Element] = []
    cursor = 0
    for match in PLACEHOLDER_PATTERN.finditer(source_text):
        before = source_text[cursor : match.start()]
        if before:
            new_runs.append(make_text_run(before))
        keys = [key.strip() for key in match.group(1).split(",") if key.strip()]
        new_runs.append(make_text_run("[", make_rpr("24", TNR, TNR, superscript=True)))
        for offset, key in enumerate(keys):
            if key not in key_to_index:
                raise ValueError(f"Citation key not found in manifest: {key}")
            if offset:
                new_runs.append(make_text_run(",", make_rpr("24", TNR, TNR, superscript=True)))
            new_runs.extend(make_ref_field_runs(key_to_bookmark[key], str(key_to_index[key])))
        new_runs.append(make_text_run("]", make_rpr("24", TNR, TNR, superscript=True)))
        cursor = match.end()
    after = source_text[cursor:]
    if after:
        new_runs.append(make_text_run(after))
    return new_runs


def paragraph_visible_run_text(paragraph: ET.Element) -> str:
    return "".join(run_text(run) for run in paragraph.findall(qn("w:r")))


def paragraph_has_preserved_citation_structure(paragraph: ET.Element) -> bool:
    if paragraph.find(qn("w:bookmarkStart")) is not None or paragraph.find(qn("w:bookmarkEnd")) is not None:
        return True
    if paragraph.find(f".//{qn('w:fldChar')}") is not None or paragraph.find(f".//{qn('w:instrText')}") is not None:
        return True
    for run in paragraph.findall(qn("w:r")):
        text = run_text(run)
        if not CITATION_PATTERN.search(text):
            continue
        rpr = run.find(qn("w:rPr"))
        vert = rpr.find(qn("w:vertAlign")) if rpr is not None else None
        if vert is not None and vert.get(qn("w:val")) == "superscript":
            return True
    return False


def document_citation_structure_counts(root: ET.Element) -> dict[str, int]:
    return {
        "field_chars": sum(1 for _ in root.iter(qn("w:fldChar"))),
        "ref_instructions": sum(1 for node in root.iter(qn("w:instrText")) if re.search(r"\bREF\b", node.text or "")),
        "bookmarks": sum(1 for _ in root.iter(qn("w:bookmarkStart"))),
    }


def assert_citation_structure_not_reduced(before: dict[str, int], after: dict[str, int]) -> None:
    reduced = [key for key, value in before.items() if after.get(key, 0) < value]
    if reduced:
        details = ", ".join(f"{key}: {before[key]} -> {after.get(key, 0)}" for key in reduced)
        raise ValueError(f"DOCX citation structure was reduced while replacing text runs: {details}")


def first_run_insert_index(paragraph: ET.Element) -> int:
    children = list(paragraph)
    for index, child in enumerate(children):
        if child.tag == qn("w:r"):
            return index
    for index, child in enumerate(children):
        if child.tag != qn("w:pPr"):
            return index
    return len(children)


def replace_citation_placeholders(root: ET.Element, references: list[dict[str, Any]]) -> None:
    for paragraph in root.iter(qn("w:p")):
        source_text = paragraph_visible_run_text(paragraph)
        if not PLACEHOLDER_PATTERN.search(source_text):
            continue
        if paragraph_has_preserved_citation_structure(paragraph):
            raise ValueError("Refusing to rebuild a paragraph that already contains superscript citations, Word fields, or bookmarks.")
        new_runs = make_citation_replacement_runs(source_text, references)
        insert_at = first_run_insert_index(paragraph)
        for run in list(paragraph.findall(qn("w:r"))):
            paragraph.remove(run)
        for offset, new_run in enumerate(new_runs):
            paragraph.insert(insert_at + offset, new_run)


def find_reference_title_paragraph(body: ET.Element) -> ET.Element | None:
    for element in body:
        if element.tag == qn("w:p"):
            text = "".join(node.text or "" for node in element.iter(qn("w:t"))).strip()
            if text == TITLE_TEXT:
                return element
    return None


def remove_existing_reference_section(body: ET.Element) -> int | None:
    title = find_reference_title_paragraph(body)
    if title is None:
        return None
    children = list(body)
    start = children.index(title)
    end = len(children)
    for index, element in enumerate(children[start + 1 :], start=start + 1):
        if element.tag == qn("w:sectPr"):
            end = index
            break
        if element.tag == qn("w:p") and looks_like_chapter_heading(element):
            end = index
            break
    for element in children[start:end]:
        body.remove(element)
    return start


def make_title_paragraph() -> ET.Element:
    paragraph = ET.Element(qn("w:p"))
    ppr = ET.SubElement(paragraph, qn("w:pPr"))
    spacing = ET.SubElement(ppr, qn("w:spacing"))
    spacing.set(qn("w:before"), "480")
    spacing.set(qn("w:after"), "240")
    jc = ET.SubElement(ppr, qn("w:jc"))
    jc.set(qn("w:val"), "center")
    paragraph.append(make_text_run(TITLE_TEXT, make_rpr("32", HEITI, TNR)))
    return paragraph


def make_reference_paragraph(item: dict[str, Any], num_id: str, ref_index: int, use_bookmark: bool = False) -> ET.Element:
    paragraph = ET.Element(qn("w:p"))
    ppr = ET.SubElement(paragraph, qn("w:pPr"))
    num_pr = ET.SubElement(ppr, qn("w:numPr"))
    ilvl = ET.SubElement(num_pr, qn("w:ilvl"))
    ilvl.set(qn("w:val"), "0")
    num = ET.SubElement(num_pr, qn("w:numId"))
    num.set(qn("w:val"), num_id)
    spacing = ET.SubElement(ppr, qn("w:spacing"))
    spacing.set(qn("w:line"), "300")
    spacing.set(qn("w:lineRule"), "auto")
    jc = ET.SubElement(ppr, qn("w:jc"))
    jc.set(qn("w:val"), "both")
    ind = ET.SubElement(ppr, qn("w:ind"))
    ind.set(qn("w:left"), "0")
    ind.set(qn("w:hangingChars"), "200")
    bookmark_id = None
    if use_bookmark:
        bookmark_id = ref_index
        bookmark = ET.SubElement(paragraph, qn("w:bookmarkStart"))
        bookmark.set(qn("w:id"), str(bookmark_id))
        bookmark.set(qn("w:name"), safe_bookmark_name(str(item["key"])))
    paragraph.append(make_text_run(str(item["gbt7714"]), make_rpr("21", SIMSUN, TNR)))
    if use_bookmark and bookmark_id is not None:
        bookmark_end = ET.SubElement(paragraph, qn("w:bookmarkEnd"))
        bookmark_end.set(qn("w:id"), str(bookmark_id))
    return paragraph


def make_reference_tail_paragraph() -> ET.Element:
    paragraph = ET.Element(qn("w:p"))
    ET.SubElement(paragraph, qn("w:pPr"))
    return paragraph


def existing_integer_ids(root: ET.Element | None, element_name: str, attr_name: str) -> list[int]:
    if root is None:
        return []
    ids = []
    for element in root.findall(qn(element_name)):
        value = element.get(qn(attr_name))
        if value and value.isdigit():
            ids.append(int(value))
    return ids


def make_numbering_root() -> ET.Element:
    return ET.Element(qn("w:numbering"))


def ensure_reference_numbering(numbering_root: ET.Element | None) -> tuple[ET.Element, str]:
    if numbering_root is None:
        numbering_root = make_numbering_root()

    abstract_ids = existing_integer_ids(numbering_root, "w:abstractNum", "w:abstractNumId")
    num_ids = existing_integer_ids(numbering_root, "w:num", "w:numId")
    abstract_id = str(max(abstract_ids, default=0) + 1)
    num_id = str(max(num_ids, default=0) + 1)

    abstract = ET.SubElement(numbering_root, qn("w:abstractNum"))
    abstract.set(qn("w:abstractNumId"), abstract_id)
    multi = ET.SubElement(abstract, qn("w:multiLevelType"))
    multi.set(qn("w:val"), "singleLevel")
    lvl = ET.SubElement(abstract, qn("w:lvl"))
    lvl.set(qn("w:ilvl"), "0")
    start = ET.SubElement(lvl, qn("w:start"))
    start.set(qn("w:val"), "1")
    fmt = ET.SubElement(lvl, qn("w:numFmt"))
    fmt.set(qn("w:val"), "decimal")
    lvl_text = ET.SubElement(lvl, qn("w:lvlText"))
    lvl_text.set(qn("w:val"), "[%1]")
    lvl_jc = ET.SubElement(lvl, qn("w:lvlJc"))
    lvl_jc.set(qn("w:val"), "left")

    num = ET.SubElement(numbering_root, qn("w:num"))
    num.set(qn("w:numId"), num_id)
    abstract_ref = ET.SubElement(num, qn("w:abstractNumId"))
    abstract_ref.set(qn("w:val"), abstract_id)
    return numbering_root, num_id


def body_default_reference_insert_index(body: ET.Element) -> int:
    children = list(body)
    for index, child in enumerate(children):
        if child.tag != qn("w:p"):
            continue
        text = "".join(node.text or "" for node in child.iter(qn("w:t"))).strip()
        compact = re.sub(r"\s+", "", text)
        if re.match(r"^(附录[0-9一二三四五六七八九十百千万]*|致谢)", compact, re.IGNORECASE):
            return index
        if re.match(r"^(Appendix|Acknowledg(?:e)?ments)\b", text, re.IGNORECASE):
            return index
    for index, child in enumerate(children):
        if child.tag == qn("w:sectPr"):
            return index
    return len(children)


def body_insert_many(body: ET.Element, index: int, elements: list[ET.Element]) -> None:
    for offset, element in enumerate(elements):
        body.insert(index + offset, element)


def append_reference_section(root: ET.Element, references: list[dict[str, Any]], num_id: str, use_bookmarks: bool = False) -> None:
    body = root.find(qn("w:body"))
    if body is None:
        raise ValueError("word/document.xml is missing w:body")
    insert_index = remove_existing_reference_section(body)
    if insert_index is None:
        insert_index = body_default_reference_insert_index(body)
    elements = [make_title_paragraph()]
    for index, item in enumerate(references, start=1):
        elements.append(make_reference_paragraph(item, num_id, index, use_bookmark=use_bookmarks))
    elements.append(make_reference_tail_paragraph())
    body_insert_many(body, insert_index, elements)


def ensure_content_type(files: dict[str, bytes], part_name: str, content_type: str) -> None:
    root = ET.fromstring(files.get("[Content_Types].xml", b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'))
    for override in root.findall(f"{{{NS['ct']}}}Override"):
        if override.get("PartName") == part_name:
            return
    override = ET.SubElement(root, f"{{{NS['ct']}}}Override")
    override.set("PartName", part_name)
    override.set("ContentType", content_type)
    files["[Content_Types].xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)


def ensure_document_relationship(files: dict[str, bytes], rel_type: str, target: str) -> None:
    rels_path = "word/_rels/document.xml.rels"
    root = ET.fromstring(files.get(rels_path, f'<Relationships xmlns="{NS["rel"]}"/>'.encode("utf-8")))
    for rel in root.findall(f"{{{NS['rel']}}}Relationship"):
        if rel.get("Type") == rel_type and rel.get("Target") == target:
            return
    existing = [rel.get("Id", "") for rel in root.findall(f"{{{NS['rel']}}}Relationship")]
    numeric = [int(value[3:]) for value in existing if value.startswith("rId") and value[3:].isdigit()]
    rel = ET.SubElement(root, f"{{{NS['rel']}}}Relationship")
    rel.set("Id", f"rId{max(numeric, default=0) + 1}")
    rel.set("Type", rel_type)
    rel.set("Target", target)
    files[rels_path] = ET.tostring(root, encoding="utf-8", xml_declaration=True)


def ensure_update_fields(files: dict[str, bytes]) -> None:
    settings_path = "word/settings.xml"
    root = ET.fromstring(files.get(settings_path, f'<w:settings xmlns:w="{NS["w"]}"/>'.encode("utf-8")))
    update = root.find(qn("w:updateFields"))
    if update is None:
        update = ET.SubElement(root, qn("w:updateFields"))
    update.set(qn("w:val"), "true")
    files[settings_path] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    ensure_content_type(files, "/word/settings.xml", SETTINGS_CONTENT_TYPE)
    ensure_document_relationship(files, SETTINGS_REL_TYPE, "settings.xml")


def verification_source_value(item: dict[str, Any]) -> str:
    values: list[str] = []
    for key in FORMAL_SOURCE_KEYS:
        value = item.get(key)
        if value is None:
            continue
        if isinstance(value, list):
            values.extend(str(part).strip() for part in value if str(part).strip())
        else:
            text = str(value).strip()
            if text:
                values.append(text)
    return " ; ".join(values)


def source_is_electronic_only(source: str) -> bool:
    return bool(ELECTRONIC_ONLY_SOURCE_PATTERN.search(source))


def validate_manifest(references: list[dict[str, Any]]) -> None:
    seen: set[str] = set()
    for index, item in enumerate(references, start=1):
        if "key" not in item:
            raise ValueError(f"Reference #{index} is missing key")
        if "gbt7714" not in item:
            raise ValueError(f"Reference #{index} is missing gbt7714")
        key = str(item["key"])
        if key in seen:
            raise ValueError(f"Duplicate reference key: {key}")
        seen.add(key)
        entry = str(item["gbt7714"])
        if re.match(r"^\s*\[\d+\]", entry):
            raise ValueError(f"GB/T 7714 entry must not include manual numbering: {key}")
        if not entry.endswith("."):
            raise ValueError(f"GB/T 7714 entry must end with an English period: {key}")
        validation_errors = validate_reference_text(entry)
        if validation_errors:
            details = "; ".join(error["message"] for error in validation_errors)
            raise ValueError(
                f"Reference {key} does not match the required seven-format policy: {details} "
                "If only electronic or preprint sources are available, explain this to the user and ask whether to skip, replace with a formally published source, or explicitly allow [EB/OL]."
            )
        source = verification_source_value(item)
        if not source:
            raise ValueError(
                f"Reference {key} is missing a queryable verification source. "
                "Manifest entries must record a DOI, publisher page, conference/journal website, database record, authoritative index, ISBN/ISSN, patent record, or equivalent formal source."
            )
        if source_is_electronic_only(source):
            raise ValueError(
                f"Reference {key} only records an electronic/preprint/project source: {source}. "
                "Use a formally published and queryable source record, or explain to the user that the item cannot be written under the current rules."
            )


def write_base_docx(
    source_docx: str | Path,
    output_docx: str | Path,
    references: list[dict[str, Any]],
    replace_placeholders: bool,
    use_bookmarks: bool,
) -> None:
    source = Path(source_docx)
    output = Path(output_docx)
    with zipfile.ZipFile(source) as zin:
        files = {name: zin.read(name) for name in zin.namelist()}
    if "word/document.xml" not in files:
        raise ValueError(f"{source} is missing word/document.xml")

    document_root = ET.fromstring(files["word/document.xml"])
    numbering_root = ET.fromstring(files["word/numbering.xml"]) if "word/numbering.xml" in files else None
    numbering_root, num_id = ensure_reference_numbering(numbering_root)

    before_citation_counts = document_citation_structure_counts(document_root)
    if replace_placeholders:
        replace_citation_placeholders(document_root, references)
        after_citation_counts = document_citation_structure_counts(document_root)
        assert_citation_structure_not_reduced(before_citation_counts, after_citation_counts)
    append_reference_section(document_root, references, num_id, use_bookmarks=use_bookmarks)
    document_root.attrib.pop(MC_IGNORABLE, None)

    document_xml = ET.tostring(document_root, encoding="utf-8", xml_declaration=True)
    unresolved = PLACEHOLDER_PATTERN.findall(document_xml.decode("utf-8", errors="ignore"))
    if replace_placeholders and unresolved:
        raise ValueError(f"Unresolved citation placeholders remain: {', '.join(unresolved)}")

    files["word/document.xml"] = document_xml
    files["word/numbering.xml"] = ET.tostring(numbering_root, encoding="utf-8", xml_declaration=True)
    ensure_content_type(files, "/word/numbering.xml", NUMBERING_CONTENT_TYPE)
    ensure_document_relationship(files, NUMBERING_REL_TYPE, "numbering.xml")
    ensure_update_fields(files)

    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)


def powershell_single_quoted(value: str | Path) -> str:
    return "'" + str(value).replace("'", "''") + "'"


def insert_numbered_item_crossrefs_with_word(docx_path: str | Path, output_docx: str | Path, references: list[dict[str, Any]]) -> None:
    powershell_exe = shutil.which("powershell") or shutil.which("pwsh")
    if powershell_exe is None:
        raise NumberedItemCrossReferenceError(
            "未找到 powershell/pwsh，无法调用 Windows Word COM 生成编号项交叉引用；可使用 --method auto 自动降级或 --method bookmark。"
        )
    refs_json = json.dumps([{"key": str(item["key"]), "gbt7714": str(item["gbt7714"])} for item in references], ensure_ascii=False)
    script = rf"""
$ErrorActionPreference = 'Stop'
$docxPath = {powershell_single_quoted(Path(docx_path).resolve())}
$outputPath = {powershell_single_quoted(Path(output_docx).resolve())}
$refs = ConvertFrom-Json @'
{refs_json}
'@
$wdRefTypeNumberedItem = 0
$wdNumberNoContext = -3
$wdCollapseEnd = 0
$wdFormatDocumentDefault = 16

function Normalize([string]$s) {{
  if ($null -eq $s) {{ return '' }}
  $x = $s -replace '\s+', ' '
  $x = $x -replace '^\s*(\[[0-9]+\]|[0-9]+[\.\)]?)\s*', ''
  return $x.Trim()
}}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $null
try {{
  $doc = $word.Documents.Open($docxPath, $false, $false, $false, '', '', $false, '', '', 0, 0, $false, $true)
  $items = $doc.GetCrossReferenceItems($wdRefTypeNumberedItem)
  if ($items.Count -lt $refs.Count) {{
    throw "编号项数量 $($items.Count) 少于参考文献数量 $($refs.Count)"
  }}

  $keyToItem = @{{}}
  for ($r = 0; $r -lt $refs.Count; $r++) {{
    $entry = Normalize([string]$refs[$r].gbt7714)
    $prefixLen = [Math]::Min(80, $entry.Length)
    $needle = if ($prefixLen -gt 0) {{ $entry.Substring(0, $prefixLen) }} else {{ $entry }}
    $found = $null
    for ($i = 1; $i -le $items.Count; $i++) {{
      $candidate = Normalize([string]$items.Item($i))
      if ($needle.Length -gt 0 -and $candidate.Contains($needle)) {{
        $found = $i
        break
      }}
    }}
    if ($null -eq $found) {{
      $found = $items.Count - $refs.Count + $r + 1
    }}
    $keyToItem[[string]$refs[$r].key] = [int]$found
  }}

  foreach ($ref in $refs) {{
    $key = [string]$ref.key
    $placeholderPattern = [regex]::Escape('[[CITE:') + '([^]]*\b' + [regex]::Escape($key) + '\b[^]]*)' + [regex]::Escape(']]')
  }}

  $search = $doc.Content
  while ($search.Find.Execute('[[CITE:')) {{
    $start = $search.Start
    $tail = $doc.Range($start, [Math]::Min($doc.Content.End, $start + 300)).Text
    $m = [regex]::Match($tail, '^\[\[CITE:([A-Za-z0-9_,.\- ]+)\]\]')
    if (-not $m.Success) {{
      $search.Start = $search.End
      $search.End = $doc.Content.End
      continue
    }}
    $placeholder = $m.Value
    $keys = @($m.Groups[1].Value.Split(',') | ForEach-Object {{ $_.Trim() }} | Where-Object {{ $_ }})
    $target = $doc.Range($start, $start + $placeholder.Length)
    $target.Text = ''
    $insert = $doc.Range($start, $start)
    $insert.Font.Name = 'Times New Roman'
    $insert.Font.Size = 12
    $insert.Font.Superscript = $true
    for ($k = 0; $k -lt $keys.Count; $k++) {{
      if ($k -gt 0) {{
        $insert.InsertAfter(',')
        $insert.SetRange($insert.End, $insert.End)
      }}
      $key = [string]$keys[$k]
      if (-not $keyToItem.ContainsKey($key)) {{
        throw "未找到引用 key 对应的编号项: $key"
      }}
      $fieldStart = $insert.Start
      $insert.InsertCrossReference($wdRefTypeNumberedItem, $wdNumberNoContext, $keyToItem[$key], $true, $false, $false, '')
      $fieldRange = $doc.Range($fieldStart, $insert.End)
      $fieldRange.Font.Name = 'Times New Roman'
      $fieldRange.Font.Size = 12
      $fieldRange.Font.Superscript = $true
      $insert.SetRange($insert.End, $insert.End)
    }}
    $end = $insert.End
    $formatRange = $doc.Range($start, $end)
    $formatRange.Font.Name = 'Times New Roman'
    $formatRange.Font.Size = 12
    $formatRange.Font.Superscript = $true
    $search = $doc.Range($end, $doc.Content.End)
  }}
  $doc.Fields.Update() | Out-Null
  $doc.SaveAs2($outputPath, $wdFormatDocumentDefault)
}} finally {{
  if ($doc -ne $null) {{ $doc.Close($false) | Out-Null }}
  $word.Quit() | Out-Null
}}
"""
    with tempfile.TemporaryDirectory() as tmpdir:
        ps1 = Path(tmpdir) / "insert_numbered_item_crossrefs.ps1"
        ps1.write_text(script, encoding="utf-8-sig")
        result = subprocess.run(
            [powershell_exe, "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(ps1)],
            text=True,
            capture_output=True,
            timeout=120,
        )
    if result.returncode != 0:
        message = (result.stderr or result.stdout or "").strip()
        raise NumberedItemCrossReferenceError(message or "Word 编号项交叉引用插入失败")


def apply_manifest_to_docx(
    source_docx: str | Path,
    output_docx: str | Path,
    references: list[dict[str, Any]],
    method: str = "auto",
) -> str:
    validate_manifest(references)
    method = method.lower()
    if method not in {"auto", "numbered-item", "bookmark"}:
        raise ValueError("method must be auto, numbered-item, or bookmark")

    output = Path(output_docx)
    if method in {"auto", "numbered-item"}:
        with tempfile.TemporaryDirectory() as tmpdir:
            temp_docx = Path(tmpdir) / "numbered_item_base.docx"
            write_base_docx(source_docx, temp_docx, references, replace_placeholders=False, use_bookmarks=False)
            try:
                insert_numbered_item_crossrefs_with_word(temp_docx, output, references)
                return "numbered-item"
            except Exception:
                if method == "numbered-item":
                    raise

    write_base_docx(source_docx, output, references, replace_placeholders=True, use_bookmarks=True)
    return "bookmark"


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Apply verified references to a DOCX with [[CITE:key]] placeholders. "
            "Uses Word numbered-item cross-references first, then bookmark REF fallback if needed."
        )
    )
    parser.add_argument("source_docx", help="input DOCX with [[CITE:key]] placeholders")
    parser.add_argument("references_json", help="verified reference manifest JSON")
    parser.add_argument("output_docx", help="output DOCX path; source is never overwritten")
    parser.add_argument("--method", choices=["auto", "numbered-item", "bookmark"], default="auto")
    args = parser.parse_args()

    references = read_manifest(args.references_json)
    method_used = apply_manifest_to_docx(args.source_docx, args.output_docx, references, method=args.method)
    print(json.dumps({"output": str(Path(args.output_docx)), "method": method_used}, ensure_ascii=False))


if __name__ == "__main__":
    main()
