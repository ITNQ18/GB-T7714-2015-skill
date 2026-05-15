# Literature Search and Verification

Use this file whenever collecting, replacing, supplementing, or verifying references. The goal is not "find something plausible"; the goal is a reference that can be independently queried and written as one of the 7 school templates.

## Search Stack

Prefer structured sources over general web search.

The model does not need native network access. Searching is delegated to `mcp-server/gb_reference_server.py`, `scripts/检索正式文献.py`, or another tool process that has network access. By default, the script and MCP tools use `http://127.0.0.1:7897` as both HTTP and HTTPS proxy unless `GBT7714_PROXY`, `HTTP_PROXY`, `HTTPS_PROXY`, or a tool argument overrides it. Use `direct` only when direct network access is intentionally available.

| Priority | Source | Use for | Notes |
|---|---|---|---|
| 1 | DOI / CrossRef | journal articles, proceedings papers, books, book chapters | Strongest general verification path. DOI metadata is preferred when complete. |
| 1 | Publisher or venue pages | journals, formal conferences, books, proceedings | IEEE, ACM, Springer, Elsevier, Wiley, Nature, Science, ACL Anthology, CVF, PMLR and comparable official venues are acceptable. |
| 1 | PubMed | biomedical journal articles | Use PMID/DOI metadata when relevant. |
| 1 | Patent office records | patents | Use CNIPA, USPTO, EPO, WIPO, Google Patents only as a patent-record index, not as a generic webpage source. |
| 2 | OpenAlex / Semantic Scholar | discovery and cross-checking | Use as candidate discovery; confirm against DOI, publisher, venue, thesis, patent, or library record before writing. |
| 2 | Library / institutional repositories | books, theses, institutional records | Accept when they expose stable bibliographic metadata. |
| 3 | arXiv, OpenReview, GitHub, official docs, CTAN, project pages | clues and model review | Keep visible during search. Do not write as final references unless the user explicitly allows the electronic template. Search for the formal published version first. |
| 3 | General web search | last resort discovery | Must be followed by structured verification before writing. |

## Workflow

1. Identify the citation claim and expected type: method, model, dataset, standard, software, thesis, patent, book, or background claim. When no reference list exists, the model should choose citation-worthy phrases/claims from the text and create internal citation points; the user does not need to provide JSON.
2. Build 2-4 queries: exact canonical title only when confidently known from the claim; method/tool/model name; method + "original paper" or "survey review" when intent requires it; method + "conference" or "journal"; Chinese title + author or institution when relevant.
3. Search structured sources first. Prefer MCP `search_and_verify_citation_points` for batch citation-point workflows, or `search_formal_references` for single-query checks; otherwise use `scripts/检索正式文献.py`, then manually open the strongest source records if needed.
4. Classify each candidate into one of the 7 templates. Prefer `[J]` or `[C]`/`In` for formal papers; for formal conference/proceedings papers, prefer template 4 `[C]. //` unless the target document explicitly uses template 3 `In`.
5. Verify required fields before writing. A candidate with missing required fields is "partial" and must not be written until completed from a reliable source. Electronic candidates may be searched and shown to the model at all times; by default they are not written. If the user allows `[EB/OL]`, they may enter the draft only after the model confirms authority, relevance, publication/update date, access date, and URL.
6. Record evidence in `参考文献/参考文献概览.md`: source URL/DOI, access path, type, and skip reason if rejected.

Structured databases can return real but irrelevant records. A DOI proves the record exists; it does not prove the record supports the sentence. When using `scripts/检索正式文献.py`, treat low `title_similarity`, low `point_relevance.title_core_coverage`, or missing original/classic signal as discovery noise unless the abstract, venue, or exact title check confirms relevance.

## Type Routing

| Template | Acceptable verification | Reject / only clue |
|---|---|---|
| `[J]` journal | DOI, publisher page, journal page, PubMed, CrossRef, CNKI/万方/维普/PubScholar public record | blog, arXiv-only page, project page |
| `[M]` book | ISBN, publisher page, library catalogue, Google Books/WorldCat as catalogue evidence | shopping pages without bibliographic metadata |
| `In` conference paper | proceedings publisher page, DOI, official proceedings page with editors/place/date/pages | event homepage without paper metadata |
| `[C]. //` proceedings extract | DOI, proceedings publisher page, ACL/CVF/PMLR/IEEE/ACM/Springer formal record, CNKI conference record | workshop webpage without proceedings metadata |
| `[D]` thesis | university repository, national thesis database, CNKI thesis record, ProQuest/institution catalogue | personal homepage PDF only |
| `[EB/OL]` electronic | only when explicitly allowed by user or school template | default is reject and search for a formal version |
| `[P]` patent | patent office record, WIPO/USPTO/EPO/CNIPA, Google Patents as index of official patent metadata | company news or product page |

## Acceptance Checklist

Before adding a candidate to the final manifest, confirm:

- It has a stable verification record: DOI, PMID, ISBN, thesis repository record, patent record, publisher/venue page, or authoritative Chinese database record.
- Its title, authors, venue, and topic match the target claim; do not accept a candidate merely because it has a DOI or high search rank.
- For citation-point workflows, `point_relevance.missing_or_risky` must be empty before a candidate enters the final manifest.
- The final reference text can fill all fields required by its template: authors, title, source, year, volume/issue or venue/proceedings, pages, publisher/place, degree institution, or patent number/date as applicable.
- It is relevant to the exact sentence or phrase being cited.
- If there is both a preprint and a formal version, the final metadata comes from the formal version.
- If only an electronic/preprint/project source exists, keep it visible as evidence for model review. Do not write it by default. Report the reason and offer skip, replace, or user-authorized `[EB/OL]`.

## Partial Success Reporting

When only some candidates pass:

```text
已核验可写入：N 条
暂不写入：M 条
暂不写入原因：
- <claim or title>: 缺少页码 / 只有 arXiv / 未找到正式出版记录 / 与正文论断不匹配
建议：替换为 <formal candidate> / 等待用户授权 [EB/OL] / 跳过该引用点
```

Do not silently drop failed candidates. Do not fill missing metadata from memory or guesswork.

## Proxy Examples

```powershell
$env:GBT7714_PROXY="http://127.0.0.1:7897"
.\.venv\Scripts\python.exe -B scripts\检索正式文献.py "Attention Is All You Need" --source crossref
.\.venv\Scripts\python.exe -B scripts\gbt.py search "Attention Is All You Need" --source crossref
```

```cmd
set GBT7714_PROXY=http://127.0.0.1:7897
.\.venv\Scripts\python.exe -B scripts\检索正式文献.py "Attention Is All You Need" --source crossref
.\.venv\Scripts\python.exe -B scripts\gbt.py search "Attention Is All You Need" --source crossref
```

```bash
export GBT7714_PROXY=http://127.0.0.1:7897
./.venv/bin/python -B scripts/检索正式文献.py "Attention Is All You Need" --source crossref
./.venv/bin/python -B scripts/gbt.py search "Attention Is All You Need" --source crossref
```

Use `scripts/gbt.py` when a terminal cannot reliably input or display Chinese filenames.
