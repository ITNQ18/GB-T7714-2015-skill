# MCP Search Tools

Use the MCP server when the model has no native network access or when reference search should be exposed as structured tools.

## Server

```text
mcp-server/gb_reference_server.py
```

Default network proxy is `http://127.0.0.1:7897`. Override with `GBT7714_PROXY` or the `proxy` tool argument. Use `direct` only when intentionally bypassing the proxy.

Install dependency in the active virtual environment:

```powershell
# From the skill root:
.\.venv\Scripts\python.exe -m pip install -r mcp-server\requirements.txt

# From the parent workspace/repo root:
.\.venv\Scripts\python.exe -m pip install -r gb-t7714-2015-skill\mcp-server\requirements.txt
```

```bash
# From the skill root:
./.venv/bin/python -m pip install -r mcp-server/requirements.txt

# From the parent workspace/repo root:
./.venv/bin/python -m pip install -r gb-t7714-2015-skill/mcp-server/requirements.txt
```

Run as stdio MCP server:

```powershell
# From the skill root:
.\.venv\Scripts\python.exe -B mcp-server\gb_reference_server.py

# From the parent workspace/repo root:
.\.venv\Scripts\python.exe -B gb-t7714-2015-skill\mcp-server\gb_reference_server.py
```

```bash
# From the skill root:
./.venv/bin/python -B mcp-server/gb_reference_server.py

# From the parent workspace/repo root:
./.venv/bin/python -B gb-t7714-2015-skill/mcp-server/gb_reference_server.py
```

The MCP server is a stdio server, not a normal CLI. Do not expect `--help` output; verify it by importing the module or by connecting through an MCP client.

## Tools

| Tool | Purpose |
|---|---|
| `search_formal_references` | Search CrossRef/OpenAlex/PubMed, optionally arXiv as clue-only, and attach GB/T write-in classification. |
| `classify_gbt7714_candidate` | Classify one candidate JSON object as writable or not-yet-writable. |
| `filter_writable_candidates` | Split a result list into default-writable and rejected/partial candidates. |
| `search_and_verify_citation_points` | Given citation points JSON, auto-build queries, batch-search, verify, and return per-point writable candidates with failure list. |
| `build_manifest_draft_from_points` | Convert citation-point batch results into a manifest draft (`key`, source evidence, empty `gbt7714` to fill before DOCX write). |

## Citation Points Input

`search_and_verify_citation_points` accepts JSON list or object. This JSON is an internal format; users do not need to write it manually. When no reference list exists, the model should read the manuscript text, choose citation-worthy phrases or claims, and generate these citation points itself.

```json
[
  {
    "id": "cp001",
    "text": "Transformer architecture improves sequence modeling quality and efficiency.",
    "title_hint": "Attention Is All You Need",
    "author_hint": "Vaswani",
    "keywords": ["Transformer", "NIPS 2017"],
    "type_hint": "conference",
    "reference_intent": "original"
  }
]
```

Only fill `title_hint` when the model is confident that the claim points to a canonical work. Otherwise leave it empty and rely on `keywords` plus `reference_intent`. Valid `reference_intent` values include `original`, `supporting`, `survey`, `tool`, and `standard`.

`build_manifest_draft_from_points` output is a draft only. Fill each `gbt7714` field, then pass the completed manifest to `scripts/应用参考文献.py`.

## Electronic Policy

Do not hide electronic/preprint/project-page results during search. They are useful evidence for the model to inspect, compare, and use as clues. The restriction applies only to final write-in.

Default behavior:

- arXiv, OpenReview, GitHub, CTAN, official docs, webpages, and project pages remain visible in search results.
- They are marked as clue/electronic candidates and do not enter the final manifest by default.
- The model may use them to discover a formal journal, conference, thesis, book, or patent record.

When the user or school template explicitly allows `[EB/OL]`, call MCP tools with `allow_electronic=true` or the script with `--allow-electronic`. Electronic candidates can then pass write-policy checks as `[EB/OL]`, but the model must still confirm authority, relevance, publication/update date, access date, and URL before writing.

## Relevance Gate

For citation-point workflows, a DOI or formal database record proves that a candidate exists, not that it supports the sentence. Candidates now include `point_relevance`:

- `title_core_coverage`: how much of the citation point's core terms appear in the candidate title.
- `matched_core_terms`: core terms found in the candidate title.
- `intent`: inferred or explicit reference intent.
- `missing_or_risky`: hard reasons the candidate must not be written yet.
- `risk_warnings`: reasons to inspect the abstract, publisher page, or PDF before writing.

If the citation point asks for an original/classic work and no strong original signal is found, the candidate must stay in `rejected_or_partial_candidates` even when it has a DOI.

## Policy

- `[J]`, `[C]`, `[M]`, `[D]`, and `[P]` can be accepted only with formal queryable records.
- Formal conference and proceedings papers prefer `[C]. //`.
- arXiv, OpenReview, GitHub, project pages, and official docs remain searchable/visible. They are write-blocked by default, and become `[EB/OL]` candidates only when the user explicitly authorizes electronic references.
- MCP search does not write DOCX files. DOCX writing and auditing remain in `scripts/应用参考文献.py` and `scripts/审计参考文献.py`.
