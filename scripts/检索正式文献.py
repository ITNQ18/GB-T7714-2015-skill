from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import sys
import time
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any

sys.dont_write_bytecode = True

try:
    import tomllib
except ModuleNotFoundError:  # Python 3.10 compatibility.
    import tomli as tomllib  # type: ignore[no-redef]


USER_AGENT = "gb-t7714-2015-skill/1.0 (bibliographic verification)"
FORMAL_TYPES = {"J", "M", "C", "D", "P"}
ELECTRONIC_TYPES = {"EB/OL"}
STOPWORDS = {
    "a",
    "an",
    "and",
    "are",
    "as",
    "at",
    "be",
    "by",
    "for",
    "from",
    "in",
    "into",
    "is",
    "it",
    "of",
    "on",
    "or",
    "that",
    "the",
    "their",
    "this",
    "to",
    "with",
    "which",
    "when",
    "while",
    "can",
    "using",
    "use",
    "used",
    "based",
    "models",
    "model",
    "method",
    "methods",
    "system",
    "systems",
    "problem",
    "problems",
    "research",
    "study",
    "approach",
    "approaches",
    "efficient",
    "efficiency",
    "enable",
    "enables",
    "support",
    "supports",
}
ORIGINAL_INTENT_TERMS = {
    "introduce",
    "introduces",
    "introduced",
    "propose",
    "proposes",
    "proposed",
    "replace",
    "replaces",
    "replaced",
    "first",
    "original",
    "seminal",
    "classic",
    "foundation",
}
SURVEY_INTENT_TERMS = {"survey", "overview", "review", "summarize", "summarizes", "taxonomy"}
FORMAL_SOURCE_NAMES = {
    "crossref",
    "openalex",
    "pubmed",
    "publisher",
    "doi",
    "journal",
    "conference",
    "proceedings",
    "patent",
    "thesis",
    "cnki",
    "wanfang",
    "pubscholar",
}
CLUE_PATTERN = re.compile(
    r"arxiv|openreview|github|ctan|official docs?|documentation|project page|webpage|website|blog|example\.com",
    re.IGNORECASE,
)


def load_config() -> dict[str, Any]:
    config_path = Path(__file__).resolve().parent.parent / "mcp-server" / "config.toml"
    if not config_path.exists():
        return {}
    with config_path.open("rb") as handle:
        return tomllib.load(handle)


_CONFIG = load_config()
DEFAULT_PROXY = _CONFIG.get("network", {}).get("default_proxy", "http://127.0.0.1:7897")
DEFAULT_SOURCES = _CONFIG.get("search", {}).get("default_sources", ["crossref", "openalex", "pubmed"])
MAX_ROWS = int(_CONFIG.get("search", {}).get("max_rows", 20))
DEFAULT_TIMEOUT = int(_CONFIG.get("network", {}).get("timeout", 20))


def configure_proxy(proxy: str | None) -> str:
    """Configure urllib proxy use and return the effective proxy label."""
    effective = (
        proxy
        or os.environ.get("GBT7714_PROXY")
        or os.environ.get("HTTPS_PROXY")
        or os.environ.get("https_proxy")
        or os.environ.get("HTTP_PROXY")
        or os.environ.get("http_proxy")
        or DEFAULT_PROXY
    )
    if effective.lower() in {"none", "off", "direct", "no"}:
        urllib.request.install_opener(urllib.request.build_opener(urllib.request.ProxyHandler({})))
        return "direct"
    urllib.request.install_opener(
        urllib.request.build_opener(
            urllib.request.ProxyHandler(
                {
                    "http": effective,
                    "https": effective,
                }
            )
        )
    )
    return effective


def fetch_json(url: str, timeout: int) -> dict[str, Any]:
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(req, timeout=timeout) as response:
        return json.loads(response.read().decode("utf-8"))


def fetch_text(url: str, timeout: int) -> str:
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(req, timeout=timeout) as response:
        return response.read().decode("utf-8", errors="replace")


def first(value: Any, default: str = "") -> str:
    if isinstance(value, list) and value:
        return str(value[0])
    if value is None:
        return default
    return str(value)


def crossref_type_to_gbt(work_type: str) -> str:
    mapping = {
        "journal-article": "J",
        "proceedings-article": "C",
        "book": "M",
        "monograph": "M",
        "book-chapter": "C",
        "dissertation": "D",
        "posted-content": "CLUE_ONLY",
    }
    return mapping.get(work_type or "", "UNKNOWN")


def title_similarity(query: str, title: str) -> float:
    def tokens(text: str) -> set[str]:
        cleaned = "".join(ch.lower() if ch.isalnum() else " " for ch in text)
        return {part for part in cleaned.split() if len(part) > 1}

    query_tokens = tokens(query)
    title_tokens = tokens(title)
    if not query_tokens or not title_tokens:
        return 0.0
    return round(len(query_tokens & title_tokens) / len(query_tokens | title_tokens), 3)


def normalize_title(text: str) -> str:
    cleaned = "".join(ch.lower() if ch.isalnum() else " " for ch in text)
    return " ".join(cleaned.split())


def title_exact_match(query: str, title: str) -> bool:
    return bool(query and title and normalize_title(query) == normalize_title(title))


def text_tokens(text: str) -> set[str]:
    normalized = normalize_title(text)
    return {
        token
        for token in normalized.split()
        if len(token) > 2 and token not in STOPWORDS and not token.isdigit()
    }


def infer_reference_intent(point: dict[str, Any]) -> str:
    explicit = str(point.get("intent") or point.get("reference_intent") or "").strip().lower()
    if explicit:
        return explicit
    blob = " ".join(
        [
            str(point.get("text") or ""),
            " ".join(str(item) for item in point.get("keywords", []) if str(item).strip()),
            str(point.get("type_hint") or ""),
        ]
    ).lower()
    tokens = text_tokens(blob)
    if tokens & SURVEY_INTENT_TERMS:
        return "survey"
    if tokens & ORIGINAL_INTENT_TERMS:
        return "original"
    if any(term in blob for term in ("tool", "software", "library", "package", "framework")):
        return "tool"
    if any(term in blob for term in ("standard", "protocol", "specification")):
        return "standard"
    return "supporting"


def core_terms_for_point(point: dict[str, Any]) -> list[str]:
    raw_terms: list[str] = []
    for value in point.get("keywords", []):
        value = str(value).strip()
        if value:
            raw_terms.append(value)
    for key in ("title_hint", "method_hint", "model_hint", "tool_hint", "dataset_hint"):
        value = str(point.get(key) or "").strip()
        if value:
            raw_terms.append(value)
    if not raw_terms:
        raw_terms.append(str(point.get("text") or ""))

    ordered: list[str] = []
    seen: set[str] = set()
    for raw in raw_terms:
        phrase = normalize_title(raw)
        if phrase and phrase not in seen and len(phrase) > 2:
            seen.add(phrase)
            ordered.append(phrase)
        for token in text_tokens(raw):
            if token not in seen:
                seen.add(token)
                ordered.append(token)
    return ordered[:12]


def phrase_matches(phrases: list[str], text: str) -> list[str]:
    normalized = normalize_title(text)
    matches: list[str] = []
    for phrase in phrases:
        if " " in phrase and phrase in normalized:
            matches.append(phrase)
    return matches


def assess_point_relevance(candidate: dict[str, Any], point: dict[str, Any]) -> dict[str, Any]:
    title = str(candidate.get("title") or "")
    candidate_blob = " ".join(
        str(candidate.get(key) or "")
        for key in ("title", "container", "publisher", "work_type")
    )
    core_terms = core_terms_for_point(point)
    core_tokens = [term for term in core_terms if " " not in term]
    title_token_set = text_tokens(title)
    blob_token_set = text_tokens(candidate_blob)
    title_matches = [token for token in core_tokens if token in title_token_set]
    blob_matches = [token for token in core_tokens if token in blob_token_set]
    phrase_hits = phrase_matches(core_terms, title)

    denominator = max(1, min(6, len(core_tokens) or len(core_terms)))
    title_coverage = round((len(title_matches) + len(phrase_hits)) / denominator, 3)
    evidence_coverage = round((len(blob_matches) + len(phrase_hits)) / denominator, 3)
    intent = infer_reference_intent(point)
    title_lower = normalize_title(title)
    classic_signal = any(term in title_lower for term in ("survey", "overview", "review"))
    title_hint = normalize_title(str(point.get("title_hint") or ""))
    title_hint_exact = bool(title_hint and title_hint == title_lower)
    title_word_count = len(title_lower.split())
    original_signal = bool(
        title_hint_exact
        or (title_word_count <= 8 and evidence_coverage >= 0.4)
    )

    missing: list[str] = []
    warnings: list[str] = []
    if title_hint and not title_hint_exact:
        missing.append("候选题名与模型给出的 canonical title_hint 不一致")
    if title_lower.startswith("book review"):
        missing.append("候选是书评/评论，不是被引用主题的一手文献")
    if title_coverage < 0.35:
        missing.append("候选标题未覆盖引用点核心概念")
    elif evidence_coverage < 0.45:
        warnings.append("候选出版信息核心词覆盖不足，需检查摘要或出版页确认支撑关系")
    if intent == "original" and not original_signal:
        missing.append("引用点像是在找原始/经典文献，但候选未显示明显原始文献信号")
    if title_word_count > 12 and not phrase_hits:
        warnings.append("候选标题较长且缺少核心短语命中，可能是应用场景论文而非通用支撑文献")
    if intent == "survey" and not classic_signal:
        warnings.append("引用点像是在找综述/背景文献，但候选标题未显示综述信号")

    score = round(min(1.0, title_coverage * 0.7 + evidence_coverage * 0.3), 3)
    return {
        "intent": intent,
        "score": score,
        "title_core_coverage": title_coverage,
        "evidence_core_coverage": evidence_coverage,
        "core_terms": core_terms,
        "matched_core_terms": sorted(set(title_matches + phrase_hits)),
        "title_hint_exact": title_hint_exact,
        "missing_or_risky": missing,
        "risk_warnings": warnings,
    }


def type_hint_to_gbt(type_hint: str) -> str | None:
    hint = type_hint.strip().lower()
    if hint in {"j", "journal", "article"}:
        return "J"
    if hint in {"c", "conference", "proceedings", "proceeding"}:
        return "C"
    if hint in {"m", "book", "monograph"}:
        return "M"
    if hint in {"d", "thesis", "dissertation"}:
        return "D"
    if hint in {"p", "patent"}:
        return "P"
    return None


def enrich_candidate_from_point(candidate: dict[str, Any], point: dict[str, Any]) -> None:
    relevance = candidate.get("point_relevance")
    if not isinstance(relevance, dict) or not relevance.get("title_hint_exact"):
        return
    hinted_type = type_hint_to_gbt(str(point.get("type_hint") or ""))
    work_type = str(candidate.get("work_type") or "").lower()
    if hinted_type and candidate.get("suggested_gbt_type") in {None, "", "UNKNOWN"} and work_type not in {
        "posted-content",
        "preprint",
    }:
        candidate["suggested_gbt_type"] = hinted_type
    has_formal_identifier = bool(candidate.get("doi") or candidate.get("pmid") or candidate.get("isbn"))
    if candidate.get("suggested_gbt_type") in FORMAL_TYPES and has_formal_identifier and work_type not in {
        "posted-content",
        "preprint",
    }:
        candidate["formal_record"] = True


def classify_candidate(candidate: dict[str, Any], allow_electronic: bool = False) -> dict[str, Any]:
    gbt_type = str(candidate.get("suggested_gbt_type") or "UNKNOWN")
    formal_record = bool(candidate.get("formal_record"))
    similarity = candidate.get("title_similarity")
    missing: list[str] = []
    warnings: list[str] = []
    is_clue_or_electronic = gbt_type in {"CLUE_ONLY", "EB/OL"} or bool(CLUE_PATTERN.search(
        " ".join(str(candidate.get(key, "")) for key in ("source", "work_type", "url", "note", "container", "publisher"))
    ))

    if gbt_type not in FORMAL_TYPES:
        if allow_electronic and is_clue_or_electronic:
            gbt_type = "EB/OL"
            warnings.append("用户允许电子文献格式；该候选仍需模型确认权威性、发布日期/更新日期和访问路径")
        else:
            missing.append("类型不是默认可写入的正式文献类型")
    if not formal_record and not (allow_electronic and gbt_type == "EB/OL"):
        missing.append("缺少正式可查询记录")
    if not (candidate.get("doi") or candidate.get("pmid") or candidate.get("url")):
        missing.append("缺少 DOI/PMID/URL 等核验入口")

    source_blob = " ".join(
        str(candidate.get(key, ""))
        for key in ("source", "work_type", "url", "note", "container", "publisher")
    )
    source_name = str(candidate.get("source", "")).lower()
    has_strong_id = bool(
        candidate.get("doi")
        or candidate.get("pmid")
        or candidate.get("isbn")
        or candidate.get("patent_number")
        or candidate.get("database_record")
    )
    is_known_formal_source = any(name in source_name for name in FORMAL_SOURCE_NAMES)
    if CLUE_PATTERN.search(source_blob) and not has_strong_id:
        if allow_electronic and gbt_type == "EB/OL":
            warnings.append("来源疑似网页/预印本/项目页；仅在用户允许电子文献且模型确认权威性后可写入")
        else:
            missing.append("来源疑似网页/预印本/项目页，仅可作为线索")
    if not has_strong_id and not is_known_formal_source:
        if allow_electronic and gbt_type == "EB/OL" and candidate.get("url"):
            warnings.append("电子候选缺少 DOI/PMID/ISBN 等强标识符；需模型确认来源权威性")
        else:
            missing.append("缺少强标识符或正式数据库来源")
    if allow_electronic and gbt_type == "EB/OL" and not (candidate.get("published_date") or candidate.get("updated_date") or candidate.get("year")):
        warnings.append("电子文献候选缺少发表/更新日期；写入前必须补齐，访问日期不能替代")
    if candidate.get("title_exact_match") is False:
        warnings.append("标题不是精确匹配，需人工确认不是同名或相似记录")
    if similarity is not None and isinstance(similarity, (int, float)) and similarity < 0.25:
        warnings.append("标题相似度较低，需人工确认相关性")
    point_relevance = candidate.get("point_relevance")
    if isinstance(point_relevance, dict):
        missing.extend(str(item) for item in point_relevance.get("missing_or_risky", []) if str(item).strip())
        warnings.extend(str(item) for item in point_relevance.get("risk_warnings", []) if str(item).strip())

    writable = not missing
    preferred_template = "[C]. //" if gbt_type == "C" else f"[{gbt_type}]" if gbt_type in FORMAL_TYPES | ELECTRONIC_TYPES else None
    return {
        "writable_by_default": writable,
        "search_visible": True,
        "write_policy": "electronic_allowed" if allow_electronic and gbt_type == "EB/OL" else "default_formal_only",
        "suggested_gbt_type": gbt_type,
        "preferred_template": preferred_template,
        "missing_or_risky": missing,
        "risk_warnings": warnings,
        "point_relevance": point_relevance if isinstance(point_relevance, dict) else None,
        "acceptance_note": (
            "可进入 manifest，但仍需补齐 GB/T 7714 字段并确认支撑正文论断。"
            if writable
            else "暂不写入；先补核验来源、正式版本或相关性证据。"
        ),
    }


def parse_citation_points(raw: Any) -> list[dict[str, Any]]:
    if isinstance(raw, list):
        points = raw
    elif isinstance(raw, dict):
        points = raw.get("citation_points") or raw.get("points") or raw.get("items") or []
    else:
        points = []
    normalized: list[dict[str, Any]] = []
    for idx, item in enumerate(points, start=1):
        if isinstance(item, str):
            normalized.append({"id": f"cp{idx:03d}", "text": item})
            continue
        if not isinstance(item, dict):
            continue
        text = str(item.get("text") or item.get("claim") or item.get("citation_text") or "").strip()
        if not text:
            continue
        point_id = str(item.get("id") or item.get("key") or item.get("point_id") or f"cp{idx:03d}").strip()
        normalized.append(
            {
                "id": point_id,
                "text": text,
                "title_hint": str(item.get("title_hint") or item.get("title") or "").strip(),
                "author_hint": str(item.get("author_hint") or item.get("author") or "").strip(),
                "keywords": item.get("keywords") if isinstance(item.get("keywords"), list) else [],
                "type_hint": str(item.get("type_hint") or item.get("type") or "").strip(),
            }
        )
    return normalized


def truncate_words(text: str, limit: int) -> str:
    words = text.split()
    return " ".join(words[:limit]).strip()


def build_queries_for_point(point: dict[str, Any], max_queries: int = 4) -> list[str]:
    title_hint = str(point.get("title_hint", "")).strip()
    author_hint = str(point.get("author_hint", "")).strip()
    text = str(point.get("text", "")).strip()
    type_hint = str(point.get("type_hint", "")).strip().lower()
    keywords = [str(x).strip() for x in point.get("keywords", []) if str(x).strip()]
    intent = infer_reference_intent(point)
    core_terms = core_terms_for_point(point)
    main_term = keywords[0] if keywords else (core_terms[0] if core_terms else "")

    queries: list[str] = []
    if title_hint:
        queries.append(title_hint)
    if title_hint and author_hint:
        queries.append(f"{title_hint} {author_hint}")
    if main_term and intent == "original":
        queries.append(f"{main_term} original paper")
        queries.append(f"{main_term} proposed")
    elif main_term and intent == "survey":
        queries.append(f"{main_term} survey review")
    elif main_term:
        queries.append(f"{main_term} formal paper")
    if keywords:
        queries.append(" ".join(keywords[:8]))

    base_text = truncate_words(text, 18)
    if base_text:
        queries.append(base_text)
    if type_hint and base_text:
        if "conference" in type_hint or "proceeding" in type_hint or type_hint == "c":
            queries.append(f"{base_text} conference proceedings")
        elif "journal" in type_hint or type_hint == "j":
            queries.append(f"{base_text} journal article")
        elif "thesis" in type_hint or type_hint == "d":
            queries.append(f"{base_text} thesis dissertation")
        elif "patent" in type_hint or type_hint == "p":
            queries.append(f"{base_text} patent")

    unique: list[str] = []
    seen: set[str] = set()
    for query in queries:
        norm = normalize_title(query)
        if not norm or norm in seen:
            continue
        seen.add(norm)
        unique.append(query)
    return unique[: max(1, max_queries)]


def candidate_identity(candidate: dict[str, Any]) -> str:
    relevance = candidate.get("point_relevance")
    if isinstance(relevance, dict) and relevance.get("title_hint_exact") and candidate.get("title"):
        return "title-hint:" + normalize_title(str(candidate.get("title") or ""))
    for key in ("doi", "pmid", "isbn", "url", "title"):
        value = str(candidate.get(key) or "").strip().lower()
        if value:
            return f"{key}:{value}"
    body = json.dumps(candidate, ensure_ascii=False, sort_keys=True)
    return "hash:" + hashlib.sha1(body.encode("utf-8")).hexdigest()


def rank_key(candidate: dict[str, Any]) -> tuple[int, float, int, float]:
    acceptance = candidate.get("gbt7714_acceptance") or {}
    writable = 1 if acceptance.get("writable_by_default") else 0
    relevance = candidate.get("point_relevance") or acceptance.get("point_relevance") or {}
    relevance_score = float(relevance.get("score") or 0.0) if isinstance(relevance, dict) else 0.0
    exact = 1 if candidate.get("title_exact_match") else 0
    similarity = float(candidate.get("title_similarity") or 0.0)
    return (writable, relevance_score, exact, similarity)


def search_citation_points(
    citation_points: list[dict[str, Any]],
    sources: list[str] | None = None,
    rows: int = 5,
    timeout: int = DEFAULT_TIMEOUT,
    proxy: str | None = None,
    max_queries: int = 4,
    top_k: int = 3,
    allow_electronic: bool = False,
) -> dict[str, Any]:
    selected_sources = sources or DEFAULT_SOURCES
    normalized_points = parse_citation_points(citation_points)
    payload: dict[str, Any] = {
        "mode": "citation_points",
        "sources": selected_sources,
        "proxy": configure_proxy(proxy),
        "summary": {
            "total_points": len(normalized_points),
            "points_with_writable": 0,
            "points_without_writable": 0,
            "total_writable_candidates": 0,
        },
        "write_policy": "electronic_allowed" if allow_electronic else "default_formal_only",
        "points": [],
        "failure_list": [],
    }

    for idx, point in enumerate(normalized_points, start=1):
        queries = build_queries_for_point(point, max_queries=max_queries)
        collected: list[dict[str, Any]] = []
        query_runs: list[dict[str, Any]] = []
        for query in queries:
            result = search_references(
                query=query,
                sources=selected_sources,
                rows=rows,
                timeout=timeout,
                proxy=proxy,
            )
            query_runs.append(
                {
                    "query": query,
                    "result_count": len(result.get("results", [])),
                    "errors": result.get("errors", []),
                }
            )
            for item in result.get("results", []):
                candidate = dict(item)
                candidate["matched_query"] = query
                candidate["point_relevance"] = assess_point_relevance(candidate, point)
                enrich_candidate_from_point(candidate, point)
                candidate["gbt7714_acceptance"] = classify_candidate(candidate, allow_electronic=allow_electronic)
                collected.append(candidate)

        deduped: list[dict[str, Any]] = []
        seen_ids: set[str] = set()
        for item in sorted(collected, key=rank_key, reverse=True):
            ident = candidate_identity(item)
            if ident in seen_ids:
                continue
            seen_ids.add(ident)
            deduped.append(item)

        writable = [item for item in deduped if item.get("gbt7714_acceptance", {}).get("writable_by_default")]
        selected = writable[: max(1, top_k)]
        rejected = [item for item in deduped if item not in selected][: max(1, top_k)]
        point_result = {
            "point_id": point.get("id") or f"cp{idx:03d}",
            "claim_text": point.get("text", ""),
            "queries": query_runs,
            "selected_writable_candidates": selected,
            "rejected_or_partial_candidates": rejected,
        }
        payload["points"].append(point_result)

        if selected:
            payload["summary"]["points_with_writable"] += 1
            payload["summary"]["total_writable_candidates"] += len(selected)
        else:
            payload["summary"]["points_without_writable"] += 1
            reason = "未检索到候选"
            if deduped:
                missing = deduped[0].get("gbt7714_acceptance", {}).get("missing_or_risky") or []
                reason = "；".join(missing[:3]) if missing else "候选存在但未通过写入核验"
            payload["failure_list"].append(
                {
                    "point_id": point_result["point_id"],
                    "claim_text": point_result["claim_text"],
                    "reason": reason,
                }
            )
    return payload


def openalex_type_to_gbt(work_type: str, primary_location: dict[str, Any] | None) -> str:
    source_type = ""
    if primary_location:
        source = primary_location.get("source") or {}
        source_type = source.get("type") or ""
    if work_type in {"article", "review"} or source_type == "journal":
        return "J"
    if work_type in {"proceedings-article", "book-chapter"} or source_type == "conference":
        return "C"
    if work_type in {"book", "monograph"} or source_type == "book series":
        return "M"
    if work_type == "dissertation":
        return "D"
    return "UNKNOWN"


def compact_authors_crossref(item: dict[str, Any]) -> list[str]:
    authors = []
    for author in item.get("author", [])[:5]:
        given = author.get("given", "")
        family = author.get("family", "")
        name = " ".join(part for part in [given, family] if part).strip()
        if name:
            authors.append(name)
    return authors


def compact_authors_openalex(item: dict[str, Any]) -> list[str]:
    authors = []
    for entry in item.get("authorships", [])[:5]:
        author = entry.get("author") or {}
        name = author.get("display_name")
        if name:
            authors.append(name)
    return authors


def year_from_crossref(item: dict[str, Any]) -> int | None:
    for key in ("published-print", "published-online", "issued", "created"):
        parts = (item.get(key) or {}).get("date-parts") or []
        if parts and parts[0] and parts[0][0]:
            try:
                return int(parts[0][0])
            except (TypeError, ValueError):
                return None
    return None


def crossref_search(query: str, rows: int, timeout: int) -> list[dict[str, Any]]:
    params = urllib.parse.urlencode({"query.bibliographic": query, "rows": rows})
    data = fetch_json(f"https://api.crossref.org/works?{params}", timeout)
    results = []
    for item in (data.get("message") or {}).get("items", []):
        work_type = item.get("type", "")
        gbt_type = crossref_type_to_gbt(work_type)
        doi = item.get("DOI", "")
        title = first(item.get("title"))
        results.append(
            {
                "source": "CrossRef",
                "title": title,
                "title_similarity": title_similarity(query, title),
                "title_exact_match": title_exact_match(query, title),
                "authors": compact_authors_crossref(item),
                "year": year_from_crossref(item),
                "container": first(item.get("container-title")),
                "publisher": item.get("publisher", ""),
                "volume": item.get("volume", ""),
                "issue": item.get("issue", ""),
                "pages": item.get("page", ""),
                "doi": doi,
                "url": f"https://doi.org/{doi}" if doi else item.get("URL", ""),
                "work_type": work_type,
                "suggested_gbt_type": gbt_type,
                "formal_record": bool(doi and gbt_type not in {"CLUE_ONLY", "UNKNOWN"}),
                "note": "Use DOI/publisher metadata; proceedings and book chapters usually map to [C].",
            }
        )
    return results


def openalex_search(query: str, rows: int, timeout: int) -> list[dict[str, Any]]:
    params = urllib.parse.urlencode({"search": query, "per-page": rows})
    data = fetch_json(f"https://api.openalex.org/works?{params}", timeout)
    results = []
    for item in data.get("results", []):
        primary_location = item.get("primary_location") or {}
        source = primary_location.get("source") or {}
        doi = (item.get("doi") or "").replace("https://doi.org/", "")
        gbt_type = openalex_type_to_gbt(item.get("type", ""), primary_location)
        title = item.get("display_name", "")
        results.append(
            {
                "source": "OpenAlex",
                "title": title,
                "title_similarity": title_similarity(query, title),
                "title_exact_match": title_exact_match(query, title),
                "authors": compact_authors_openalex(item),
                "year": item.get("publication_year"),
                "container": source.get("display_name", ""),
                "publisher": source.get("host_organization_name", ""),
                "doi": doi,
                "url": item.get("doi") or item.get("id", ""),
                "work_type": item.get("type", ""),
                "suggested_gbt_type": gbt_type,
                "formal_record": bool(doi and gbt_type not in {"UNKNOWN"}),
                "note": "Discovery/cross-check source; verify against DOI, publisher, venue, thesis, or patent record before writing.",
            }
        )
    return results


def pubmed_search(query: str, rows: int, timeout: int) -> list[dict[str, Any]]:
    params = urllib.parse.urlencode(
        {"db": "pubmed", "term": query, "retmax": rows, "retmode": "json", "sort": "relevance"}
    )
    data = fetch_json(f"https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi?{params}", timeout)
    ids = ((data.get("esearchresult") or {}).get("idlist") or [])[:rows]
    if not ids:
        return []
    time.sleep(0.35)
    fetch_params = urllib.parse.urlencode({"db": "pubmed", "id": ",".join(ids), "retmode": "xml", "rettype": "abstract"})
    xml_text = fetch_text(f"https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi?{fetch_params}", timeout)
    root = ET.fromstring(xml_text)
    results = []
    for article in root.findall(".//PubmedArticle"):
        citation = article.find("MedlineCitation")
        art = citation.find("Article") if citation is not None else None
        if art is None:
            continue
        pmid_el = citation.find("PMID") if citation is not None else None
        title_el = art.find("ArticleTitle")
        title = "".join(title_el.itertext()).strip() if title_el is not None else ""
        journal_el = art.find("Journal/Title")
        year_el = art.find("Journal/JournalIssue/PubDate/Year")
        doi = ""
        for node in art.findall("ELocationID"):
            if node.get("EIdType") == "doi" and node.text:
                doi = node.text.strip()
                break
        results.append(
            {
                "source": "PubMed",
                "title": title,
                "title_similarity": title_similarity(query, title),
                "title_exact_match": title_exact_match(query, title),
                "authors": [],
                "year": int(year_el.text) if year_el is not None and year_el.text and year_el.text.isdigit() else None,
                "container": journal_el.text.strip() if journal_el is not None and journal_el.text else "",
                "doi": doi,
                "pmid": pmid_el.text.strip() if pmid_el is not None and pmid_el.text else "",
                "url": f"https://pubmed.ncbi.nlm.nih.gov/{pmid_el.text.strip()}/" if pmid_el is not None and pmid_el.text else "",
                "work_type": "journal-article",
                "suggested_gbt_type": "J",
                "formal_record": True,
                "note": "PubMed record; use DOI or journal metadata for [J].",
            }
        )
    return results


def arxiv_search(query: str, rows: int, timeout: int) -> list[dict[str, Any]]:
    params = urllib.parse.urlencode({"search_query": query, "start": 0, "max_results": rows})
    xml_text = fetch_text(f"https://export.arxiv.org/api/query?{params}", timeout)
    ns = {"atom": "http://www.w3.org/2005/Atom"}
    root = ET.fromstring(xml_text)
    results = []
    for entry in root.findall("atom:entry", ns):
        title = entry.find("atom:title", ns)
        title_text = " ".join((title.text or "").split()) if title is not None else ""
        published = entry.find("atom:published", ns)
        identifier = entry.find("atom:id", ns)
        results.append(
            {
                "source": "arXiv",
                "title": title_text,
                "title_similarity": title_similarity(query, title_text),
                "title_exact_match": title_exact_match(query, title_text),
                "authors": [a.find("atom:name", ns).text for a in entry.findall("atom:author", ns) if a.find("atom:name", ns) is not None],
                "year": int((published.text or "")[:4]) if published is not None and (published.text or "")[:4].isdigit() else None,
                "published_date": (published.text or "")[:10] if published is not None and published.text else "",
                "url": identifier.text if identifier is not None else "",
                "work_type": "preprint",
                "suggested_gbt_type": "CLUE_ONLY",
                "formal_record": False,
                "note": "Clue only under this skill. Search for DOI, journal, or proceedings version before writing.",
            }
        )
    return results


def run_source(name: str, query: str, rows: int, timeout: int) -> tuple[list[dict[str, Any]], str | None]:
    try:
        if name == "crossref":
            return crossref_search(query, rows, timeout), None
        if name == "openalex":
            return openalex_search(query, rows, timeout), None
        if name == "pubmed":
            return pubmed_search(query, rows, timeout), None
        if name == "arxiv":
            return arxiv_search(query, rows, timeout), None
    except Exception as exc:
        return [], str(exc)
    return [], f"unknown source: {name}"


def search_references(
    query: str,
    sources: list[str] | None = None,
    rows: int = 5,
    timeout: int = 20,
    proxy: str | None = None,
) -> dict[str, Any]:
    selected_sources = sources or DEFAULT_SOURCES
    selected_rows = max(1, min(rows, MAX_ROWS))
    effective_proxy = configure_proxy(proxy)
    payload: dict[str, Any] = {
        "query": query,
        "sources": selected_sources,
        "proxy": effective_proxy,
        "results": [],
        "errors": [],
    }

    for source in selected_sources:
        results, error = run_source(source, query, selected_rows, timeout)
        payload["results"].extend(results)
        if error:
            payload["errors"].append({"source": source, "error": error})
        time.sleep(0.2)

    return payload


def main() -> int:
    parser = argparse.ArgumentParser(description="Search candidate formal references for GB/T 7714 workflows.")
    parser.add_argument("query", nargs="?", help="title, method name, author/title, or topic query")
    parser.add_argument("--source", action="append", choices=["crossref", "openalex", "pubmed", "arxiv"], help="repeatable source; default: crossref, openalex, pubmed")
    parser.add_argument("--rows", type=int, default=5, help="results per source")
    parser.add_argument("--timeout", type=int, default=DEFAULT_TIMEOUT)
    parser.add_argument("--proxy", help="proxy URL; defaults to env proxy or http://127.0.0.1:7897; use 'direct' to disable")
    parser.add_argument("--citation-points-json", help="path to JSON file: list or object with citation_points/points/items")
    parser.add_argument("--top-k", type=int, default=3, help="max writable candidates to keep for each citation point")
    parser.add_argument("--max-queries", type=int, default=4, help="max generated queries for each citation point")
    parser.add_argument("--output", help="optional JSON output path")
    parser.add_argument("--allow-electronic", action="store_true", help="allow [EB/OL] candidates to pass write-policy checks when the user/school permits electronic references")
    args = parser.parse_args()
    if args.citation_points_json:
        raw = json.loads(Path(args.citation_points_json).read_text(encoding="utf-8-sig"))
        payload = search_citation_points(
            citation_points=parse_citation_points(raw),
            sources=args.source,
            rows=args.rows,
            timeout=args.timeout,
            proxy=args.proxy,
            max_queries=max(1, int(args.max_queries)),
            top_k=max(1, int(args.top_k)),
            allow_electronic=bool(args.allow_electronic),
        )
        rendered = json.dumps(payload, ensure_ascii=False, indent=2)
        if args.output:
            Path(args.output).write_text(rendered + "\n", encoding="utf-8")
        else:
            print(rendered)
        return 0 if payload["summary"]["points_with_writable"] > 0 else 1

    if not args.query:
        parser.error("query is required unless --citation-points-json is provided")

    payload = search_references(
        query=args.query,
        sources=args.source,
        rows=args.rows,
        timeout=args.timeout,
        proxy=args.proxy,
    )
    rendered = json.dumps(payload, ensure_ascii=False, indent=2)
    if args.output:
        Path(args.output).write_text(rendered + "\n", encoding="utf-8")
    else:
        print(rendered)
    return 0 if payload["results"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
