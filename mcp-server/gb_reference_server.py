from __future__ import annotations

import importlib.util
import json
import os
import re
from pathlib import Path
from typing import Any

try:
    import tomllib
except ModuleNotFoundError:  # Python 3.10 compatibility.
    import tomli as tomllib  # type: ignore[no-redef]

try:
    from mcp.server.fastmcp import FastMCP
except Exception:  # pragma: no cover - allow local import without mcp runtime
    try:
        from mcp.server import FastMCP
    except Exception:
        FastMCP = None  # type: ignore[assignment]


SERVER_DIR = Path(__file__).resolve().parent
SKILL_DIR = SERVER_DIR.parent
CONFIG_PATH = SERVER_DIR / "config.toml"
FORMAL_TYPES = {"J", "M", "C", "D", "P"}
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


def _resolve_search_script() -> Path:
    direct = SKILL_DIR / "scripts" / "检索正式文献.py"
    if direct.exists():
        return direct
    scripts_dir = SKILL_DIR / "scripts"
    for candidate in scripts_dir.glob("*.py"):
        if "检索正式文献" in candidate.name:
            return candidate
    raise RuntimeError(f"Cannot find search script under {scripts_dir}")


SEARCH_SCRIPT = _resolve_search_script()


def _load_config() -> dict[str, Any]:
    if not CONFIG_PATH.exists():
        return {}
    with CONFIG_PATH.open("rb") as handle:
        return tomllib.load(handle)


_CONFIG = _load_config()
DEFAULT_PROXY = _CONFIG.get("network", {}).get("default_proxy", "http://127.0.0.1:7897")
DEFAULT_SOURCES = _CONFIG.get("search", {}).get("default_sources", ["crossref", "openalex", "pubmed"])
MAX_ROWS = int(_CONFIG.get("search", {}).get("max_rows", 20))
DEFAULT_TIMEOUT = int(_CONFIG.get("network", {}).get("timeout", 20))


def _load_search_module():
    spec = importlib.util.spec_from_file_location("gbt7714_search", SEARCH_SCRIPT)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Cannot load search script: {SEARCH_SCRIPT}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


_search_module = _load_search_module()


class _NoopMCP:
    def tool(self):
        def decorator(func):
            return func

        return decorator

    def run(self, *args, **kwargs):
        raise RuntimeError("缺少 mcp 运行依赖。请先安装 mcp-server/requirements.txt。")


mcp = FastMCP("gb-t7714-reference-search") if FastMCP is not None else _NoopMCP()


def _json(data: Any) -> str:
    return json.dumps(data, ensure_ascii=False, indent=2)


def _parse_sources(sources: list[str] | str | None) -> list[str] | None:
    if sources is None:
        return DEFAULT_SOURCES
    if isinstance(sources, str):
        values = [part.strip() for part in sources.split(",") if part.strip()]
    else:
        values = [str(part).strip() for part in sources if str(part).strip()]
    valid = {"crossref", "openalex", "pubmed", "arxiv"}
    invalid = [value for value in values if value not in valid]
    if invalid:
        raise ValueError(f"无效来源: {invalid}; 可选: {sorted(valid)}")
    return values or None


def _effective_proxy(proxy: str | None) -> str:
    return (
        proxy
        or os.environ.get("GBT7714_PROXY")
        or os.environ.get("HTTPS_PROXY")
        or os.environ.get("https_proxy")
        or os.environ.get("HTTP_PROXY")
        or os.environ.get("http_proxy")
        or DEFAULT_PROXY
    )


def _local_classify_candidate(candidate: dict[str, Any]) -> dict[str, Any]:
    gbt_type = str(candidate.get("suggested_gbt_type") or "UNKNOWN")
    formal_record = bool(candidate.get("formal_record"))
    similarity = candidate.get("title_similarity")
    missing: list[str] = []
    warnings: list[str] = []

    if gbt_type not in FORMAL_TYPES:
        missing.append("类型不是默认可写入的正式文献类型")
    if not formal_record:
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
        missing.append("来源疑似网页/预印本/项目页，仅可作为线索")
    if not has_strong_id and not is_known_formal_source:
        missing.append("缺少强标识符或正式数据库来源")
    if candidate.get("title_exact_match") is False:
        warnings.append("标题不是精确匹配，需人工确认不是同名或相似记录")
    if similarity is not None and isinstance(similarity, (int, float)) and similarity < 0.25:
        warnings.append("标题相似度较低，需人工确认相关性")

    writable = not missing
    preferred_template = "[C]. //" if gbt_type == "C" else f"[{gbt_type}]" if gbt_type in FORMAL_TYPES else None
    return {
        "writable_by_default": writable,
        "suggested_gbt_type": gbt_type,
        "preferred_template": preferred_template,
        "missing_or_risky": missing,
        "risk_warnings": warnings,
        "acceptance_note": (
            "可进入 manifest，但仍需补齐 GB/T 7714 字段并确认支撑正文论断。"
            if writable
            else "暂不写入；先补核验来源、正式版本或相关性证据。"
        ),
    }


def classify_candidate(candidate: dict[str, Any], allow_electronic: bool = False) -> dict[str, Any]:
    classifier = getattr(_search_module, "classify_candidate", None)
    if callable(classifier):
        return classifier(candidate, allow_electronic=allow_electronic)
    return _local_classify_candidate(candidate)


@mcp.tool()
def search_formal_references(
    query: str,
    sources: list[str] | str | None = None,
    rows: int = 5,
    timeout: int = DEFAULT_TIMEOUT,
    proxy: str | None = None,
    allow_electronic: bool = False,
) -> str:
    """检索正式文献候选并返回写入核验结果。"""
    if not query or not query.strip():
        return _json({"error": "查询词为空。请提供标题、方法名或主题词。"})
    effective_proxy = _effective_proxy(proxy)
    try:
        payload = _search_module.search_references(
            query=query.strip(),
            sources=_parse_sources(sources),
            rows=max(1, min(int(rows), MAX_ROWS)),
            timeout=max(5, int(timeout)),
            proxy=effective_proxy,
        )
    except Exception as exc:
        return _json(
            {
                "error": "检索失败，请检查代理、网络或来源参数。",
                "detail": str(exc),
                "proxy": effective_proxy,
            }
        )

    for item in payload.get("results", []):
        item["gbt7714_acceptance"] = classify_candidate(item, allow_electronic=allow_electronic)
    return _json(payload)


@mcp.tool()
def classify_gbt7714_candidate(candidate_json: str, allow_electronic: bool = False) -> str:
    """对单条候选进行 GB/T 7714 写入核验分类。"""
    try:
        candidate = json.loads(candidate_json)
    except json.JSONDecodeError as exc:
        return _json({"error": f"candidate_json 不是合法 JSON: {exc}"})
    if not isinstance(candidate, dict):
        return _json({"error": "candidate_json 必须是对象。"})
    return _json(classify_candidate(candidate, allow_electronic=allow_electronic))


@mcp.tool()
def filter_writable_candidates(results_json: str, allow_electronic: bool = False) -> str:
    """按默认可写入/暂不写入切分候选列表。"""
    try:
        data = json.loads(results_json)
    except json.JSONDecodeError as exc:
        return _json({"error": f"results_json 不是合法 JSON: {exc}"})
    results = data.get("results") if isinstance(data, dict) else data
    if not isinstance(results, list):
        return _json({"error": "输入应为结果列表，或包含 results 的对象。"})

    writable: list[dict[str, Any]] = []
    rejected: list[dict[str, Any]] = []
    for item in results:
        if not isinstance(item, dict):
            continue
        classification = classify_candidate(item, allow_electronic=allow_electronic)
        enriched = {**item, "gbt7714_acceptance": classification}
        if classification["writable_by_default"]:
            writable.append(enriched)
        else:
            rejected.append(enriched)
    return _json({"writable": writable, "not_yet_writable": rejected})


@mcp.tool()
def search_and_verify_citation_points(
    citation_points_json: str,
    sources: list[str] | str | None = None,
    rows: int = 5,
    timeout: int = DEFAULT_TIMEOUT,
    proxy: str | None = None,
    max_queries: int = 4,
    top_k: int = 3,
    allow_electronic: bool = False,
) -> str:
    """针对一组引用点批量检索并核验，输出可写入与失败清单。"""
    try:
        raw = json.loads(citation_points_json)
    except json.JSONDecodeError as exc:
        return _json({"error": f"citation_points_json 不是合法 JSON: {exc}"})

    normalizer = getattr(_search_module, "parse_citation_points", None)
    if not callable(normalizer):
        return _json({"error": "检索脚本缺少 parse_citation_points()，请更新 scripts/检索正式文献.py。"})
    points = normalizer(raw)
    if not points:
        return _json({"error": "引用点列表为空。请提供 list，或对象中的 citation_points/points/items。"})

    search_batch = getattr(_search_module, "search_citation_points", None)
    if not callable(search_batch):
        return _json({"error": "检索脚本缺少 search_citation_points()，请更新 scripts/检索正式文献.py。"})

    effective_proxy = _effective_proxy(proxy)
    try:
        payload = search_batch(
            citation_points=points,
            sources=_parse_sources(sources),
            rows=max(1, min(int(rows), MAX_ROWS)),
            timeout=max(5, int(timeout)),
            proxy=effective_proxy,
            max_queries=max(1, int(max_queries)),
            top_k=max(1, int(top_k)),
            allow_electronic=allow_electronic,
        )
    except Exception as exc:
        return _json(
            {
                "error": "批量引用点检索失败，请检查代理、网络或输入格式。",
                "detail": str(exc),
                "proxy": effective_proxy,
            }
        )
    return _json(payload)


@mcp.tool()
def build_manifest_draft_from_points(batch_result_json: str) -> str:
    """从 search_and_verify_citation_points 结果生成 manifest 草稿。"""
    try:
        payload = json.loads(batch_result_json)
    except json.JSONDecodeError as exc:
        return _json({"error": f"batch_result_json 不是合法 JSON: {exc}"})

    points = payload.get("points") if isinstance(payload, dict) else None
    if not isinstance(points, list):
        return _json({"error": "输入必须是 search_and_verify_citation_points 的 JSON 输出。"})

    draft: list[dict[str, Any]] = []
    failures: list[dict[str, Any]] = []
    for idx, point in enumerate(points, start=1):
        if not isinstance(point, dict):
            continue
        point_id = str(point.get("point_id") or f"cp{idx:03d}")
        claim_text = str(point.get("claim_text") or "")
        selected = point.get("selected_writable_candidates") or []
        if not selected:
            failures.append(
                {
                    "point_id": point_id,
                    "claim_text": claim_text,
                    "reason": "无可写入候选，需补检索或人工确认。",
                }
            )
            continue
        for rank, candidate in enumerate(selected, start=1):
            if not isinstance(candidate, dict):
                continue
            draft.append(
                {
                    "key": f"{point_id}_{rank}",
                    "point_id": point_id,
                    "claim_text": claim_text,
                    "type": candidate.get("suggested_gbt_type"),
                    "title": candidate.get("title"),
                    "authors": candidate.get("authors"),
                    "year": candidate.get("year"),
                    "published_date": candidate.get("published_date"),
                    "updated_date": candidate.get("updated_date"),
                    "container": candidate.get("container"),
                    "doi": candidate.get("doi"),
                    "pmid": candidate.get("pmid"),
                    "url": candidate.get("url"),
                    "verification_source": candidate.get("source"),
                    "verification_url": candidate.get("url") or (f"https://doi.org/{candidate.get('doi')}" if candidate.get("doi") else ""),
                    "gbt7714": "",
                    "status": "draft",
                    "note": "请补全 gbt7714 后再交给 scripts/应用参考文献.py 写入 DOCX；若 type 为 EB/OL，必须补齐发表/更新日期、引用日期和访问路径。",
                }
            )
    return _json({"manifest_draft": draft, "not_ready_points": failures})


if __name__ == "__main__":
    mcp.run(transport="stdio")
