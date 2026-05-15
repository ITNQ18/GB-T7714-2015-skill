# GB/T 7714-2015 Reference Skill

一个面向中文论文、毕业设计和课程报告的 Codex 技能，用于按 GB/T 7714-2015 和常见学校模板处理参考文献、正文引用、Word 编号项交叉引用和文献真实性核验。

这个技能的目标不是“生成看起来像参考文献的文本”，而是帮助模型在写入前完成真实检索、来源核验、格式审计和引用覆盖检查。

## 功能概览

- 按 GB/T 7714-2015 和学校 7 类模板整理参考文献。
- 自动识别正文中需要引用的位置，并为引用点检索候选文献。
- 通过 CrossRef、OpenAlex、PubMed、arXiv 等来源检索和核验文献。
- 支持无联网能力模型通过 MCP + 脚本代理检索文献。
- 默认优先正式期刊、会议论文、论文集、专著、学位论文和专利。
- 默认不写入 `[EB/OL]`，但检索阶段保留电子来源给模型判断；用户允许时可启用电子文献写入策略。
- 生成真实 Word 编号项参考文献表，正文引用使用上标。
- 优先使用 Word“编号项/段落编号”交叉引用；无法使用 Word COM 时自动降级为书签 REF。
- 审计 DOCX 中的引用覆盖、越界引用、未引用文献、长串引用、字体字号、上标、编号和格式问题。
- 提供英文入口 `scripts/gbt.py`，便于在 cmd、bash、远程终端或非 UTF-8 环境下运行。

## 适用场景

- 中文毕业论文、毕业设计、课程报告参考文献整理。
- 已有 DOCX 的参考文献格式审计和修复。
- 从正文自动补充真实、可核验的参考文献。
- 将普通 `[1]`、`[2]` 引用升级为 Word 字段交叉引用。
- 检查参考文献是否都被正文引用。
- 拆分过密的长串引用，让每个引用支撑更具体的短语或论断。

## 目录结构

```text
.
├── SKILL.md                    # Codex 技能规则入口
├── agents/openai.yaml          # 技能展示元数据
├── requirements.txt            # 脚本运行依赖
├── scripts/
│   ├── gbt.py                  # 英文调度器
│   ├── 检索正式文献.py
│   ├── 应用参考文献.py
│   ├── 审计参考文献.py
│   ├── 抽取文档引用.py
│   ├── 生成参考文献概览.py
│   ├── 重排序号.py
│   └── 拆分引用组.py
├── mcp-server/
│   ├── gb_reference_server.py  # MCP 文献检索工具
│   ├── config.toml
│   └── requirements.txt
└── references/
    ├── literature-search.md
    └── mcp-search.md
```

## 安装

建议使用 Python 3.10+。Python 3.11+ 可直接使用标准库 `tomllib`；Python 3.10 会通过 `requirements.txt` 安装 `tomli`。

### PowerShell

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

### cmd

```cmd
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

### macOS / Linux

```bash
python3 -m venv .venv
./.venv/bin/python -m pip install --upgrade pip
./.venv/bin/python -m pip install -r requirements.txt
```

## 快速使用

如果终端可以正常输入中文文件名，可以直接运行中文脚本。更推荐使用英文调度器 `scripts/gbt.py`：

```bash
python -B scripts/gbt.py search "Monte Carlo Tree Search" --source crossref
python -B scripts/gbt.py audit target.docx
python -B scripts/gbt.py extract target.docx -o extract.json
python -B scripts/gbt.py apply source.docx references.json output.docx
```

PowerShell 示例：

```powershell
.\.venv\Scripts\python.exe -B scripts\gbt.py search "Attention Is All You Need" --source crossref
.\.venv\Scripts\python.exe -B scripts\gbt.py audit target.docx
```

macOS / Linux 示例：

```bash
./.venv/bin/python -B scripts/gbt.py search "Attention Is All You Need" --source crossref
./.venv/bin/python -B scripts/gbt.py audit target.docx
```

## 文献检索与真实性策略

本技能把“检索可见”和“最终写入”分开处理：

- 检索阶段保留 arXiv、OpenReview、GitHub、项目主页、官方文档、网页等线索，供模型阅读和判断。
- 默认写入阶段只接受正式、可查询、可核验的文献记录。
- 电子来源默认不写入 `[EB/OL]`，除非用户或学校模板明确允许。
- 用户允许 `[EB/OL]` 时，可启用 `--allow-electronic` 或 MCP 的 `allow_electronic=true`。
- 即使允许电子文献，也必须补齐作者、题名、发表或更新日期、引用日期和访问路径。

示例：

```bash
python -B scripts/gbt.py search "PGF TikZ vector graphics" --source crossref --source arxiv
python -B scripts/gbt.py search "PGF TikZ vector graphics" --source arxiv --allow-electronic
```

## MCP 检索工具

当模型本身没有联网能力时，可以通过 MCP server 暴露文献检索工具。

安装 MCP 依赖：

```bash
./.venv/bin/python -m pip install -r mcp-server/requirements.txt
```

Windows PowerShell：

```powershell
.\.venv\Scripts\python.exe -m pip install -r mcp-server\requirements.txt
.\.venv\Scripts\python.exe -B mcp-server\gb_reference_server.py
```

macOS / Linux：

```bash
./.venv/bin/python -B mcp-server/gb_reference_server.py
```

MCP server 提供的主要工具：

- `search_formal_references`：按查询词检索候选文献。
- `search_and_verify_citation_points`：按正文引用点批量检索并核验。
- `classify_gbt7714_candidate`：对单条候选做写入前判断。
- `filter_writable_candidates`：区分可写入和待确认候选。
- `build_manifest_draft_from_points`：从引用点检索结果生成 manifest 草稿。

默认代理为：

```text
http://127.0.0.1:7897
```

可用 `GBT7714_PROXY`、`HTTP_PROXY`、`HTTPS_PROXY` 或脚本参数 `--proxy` 覆盖；使用 `--proxy direct` 可禁用代理。

## DOCX 写入与审计

核心 DOCX 流程：

1. 准备带 `[[CITE:key]]` 占位符的 DOCX。
2. 准备已核验的 `references.json`。
3. 运行 `apply` 写入正文引用和参考文献表。
4. 运行 `audit` 审计编号、字段、引用覆盖和格式。

示例：

```bash
python -B scripts/gbt.py apply draft.docx references.json final.docx
python -B scripts/gbt.py audit final.docx -o audit.json
```

`应用参考文献.py` 默认 `--method auto`：

- Windows + Microsoft Word + PowerShell/COM 可生成 Word 编号项交叉引用。
- macOS/Linux 或无 Word 环境会降级为书签 REF fallback。
- 如果强制要求 Word 编号项交叉引用，可使用 `--method numbered-item`。
- 如果希望直接使用 fallback，可使用 `--method bookmark`。

## 跨平台说明

大多数脚本是纯 Python + OOXML 处理，可在 Windows、macOS 和 Linux 上运行。

需要注意：

- 中文脚本名在旧版 cmd、远程终端或非 UTF-8 shell 中可能不方便输入，建议使用 `scripts/gbt.py`。
- Word 编号项交叉引用依赖 Windows + Microsoft Word + PowerShell/COM。
- 网络检索依赖外部数据库可访问性，代理默认端口为 `7897`。
- MCP server 是 stdio server，不是普通 CLI；不要用 `--help` 判断是否运行成功，应由 MCP 客户端连接或导入模块验证。

## 许可证

本项目使用 MIT License，见 [LICENSE](LICENSE)。
