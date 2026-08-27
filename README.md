# ob-reference-check

论文投稿前的**参考文献全面体检**工具。专为担心 AI 辅助写作产生幻觉引用的研究者设计（组织行为学等社会科学领域友好）。

一次检查覆盖 8 类问题：

| # | 检查项 | 执行者 |
|---|--------|--------|
| 1 | 存在性验证（AI 编造文献拦截） | 脚本 |
| 2 | 元数据核对（作者/年份/期刊/卷期页/DOI） | 脚本 |
| 3 | 引用恰当性（文献是否真的支撑论述，摘要级比对） | Claude |
| 4 | 双向对应（正文引用 ↔ 文献列表） | 脚本 |
| 5 | 重复条目 | 脚本 |
| 6 | 引用格式内部一致性 | Claude |
| 7 | 时间线异常（引用了不可能存在的年份） | 脚本 |
| 8 | preprint 版本问题 | 脚本 |

## 快速开始（Claude Code 用户）

安装到 `~/.claude/skills/ob-reference-check/` 后，对话中说"检查这篇论文的参考文献 `<文件>`"即可，或 `/ob-reference-check`。

## 手动使用（不需要 Claude Code）

```bash
# 1. 安装依赖（Python 3.9+）
python3 -m venv .venv && .venv/bin/pip install -r ob-reference-check/scripts/requirements.txt

# 2. 运行检查
.venv/bin/python ob-reference-check/scripts/refcheck.py 你的论文.docx   # 也支持 .pdf / .md
```

产出（论文同目录）：
- `你的论文_refcheck_YYYYMMDD.html` — 检查报告（自包含 HTML，双击/浏览器打开，每条结果带 DOI/OpenAlex 可点击复核链接）
- `你的论文_refcheck_YYYYMMDD.json` — 结构化数据（含摘要、分诊结果，供进一步分析）

手动模式完成第 1/2/4/5/7/8 类检查（零 LLM token）；第 3/6 类需要 Claude 层。

## 特性

- **免费无 key 可用**：OpenAlex + Crossref 基础 API 免费。设置 `OPENALEX_API_KEY` / `CROSSREF_API_KEY` / `SEMANTIC_API_KEY`（Semantic Scholar）环境变量可加速（不要硬编码到任何地方）
- **三源轮换分摊额度**：大论文上百条文献时，每条文献轮换首选源（OpenAlex → Crossref → Semantic Scholar），免费日额度摊到三家；某个源限流会自动熔断，剩余源接管
- **全局缓存**：`~/.reference_check/cache/`，改稿重查秒命中已验证文献
- **三格式支持**：Word (.docx) / PDF / Markdown，纯 Python 解析，无 pandoc 依赖
- **数据源兜底**：OpenAlex 模糊检索 → Crossref DOI 核对，自动熔断限流源
- **A/B/C 分诊**：承重引用逐条深查、顺带提及轻查、引用堆砌只查真伪——token 花在刀刃上

## 已知限制

- 条目解析是启发式（覆盖 author-year 类 style：APA/Harvard 等），`parse_ok: false` 的条目由 Claude 层兜底或人工核对
- 恰当性判断基于**摘要**，不含全文；标"存疑"的建议人工读原文
- 新发表文献（<3 个月）可能未被数据库收录，报告标"无法验证"而非"编造"
- 扫描版 PDF（无文字层）会明确报错，不会假装检查过

## 开发

```bash
python3 -m venv .venv && .venv/bin/pip install -r ob-reference-check/scripts/requirements.txt
.venv/bin/python ob-reference-check/scripts/refcheck.py tests/fixtures/test_paper.md
```

测试论文 `tests/fixtures/test_paper.md` 内含已知埋点：1 条编造文献（R2）、正文引用 Pop et al. 2015 缺失于列表、Mayer 1995 页码错误。设计文档见 `docs/design.md`。
