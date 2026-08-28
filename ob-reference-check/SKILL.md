---
name: ob-reference-check
description: 论文投稿前的参考文献全面体检。当用户提到"检查参考文献"、"核对引用"、"文献是不是编造的"、"AI 写的引用靠谱吗"、"投稿前检查论文"、"reference check"、"查重文献"、"这个引用存在吗"，或者用户提供了论文文件并表达对引用真实性/准确性的担忧时，必须使用本 skill。也适用于检查 AI 辅助写作产生的幻觉引用、元数据错误、引用与文献列表不对应、引用不当（文献不支撑论述）等问题。即使用户只要求检查一条引用，也可以用本 skill 的脚本快速验证。
---

# ob-reference-check — 参考文献系统检查

论文投稿前把**所有**参考文献相关问题一次查出：AI 编造文献、元数据错误、引用不当、正文与列表不对应、格式不一致、列表内部一致性问题（DOI 错挂、排序、拼写）。

设计核心：**机械工作交给脚本（零 token、可复现），Claude 只做机器做不了的判断**（恰当性、误报分诊、格式审查）。不要跳过脚本手动逐条搜索——那既慢又不稳定。定稿也由脚本渲染——**绝不手改 HTML**（regex 同步五处状态必然漏，2026-08-28 交付 bug 教训）。

## 工作流程

### 第 0 步：环境准备（仅首次）

```bash
# 在 skill 的 scripts/ 目录（或任意可写位置）建 venv 并装依赖
python3 -m venv <skill_dir>/scripts/.venv
<skill_dir>/scripts/.venv/bin/pip install -r <skill_dir>/scripts/requirements.txt
```

之后每次用 `<skill_dir>/scripts/.venv/bin/python` 运行。如果用户机器上已有这些包（python-docx / pymupdf），直接 `python3` 也行。

**已知沙箱限制**（Claude Code 环境，2026-08-28 实测）：

- `pip install` 可能 SSL 失败 → 需要非沙箱运行
- 写论文所在目录（如坚果云同步目录）会被沙箱拦截 → 脚本跑检索前会先探测可写性并 fail fast，此时改用非沙箱重跑（结果有缓存，第二次很快），或加 `--outdir` 输出到可写目录

### 第 1 步：运行脚本（机械层）

```bash
<python> <skill_dir>/scripts/refcheck.py <论文文件>
```

- 支持格式：`.docx` / `.pdf` / `.md` / `.txt`
- 产出两个文件（论文同目录，或 `--outdir` 指定）：
  - `<论文名>_refcheck_YYYYMMDD.html` — **自动初筛底稿**（自包含 HTML，默认不自动打开；它不是最终交付报告）
  - `<论文名>_refcheck_YYYYMMDD.json` — 结构化数据（供后续复核；含本次可用数据源能力，不含任何 Key 值）
- 可选 Key：`OPENALEX_API_KEY` 与 `SEMANTIC_API_KEY`（亦兼容 `SEMANTIC_SCHOLAR_API_KEY` / `S2_API_KEY`）。有 Key 时应使用，以提高命中率、稳定性和摘要覆盖；无 Key 时仍必须使用 OpenAlex 与 Crossref 的公开接口完成初筛，Semantic Scholar 仅跳过，不能降低结论门槛。Crossref 不要求 Key。
- 只有用户明确要求查看初筛底稿时才加 `--open-draft`。默认不要打开、不要向用户转述脚本的逐条过程输出。
- 网络不可用时加 `--offline` 只用本地缓存（`~/.reference_check/cache/`）

脚本自动覆盖的机械检查：存在性三源检索（全部失败且有 DOI 时 Crossref DOI 直查兜底）、元数据逐项比对（venue 词级容差、近似差异标 `level: "near"` 交 AI 裁决）、双向对应、重复、时间线、preprint、**列表内交叉检测**（DOI 互换错挂 / 同作者组排序 / 标题残留编号，见 JSON `cross_checks` 字段）。

**自动异常不是结论**：`not_found` 只表示"自动未匹配"，可能来自解析、索引覆盖、在线发表版本或限流。不得向用户称为"文献不存在""疑似编造"或"最高风险项"；必须进入第 2 步二次复核。

### 第 2 步：汇总异常并二次复核（内部完成，不展示中间报告）

读 `*_refcheck_*.json`。脚本已把正文引用分为：

- **A 类（承重引用）**：出现在假设/理论推导句中 → 需要逐条深查
- **B 类（顺带提及）**：背景性引用 → 轻查（真实 + 主题相关）
- **C 类（引用堆砌）**：单括号 ≥3 条并排 → 跳过恰当性（堆砌处本就不逐条深支撑）

分诊由脚本按句子位置启发式定档，**Claude 可对分诊纠偏**：读到明显承重但被标 B/C 的引用（或反之）时，按其实际角色处理，不受脚本标签限制。

默认按此分诊直接完成检查，不要向用户展示 A/B/C 清单、自动命中数、初筛 HTML 或逐条检索过程。只有用户明确要求调整审读范围、或某个范围选择会实质改变工作量时，才询问分诊。

对**每一项自动异常**做二次复核后，才允许写入最终报告：

1. `not_found`：优先用批量子命令 `<python> refcheck.py --verify-doi R7,R16 <json路径>`（一次 Crossref DOI 直查，替代逐条 WebFetch）；无 DOI 的用精确标题、作者+年份和出版商页/多数据库检索复核。最终仅可归为"已确认存在""仍无法核实"或"确认需修正"；没有充分证据时不得称为编造。`note` 含"疑似 DOI 错挂"的条目还要看 `cross_checks.doi_swaps` 是否与列表内其他条目互换。
2. 元数据不一致：区分真实笔误与 online-first/正式卷期、页码简写、DOI 迁移或数据库字段缺失。只有真实笔误才写"需修改"。`level: "near"` 的差异是脚本按词级比对放行的疑似缩写/拼写变体，逐条人工裁决。
3. 正文—列表不对应、解析可疑、重复、时间线和预印本提醒：逐项确认原文和记录，排除解析误判后再报告。
4. `prior_verdict` 字段（若有）：该 DOI 在历史检查中已有人工结论，直接沿用或快速复核后沿用，不必从零分诊。

**JSON 字段速查**：`entries[].{id, raw, authors, year, title, venue, volume, issue, pages, doi, parse_ok}`；`citations[].{authors, year, sentence, section, triage}`（引用对应关系按 authors+year 匹配，无 refs 字段）；`verification[R*].{status, mismatches[], record, links, abstract, prior_verdict?}`；`cross_checks.{doi_swaps[], ordering[], title_artifacts[]}`。

### 第 3 步：A 类深查（引用恰当性）

对每条 A 类引用，用 JSON 里的两个字段做判断：

- `citations[i].sentence` — 正文引用处的完整句子
- `verification[R*].abstract` — 数据库返回的文献摘要

判断标准（写进报告时必须引用证据）：

1. 摘要中是否能找到对句子所述关系的**直接或合理间接支撑**
2. 结论方向是否一致（论文说正相关，文献是否真的是正相关）
3. 构念是否对齐（论文的 X/Y 与文献研究的变量是否同名同义）

输出三档：

- ✅ 支撑 — 摘要明确支持论述
- ⚠️ 存疑 — 摘要与论述有出入（写明哪里对不上，这是最有价值的发现）
- ❓ 摘要不足 — 文献真实但摘要信息不够判断（建议人工读原文）

**保守原则**：宁可标"存疑/摘要不足"让用户复核，不要替用户拍板"没问题"。这个 skill 的价值在于拦截风险，漏报比误报代价高。

如果某条 A 类引用自动未匹配，先完成上述存在性二次复核；在最终状态仍为"无法核实"时，才标为"引用恰当性无法判断"。

### 第 4 步：B 类轻查

B 类引用只需确认：对应文献 status 是 `found` 且摘要主题与句子领域相关。扫一遍即可，只报告明显跑题的（如论文谈信任、文献是纯算法论文）。C 类跳过。

### 第 5 步：格式一致性与列表 sanity 审查

读参考文献列表原文（JSON 的 `entries[].raw`），检查同一 style 内部统一性：

- et al. 使用规则是否一致（以论文主体风格为准，跨 style 混用才报告）
- `&` vs `and` 是否混用
- 年份括号、卷期斜体（markdown 中表现为 `*`）、页码范围符号（- vs –）是否统一
- 期刊名缩写与全称是否混用

只报告**不一致**（内部矛盾），不要求符合某个特定 style——用户的论文未必是 APA，都是 OB 常用 style。

**书目 sanity 清单**（脚本交叉检测之外，通读一遍列表人工兜底——2026-08-28 复盘：5 个真问题靠通读发现，机器召回仅 29%）：

- [ ] 期刊名单复数/拼写（Review vs Reviews——脚本词级比对已覆盖，但新期刊名仍可能漏）
- [ ] DOI 前缀与期刊匹配（AMJ 论文配 10.5465/amj、Annals 配 10.5465/19416520 等）
- [ ] 标题开头是否有残留章节号/编号
- [ ] 同一作者多条是否按年份升序、同作者同年是否加 a/b 后缀
- [ ] 章节类条目（书章）是否缺编者/卷号（数据库常不收录，需出版商页确认）

### 第 6 步：报告定稿（产出 final.json，脚本渲染，绝不手改 HTML）

把第 2-5 步的最终结论写成 `<论文名>_final.json`（论文同目录）：

```json
{
  "verdicts": [
    {
      "id": "R52",
      "final_status": "warn",
      "verdict": "期刊名拼写错误",
      "evidence": "出版商页面为 Neuroscience & Biobehavioral Reviews",
      "action": "Review → Reviews"
    },
    {
      "id": "R26",
      "final_status": "ok",
      "note": "曾自动未匹配，复核确认存在（直接引语页码后缀导致）"
    },
    {
      "id": "R45",
      "category": "appropriateness",
      "final_status": "warn",
      "verdict": "作为多层级中介例证存疑",
      "evidence": "摘要为跨层次调节而非中介",
      "action": "读原文确认或换例证文献"
    }
  ]
}
```

字段规则（脚本会校验，不合格直接报错）：

- `final_status`：`ok` / `warn`（确认需修改/存疑）/ `info`（建议人工核对）
- `category`（可选，默认 `bibliography`）：`bibliography` / `correspondence` / `appropriateness` / `format`
- **warn/info 必填 `verdict` + `evidence` + `action`**——没有证据不得下结论
- 自动初筛有异常（not_found / 差异）的条目**必须有显式结论**，不允许静默跳过
- `note`（可选，**不渲染**）：仅供 AI 层备忘复核理由，最终报告不展示任何"曾自动标记…误报"类过程信息（2026-08-28 用户反馈：看不出指哪条且属中间过程）

然后渲染并交付：

```bash
<python> <skill_dir>/scripts/refcheck.py --finalize <论文文件 | 初筛json | final.json> [--final xx.json]
open <论文名>_refcheck_YYYYMMDD_final.html
```

三种入参等价（脚本会自动定位初筛数据和结论文件）：论文文件、`*_refcheck_*.json`、或直接给 `*_final.json`。

脚本按 final.json 数据驱动渲染最终报告（概览卡片/nav 计数/各节/附录徽章全部同源），并把结论按 DOI 回流 `~/.reference_check/verdicts.json`（下次检查同批文献自动作为 `prior_verdict` 提供）。

**最终报告信息架构**（用户 2026-08-28 反馈，模板已固化）：概览只留行动导向卡片（必须修改/存疑核对/格式/其余确认），无重复段落文字；无过程性栏目（"二次复核""排除误报 N 项""曾自动标记…误报"之类一律不出现，note 字段不渲染）；总览下方有"本次检查范围"小节（脚本自动生成，**完整列出 8 类检查——存在性 / 元数据 / 正文对应 / 重复 / 时间线与预印本 / 列表内交叉检测 / 恰当性 / 格式，零命中的类别也保留"未发现问题"，让用户看到覆盖面而放心**）；scope 状态语义：**绿勾只表示"未发现问题"，有发现的类别用琥珀色 ⚠️ + 计数并标注"见下方详情"**——8 格全绿会和"必须处理 N 项"自相矛盾；复核入口每条只保留 DOI + Scholar 两个按钮（无 DOI 时仅 Scholar，Scholar 比数据库直达对作者更直观）；附录带 ⚠️/❓ 图例说明；无问题条目默认折叠。

只在定稿后做一次交付：向用户输出**一段执行摘要**（不超过 10 行），说明最终必须处理项和仍待作者读原文的项；随后用 `open` 命令打开最终 HTML。除非用户主动要求，否则不要输出过程性结果或打开初筛底稿。

## 边界与故障处理

- **解析可疑的条目**（JSON 中 `parse_ok: false`）：脚本启发式失败时，Claude 直接读 `raw` 字符串人工解析，把修正后的字段用于判断
- **扫描版 PDF**：脚本报错退出时，建议用户转 Word/Markdown 后重跑，不要用 OCR 结果凑合
- **OpenAlex 限流**（脚本统计中 `failed` 高）：正常现象，Crossref 兜底已覆盖；无需干预
- **输出目录不可写**：脚本启动时 fail fast 并提示——加 `--outdir` 或非沙箱重跑（检索结果已缓存）
- **合法的"查不到"**：新发表文献（近 3 个月）可能尚未被收录，报告中标"❓ 无法验证"并注明可能是收录延迟，不要断言是编造
- **用户只问单条引用**：直接跑脚本（它会检查全部但报告里指向该条），或 `--verify-doi R12` 一次直查——不要凭记忆回答文献是否存在，**必须查库**
