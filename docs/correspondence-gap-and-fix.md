# 产品缺口分析：正文引用但列表缺失 → 最终报告不显示

**发现日期：** 2026-08-28（README 制作过程中跑测试论文暴露）

## 缺口是什么

脚本明明**能查出**"正文引用了、但参考文献列表里没有"的问题，初筛底稿里也**画出来了**，但这类问题**到不了最终报告**——用户永远看不到。

## 数据流断在哪

```
check_correspondence()                    初筛 JSON                   初筛底稿 HTML        最终报告 HTML
scripts/refcheck.py:884                   data["correspondence"]      build_draft_report   build_final_report
                                         .cited_but_missing_in_list  :1326-1332 ✅ 渲染    ❌ 从不读 correspondence
                                                  │
                                                  ▼
                                         final.json verdicts 校验
                                         _validate_verdicts :1560
                                         id 必须 ∈ entries :1566
                                         ❌ 这类问题没有 entry id，写不进去
```

三处断点：

1. **校验层（根因）**：`_validate_verdicts`（refcheck.py:1560）要求每条 verdict 的 `id` 必须是文献列表里的条目（`R1`–`Rn`）。但"正文引用但列表缺失"的问题主体是**一条 citation（正文引用）**，它恰恰**没有**对应的列表条目——没有合法 id 可写，校验直接报错"不存在于文献列表"。
2. **渲染层**：`build_final_report`（refcheck.py:1591）只消费 `verdicts`，从不读 `data["correspondence"]`。即使数据在 JSON 里躺着，报告也不渲染。
3. **流程层**：SKILL.md 第 6 步的字段规则没有教 AI 层如何给这类问题写结论，AI 即使发现了也没有出口。

## 为什么重要

| 问题 | 当前后果 |
|------|---------|
| 测试论文里的 Pop et al. 2015（正文引用、列表缺失） | 初筛底稿可见，**最终报告完全不可见** |
| 真实用户场景 | 作者补写段落时随手引了一篇，忘了加进列表——这是**投稿前最容易发生、审稿人一眼就能看到**的问题，恰恰是工具最该拦的 |
| 反方向（列表有、正文从未引用） | 有 entry id，现有机制能覆盖，不受此缺口影响 |

## 修复方案（推荐 A）

### 方案 A：扩展 verdicts schema，支持 citation 型 verdict ✅ 推荐

给"正文引用"发合法身份证：引用在初筛 JSON 的 `citations[]` 数组里有下标，用 `C1`、`C2`… 作为 id 前缀。

**schema 变化**（向后兼容，纯新增）：

```json
{
  "verdicts": [
    {
      "id": "C3",
      "category": "correspondence",
      "final_status": "warn",
      "verdict": "正文引用了 Pop et al. (2015)，但参考文献列表中没有对应条目",
      "evidence": "引用出现在方法部分：\"…following Pop et al. (2015)…\"；列表按 姓+年份 检索无匹配",
      "action": "补充完整条目到参考文献列表，或删除该正文引用"
    }
  ]
}
```

**三处改动：**

1. **校验层**（`_validate_verdicts`）：`id` 匹配 `^C\d+$` 时按下标查 `citations[]`，合法；同时把校验的"自动异常必须有显式结论"规则扩展到 `cited_but_missing_in_list` 的每一条——AI 必须逐条复核（排除匹配误报，如引用年份与条目年份不一致导致的假阳性）后写结论
2. **渲染层**（`build_final_report`）：correspondence 类 verdict 渲染成专属卡片组"正文引用但列表缺失"，显示引用句原文 + 建议行动；Scholar 复核链接用 作者+年份 构造
3. **SKILL.md 第 6 步**：补字段规则——`C\d+` id 指向 citations 下标；cited_but_missing 的每条必须有显式结论（确认缺失→warn；系匹配误报→ok + note 说明实际对应哪条）

**优点**：与现有架构同构（数据驱动渲染、结论必须带证据、AI 复核后才下结论），不引入特殊通道。
**代价**：三处改动 + SKILL.md 同步，约 1-2 小时。

### 方案 B：脚本自动渲染，不走 verdicts

`build_final_report` 直接读 `data["correspondence"]`，有就画一节，无需 AI 参与。

**优点**：改动最小（只改渲染层）。
**缺点**：绕过了"AI 二次复核"设计——这类检测有假阳性（姓名拼写差异、年份不一致、et al. 匹配失败都会造成"假缺失"），未复核直接渲染会与"其余确认 N 条"的卡片自相矛盾，违背"结论必须有证据"的原则。**不推荐**，但可作为 A 的兜底：AI 忘写 C 类 verdict 时，校验会拦截（改动 1 保证了这一点）。

## 建议落地顺序

1. `refcheck.py`：校验层 + 渲染层（一次改完，跑 `tests/fixtures/test_paper.md` 回归——Pop et al. 2015 会自然成为验收用例）
2. `SKILL.md` 第 6 步字段规则同步
3. `assets/report-demo.png` 重截（新报告多一个"正文引用缺失"卡片，README 演示更完整）
4. README 的"8 类检查"描述不受影响（正文对应本来就是第 3 类，这次是让它真正闭环）

> 📖 **为什么这类问题没有 entry id？**
> 最终报告的每条结论都锚定一个文献列表条目（R1–Rn），这样证据、修改建议、"其余确认"的计数才能对得上。但"正文引用但列表缺失"的主体是一条正文引用，列表里恰恰没有它——所以要么给它发新身份证（方案 A 的 C1、C2…），要么绕过结论体系直接渲染（方案 B）。发身份证保留了"AI 复核 + 必须带证据"的质量闸门。
