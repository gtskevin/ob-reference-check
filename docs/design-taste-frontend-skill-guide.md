# design-taste-frontend（taste-skill）使用指南

> 安装位置：`~/.claude/skills/design-taste-frontend/`（全局，所有项目可用）
> 来源：[leonxlnx/taste-skill](https://github.com/leonxlnx/taste-skill) · 81k stars · MIT · 当前版本 v2（实验性）

---

## 一句话总结

**防止 AI 生成"一眼假"的模板化 UI。** 它是一份约 1200 行的强约束设计规范，注入到 Claude 的设计工作中，专门拦截 LLM 的默认审美惯性。

> 📖 **什么是 "AI slop"（AI 垃圾审美）？**
> 你让任何 AI 做落地页，大概率得到同一套东西：紫色渐变背景、居中大标题、三张一模一样的功能卡片、Inter 字体。就像全国连锁快餐店——每家店装修都一样，没有记忆点。这个 skill 就是强制 AI"绕开连锁店 defaults，做一家有主见的独立餐厅"。

---

## 它解决什么问题

LLM 做设计时的通病是**跳到默认审美**而不读需求。这个 skill 强制执行一套工作流程：

| 问题（AI 默认行为） | skill 的对策 |
|---|---|
| 不看需求就动手 | **Step 0 强制"设计判读"**：先输出一行 "Reading this as: <页面类型> for <受众>..." 再写代码 |
| 紫色渐变 + 玻璃拟态滥用 | **LILA RULE**：AI 紫作为默认被禁；中性底色 + 单一高对比强调色 |
| Inter 字体 + 居中 hero + 三等分卡片 | 反中心偏差、禁三等宽卡片、字体优先 Geist/Satoshi 等 |
| 衬线体滥用（"创意 = 衬线"的 AI 条件反射） | Serif 纪律：默认禁用，明确点名封杀 Fraunces 和 Instrument Serif |
| 高端消费品 = 米色+黄铜+深棕（AI 固定配色） | 整个色系家族列为默认禁用，强制轮换（冷奢银灰/森林绿/钴蓝+奶油等） |
| em-dash（—）满屏飞 | **零容忍禁令**：页面任何位置出现一个 `—` 即判定失败 |
| 每个区块上方都顶个小标签（eyebrow） | 机械计数规则：每 3 个 section 最多 1 个 eyebrow |
| 假截图（用 div 拼假产品界面） | 直接封杀：用真图、生成图，或明确留占位符 |

---

## 核心机制：三个"旋钮"（Dials）

生成前设定三个全局变量，所有布局/动效/密度决策由它们驱动：

| 旋钮 | 范围 | 低 → 高 | 默认 |
|---|---|---|---|
| `DESIGN_VARIANCE` | 1-10 | 完美对称 → 不对称/艺术化混乱 | 8 |
| `MOTION_INTENSITY` | 1-10 | 静态 → 电影级物理动效 | 6 |
| `VISUAL_DENSITY` | 1-10 | 美术馆留白 → 驾驶舱信息密度 | 4 |

skill 内置了"需求 → 旋钮值"的映射表，例如你说 "Linear 风格极简"，它会自动设成约 `5-6 / 3-4 / 2-3`；说 "Awwwards 实验性"，则拉到 `9-10 / 8-10 / 3-4`。

---

## 工作流程（skill 触发后 Claude 会怎么做）

1. **判读需求（Section 0）**：页面类型、氛围词、参考信号、受众、已有品牌资产。歧义时只问**一个**澄清问题，不连环追问
2. **设旋钮（Section 1）**：根据判读结果定三旋钮值
3. **选基建（Section 2）**：如果需求匹配某个真实设计系统（Fluent/Material/Carbon/shadcn 等），用官方包，不手造 CSS；如果是"美学风格"而非系统（粗野主义、杂志编辑风等），诚实标注并自建
4. **按规范生成（Section 3-10）**：字体、配色、布局多样性、交互状态、动效骨架代码（GSAP sticky-stack 等官方代码模板）
5. **发布前检查（Section 14）**：跑一遍约 60 项的 Pre-Flight 清单，任何一项不过 = 未完成

---

## 怎么用

### 自动触发
skill 的 description 覆盖 landing pages / portfolios / redesigns。你正常提设计需求即可：

```
帮我做一个学术工具产品的落地页，Linear 风格，克制一点
```

Claude 会自动加载这个 skill 并走上面的流程。

### 显式调用 + 手动调旋钮
```
用 design-taste-frontend 做这个页面，
DESIGN_VARIANCE 3 MOTION_INTENSITY 2 VISUAL_DENSITY 5   ← 信任优先、政府服务风格
```

### 用在改版（redesign）场景
skill 会先区分模式：**保留品牌**（先审计现有 token，渐进现代化）vs **彻底翻新**（视觉重新来，内容保留）。它会先审计再动手，且不擅自改 URL 结构、导航文案、logo。

---

## 适用边界（重要）

**适用：** 落地页、作品集、营销站、redesign、博客/编辑类页面

**明确不适用（Section 13）：**
- Dashboard / 管理后台 / 密集产品 UI（它指向 Fluent、Carbon 等专业系统）
- 数据表格（指向 TanStack Table）
- 多步表单向导
- 代码编辑器、原生移动端、实时协作 UI

> 📖 **为什么要划边界？**
> 这套规范的核心是"第一印象优先、内容做减法"——落地页的逻辑。而 dashboard 的逻辑是"信息效率优先"。把落地页美学套到 dashboard 上，就像用婚纱设计思路做工装，方向就错了。

---

## 技术栈偏好（生成代码时）

- React / Next.js（默认 Server Components，动效隔离在 `'use client'` 叶子组件）
- Tailwind v4 + Motion（原 Framer Motion，从 `motion/react` 导入）
- 图标：Phosphor / HugeIcons / Radix / Tabler（不鼓励 Lucide，禁止手绘 SVG 图标）
- 强制：`min-h-[100dvh]`（不用 `h-screen`）、暗色模式双模式、`prefers-reduced-motion` 支持、WCAG AA 对比度

---

## 和你已有 skill 的关系

你本地已装了 `interface-design`、`ui-design-review`、`frontend-design` 等同类 skill。分工建议：

| 场景 | 用哪个 |
|---|---|
| 从零做落地页/作品集 | `design-taste-frontend`（规则最全、最严格） |
| 审查已有 UI | `ui-design-review` 或本 skill 的 redesign 协议 |
| dashboard / 产品 UI | `interface-design`（本 skill 明确不管这类） |

同时触发多个设计类 skill 时，Claude 会按 description 相关度选一个主导，但显式点名（"用 design-taste-frontend"）最可控。

---

## 快速验证它在工作

生成页面后检查这几个信号：
1. Claude 开头是否输出了一行 **"Reading this as: ..."** 的设计判读
2. 页面无紫色渐变、无三等分卡片、无 em-dash
3. 是否说明了三个旋钮的取值
4. 图片是否用了真实来源而非 div 拼的假截图
