<div align="center">

<img src="assets/banner.svg" alt="ob-reference-check — a full reference-list health check before you submit" width="800">

English · **[中文](README.md)**

[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg?style=for-the-badge)](LICENSE)
[![Python 3.9+](https://img.shields.io/badge/Python-3.9%2B-blue.svg?style=for-the-badge)](ob-reference-check/scripts/requirements.txt)
[![Formats](https://img.shields.io/badge/Supports-.docx%20%2F%20.pdf%20%2F%20.md-orange.svg?style=for-the-badge)](#-quick-start)
[![Free APIs](https://img.shields.io/badge/Data%20sources-free%2C%20no%20signup-brightgreen.svg?style=for-the-badge)](#-faq)

</div>

> [!NOTE]
> **Who is this for?** You write papers with AI assistance (who doesn't now), and before submission you need certainty that your reference list contains no fabricated citations, metadata errors, or misused sources. Manually checking a hundred references takes a full day and is easy to skip. **ob-reference-check has your AI assistant do it systematically** — 8 categories of checks, cross-validated against three academic databases, ending with a short list of exactly what to fix.

---

## 🚨 Why you need this

AI writing tools hallucinate references that **look completely real** — plausible authors, year, journal, volume, pages. Everything except an actual paper. Once this reaches a journal:

- **Desk rejection** — reviewers spot it with one search
- **Academic integrity risk** — researchers have been dismissed over AI-fabricated references

And even when references are real, there may be wrong page numbers, misspelled journal names, in-text citations missing from the list, or sources that don't actually support your claims. **Manual checking is slow, exhausting, and precisely the step people skip.**

## ✨ Highlights

|     | Capability                         | What it means for you                                                                                                                                            |
| --- | ---------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 🔍  | **8 categories in one pass**       | Existence, metadata, text–list correspondence, duplicates, timeline, preprints, cross-consistency, citation appropriateness — [see the full list below](#checks) |
| 🧠  | **Fabricated-reference detection** | Every entry cross-checked against OpenAlex / Crossref / Semantic Scholar — fake citations have nowhere to hide                                                   |
| ⚖️  | **Appropriateness deep-dive**      | Not just "does it exist" — compares each source's abstract against what you claim, because a real paper that doesn't support your point is equally dangerous     |
| 📄  | **Reads your manuscript directly** | .docx / .pdf / .md — parses the reference list and in-text citations, no format conversion needed                                                                |
| 🈵  | **Free, no signup**                | Runs entirely on public academic APIs; no API keys required                                                                                                      |
| 📋  | **Action-oriented report**         | The final report tells you exactly which entries to fix and how; everything else is confirmed clean at a glance                                                  |

## 🔍 What exactly gets checked: the 8 categories

<a id="checks"></a>

Every category is designed around the risk it intercepts. Categories 1–6 run automatically in the script (free, zero tokens); 7–8 are the AI layer's judgment:

| #   | Category                     | What it checks                                                                                                                            | What it intercepts                                                                          |
| --- | ---------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------- |
| 1   | **Existence**                | Every entry cross-checked against OpenAlex / Crossref / Semantic Scholar; misses get a second pass (publisher page) before any conclusion | **AI-fabricated references** — plausible authors, year, journal, pages… and no actual paper |
| 2   | **Metadata field-by-field**  | Authors, year, journal, volume, issue, pages, DOI compared against database records                                                       | Wrong page ranges, misspelled journal names (Review vs Reviews), year mismatches            |
| 3   | **Text–list correspondence** | Forward: cited in text but missing from the list; reverse: listed but never cited                                                         | Cited a paper while drafting a new paragraph, forgot to add it to the list                  |
| 4   | **Duplicates**               | The same reference appearing twice                                                                                                        | Duplicates pasted in while reusing paragraphs                                               |
| 5   | **Timeline & preprints**     | Publication year later than the current year (impossible); preprint citations where a published version may exist                         | AI hallucinations often "invent" future years; preprints reviewers ask you to update        |
| 6   | **Within-list cross-checks** | Two entries with swapped DOIs; same-author ordering and missing a/b suffixes; leftover numbering in titles                                | A swapped DOI silently points to someone else's paper; ordering costs style points          |
| 7   | **Appropriateness**          | Hypothesis-bearing citations checked one by one: does the abstract support the claim, same direction, same constructs                     | **A real paper that doesn't support your point** — you say positive, it found negative      |
| 8   | **Format consistency**       | et al. usage, `&` vs `and`, year parentheses, volume italics, page-range symbols, journal abbreviations                                   | Only internal inconsistencies — it doesn't push any particular style guide                  |

> 📖 **Why "no match" ≠ "fabricated"?** A lookup can miss because of indexing lag (papers < 3 months old) or online-first versions. The script only reports "not matched"; only after the AI's second pass still fails does the report say "could not be verified" — it will never call a reference fabricated for you, but it will never silently skip one either.

## 📊 What the report looks like

After one run you get a self-contained HTML report (this example is from the built-in test paper, [which you can run yourself](#manual-run-no-ai-assistant-needed)):

<img src="assets/report-demo.png" alt="Sample final report: summary cards, 8-category scope grid, must-fix issue cards" width="800">

The summary cards tell you at a glance: **3 must-fix / 0 needs-review / 0 formatting / 8 confirmed clean**. The example shows one of each classic problem: a fabricated (AI-hallucinated) citation, a wrong page range, and an in-text citation missing from the reference list. Each issue comes with evidence (what the database says vs. what your paper says) and a concrete fix, plus one-click DOI and Google Scholar links for your own verification.

## 🚀 Quick Start

### Method 1: Tell Your AI (Recommended)

Send this message to your AI coding assistant and it will install everything:

| Platform         | Copy this and send to your AI                                                                                    |
| ---------------- | ---------------------------------------------------------------------------------------------------------------- |
| **Claude Code**  | `Install the ob-reference-check skill from https://github.com/gtskevin/ob-reference-check`                       |
| **OpenAI Codex** | `Install the ob-reference-check skill from https://github.com/gtskevin/ob-reference-check`                       |
| **Gemini CLI**   | `Install the ob-reference-check skill from https://github.com/gtskevin/ob-reference-check`                       |
| **Cursor**       | Download the repo files into `.cursor/rules/`                                                                    |
| **Windsurf**     | Download the repo files into `.windsurf/rules/`                                                                  |
| **Other AI**     | Put the contents of [SKILL.md](ob-reference-check/SKILL.md) into your AI's custom instructions / rules directory |

> 💡 **Never used a terminal?** Once installed, just say one thing to your AI:
>
> **"Check the references in my paper: my-paper.docx"**
>
> The AI runs the full check and opens the report. No code required.

### Method 2: Manual Install

<details>
<summary>Claude Code / Codex / Gemini CLI</summary>

```bash
git clone https://github.com/gtskevin/ob-reference-check.git
# Claude Code
cp -r ob-reference-check/ob-reference-check ~/.claude/skills/
# OpenAI Codex
cp -r ob-reference-check/ob-reference-check ~/.codex/skills/
```

Or download just the core files without cloning:

```bash
mkdir -p ~/.claude/skills/ob-reference-check/scripts
curl -sL https://raw.githubusercontent.com/gtskevin/ob-reference-check/main/ob-reference-check/SKILL.md -o ~/.claude/skills/ob-reference-check/SKILL.md
curl -sL https://raw.githubusercontent.com/gtskevin/ob-reference-check/main/ob-reference-check/scripts/refcheck.py -o ~/.claude/skills/ob-reference-check/scripts/refcheck.py
curl -sL https://raw.githubusercontent.com/gtskevin/ob-reference-check/main/ob-reference-check/scripts/requirements.txt -o ~/.claude/skills/ob-reference-check/scripts/requirements.txt
```

</details>

<details>
<summary>Manual run (no AI assistant needed)</summary>

The script covers the mechanical checks (6 of the 8 categories) on its own; all you need is Python 3.9+:

```bash
git clone https://github.com/gtskevin/ob-reference-check.git
cd ob-reference-check
python3 -m venv .venv && .venv/bin/pip install -r ob-reference-check/scripts/requirements.txt
.venv/bin/python ob-reference-check/scripts/refcheck.py your-paper.docx   # also .pdf / .md
```

Outputs (next to your paper):

- `your-paper_refcheck_YYYYMMDD.html` — automated screening draft
- `your-paper_refcheck_YYYYMMDD.json` — structured data

> ⚠️ Appropriateness deep-dive and format-consistency review need the AI layer — that's exactly the value of installing it as a skill. To preview the flow, run the built-in test paper `tests/fixtures/test_paper.md`.

</details>

## ⚙️ How it works

Core design: **mechanical checks go to a script (zero tokens, reproducible, free); the AI only makes the judgments machines can't.**

```
Your paper (.docx/.pdf/.md)
   │
   ▼
① Script screening (free, zero tokens)
   Parse reference list + in-text citations
   ├─ Existence: each entry looked up across OpenAlex / Crossref / Semantic Scholar
   ├─ Metadata: authors/year/journal/volume/pages/DOI compared field by field
   ├─ Correspondence: in-text citations ↔ list entries, both directions
   ├─ Duplicates / timeline anomalies / preprint versions
   └─ Within-list cross-checks: swapped DOIs, author ordering, leftover numbering
   │
   ▼
② AI review and deep-dive
   ├─ Second pass on anomalies: "no match" ≠ "doesn't exist" — the AI checks publisher pages before concluding
   ├─ Appropriateness: compares what you claim vs. what the source's abstract actually studied
   │   A/B/C triage — hypothesis-bearing citations get deep checks, passing mentions get light checks
   └─ Format consistency: et al. usage, &/and, journal abbreviations, internal uniformity
   │
   ▼
③ Final report (HTML)
   N must-fix entries (with evidence + suggestion) / needs review / formatting / rest confirmed
```

Details worth knowing:

- **Evidence required**: every conclusion in the report carries a verdict + evidence + concrete fix. When unsure, it flags "needs your review" rather than guessing — a missed problem costs far more than a false alarm
- **Rotated sources spread quotas**: for a 100-reference paper, free API quotas are spread across three databases; a rate-limited source trips a breaker and the others take over
- **Global cache**: `~/.reference_check/cache/` — re-checking a revised draft hits cache instantly
- **Verdict memory**: your confirmed conclusions are stored per DOI and reused the next time you check the same references

## 🤔 Compared to alternatives

|                                    | Manual checking                  | Asking ChatGPT directly                                     | ob-reference-check                                                    |
| ---------------------------------- | -------------------------------- | ----------------------------------------------------------- | --------------------------------------------------------------------- |
| Fake references                    | Works, but hours for 100 entries | ❌ answers from "memory" — the very thing that hallucinates | ✅ live three-source database lookup                                  |
| Metadata comparison                | Fatigue leads to misses          | ❌ unreliable without live lookup                           | ✅ automatic field-by-field comparison                                |
| Does the source support the claim? | ✅ requires reading the paper    | Partial, no systematic flow                                 | ⚠️ abstract-level comparison (see [limitations](#-known-limitations)) |
| Text–list correspondence           | ✅ tedious                       | ❌ misses things in long text                               | ✅ automatic bidirectional matching                                   |
| Cost                               | A full day of yours              | Subscription                                                | Subscription + free academic APIs                                     |
| Reproducibility                    | Varies by person                 | Different every time                                        | ✅ script screening is fully reproducible                             |

> Honest note: this tool **does not replace reading the source papers yourself**. Appropriateness judgments are abstract-level; entries flagged "needs review" deserve a human read. What it does is pinpoint exactly where your time is best spent.

## ⚠️ Known limitations

- Entry parsing covers author-year styles (APA / Harvard etc.); entries that fail parsing are handled by the AI layer or flagged for manual review
- Appropriateness judgment is based on **abstracts**, not full texts
- Very recent publications (< 3 months) may not be indexed yet — the report says "could not verify" rather than "fabricated"
- Scanned PDFs (no text layer) fail loudly with a clear error — it will never pretend to have checked

## ❓ FAQ

<details>
<summary>Does my paper get uploaded anywhere?</summary>

No. The script only sends **individual reference metadata** (title/authors/year) to public academic database APIs for lookup; your manuscript is parsed locally. The final report is a local HTML file — nothing is stored in the cloud.

</details>

<details>
<summary>Do I really not need API keys?</summary>

Correct. OpenAlex and Crossref public endpoints are free. Setting `OPENALEX_API_KEY` / `SEMANTIC_API_KEY` environment variables improves hit rate and stability, but is purely optional.

</details>

<details>
<summary>I'm not in organizational behavior — can I still use it?</summary>

Yes. The checks (existence, metadata, correspondence, format) work for any paper with a reference list. The appropriateness examples lean toward management/psychology phrasing, but the method is general.

</details>

<details>
<summary>How long / how expensive is a run?</summary>

Script screening typically finishes in minutes (depending on reference count and network) and costs zero AI tokens. The AI review stage costs tokens proportional to the number of anomalies — most references verify once and are cached, so re-checks are cheap.

</details>

<details>
<summary>The AI says a reference "could not be verified" — now what?</summary>

The report suggests a fix: usually delete or replace the citation. If you're confident the paper exists (e.g., it's very recent), click the Google Scholar link in the report to confirm manually — the report is designed so each manual re-check takes about 30 seconds.

</details>

## 📄 License

[MIT](LICENSE)

## 🙏 Acknowledgments

Data sources: [OpenAlex](https://openalex.org) · [Crossref](https://www.crossref.org) · [Semantic Scholar](https://www.semanticscholar.org)
