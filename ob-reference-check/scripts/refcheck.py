#!/usr/bin/env python3
"""
ob-reference-check — 论文参考文献系统检查（机械层）

用法:
    python refcheck.py <论文文件 .docx/.pdf/.md> [选项]
    python refcheck.py --verify-doi R7,R16 <论文文件或 *_refcheck_*.json>
    python refcheck.py --prepare-final <*_refcheck_*.json> [--final xx.json]
    python refcheck.py --finalize <论文文件或 *_refcheck_*.json> [--final xx.json] [--source 稿件]

做什么（脚本层，零 LLM token）:
    1. 解析 Word / PDF / Markdown 三种格式
    2. 提取参考文献列表条目 + 正文引用标记
    3. OpenAlex / Crossref / Semantic Scholar 检索初筛 + 元数据比对（带全局缓存）
       检索全部失败时若条目带 DOI，用 Crossref DOI 直查兜底
    4. 机械检查: 双向对应 / 重复条目 / 时间线异常 / preprint 版本
       + 列表内交叉检测（DOI 互换错挂 / 同作者排序 / 标题残留编号）
    5. A/B/C 分诊初筛（承重引用 / 顺带提及 / 引用堆砌）
    6. 生成自包含 HTML 初筛底稿（默认不打开）+ .json 数据文件（供人工复核）
    7. --verify-doi: 批量 DOI 直查（替代人工逐条检索）
    8. --finalize: 读 Claude 复核结论 final.json，数据驱动渲染最终报告
       （带证据的书目结论按条目指纹保存，下次运行可作为 prior_verdict）

不做什么:
    - 引用恰当性判断（层 3）→ 由 Claude 读 .json 中的句子+摘要完成
    - 引用格式一致性审查 → 由 Claude 完成
    - final.json 的结论本身 → 由 Claude 复核后产出，脚本只校验+渲染
"""

import argparse
import datetime
import glob
import difflib
import hashlib
import html as html_mod
import json
import os
import re
import subprocess
import sys
import time
import urllib.error
import urllib.parse
import urllib.request

CACHE_DIR = os.path.expanduser("~/.reference_check/cache")
USER_AGENT = "ob-reference-check/1.0 (mailto:ob-refcheck@example.com)"

# 语义化版本（发布到 GitHub 后供更新检查比对；详见 SKILL.md 分发说明）
__version__ = "1.4.0"

# ---------------------------------------------------------------------------
# 1. 文档解析（三格式 → 段落列表，每段带 heading 信息）
# ---------------------------------------------------------------------------

def parse_document(path):
    """返回 list[ {text, heading, level} ]，heading 为 None 表示正文段落。"""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx":
        return _parse_docx(path)
    if ext == ".pdf":
        return _parse_pdf(path)
    if ext in (".md", ".markdown", ".txt"):
        return _parse_md(path)
    sys.exit(f"[错误] 不支持的格式: {ext}（支持 .docx / .pdf / .md / .txt）")


def _parse_docx(path):
    import docx  # python-docx
    d = docx.Document(path)
    out = []
    for p in d.paragraphs:
        text = p.text.strip()
        if not text:
            continue
        if p.style and p.style.name and p.style.name.startswith("Heading"):
            try:
                level = int(p.style.name.split()[-1])
            except ValueError:
                level = 1
            out.append({"text": text, "heading": text, "level": level})
        else:
            out.append({"text": text, "heading": None, "level": 0})
    return out


def _parse_pdf(path):
    import fitz  # PyMuPDF
    doc = fitz.open(path)
    if not any(page.get_text().strip() for page in doc):
        sys.exit("[错误] PDF 无文字层（疑似扫描版）。请先用 OCR 处理，"
                 "或导出为文字版 PDF —— 脚本不会在无文字时假装检查过。")
    paragraphs = []
    for page in doc:
        for block in page.get_text("blocks"):
            text = block[4].strip()
            if not text:
                continue
            # PDF 没有 heading 结构。block 可能把标题和正文合并在一起
            # （行距近时常见），所以逐行扫描: 独立成行且匹配标题模式的行拆出来
            lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
            buf = []
            for ln in lines:
                is_heading = (len(ln) < 80 and re.match(
                    r"^(\d+(\.\d+)*\s+\S|References\b|参考文献)", ln, re.IGNORECASE))
                if is_heading:
                    if buf:
                        paragraphs.append({"text": " ".join(buf), "heading": None,
                                           "level": 0})
                        buf = []
                    paragraphs.append({"text": ln, "heading": ln, "level": 2})
                else:
                    buf.append(ln)
            if buf:
                paragraphs.append({"text": " ".join(buf), "heading": None, "level": 0})
    return paragraphs


def _parse_md(path):
    with open(path, encoding="utf-8") as f:
        lines = f.read().splitlines()
    out = []
    for line in lines:
        text = line.strip()
        if not text:
            continue
        m = re.match(r"^(#{1,6})\s+(.*)", text)
        if m:
            out.append({"text": m.group(2).strip(), "heading": m.group(2).strip(),
                        "level": len(m.group(1))})
        else:
            out.append({"text": text, "heading": None, "level": 0})
    return out

# ---------------------------------------------------------------------------
# 2. 参考文献列表提取与条目解析
# ---------------------------------------------------------------------------

REF_HEADINGS = re.compile(
    r"^(references?|reference list|bibliography|works cited|参考文献|文献列表)\s*$",
    re.IGNORECASE)


def split_references(paragraphs):
    """定位 References 标题，返回 (条目列表, 标题所在段索引)。"""
    ref_para_idx = None
    for i, p in enumerate(paragraphs):
        if p["heading"] and REF_HEADINGS.match(p["heading"].strip()):
            ref_para_idx = i
            break
        # md/txt 里标题可能只是普通段落
        if p["heading"] is None and REF_HEADINGS.match(p["text"].strip()):
            ref_para_idx = i
            break
    if ref_para_idx is None:
        return None, None
    tail = paragraphs[ref_para_idx + 1:]
    entries = []
    current = []
    for p in tail:
        t = p["text"]
        # In some DOCX manuscripts, tables and figures follow the reference
        # list in the same document.  They are not reference entries.
        if re.match(r"^(?:Table|Figure)\s+\d+\b", t, re.IGNORECASE):
            break
        # 新条目开头: 大写字母开头 + 含 4 位年份（author-year style 通用特征）
        if _looks_like_entry_start(t) and current:
            entries.append(" ".join(current))
            current = [t]
        else:
            current.append(t)
    if current:
        entries.append(" ".join(current))
    return entries, ref_para_idx


_AUTHOR_START = re.compile(r"^[A-Za-zÀ-ÿ一-鿿][\w'’\-,. &]+")


def _looks_like_entry_start(t):
    return bool(_AUTHOR_START.match(t) and re.search(r"\d{4}", t) and len(t) > 20)


# DOI suffixes may legitimately contain parentheses (for example,
# ``10.1016/0147-1767(86)90007-6``); stop at whitespace rather than treating
# the first closing parenthesis as the end of the DOI.
DOI_RE = re.compile(r"(10\.\d{4,9}/[^\s\"'<>\]]+)")


def parse_entry(raw, idx):
    """把一条参考文献解析成结构化字段。尽力而为的启发式，失败字段置 None。"""
    raw = re.sub(r"\s+", " ", raw).strip().rstrip(".")
    e = {"id": f"R{idx}", "raw": raw, "authors": None, "year": None,
         "title": None, "venue": None, "volume": None, "issue": None,
         "pages": None, "doi": None, "parse_ok": True}

    m = DOI_RE.search(raw)
    if m:
        e["doi"] = m.group(1).rstrip(".")

    # 年份: 优先 (2020) / (2020a)，其次任意 4 位年份
    ym = re.search(r"[ (.](\d{4})([a-z]?)[).,]", raw + " ")
    if ym:
        e["year"] = int(ym.group(1))
        e["year_suffix"] = ym.group(2)
    else:
        ym2 = re.search(r"\b(19|20)\d{2}\b", raw)
        if ym2:
            e["year"] = int(ym2.group(0))
        ym = ym2  # 用于下面的切分定位

    # 作者: 年份之前的部分
    if ym:
        authors_str = raw[:ym.start()].rstrip(" ,(.")
    else:
        authors_str = raw.split(".")[0]
    # 过滤掉缩写名（"A. B."无连续小写字母），只留姓氏
    e["authors"] = [a.strip() for a in re.split(r",|&|;| and ", authors_str)
                    if re.search(r"[A-ZÀ-Ÿ][a-zà-ÿ]{1,}", a.strip())]

    # 标题/书目边界: 先剥 URL（R52 教训），再在最左侧找卷期页模式，然后
    # 回溯到它之前最后一个句末终止符（.?!）。
    # 问号既可能是标题结尾（其后是期刊名，R16: "...open? Academy of
    # Management Journal, 50(4)..."），也可能是标题内部设问（其后还有
    # 副标题，R37: "Who gets credit for input? Demographic and structural
    # status..."）——以卷期模式的位置为准，不能以第一个终止符为准。
    rest = raw[ym.end():] if ym else raw[min(len(authors_str), len(raw)):]
    rest = rest.lstrip("). ").strip()
    rest_nourl = re.sub(r"https?://\S*", " ", rest)

    vm = None
    for pat in (r"[, ]\d+\s*\(\d+\)",      # 50(4)
                r"[, ]\d+\s*,\s*\d+\s*[-–—]",  # 28, 285-305（无期号）
                r"[, ]\d+\s*:"):            # 28: 1-15
        m = re.search(pat, rest_nourl)
        if m and (vm is None or m.start() < vm.start()):
            vm = m

    if vm:
        before = rest_nourl[:vm.start()].rstrip()
        terms = list(re.finditer(r"[.?!](?=\s+[A-ZÀ-Ÿ])|[.?!]$", before))
        if terms:
            t = terms[-1]
            term_char = t.group(0)[0]  # 问句式标题保留结尾的 "?"（APA 习惯）
            e["title"] = before[:t.start()].strip() + (
                term_char if term_char in "?!" else "")
            venue_part = before[t.end():]
        else:
            e["title"] = before.strip(" ,.")
            venue_part = ""
        bib = rest_nourl[vm.start():]
        mv = re.match(r"[, ](\d+)\s*\((\d+)\)", bib)
        if mv:
            e["volume"], e["issue"] = mv.group(1), mv.group(2)
        else:
            mv2 = re.match(r"[, ](\d+)\s*[,:]", bib)
            if mv2:
                e["volume"] = mv2.group(1)
        e["venue"] = venue_part.strip(" ,.") or None
    else:
        # 无卷期页模式（书籍、章节等）: 标题到第一个句末终止符
        title_m = re.match(r"^(.*?)(?:[.?!](?:\s+(?=[A-ZÀ-Ÿ])|$))",
                           rest_nourl)
        e["title"] = (title_m.group(1).strip() if title_m
                      else rest_nourl.split(".")[0]).strip()
        if title_m and title_m.group(0)[-1:] in ("?", "!"):
            e["title"] += title_m.group(0)[-1]
        venue_part = rest_nourl[len(e["title"]):]
        v = DOI_RE.sub("", venue_part).strip(" ,.?! ")
        e["venue"] = v if 0 < len(v) < 120 else None
    if len(e["title"]) < 8:
        e["title"] = rest_nourl[:120]
        e["parse_ok"] = False  # 标题解析可疑，让 Claude 兜底
    pm = re.search(r"(\d+\s*[-–—]\s*\d+)", rest_nourl)
    if pm:
        e["pages"] = re.sub(r"\s*", "", pm.group(1)).replace("–", "-").replace("—", "-")

    if not e["authors"] or e["year"] is None:
        e["parse_ok"] = False
    return e

# ---------------------------------------------------------------------------
# 3. 正文引用提取 + A/B/C 分诊
# ---------------------------------------------------------------------------

PAREN_CITE = re.compile(r"[(（]([^()（）]*?\d{4}[^()（）]*?)[)）]")
NARR_CITE = re.compile(
    r"([A-ZÀ-ÿ][\w'’\-]+(?:\s+(?:et al\.?|&|and)\s+[A-ZÀ-ÿ][\w'’\-]+|\s+et al\.?)*)"
    r"\s*[(（]\s*(\d{4}[a-z]?)")
CITE_SPLIT = re.compile(r"[;；]")
# 括号内单条引用: "Author, 2020" / "Author & B, 2020" / "Author et al., 2020, 2021"
INNER_CITE = re.compile(r"^(.*?),?\s*((?:\d{4}[a-z]?(?:\s*,\s*\d{4}[a-z]?)*))$")

HYPOTHESIS_MARKERS = re.compile(
    r"hypothes|we predict|we propose|we expect|we argue|theory suggest|"
    r"following [A-Z][\w'’-]+'?s? (theory|model|framework)|"
    r"based on .*?(theory|model|framework)|\bH\d\b|假设", re.IGNORECASE)


def split_sentences(text):
    return re.split(r"(?<=[.!?])\s+(?=[A-ZÀ-ÿ一-鿿\"'])", text)


def extract_citations(body_paragraphs):
    """返回 list[citation dict]。body_paragraphs 不含参考文献列表部分。"""
    citations = []
    section = "正文"
    for paragraph_index, p in enumerate(body_paragraphs, 1):
        if p["heading"]:
            section = p["heading"]
            continue
        text = p["text"]
        for sent in split_sentences(text):
            found = []
            for m in PAREN_CITE.finditer(sent):
                inner = m.group(1)
                parts = CITE_SPLIT.split(inner)
                for part in parts:
                    part = part.strip()
                    # 直接引语的页码/章节后缀不算作者或年份:
                    # (Galinsky & Moskowitz, 2000, p. 710) 曾因 "p. 710"
                    # 后缀整体匹配失败而被误报为"正文引了但列表没有"
                    part = re.sub(
                        r",?\s*(?:pp?\.|chap\.?|chapter)\s*[\w.\-–\s]*$",
                        "", part, flags=re.IGNORECASE)
                    im = INNER_CITE.match(part)
                    if not im:
                        continue
                    author_part, year_part = im.group(1).strip(), im.group(2)
                    years = re.findall(r"\d{4}[a-z]?", year_part)
                    surnames = _surnames_from_inline(author_part)
                    if not surnames:
                        # Do not treat standalone years or registration IDs
                        # (e.g., AsPredicted #237396) as author-year citations.
                        continue
                    for y in years:
                        found.append({"authors": surnames, "year": y,
                                      "raw": part, "narrative": False,
                                      "group_size": len(parts)})
            for m in NARR_CITE.finditer(sent):
                surnames = _surnames_from_inline(m.group(1))
                found.append({"authors": surnames, "year": m.group(2),
                              "raw": f"{m.group(1)} ({m.group(2)})", "narrative": True,
                              "group_size": 1})
            for c in found:
                c["sentence"] = sent
                c["section"] = section
                c["locator"] = p.get("locator") or f"正文段落 {paragraph_index}"
                citations.append(c)

    # 承重角色优先；并引数量只作分组提示，C 类仍需轻查。
    for c in citations:
        if HYPOTHESIS_MARKERS.search(c["sentence"]):
            c["triage"] = "A"
        elif c["group_size"] >= 3:
            c["triage"] = "C"
        else:
            c["triage"] = "B"
    return citations


def _surnames_from_inline(s):
    """'Bakker & Demerouti' / 'Bakker et al.' / 'Bakker, Demerouti, & Xie' → 姓氏列表"""
    s = re.sub(r"\bet al\.?", "", s)
    # 注意: & 不是单词字符，不能用 \b&\b 匹配
    s = re.sub(r"\s*&\s*|;|,", ",", s)
    s = re.sub(r"\s+and\s+", ",", s, flags=re.IGNORECASE)
    # 只去掉开头的引导词（see Smith, 2020），避免把作者 "See" 误杀
    s = re.sub(r"^(?:see|e\.g\.,?|cf\.?)\s*,?\s*", "", s.strip(), flags=re.IGNORECASE)
    names = []
    for part in s.split(","):
        part = re.sub(r"[’']s$", "", part.strip(" ."), flags=re.IGNORECASE)
        words = [w for w in part.split() if w]
        if not words:
            continue
        surname = words[-1].strip(".")
        if re.match(r"^[A-ZÀ-ÿ]", surname) and not re.fullmatch(r"(19|20)\d{2}[a-z]?", surname):
            names.append(surname)
    return names

# ---------------------------------------------------------------------------
# 4. OpenAlex / Crossref 验证（带缓存 + 限速退避）
# ---------------------------------------------------------------------------

def _env_any(*names):
    """大小写不敏感地读环境变量（用户可能写成 OpenAlex_API_KEY 等）。"""
    for n in names:
        v = os.environ.get(n)
        if v:
            return v
    for k, v in os.environ.items():
        if k.upper() in names:
            return v
    return None


class Verifier:
    def __init__(self, offline=False, cache_dir=CACHE_DIR):
        self.offline = offline
        self.cache_dir = cache_dir
        os.makedirs(cache_dir, exist_ok=True)
        self.openalex_key = _env_any("OPENALEX_API_KEY")
        self.s2_key = _env_any("SEMANTIC_API_KEY", "SEMANTIC_SCHOLAR_API_KEY",
                               "S2_API_KEY")
        # 只记录能力等级，不写入或输出任何密钥。无 Key 时仍使用前两家的公开接口。
        self.source_capabilities = {
            "openalex": "key" if self.openalex_key else "public",
            "crossref": "public",
            "semantic_scholar": "key" if self.s2_key else "not_configured",
        }
        self.stats = {"cache_hit": 0, "api_openalex": 0, "api_crossref": 0,
                      "api_s2": 0, "failed": 0}
        self._openalex_dead = False  # 熔断: 连续 429 后本次运行不再请求 OpenAlex
        self._crossref_dead = False
        self._s2_dead = False
        self._rotate = 0  # 每条文献轮换首选源，分摊三家日额度

    # -- HTTP --
    def _get(self, url, headers=None):
        last_err = None
        for attempt in range(4):
            try:
                hdrs = {"User-Agent": USER_AGENT}
                if headers:
                    hdrs.update(headers)
                req = urllib.request.Request(url, headers=hdrs)
                with urllib.request.urlopen(req, timeout=30) as resp:
                    return json.loads(resp.read().decode("utf-8"))
            except urllib.error.HTTPError as ex:
                last_err = ex
                if ex.code == 404:
                    raise  # 资源确定不存在，重试无意义（DOI 直查未知 DOI 时常见）
                if ex.code in (429, 403):
                    time.sleep(3)  # 限流: 短退避即可，熔断逻辑会止损
                else:
                    time.sleep(2 ** attempt)
            except Exception as ex:
                last_err = ex
                time.sleep(2 ** attempt)
        raise last_err

    # -- cache --
    def _cache_path(self, key):
        h = hashlib.sha1(key.encode("utf-8")).hexdigest()
        return os.path.join(self.cache_dir, h + ".json")

    def _cache_get(self, key):
        p = self._cache_path(key)
        if os.path.exists(p):
            try:
                with open(p, encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                return None
        return None

    def _cache_put(self, key, obj):
        p = self._cache_path(key)
        with open(p, "w", encoding="utf-8") as f:
            json.dump({"cached_at": datetime.date.today().isoformat(),
                       "data": obj}, f, ensure_ascii=False)

    # -- public --
    def verify(self, entry):
        """返回 {status, confidence, source, record, mismatches, links, abstract}"""
        title = entry.get("title") or ""
        # 缓存文献记录，不复用相对于旧稿的字段差异。身份变化重新检索。
        identity = {"title": title.casefold(), "doi": (entry.get("doi") or "").lower(),
                    "authors": entry.get("authors")}
        key = "v6:" + json.dumps(identity, ensure_ascii=False, sort_keys=True)
        cached = self._cache_get(key)
        if cached and (self.offline or not self._stale(cached)):
            self.stats["cache_hit"] += 1
            result = dict(cached["data"])
            if result.get("status") == "found" and result.get("record"):
                result = self._found_result(entry, result["record"])
            result.setdefault("mismatches", [])
            result.setdefault("links", self._search_links(entry))
            if self._stale(cached):
                result["cache_stale"] = True
                result["note"] = "离线使用过期缓存，须重新核实"
            return result
        if self.offline:
            return self._unverified(entry, "离线模式且无缓存")
        result = self._verify_online(entry, title)
        if result.get("status") != "error":
            cache_data = dict(result)
            for field in ("mismatches", "confidence", "links"):
                cache_data.pop(field, None)
            self._cache_put(key, cache_data)
        return result

    @staticmethod
    def _stale(cached):
        try:
            age = (datetime.date.today()
                   - datetime.date.fromisoformat(cached["cached_at"])).days
        except (KeyError, ValueError, TypeError):
            return True
        ttl = 180 if cached.get("data", {}).get("status") == "found" else 1
        return age > ttl

    def _found_result(self, entry, rec):
        mismatches = _compare_metadata(entry, rec)
        return {"status": "found", "confidence": "medium" if mismatches else "high",
                "source": rec.get("_source"), "record": rec,
                "mismatches": mismatches, "links": self._record_links(rec, entry),
                "abstract": rec.get("abstract")}

    def _verify_online(self, entry, title):
        if not title or len(title) < 8:
            return self._unverified(entry, "标题解析失败，无法检索")
        # 轮换首选源（大论文上百条文献时把额度摊到三家），熔断的源返回 None 自动跳过
        searchers = [self._search_openalex, self._search_crossref, self._search_s2]
        start = self._rotate % len(searchers)
        self._rotate += 1
        rec = None
        for s in searchers[start:] + searchers[:start]:
            rec = s(title, entry)
            if rec is not None:
                break
        if rec is None:
            # DOI 直查兜底（C6①）: 检索噪音/解析问题不应把带有效 DOI 的条目
            # 判为未匹配。若 DOI 查到的标题与论文标题对不上，说明 DOI 指向
            # 另一篇文献（疑似错挂），保留 not_found 但附明确线索。
            if entry.get("doi"):
                rec = self._crossref_doi_lookup(entry["doi"])
                if rec is not None:
                    t = (rec.get("title") or "").lower()
                    tl = title.lower()
                    # 数据库标题常是短版（无副标题）或长版（含副标题），
                    # 纯比例打分会把同一篇文献判成两篇（R7 教训:
                    # "Looking Out From the Top" vs 全标题 ratio≈0.45）。
                    # 比例 + 双向 containment 任一命中即视为同一篇。
                    same = (difflib.SequenceMatcher(None, tl, t).ratio() >= 0.6
                            or t in tl or tl in t)
                    if not same:
                        return {"status": "not_found", "confidence": "low",
                                "source": "crossref-doi", "record": rec,
                                "mismatches": [], "links": self._search_links(entry),
                                "abstract": None,
                                "note": "论文所写 DOI 指向另一篇文献（标题不符），"
                                        "疑似 DOI 错挂；必须人工复核"}
        if rec is None:
            # 公共索引的覆盖、检索限流和条目解析都会造成漏检；自动未匹配不是
            # “文献不存在”的证据，必须由 Skill 层二次复核后才能形成最终结论。
            return {"status": "not_found", "confidence": "low",
                    "source": "openalex+crossref+semanticscholar", "record": None,
                    "mismatches": [], "links": self._search_links(entry),
                    "abstract": None,
                    "note": "自动检索未匹配，必须人工复核；不能据此断言为编造或真实缺失"}
        return self._found_result(entry, rec)

    def _search_openalex(self, title, entry):
        if self._openalex_dead:
            return None
        q = urllib.parse.quote(title)
        url = (f"https://api.openalex.org/works?filter=title.search:{q}"
               f"&per-page=5&select=id,doi,title,display_name,publication_year,"
               f"authorships,primary_location,biblio,type,abstract_inverted_index")
        if self.openalex_key:
            url += f"&api_key={self.openalex_key}"
        try:
            self.stats["api_openalex"] += 1
            data = self._get(url)
            self._openalex_429_streak = 0
        except urllib.error.HTTPError as ex:
            self.stats["failed"] += 1
            if ex.code in (429, 403):
                streak = getattr(self, "_openalex_429_streak", 0) + 1
                self._openalex_429_streak = streak
                if streak >= 2:
                    # 熔断: IP 级限流/日限额，重试无意义，本次运行走 Crossref
                    self._openalex_dead = True
            return None
        except Exception:
            self.stats["failed"] += 1
            return None
        best, best_score = None, 0.0
        for w in data.get("results", []):
            t = (w.get("display_name") or w.get("title") or "")
            score = difflib.SequenceMatcher(None, title.lower(), t.lower()).ratio()
            if score > best_score:
                best, best_score = w, score
        if best is None or best_score < 0.75:
            return None
        src = ""
        bib = best.get("biblio") or {}
        if best.get("primary_location") and best["primary_location"].get("source"):
            src = best["primary_location"]["source"].get("display_name") or ""
        rec = {
            "_source": "openalex",
            "id": best.get("id"),
            "title": best.get("display_name") or best.get("title"),
            "year": best.get("publication_year"),
            "authors": [a.get("author", {}).get("display_name", "")
                        for a in best.get("authorships", [])],
            "venue": src,
            "volume": bib.get("volume") or None,
            "issue": bib.get("issue") or None,
            "pages": (f"{bib['first_page']}-{bib['last_page']}"
                      if bib.get("first_page") and bib.get("last_page")
                      else bib.get("first_page")),
            "doi": (best.get("doi") or "").replace("https://doi.org/", "") or None,
            "type": best.get("type"),
            "similarity": round(best_score, 3),
            "abstract": _invert_index_to_text(best.get("abstract_inverted_index")),
        }
        return rec

    def _search_crossref(self, title, entry):
        if self._crossref_dead:
            return None
        q = urllib.parse.quote(entry["raw"][:500])
        url = f"https://api.crossref.org/works?query.bibliographic={q}&rows=5"
        try:
            self.stats["api_crossref"] += 1
            data = self._get(url)
        except urllib.error.HTTPError as ex:
            self.stats["failed"] += 1
            if ex.code in (429, 403):
                streak = getattr(self, "_crossref_429_streak", 0) + 1
                self._crossref_429_streak = streak
                if streak >= 2:
                    self._crossref_dead = True
            return None
        except Exception:
            self.stats["failed"] += 1
            return None
        best, best_score = None, 0.0
        for w in data.get("message", {}).get("items", []):
            t = (w.get("title") or [""])[0]
            score = difflib.SequenceMatcher(None, title.lower(), t.lower()).ratio()
            if score > best_score:
                best, best_score = w, score
        if best is None or best_score < 0.75:
            return None
        cr = best.get("container-title") or [""]
        rec = {
            "_source": "crossref",
            "title": (best.get("title") or [""])[0],
            "year": _crossref_year(best),
            "authors": [a.get("family", "") for a in best.get("author", [])],
            "venue": cr[0] if cr else "",
            "volume": best.get("volume") or None,
            "issue": best.get("issue") or None,
            "pages": best.get("page") or None,
            "doi": best.get("DOI"),
            "type": best.get("type"),
            "similarity": round(best_score, 3),
            "abstract": _strip_html(best.get("abstract") or ""),
        }
        return rec

    def _crossref_doi_lookup(self, doi):
        """Crossref DOI 直查（C6①: title 检索失败时的兜底，比检索式可靠）。"""
        if self._crossref_dead:
            return None
        url = f"https://api.crossref.org/works/{urllib.parse.quote(doi)}"
        try:
            self.stats["api_crossref"] += 1
            data = self._get(url)
        except Exception:
            self.stats["failed"] += 1
            return None
        item = data.get("message") or {}
        t = (item.get("title") or [""])[0]
        if not t:
            return None
        cr = item.get("container-title") or [""]
        return {
            "_source": "crossref-doi",
            "title": t,
            "year": _crossref_year(item),
            "authors": [a.get("family", "") for a in item.get("author", [])],
            "venue": cr[0] if cr else "",
            "volume": item.get("volume") or None,
            "issue": item.get("issue") or None,
            "pages": item.get("page") or None,
            "doi": item.get("DOI"),
            "type": item.get("type"),
            "similarity": 1.0,
            "abstract": _strip_html(item.get("abstract") or ""),
        }

    def _search_s2(self, title, entry):
        """Semantic Scholar（需要 API key；免费池太容易被限流，无 key 不启用）。"""
        if not self.s2_key or self._s2_dead:
            return None
        q = urllib.parse.quote(title)
        url = ("https://api.semanticscholar.org/graph/v1/paper/search?"
               f"query={q}&limit=5&fields=title,year,authors,venue,journal,"
               f"externalIds,abstract")
        try:
            self.stats["api_s2"] += 1
            data = self._get(url, headers={"x-api-key": self.s2_key})
        except urllib.error.HTTPError as ex:
            self.stats["failed"] += 1
            if ex.code in (429, 403):
                streak = getattr(self, "_s2_429_streak", 0) + 1
                self._s2_429_streak = streak
                if streak >= 2:
                    self._s2_dead = True
            return None
        except Exception:
            self.stats["failed"] += 1
            return None
        best, best_score = None, 0.0
        for w in data.get("data") or []:
            t = w.get("title") or ""
            score = difflib.SequenceMatcher(None, title.lower(), t.lower()).ratio()
            if score > best_score:
                best, best_score = w, score
        if best is None or best_score < 0.75:
            return None
        jr = best.get("journal") or {}
        rec = {
            "_source": "s2",
            "id": best.get("paperId"),
            "title": best.get("title"),
            "year": best.get("year"),
            "authors": [a.get("name", "") for a in best.get("authors", [])],
            "venue": best.get("venue") or jr.get("name") or "",
            "volume": jr.get("volume") or None,
            "issue": None,
            "pages": jr.get("pages") or None,
            "doi": (best.get("externalIds") or {}).get("DOI"),
            "type": "journal-article",
            "similarity": round(best_score, 3),
            "abstract": best.get("abstract") or "",
        }
        return rec

    @staticmethod
    def _unverified(entry, reason):
        return {"status": "unverified", "confidence": "none", "source": None,
                "record": None, "mismatches": [], "links": [],
                "abstract": None, "note": reason}

    @staticmethod
    def _search_links(entry):
        q = urllib.parse.quote((entry.get("title") or entry["raw"])[:200])
        return {"openalex_search": f"https://openalex.org/works?search={q}",
                "google_scholar": f"https://scholar.google.com/scholar?q={q}"}

    @staticmethod
    def _record_links(rec, entry):
        links = {}
        if rec.get("doi"):
            links["doi"] = f"https://doi.org/{rec['doi']}"
        if rec.get("_source") == "openalex" and rec.get("id"):
            links["openalex"] = rec["id"]
        if rec.get("_source") == "s2" and rec.get("id"):
            links["s2"] = f"https://www.semanticscholar.org/paper/{rec['id']}"
        if "openalex" not in links:
            q = urllib.parse.quote(rec.get("title") or "")
            links["google_scholar"] = f"https://scholar.google.com/scholar?q={q}"
        return links


def _crossref_year(item):
    for k in ("published-print", "published-online", "issued"):
        d = item.get(k, {}).get("date-parts", [[]])[0]
        if d and d[0]:
            return d[0]
    return None


def _invert_index_to_text(idx):
    if not idx:
        return None
    pos = {}
    for word, positions in idx.items():
        for p in positions:
            pos[p] = word
    return " ".join(pos[k] for k in sorted(pos))


def _strip_html(s):
    return re.sub(r"<[^>]+>", "", s).strip() or None


PREPRINT_MARKERS = re.compile(
    r"arxiv|psyarxiv|biorxiv|medrxiv|ssrn|osf\.io|preprint", re.IGNORECASE)


VENUE_STOPWORDS = {"the", "of", "and", "for", "in", "a", "an", "on"}


def _venue_words(v):
    v = re.sub(r"[^a-z0-9]+", " ", v.lower())
    return [w for w in v.split() if w not in VENUE_STOPWORDS]


def _abbr_word(short, long):
    """short 是 long 的缩写（首字母/截断前缀）吗？单复数屈折不算缩写。"""
    if short == long:
        return True
    if not short or not long.startswith(short):
        return False
    if long in (short + "s", short + "es"):
        # review → reviews 属拼写出入而非缩写（R52 教训），不静默放行
        return False
    return True


def _venue_equivalent(paper, db):
    """期刊名词级等价判断: 全词相等，或一侧是另一侧的逐词缩写前缀。"""
    wp, wd = _venue_words(paper), _venue_words(db)
    if not wp or not wd:
        return False
    if wp == wd:
        return True
    short, long_ = (wp, wd) if len(wp) <= len(wd) else (wd, wp)
    i = 0
    for w in long_:
        if i < len(short) and _abbr_word(short[i], w):
            i += 1
    return i == len(short)


def _compare_metadata(entry, rec):
    """逐项比对，返回 mismatch 列表 [{field, paper, database, level?}]

    level 含义: 无 level = 明显差异; "near" = 高度相似（疑似缩写变体或
    单字符拼写出入）——统一交给 AI 层裁决，不做静默容差（F2 教训:
    容差规则在代码里替用户做语义判断，会吞掉真错误）。
    """
    out = []

    def add(field, paper_val, db_val, threshold=0.75):
        p = (paper_val or "").strip() if isinstance(paper_val, str) else paper_val
        d = (db_val or "").strip() if isinstance(db_val, str) else db_val
        if p in (None, "") and d in (None, ""):
            return
        if p in (None, "") or d in (None, ""):
            if field in ("volume", "issue", "pages", "doi"):
                return  # 论文漏写这些字段不算错
            out.append({"field": field, "paper": p, "database": d})
            return
        if isinstance(p, int) or isinstance(d, int):
            if p != d:
                out.append({"field": field, "paper": p, "database": d})
            return
        pl, dl = str(p).lower(), str(d).lower()
        if field == "venue":
            if _venue_equivalent(pl, dl):
                return
            sim = difflib.SequenceMatcher(None, pl, dl).ratio()
            mm = {"field": field, "paper": p, "database": d}
            if sim < threshold:
                out.append(mm)
            else:
                mm["level"] = "near"
                out.append(mm)
            return
        if field == "first_author":
            # 姓氏无缩写惯例，精确匹配；"Li" ⊂ "Lin" 的子串容差会吞掉真错误
            if pl != dl:
                out.append({"field": field, "paper": p, "database": d})
            return
        if field in ("volume", "issue"):
            if pl.isdigit() and dl.isdigit():
                same = int(pl) == int(dl)
            else:
                same = pl == dl
            if not same:
                out.append({"field": field, "paper": p, "database": d})
            return
        if pl in dl or dl in pl:
            return  # 简称/缩写容差（如 "Human Factors" ⊂ 数据库全称）
        sim = difflib.SequenceMatcher(None, pl, dl).ratio()
        if sim < threshold:
            out.append({"field": field, "paper": p, "database": d})

    add("year", entry["year"], rec["year"])
    normalize_title = lambda text: re.sub(r"[\W_]+", " ", str(text or "").casefold()).strip()
    if entry.get("title") and rec.get("title"):
        if normalize_title(entry["title"]) != normalize_title(rec["title"]):
            out.append({"field": "title", "paper": entry["title"],
                        "database": rec["title"], "level": "near"})
    # 数据库可能提供全名；姓氏序列的差异交人工确认，不能直接判错。
    surname = lambda name: re.sub(r"\W+", "", name.split()[-1].casefold()) if name.split() else ""
    pa = [surname(a) for a in entry.get("authors") or []]
    da = [surname(a) for a in rec.get("authors") or []]
    if pa and da and pa != da:
        out.append({"field": "authors", "paper": entry["authors"],
                    "database": rec["authors"], "level": "near"})
    # 第一作者姓氏
    first_db = (rec.get("authors") or [""])[0].split()[-1] if rec.get("authors") else ""
    first_paper = entry["authors"][0].split()[-1] if entry.get("authors") else ""
    add("first_author", first_paper, first_db)
    add("venue", entry["venue"], rec["venue"], threshold=0.6)
    add("volume", entry["volume"], rec["volume"])
    add("issue", entry["issue"], rec["issue"])
    if entry.get("pages") and rec.get("pages") and "-" in str(rec["pages"]):
        # 数据库只有起始页（无范围）时不比较，避免误报
        norm = lambda x: re.sub(r"[^0-9]", "", str(x))
        if norm(entry["pages"]) != norm(rec["pages"]):
            out.append({"field": "pages", "paper": entry["pages"],
                        "database": rec["pages"]})
    if entry.get("doi") and rec.get("doi"):
        if entry["doi"].lower().rstrip(".") != rec["doi"].lower().rstrip("."):
            out.append({"field": "doi", "paper": entry["doi"], "database": rec["doi"]})
    return out

# ---------------------------------------------------------------------------
# 5. 机械检查: 对应 / 重复 / 时间线 / preprint
# ---------------------------------------------------------------------------

def _author_key(name):
    return re.sub(r"\W+", "", name.split()[-1].casefold()) if name.split() else ""


def entry_key(e):
    first = (e.get("authors") or [""])[0]
    year = str(e.get("year") or "") + e.get("year_suffix", "")
    return (_author_key(first), year)


def check_correspondence(entries, citations):
    cited_ids, uncertain_ids = set(), set()
    unmatched, ambiguous = [], []
    for i, c in enumerate(citations):
        c["cid"] = f"C{i + 1}"
        c["ref_ids"] = []
        c["candidate_ref_ids"] = []
        first = _author_key((c.get("authors") or [""])[0])
        year = str(c.get("year") or "")
        candidates = [e for e in entries if entry_key(e) == (first, year)]
        # 未写后缀时保留候选，不把同年多条一并标作已引用。
        if not candidates and year.isdigit():
            candidates = [e for e in entries if entry_key(e)[0] == first
                          and str(e.get("year")) == year]
        if len(candidates) > 1 and len(c.get("authors", [])) > 1:
            authors = [_author_key(a) for a in c["authors"]]
            narrowed = [e for e in candidates if [_author_key(a) for a in e.get("authors", [])] == authors]
            if narrowed:
                candidates = narrowed
        if len(candidates) == 1:
            c["ref_ids"] = [candidates[0]["id"]]
            cited_ids.update(c["ref_ids"])
        elif candidates:
            c["candidate_ref_ids"] = [e["id"] for e in candidates]
            uncertain_ids.update(c["candidate_ref_ids"])
            ambiguous.append(c)
        else:
            unmatched.append(c)
    return {"cited_but_missing_in_list": unmatched,
            "ambiguous_citations": ambiguous,
            "listed_but_never_cited": [e["id"] for e in entries
                                       if e["id"] not in cited_ids | uncertain_ids]}


def check_duplicates(entries):
    dups = []
    seen_titles = {}
    seen_dois = {}
    for e in entries:
        tkey = re.sub(r"\W+", "", (e["title"] or "").lower())[:60]
        if len(tkey) > 20 and tkey in seen_titles:
            dups.append({"ids": [seen_titles[tkey], e["id"]], "by": "title"})
        elif len(tkey) > 20:
            seen_titles[tkey] = e["id"]
        if e["doi"]:
            dk = e["doi"].lower().rstrip(".")
            if dk in seen_dois:
                dups.append({"ids": [seen_dois[dk], e["id"]], "by": "doi"})
            else:
                seen_dois[dk] = e["id"]
    return dups


def check_timeline(entries, today=None):
    today = today or datetime.date.today()
    issues = []
    for e in entries:
        if e["year"] and e["year"] > today.year:
            issues.append({"id": e["id"],
                           "issue": f"出版年份 {e['year']} 晚于当前年份，需核实是否为预排卷期或年份笔误"})
    return issues


def check_preprint(entries):
    issues = []
    for e in entries:
        venue = e["venue"] or ""
        if PREPRINT_MARKERS.search(venue):
            issues.append({"id": e["id"],
                           "issue": f"引用的是 preprint（{venue}）——正式版可能已发表，建议更新"})
    return issues


def check_doi_swaps(entries, results):
    """列表内 DOI 错挂/互换检测（C1, R55/R56 教训）。

    逐条比对发现不了"A 的 DOI 指向 B、B 的 DOI 指向 A"这类列表内部
    一致性错误——单条各自与数据库比对只会显示为普通 DOI 差异。
    """
    paper_doi, db_doi = {}, {}
    for e in entries:
        paper_doi[e["id"]] = (e.get("doi") or "").lower().rstrip(".")
        rec = (results.get(e["id"], {}) or {}).get("record") or {}
        db_doi[e["id"]] = (rec.get("doi") or "").lower().rstrip(".")
    by_db = {v: k for k, v in db_doi.items() if v}
    issues, seen = [], set()
    for a, pdoi in paper_doi.items():
        if not pdoi or pdoi == db_doi.get(a):
            continue
        b = by_db.get(pdoi)
        if not b or b == a:
            continue
        key = tuple(sorted((a, b)))
        if key in seen:
            continue
        seen.add(key)
        if paper_doi.get(b) and paper_doi.get(b) == db_doi.get(a):
            issues.append({"ids": [a, b], "level": "warn",
                           "issue": f"疑似 DOI 互换错挂：{a} 所写 DOI 实为 {b} "
                                    f"的文献，{b} 所写 DOI 实为 {a} 的文献"})
        else:
            issues.append({"ids": [a, b], "level": "warn",
                           "issue": f"疑似 DOI 错挂：{a} 所写 DOI 实为 {b} 的文献"})
    return issues


def check_ordering(entries):
    """排序一致性检查（C3, R55/R56 教训）。

    APA 规则: 完整作者名单相同的多条才按年份升序；同第一作者但合作者
    不同的按第二作者字母序——后者不做机械判断（易误报），只查前者。
    """
    issues, groups = {}, {}
    for e in entries:
        if not e.get("authors") or e.get("year") is None:
            continue
        key = tuple(re.sub(r"\W+", "", a.split()[-1].lower())
                    for a in e["authors"])
        groups.setdefault(key, []).append(e)
    for items in groups.values():
        if len(items) < 2:
            continue
        years = [i["year"] for i in items]
        if years != sorted(years):
            seq = " → ".join(f"{i['id']}({i['year']})" for i in items)
            issues_key = items[0]["authors"][0]
            issues[issues_key] = {
                "ids": [i["id"] for i in items], "level": "info",
                "issue": f"同一作者组（{', '.join(items[0]['authors'])}）的多条文献"
                         f"未按年份升序排列: {seq}"}
    return list(issues.values())


def check_title_artifacts(entries):
    """标题以「数字+空格」开头 → 疑似章节号残留（C4, R49 教训）。"""
    issues = []
    for e in entries:
        if e.get("title") and re.match(r"^\d{1,3}\s+\S", e["title"]):
            issues.append({"id": e["id"], "level": "info",
                           "issue": f"标题以数字开头（{e['title'][:30]}…），"
                                    f"疑似章节号残留，建议核对"})
    return issues

# ---------------------------------------------------------------------------
# 6. 报告生成（独立 HTML，自包含样式，浏览器直接打开）
# ---------------------------------------------------------------------------

STATUS_ICON = {"found": "✅", "not_found": "⚠️", "unverified": "❓"}

REPORT_CSS = """
:root {
  --ok: #166534; --bad: #a12a2a; --warn: #9a5a10; --info: #285f8f;
  --gray: #64706f; --ink: #172321; --muted-ink: #44514e; --border: #d9ddd6;
  --bg: #f3f2ed; --paper: #fffefa; --tint: #f7f7f2; --nav-h: 48px;
  --serif: Georgia, "Times New Roman", "Songti SC", "Noto Serif SC", serif;
  --mono: ui-monospace, "SF Mono", SFMono-Regular, Menlo, monospace;
}
* { box-sizing: border-box; margin: 0; padding: 0; }
@media (prefers-reduced-motion: no-preference) {
  html { scroll-behavior: smooth; }
}
body {
  font-family: -apple-system, "PingFang SC", "Microsoft YaHei", "Segoe UI", sans-serif;
  background: var(--bg); color: var(--ink); font-size: 14px; line-height: 1.78;
  padding: 36px 20px 72px;
}
.container { max-width: 1120px; margin: 0 auto; }
header.paper { border-top: 4px solid #315e55; padding: 20px 0 25px; margin-bottom: 12px; }
.eyebrow { color: #315e55; font-size: 11px; font-weight: 700; letter-spacing: .13em;
  text-transform: uppercase; margin-bottom: 9px; }
header.paper h1 { font-family: var(--serif); font-size: clamp(28px, 4vw, 38px); font-weight: 700;
  letter-spacing: .015em; line-height: 1.15; }
header.paper .meta { color: var(--muted-ink); font-size: 13px; margin-top: 9px;
  font-variant-numeric: tabular-nums; }
header.paper .context { display: flex; flex-wrap: wrap; gap: 7px 18px; margin-top: 17px;
  font-size: 12.5px; color: var(--muted-ink); }
header.paper .context span { padding-left: 10px; border-left: 1px solid #a8b7b0; }
header.paper .context span:first-child { padding-left: 0; border-left: 0; }

/* 目录：桌面端固定在内容左侧，长报告中始终可见。 */
.report-layout { display: grid; grid-template-columns: 190px minmax(0, 1fr); gap: 34px; align-items: start; }
.topnav { position: sticky; top: 24px; display: block; margin: 0; padding: 15px 0;
  border-top: 2px solid #315e55; border-bottom: 1px solid var(--border); box-shadow: none; background: transparent; }
.topnav::before { display: block; margin: 0 0 9px; content: "报告目录"; }
.topnav a { display: block; padding: 8px 2px; font-size: 14px; border-bottom: 1px solid #e3e5de;
  font-weight: 600; color: var(--muted-ink); text-decoration: none; white-space: nowrap; }
.topnav a:hover { border-bottom-color: #315e55; color: var(--ink); text-decoration: none; }
.report-main { min-width: 0; }
.topnav a .cnt { font-size: 12px; font-weight: 700; margin-left: 3px;
  color: #84908c; font-variant-numeric: tabular-nums; }
.topnav a .cnt.hot-bad { color: var(--bad); }
.topnav a .cnt.hot-warn { color: var(--warn); }
.topnav a .cnt.hot-info { color: var(--info); }

/* 汇总：只呈现读者需要先决定的四类信息 */
.overview-heading { display: flex; justify-content: space-between; gap: 18px; align-items: end;
  margin-bottom: 12px; }
.overview-heading h2 { font-family: var(--serif); font-size: 20px; line-height: 1.2; }
.overview-heading p { color: var(--gray); font-size: 12.5px; text-align: right; }
.summary-grid { display: grid; grid-template-columns: repeat(4, 1fr); margin-bottom: 32px;
  background: var(--paper); border: 1px solid var(--border); }
.stat-card { padding: 17px 18px 16px; border-right: 1px solid var(--border); }
.stat-card:last-child { border-right: 0; }
.stat-card .num { font-size: 26px; font-weight: 700;
  font-variant-numeric: tabular-nums; line-height: 1; }
.stat-card .num.zero { font-size: 22px; color: #82908a; }
.stat-card .label { font-size: 12px; color: var(--gray); margin-top: 7px; }
.num.ok { color: var(--ok); } .num.bad { color: var(--bad); }
.num.warn { color: var(--warn); } .num.info { color: var(--info); }

/* 核验范围：即使没有异常，也让用户知道检查过什么 */
.scope { background: var(--tint); border-left-color: #78918a; }
.scope-intro { color: var(--muted-ink); font-size: 13.5px; margin: 0 0 14px; }
.scope-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 1px; background: var(--border);
  border: 1px solid var(--border); margin: 6px 0 12px; }
.scope-item { background: var(--paper); padding: 13px 14px; min-height: 88px; }
.scope-item strong { display: block; font-size: 13.5px; margin-bottom: 4px; }
.scope-item span { display: block; color: var(--gray); font-size: 12px; line-height: 1.55; }
.scope-item .scope-ok { color: var(--ok); font-weight: 600; }
.scope-item .scope-bad { color: var(--bad); font-weight: 600; }
.scope-item .scope-warn { color: var(--warn); font-weight: 600; }
.scope-item .scope-info { color: var(--info); font-weight: 600; }
.scope-item .scope-next { color: var(--info); font-weight: 600; }

/* 区块：报刊式分栏与风险色边线，避免通用后台卡片堆叠 */
section { background: var(--paper); border-top: 1px solid var(--border); border-bottom: 1px solid var(--border);
  padding: 23px 28px 11px; margin-bottom: 18px; border-left: 4px solid #9aa8a2; }
section, #overview { scroll-margin-top: 72px; }
section.critical { border-left-color: var(--bad); }
section.attention { border-left-color: var(--warn); }
section.notice { border-left-color: var(--info); }
section.review { border-left-color: #78918a; }
section h2 { font-family: var(--serif); font-size: 20px; font-weight: 700;
  padding-bottom: 11px; margin-bottom: 4px; line-height: 1.3;
  border-bottom: 1px solid var(--border); }
section h2 .cnt2 { font-family: -apple-system, "PingFang SC", sans-serif;
  font-size: 13px; font-weight: 700; color: var(--gray);
  margin-left: 8px; font-variant-numeric: tabular-nums; }
section h3 { font-size: 14px; font-weight: 600; margin: 10px 0 2px; color: #334155; }

/* 详情行：不再套卡片，保持长文本的阅读节奏 */
.item { padding: 15px 4px 15px 15px; border-bottom: 1px solid var(--border);
  border-left: 2px solid transparent; }
.item:last-child, .item:last-of-type { border-bottom: none; }
.item.crit { border-left-color: var(--bad); }
.item.warn { border-left-color: var(--warn); }
.item.info { border-left-color: var(--info); }
.item h3 { margin: 0 0 5px; font-size: 14px; font-weight: 600;
  word-break: break-all; }
.item p { font-size: 13.5px; }
.muted { color: var(--gray); font-size: 12.5px; }

/* 徽章：状态可读，颜色只是第二通道 */
.badge { display: inline-block; font-size: 12px; padding: 1px 9px;
  border-radius: 3px; font-weight: 600; margin-right: 6px;
  white-space: nowrap; background: #fffefa; border: 1px solid var(--border);
  color: var(--gray); }
.badge.ok { color: var(--ok); border-color: #b5cfb3; }
.badge.bad { color: var(--bad); border-color: #dfbcbc; }
.badge.warn { color: var(--warn); border-color: #dfc89e; }
.badge.info { color: var(--info); border-color: #b6cce0; }

/* 表格 */
table { width: 100%; border-collapse: collapse; font-size: 13.5px; margin: 8px 0; }
th { text-align: left; background: var(--tint); padding: 9px 12px; font-size: 12.5px;
  color: var(--muted-ink); border-bottom: 1px solid var(--border); }
td { padding: 8px 12px; border-bottom: 1px solid var(--border); vertical-align: top;
  font-variant-numeric: tabular-nums; }
td:first-child { font-family: var(--mono); font-size: 12.5px; }
td.val-paper { color: var(--bad); }
td.val-db { color: var(--ok); font-weight: 700; }

/* 附录长表: 表头吸顶 */
.appendix thead th { position: sticky; top: var(--nav-h); z-index: 10;
  background: var(--tint); box-shadow: 0 1px 0 var(--border); }
.appendix td.ref-text { font-size: 12.5px; }
.appendix tbody tr:nth-child(even) { background: #fcfcf8; }

/* 复核入口：小而明确，不与主结论竞争 */
.lnk { display: inline-block; font-size: 12px; padding: 2px 10px;
  border-radius: 3px; border: 1px solid var(--border); color: var(--muted-ink);
  background: #fffefa; text-decoration: none; margin: 1px 3px 1px 0;
  white-space: nowrap; }
.lnk:hover { border-color: #74827c; color: var(--ink); text-decoration: none; }
.lnk.doi { background: #24473f; color: #fff; border-color: #24473f; }
.lnk.doi:hover { background: #172f2a; }
a { color: var(--info); text-decoration: none; }
a:focus-visible { outline: 2px solid var(--info); outline-offset: 2px;
  border-radius: 4px; }

.todo { border-left: 2px solid #9aa8a2; padding: 13px 16px;
  color: var(--gray); font-size: 13.5px; margin: 10px 0 14px; background: var(--tint); }
footer { text-align: center; color: #7d8884; font-size: 12px; margin-top: 30px;
  line-height: 1.6; }
footer .foot-warn { color: #b45309; font-weight: 600; }
@media (max-width: 640px) {
  body { padding: 22px 12px 44px; }
  header.paper { padding-top: 16px; }
  header.paper h1 { font-size: 29px; }
  .report-layout { display: block; }
  .topnav { position: sticky; top: 0; display: flex; flex-wrap: wrap; margin-bottom: 22px;
    padding: 10px 0; border-top: 1px solid var(--border); background: rgba(243,242,237,.96); }
  .topnav::before { display: inline; margin: 0 7px 0 0; }
  .topnav a { display: inline-block; padding: 5px 0; border-bottom: 2px solid transparent; }
  .overview-heading { display: block; }
  .overview-heading p { text-align: left; margin-top: 5px; }
  .summary-grid { grid-template-columns: repeat(2, 1fr); }
  .stat-card:nth-child(2) { border-right: 0; }
  .stat-card:nth-child(-n+2) { border-bottom: 1px solid var(--border); }
  section { padding: 17px 16px 7px; }
  section, #overview { scroll-margin-top: 180px; }
  .scope-grid { grid-template-columns: 1fr; }
  .appendix thead th { position: static; }
  table { display: block; overflow-x: auto; white-space: nowrap; }
  td { white-space: normal; }
}
@media print {
  body { background: #fff; padding: 0; }
  .report-layout { display: block; }
  .topnav { display: none; }
  section { border: none; page-break-inside: avoid; }
}
"""


def _badge(text, kind):
    return f'<span class="badge {kind}">{html_mod.escape(str(text))}</span>'


def _links_html(links):
    parts = []
    if links.get("doi"):
        parts.append(f'<a class="lnk doi" href="{html_mod.escape(links["doi"])}"'
                     f' target="_blank">DOI ↗</a>')
    if links.get("openalex"):
        parts.append(f'<a class="lnk" href="{html_mod.escape(links["openalex"])}"'
                     f' target="_blank">OpenAlex ↗</a>')
    if links.get("s2"):
        parts.append(f'<a class="lnk" href="{html_mod.escape(links["s2"])}"'
                     f' target="_blank">Semantic Scholar ↗</a>')
    if links.get("openalex_search"):
        parts.append(f'<a class="lnk" href="{html_mod.escape(links["openalex_search"])}"'
                     f' target="_blank">OpenAlex 搜索 ↗</a>')
    if links.get("google_scholar"):
        parts.append(f'<a class="lnk" href="{html_mod.escape(links["google_scholar"])}"'
                     f' target="_blank">Scholar 搜索 ↗</a>')
    return "".join(parts) if parts else '<span class="muted">（无链接）</span>'


def _report_links(result):
    """为报告统一复核入口，兼容旧缓存中保存的 OpenAlex 搜索链接。"""
    links = dict(result.get("links") or {})
    if result.get("status") != "found" or "openalex_search" not in links:
        return links
    record = result.get("record") or {}
    title = record.get("title") or ""
    if not title:
        return links
    links.pop("openalex_search")
    q = urllib.parse.quote(title)
    links["google_scholar"] = f"https://scholar.google.com/scholar?q={q}"
    return links


def _final_links(result, entry):
    """最终报告的复核入口（2026-08-28 用户反馈）：只保留 DOI 与 Google
    Scholar 两个按钮——此前一格里最多并排 4 个数据库链接，对作者太多且
    不直观；无 DOI 的条目仅保留 Scholar 搜索。"""
    links = _report_links(result)
    out = {}
    if links.get("doi"):
        out["doi"] = links["doi"]
    scholar = links.get("google_scholar")
    if not scholar:
        title = ((result.get("record") or {}).get("title")
                 or entry.get("title") or entry.get("raw", ""))
        scholar = ("https://scholar.google.com/scholar?q="
                   + urllib.parse.quote(title[:120]))
    out["google_scholar"] = scholar
    return out


def _nav_cnt(n, hot=""):
    if n is None:  # 待 Claude 层填充，不显示计数
        return ""
    if not n:
        return '<span class="cnt">0</span>'
    cls = f"cnt {hot}" if hot else "cnt"
    return f'<span class="{cls}">{n}</span>'


def _entry_by_id(entries, eid):
    for e in entries:
        if e["id"] == eid:
            return e
    return {"raw": "?"}


def _is_cid(vid):
    """id 是否为正文引用编号（C1、C2…，指向 citations 下标）。"""
    return bool(re.fullmatch(r"C\d+", str(vid or "")))


def _citation_by_cid(citations, cid):
    """按 C 编号取 citation；越界返回 None。"""
    if not _is_cid(cid):
        return None
    n = int(cid[1:])
    return citations[n - 1] if 1 <= n <= len(citations) else None


def _citation_cid(citations, c):
    """按内容定位 citation 的 C 编号（兼容旧缓存 JSON 无 cid 字段）。"""
    for i, x in enumerate(citations):
        if (x.get("authors") == c.get("authors")
                and x.get("year") == c.get("year")
                and x.get("sentence") == c.get("sentence")):
            return f"C{i + 1}"
    return None


def build_report(paper_path, entries, citations, results, corr, dups,
                 timeline, preprints, verifier_stats, cross=None):
    today = datetime.date.today().isoformat()
    stem = os.path.splitext(os.path.basename(paper_path))[0]
    # Report fields can legitimately be absent (for example, a database
    # record may not expose an issue or page range).  Normalize those values
    # before HTML escaping so report generation never fails on None/int data.
    esc = lambda value: html_mod.escape("" if value is None else str(value))
    n_found = sum(1 for r in results.values() if r["status"] == "found")
    n_notfound = sum(1 for r in results.values() if r["status"] == "not_found")
    n_unver = sum(1 for r in results.values() if r["status"] == "unverified")
    n_mm = sum(1 for r in results.values() if r.get("mismatches"))
    n_corr_cited = len(corr["cited_but_missing_in_list"])
    n_corr_listed = len(corr["listed_but_never_cited"])
    n_misc = len(dups) + len(timeline) + len(preprints)

    # ---- 详情区块（只构建有内容的，空的整个不渲染，导航同步省略）----
    sections = []  # (id, 标题, 短标签, 数量, hot, 内容)

    nf_rows = []
    for eid, r in results.items():
        if r["status"] != "not_found":
            continue
        e = _entry_by_id(entries, eid)
        nf_rows.append(f"""<div class="item warn">
<h3>{_badge(eid, 'warn')} {esc(e['raw'][:150])}</h3>
<p class="muted">自动检索未匹配，不构成“文献不存在”结论。请以 DOI、出版商页或多数据库检索完成二次复核。</p>
<p>复核：{_links_html(_report_links(r))}</p>
</div>""")
    if nf_rows:
        sections.append(("sec-notfound", "⚠️ 自动未匹配（需复核）", "待复核",
                         n_notfound, "hot-warn", "".join(nf_rows)))

    mm_rows = []
    for eid, r in results.items():
        if not r.get("mismatches"):
            continue
        e = _entry_by_id(entries, eid)
        rows = "".join(
            f"<tr><td>{esc(m['field'])}</td>"
            f"<td class='val-paper'>{esc(m['paper'])}</td>"
            f"<td class='val-db'>{esc(m['database'])}</td></tr>"
            for m in r["mismatches"])
        mm_rows.append(f"""<div class="item warn">
<h3>{_badge(eid, 'warn')} {esc(e['raw'][:150])}</h3>
<table><tr><th>字段</th><th>论文所写</th><th>数据库记录（应改为此值）</th></tr>
{rows}</table>
<p>复核：{_links_html(_report_links(r))}</p>
</div>""")
    if mm_rows:
        sections.append(("sec-meta", "⚠️ 元数据不一致", "元数据",
                         n_mm, "hot-warn", "".join(mm_rows)))

    corr_rows = []
    if corr["cited_but_missing_in_list"]:
        corr_rows.append("<h3>正文引了，但列表里没有</h3>")
        for c in corr["cited_but_missing_in_list"]:
            cid = c.get("cid") or _citation_cid(citations, c)
            corr_rows.append(f"""<div class="item info">
<p>{_badge(cid or "?", "info")} “{esc(c['sentence'][:180])}”</p>
<p class="muted">→ 缺失条目：({esc(', '.join(c['authors']))}, {esc(c['year'])})</p>
</div>""")
    if corr["listed_but_never_cited"]:
        corr_rows.append("<h3>列表有，但正文从未引用</h3>")
        for eid in corr["listed_but_never_cited"]:
            e = _entry_by_id(entries, eid)
            corr_rows.append(f'<div class="item info"><p>{_badge(eid, "info")} '
                             f'{esc(e["raw"][:150])}</p></div>')
    if corr_rows:
        sections.append(("sec-corr", "🔗 对应关系问题", "对应关系",
                         n_corr_cited + n_corr_listed, "hot-info",
                         "".join(corr_rows)))

    misc_rows = []
    for d in dups:
        misc_rows.append(f'<div class="item warn"><p>{_badge("重复", "warn")} '
                         f'{esc(" & ".join(d["ids"]))}（按{esc(d["by"])}重复）</p></div>')
    for t in timeline:
        misc_rows.append(f'<div class="item warn"><p>{_badge(t["id"], "warn")} '
                         f'{esc(t["issue"])}</p></div>')
    for p_ in preprints:
        misc_rows.append(f'<div class="item warn"><p>{_badge(p_["id"], "warn")} '
                         f'{esc(p_["issue"])}</p></div>')
    if misc_rows:
        sections.append(("sec-misc", "⧉ 重复条目 / ⏰ 时间线 / 📄 preprint", "重复/时间线",
                         n_misc, "hot-warn", "".join(misc_rows)))

    # 列表内部一致性（交叉检测: 逐条比对发现不了的互换/排序/残留问题）
    cross = cross or {}
    list_rows = []
    for s in cross.get("doi_swaps", []):
        list_rows.append(f'<div class="item warn"><p>{_badge("错挂", "warn")} '
                         f'{esc(s["issue"])}</p></div>')
    for s in cross.get("ordering", []):
        list_rows.append(f'<div class="item info"><p>{_badge("排序", "info")} '
                         f'{esc(s["issue"])}</p></div>')
    for s in cross.get("title_artifacts", []):
        list_rows.append(f'<div class="item info"><p>{_badge(s["id"], "info")} '
                         f'{esc(s["issue"])}</p></div>')
    if list_rows:
        sections.append(("sec-list", "⚠️ 列表内部一致性", "列表一致性",
                         len(list_rows), "hot-warn", "".join(list_rows)))

    # 需要语义判断的项目：提供明确的后续审读入口。
    sections.append(("sec-appro", "进一步审读：引用是否支撑论述", "论述支撑", None, "",
                     '<div class="todo"><strong>可继续由 AI 审读。</strong><br>'
                     '本报告已完成自动核验。若希望继续检查，请在当前对话中回复：'
                     '<b>“继续审读引用恰当性和格式一致性。”</b><br>'
                     'AI 会先列出需要优先审读的关键引用；你可以确认或调整范围后，再继续比较正文论述与文献摘要，'
                     '并检查格式内部是否一致。仅凭题名和书目信息，不能判断文献是否真正支撑正文论述。'
                     '</div>'))
    sections.append(("sec-format", "进一步审读：参考文献格式是否统一", "格式统一", None, "",
                     '<div class="todo"><strong>可继续由 AI 审读。</strong><br>'
                     '本报告已完成自动核验。若希望继续检查，请在当前对话中回复：'
                     '<b>“继续审读引用恰当性和格式一致性。”</b><br>'
                     'AI 会先列出需要优先审读的关键引用；你可以确认或调整范围后，再继续比较正文论述与文献摘要，'
                     '并检查格式内部是否一致。目标期刊的格式要求是最终依据。'
                     '</div>'))

    # ---- 组装 ----
    H = []
    nav_links = ['<a href="#overview">总览</a>', '<a href="#scope">核验范围</a>']
    for sid, _title, short, cnt, hot, _body in sections:
        nav_links.append(f'<a href="#{sid}">{short}{_nav_cnt(cnt, hot)}</a>')
    nav_links.append(f'<a href="#appendix">附录 · 全部 {len(entries)} 条</a>')

    H.append(f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>参考文献检查报告 — {esc(stem)}</title>
<style>{REPORT_CSS}</style>
</head>
<body>
<div class="container">
<header class="paper">
  <div class="eyebrow">Reference integrity review</div>
  <h1>参考文献检查报告</h1>
  <div class="meta">{esc(stem)} ｜ 检查日期 {today} ｜
  自动初筛结果（尚非最终结论）</div>
  <div class="context"><span>文献列表 {len(entries)} 条</span>
  <span>正文引用 {len(citations)} 处</span><span>支持 DOCX / PDF / Markdown</span></div>
</header>

<div class="report-layout">
<nav class="topnav" id="contents" aria-label="报告目录">{''.join(nav_links)}</nav>
<main class="report-main">
<div id="overview" class="overview-heading"><h2>核验概览</h2>
<p>本页是自动初筛底稿；所有异常项须人工复核后才可形成最终结论。</p></div>
<div class="summary-grid">""")

    n_list_issues = (len(cross.get("doi_swaps", [])) + len(cross.get("ordering", []))
                     + len(cross.get("title_artifacts", [])))
    n_ok = n_found - n_mm
    n_other = n_unver + n_corr_cited + n_corr_listed + n_misc + n_list_issues

    def highest_severity(bad=0, warn=0, info=0):
        if bad:
            return "bad"
        if warn:
            return "warn"
        if info:
            return "info"
        return "zero"

    cards = [
        (n_ok, "ok" if n_ok else "zero", "✅ 确认存在且一致"),
        (n_notfound, highest_severity(warn=n_notfound), "⚠️ 自动未匹配（需复核）"),
        (n_mm, highest_severity(warn=n_mm), "⚠️ 文献信息需修正"),
        (n_other, highest_severity(warn=n_misc,
                                   info=n_unver + n_corr_cited + n_corr_listed),
         "🔗 对应、版本或待复核问题"),
    ]
    for num, cls, label in cards:
        H.append(f'<div class="stat-card"><div class="num {cls}">{num}</div>'
                 f'<div class="label">{label}</div></div>')
    H.append('</div>')

    def scope_status(count, level="warn"):
        if count:
            return f'<span class="scope-{level}">发现 {count} 项：见下方详情</span>'
        return f'<span class="scope-ok">已检查：未发现问题</span>'

    def existence_status():
        status = []
        if n_notfound:
            status.append(f'<span class="scope-warn">自动未匹配 {n_notfound} 项：须人工复核</span>')
        if n_unver:
            status.append(f'<span class="scope-info">无法确认 {n_unver} 项：见下方详情</span>')
        return "".join(status) or '<span class="scope-ok">已检查：未发现问题</span>'

    scope_items = [
        ("文献是否存在", existence_status(),
         "通过学术数据库核对每一条参考文献。"),
        ("书目信息是否准确", scope_status(n_mm),
         "核对作者、年份、期刊、卷期页码和 DOI。"),
        ("正文与文献列表是否对应", scope_status(n_corr_cited + n_corr_listed, "info"),
         "查找正文漏列、列表未引和错配引用。"),
        ("是否有重复条目", scope_status(len(dups)),
         "按 DOI、标题等信息识别重复参考文献。"),
        ("时间线与预印本提醒", scope_status(len(timeline) + len(preprints)),
         "识别未来年份和可能已有正式版本的预印本。"),
        ("引用是否真正支撑论述", '<span class="scope-next">可继续由 AI 审读</span>',
         "需要结合正文和文献内容判断，不能只靠数据库自动决定。"),
        ("格式是否前后一致", '<span class="scope-next">可继续由 AI 审读</span>',
         "需按目标期刊要求检查格式规则是否统一。"),
    ]
    scope_html = ['<section id="scope" class="scope review"><h2>本次核验范围</h2>',
                  '<p class="scope-intro">没有显示问题不等于没有检查。本页仅为自动初筛；'
                  '标为“可继续由 AI 审读”的项目需要结合论文内容或目标期刊要求判断。</p>',
                  '<div class="scope-grid">']
    for title, status, detail in scope_items:
        scope_html.append(f'<div class="scope-item"><strong>{title}</strong>{status}'
                          f'<span>{detail}</span></div>')
    scope_html.append('</div></section>')
    H.append(''.join(scope_html))

    for sid, title, _short, cnt, _hot, body in sections:
        cnt_html = f'<span class="cnt2">{cnt}</span>' if cnt is not None else ""
        section_class = {"hot-bad": "critical", "hot-warn": "attention",
                         "hot-info": "notice"}.get(_hot, "review")
        H.append(f'<section id="{sid}" class="{section_class}"><h2>{title}{cnt_html}</h2>{body}</section>')

    # 附录: 全部文献核验清单（问题条目排前面）
    H.append('<section id="appendix" class="review">'
             '<h2>附录 · 全部文献核验清单</h2>'
             '<p class="muted">逐条可点击复核：✅/⚠️ 的 DOI 直达出版商页面；'
             '❌ 的检索链接确认是否真的查无此文。问题条目已排在前面。</p>')

    def _row_rank(e):
        r = results.get(e["id"], {})
        if r.get("status") == "not_found":
            return 0
        if r.get("mismatches") or r.get("status") == "unverified":
            return 1
        return 2

    H.append('<table class="appendix"><thead><tr><th>ID</th><th>文献</th>'
             '<th>状态</th><th>复核</th></tr></thead><tbody>')
    for e in sorted(entries, key=_row_rank):
        r = results.get(e["id"], {})
        status = r.get("status", "unverified")
        icon = {"found": "ok", "not_found": "warn", "unverified": "gray"}[status]
        label = STATUS_ICON.get(status, "❓")
        if r.get("mismatches"):
            icon, label = "warn", "⚠️"
        desc = esc(e["raw"][:90] + ("…" if len(e["raw"]) > 90 else ""))
        H.append(f'<tr><td>{esc(e["id"])}</td>'
                 f'<td class="ref-text">{desc}</td>'
                 f'<td>{_badge(label, icon)}</td>'
                 f'<td>{_links_html(_report_links(r))}</td></tr>')
    H.append("</tbody></table></section>")

    H.append("""<footer>由 ob-reference-check 生成 · 自动初筛底稿，须经人工复核后再交付最终结论</footer>
</main>
</div>
</div></body></html>""")
    return "\n".join(H)


# ---------------------------------------------------------------------------
# 6b. 定稿机制（P0-1）: Claude 复核结论 final.json → 脚本重渲染最终报告
#
# 废除「Claude 手改 HTML」：本次交付 bug（16 行徽章改 class 没换 emoji）的
# 根源是五处状态（概览卡片/nav 计数/各节状态/附录徽章/说明文字）靠人肉
# regex 同步。改为单一数据源（final.json）驱动，模板保证一致性。
# ---------------------------------------------------------------------------

VERDICT_STORE = os.path.expanduser("~/.reference_check/verdicts.json")
FINAL_ICON = {"ok": "✅", "warn": "⚠️", "info": "❓"}


def load_verdict_store():
    try:
        with open(VERDICT_STORE, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_verdict_store(store):
    os.makedirs(os.path.dirname(VERDICT_STORE), exist_ok=True)
    tmp = VERDICT_STORE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(store, f, ensure_ascii=False, indent=1)
    os.replace(tmp, VERDICT_STORE)


CHECK_TITLES = {
    'existence': '文献存在性核验', 'metadata': '书目元数据逐项比对',
    'correspondence': '正文引用 — 文献列表对应', 'duplicates': '重复条目检测',
    'timeline': '时间线与预印本检查', 'cross_checks': '列表内部一致性交叉检测',
    'appropriateness': '引用恰当性深查', 'format': '格式一致性与书目通读',
}


def verdict_check(v):
    return v.get('check') or {'bibliography': 'metadata'}.get(
        v.get('category', 'bibliography'), v.get('category', 'bibliography'))


def _entry_fingerprint(entry):
    fields = ('raw', 'authors', 'year', 'year_suffix', 'title', 'venue',
              'volume', 'issue', 'pages', 'doi')
    value = {k: entry.get(k) for k in fields}
    return hashlib.sha256(json.dumps(value, sort_keys=True, ensure_ascii=False).encode()).hexdigest()


def updated_verdict_store(store, data, verdicts):
    """Only entry-bound bibliography/format judgments may be reused; never claim support."""
    updated = dict(store)
    for e in data['entries']:
        doi = (e.get('doi') or '').lower().rstrip('.')
        related = [v for v in verdicts if v['id'] == e['id']
                   and v.get('category', 'bibliography') in ('bibliography', 'format')
                   and verdict_check(v) in ('existence', 'metadata', 'format')]
        if doi and related:
            updated[doi] = {'schema_version': 2, 'entry_fingerprint': _entry_fingerprint(e),
                            'verdicts': related, 'date': datetime.date.today().isoformat()}
    return updated


def prior_verdict_for(entry, store):
    value = store.get((entry.get('doi') or '').lower().rstrip('.'))
    if not isinstance(value, dict) or value.get('schema_version') != 2:
        return None
    if value.get('entry_fingerprint') != _entry_fingerprint(entry):
        return None
    try:
        if (datetime.date.today() - datetime.date.fromisoformat(value['date'])).days > 180:
            return None
    except (KeyError, ValueError, TypeError):
        return None
    # R numbers are local to a manuscript, not reusable identity.
    return dict(value, verdicts=[dict(v, id=entry['id']) for v in value['verdicts']])


def _check_total(data, check):
    if check == 'appropriateness':
        return len(data.get('citations', []))
    if check == 'correspondence':
        return len(data['entries']) + len(data.get('citations', []))
    return len(data['entries'])


def _required_dispositions(data, check):
    """Return IDs needing explicit decisions in this check, including cleared alarms."""
    issues = []
    if check == 'duplicates':
        issues = data.get('duplicates', [])
    elif check == 'timeline':
        issues = data.get('timeline', []) + data.get('preprints', [])
    elif check == 'cross_checks':
        issues = [x for xs in data.get('cross_checks', {}).values() for x in xs]
    elif check == 'correspondence':
        corr = data.get('correspondence', {})
        issues = [{'id': c.get('cid') or _citation_cid(data.get('citations', []), c)}
                  for c in corr.get('cited_but_missing_in_list', []) + corr.get('ambiguous_citations', [])]
        issues += [{'id': eid} for eid in corr.get('listed_but_never_cited', [])]
    elif check in ('existence', 'metadata'):
        for e in data['entries']:
            r = data.get('verification', {}).get(e['id'], {})
            if ((check == 'existence' and (r.get('status') != 'found' or r.get('cache_stale')))
                    or (check == 'metadata' and (r.get('mismatches') or e.get('parse_ok') is False))):
                issues.append({'id': e['id']})
    return {eid for issue in issues for eid in issue.get('ids', [issue.get('id')]) if eid}


def _validate_checks(data, verdicts, checks):
    if not isinstance(checks, dict):
        sys.exit('[错误] checks 必须是对象')
    errors = []
    for check, state in checks.items():
        if check not in CHECK_TITLES or not isinstance(state, dict):
            errors.append(f'{check}: 未知检查维度或状态结构无效')
            continue
        status = state.get('status')
        total = _check_total(data, check)
        reviewed = state.get('reviewed', 0)
        if status not in ('completed', 'partial', 'not_run', 'blocked'):
            errors.append(f'{check}: 无效状态')
        if type(reviewed) is not int or not 0 <= reviewed <= total:
            errors.append(f'{check}: reviewed 须在 0..{total} 之间')
        if status in ('completed', 'partial', 'blocked') and not state.get('note'):
            errors.append(f'{check}: 必须记录检查依据或限制 note')
        if status == 'not_run' and reviewed != 0:
            errors.append(f'{check}: 未检查时 reviewed 必须为 0')
        if check == 'appropriateness':
            reviewed_ids = {v['id'] for v in verdicts if verdict_check(v) == check and _is_cid(v['id'])}
            if reviewed != len(reviewed_ids):
                errors.append(f'{check}: reviewed 必须等于逐处引用结论数 {len(reviewed_ids)}')
        if status == 'completed':
            if reviewed != total:
                errors.append(f'{check}: 完成状态须覆盖 {total} 项')
            resolved = {v['id'] for v in verdicts if verdict_check(v) == check}
            missing = _required_dispositions(data, check) - resolved
            if missing:
                errors.append(f'{check}: 异常缺少本维度结论 {sorted(missing)}')
            # A completed check may contain unresolved findings; it means the work
            # was performed, not that the bibliography is correct.
    if errors:
        sys.exit('[错误] 检查范围校验失败:\n  - ' + '\n  - '.join(errors))


def _validate_verdicts(entries, verification, verdicts,
                       citations=None, correspondence=None):
    r"""final.json 校验（F5: warn/info 结论必须带证据，防未查库断言）。

    id 支持两种：条目号（R*，指向文献列表）与正文引用号（C\d+，指向
    citations 下标）——后者让"正文引用但列表缺失"类问题获得结论出口
    （2026-08-28 缺口修复，docs/correspondence-gap-and-fix.md）。
    """
    by_id = {e["id"]: e for e in entries}
    citations = citations or []
    errors, seen_cid = [], set()
    unique = set()
    seen_checks = set()
    for v in verdicts:
        vid = v.get("id")
        category = v.get("category", "bibliography")
        check = verdict_check(v)
        if category not in ("bibliography", "correspondence", "appropriateness", "format") or check not in CHECK_TITLES:
            errors.append(f"{vid}: 未知 category/check")
        allowed = {"bibliography": {"existence", "metadata", "duplicates", "timeline", "cross_checks"},
                   "correspondence": {"correspondence"}, "appropriateness": {"appropriateness"}, "format": {"format"}}
        if check not in allowed.get(category, set()):
            errors.append(f"{vid}: category 与 check 不一致")
        key = (vid, check, v.get("ref_id"))
        if key in unique:
            errors.append(f"{vid}: 同一检查维度重复结论")
        unique.add(key)
        seen_checks.add((vid, check))
        if category == "appropriateness":
            if not _is_cid(vid):
                errors.append(f"{vid}: 恰当性结论必须绑定 C 编号")
            if v.get("ref_id") not in by_id and v.get("final_status") != "info":
                errors.append(f"{vid}: 恰当性结论必须指定有效 ref_id")
            citation = _citation_by_cid(citations, vid) if _is_cid(vid) else None
            if citation and citation.get("ref_ids") and v.get("ref_id") not in citation["ref_ids"]:
                errors.append(f"{vid}: ref_id 与本次引用对应关系不一致，先修正对应记录")
            if not v.get("evidence"):
                errors.append(f"{vid}: 恰当性结论必须有证据，即使为 ok")
            if v.get("final_status") == "ok" and v.get("evidence_level") not in ("abstract", "full_text"):
                errors.append(f"{vid}: 支撑结论须注明 evidence_level=abstract/full_text")
            if v.get("final_status") == "ok" and v.get("evidence_level") == "abstract":
                if not (verification.get(v.get("ref_id"), {}).get("abstract") or v.get("source_excerpt")):
                    errors.append(f"{vid}: 摘要为空，不能判定支撑")
        if _is_cid(vid):
            if _citation_by_cid(citations, vid) is None:
                errors.append(f"{vid}: 不存在于正文引用"
                              f"（citations 共 {len(citations)} 处）")
                continue
            if category == "correspondence":
                seen_cid.add(vid)
            if v.get("final_status") not in FINAL_ICON:
                errors.append(f"{vid}: final_status 必须是 ok/warn/info，"
                              f"当前 {v.get('final_status')!r}")
            if v.get("final_status") in ("warn", "info"):
                for k in ("verdict", "evidence", "action"):
                    if not v.get(k):
                        errors.append(f"{vid}: {k} 为必填"
                                      f"（warn/info 结论必须带证据与行动项）")
            continue
        if vid not in by_id:
            errors.append(f"{vid}: 不存在于文献列表")
            continue
        if v.get("final_status") not in FINAL_ICON:
            errors.append(f"{vid}: final_status 必须是 ok/warn/info，"
                          f"当前 {v.get('final_status')!r}")
        if v.get("final_status") in ("warn", "info"):
            for k in ("verdict", "evidence", "action"):
                if not v.get(k):
                    errors.append(f"{vid}: {k} 为必填"
                                  f"（warn/info 结论必须带证据与行动项）")
    for e in entries:
        eid = e["id"]
        r = verification.get(eid, {})
        required = []
        if r.get("status") in ("not_found", "unverified") or r.get("cache_stale"):
            required.append("existence")
        if r.get("mismatches") or e.get("parse_ok") is False:
            required.append("metadata")
        for check in required:
            # Legacy bibliography verdicts can resolve an existence alarm only
            # when they do not explicitly claim to be a different check.
            legacy = check == "existence" and any(
                v.get("id") == eid and "check" not in v
                and v.get("category", "bibliography") == "bibliography" for v in verdicts)
            if (eid, check) not in seen_checks and not legacy:
                errors.append(f"{eid}: {check} 自动异常必须有本维度显式复核结论")
    # 每条"正文引用但列表缺失"必须有显式结论（同"自动异常必须有结论"规则）：
    # 确认缺失→warn；复核为匹配误报（年份不一致/拼写差异等）→ok
    for c in ((correspondence or {}).get("cited_but_missing_in_list", [])
              + (correspondence or {}).get("ambiguous_citations", [])):
        cid = c.get("cid") or _citation_cid(citations, c)
        if not cid or cid in seen_cid:
            continue
        who = f"({', '.join(c.get('authors', []))}, {c.get('year')})"
        errors.append(f"{cid}: 正文引用 {who} 缺失或存在歧义，"
                      f"必须有显式结论（确认缺失→warn；匹配误报→ok）")
    if errors:
        sys.exit("[错误] final.json 校验失败:\n  - " + "\n  - ".join(errors))


def build_final_report(data, verdicts, checks=None):
    """按复核结论渲染最终报告（P1-1~4 信息架构）。

    - 概览只留行动导向卡片，无重复段落文字（P1-1）
    - 复核过程与误报备注不展示——用户只看最终结论（P1-2，
      2026-08-28 用户反馈：行内"曾自动标记…误报"看不出指哪条，删）
    - 无问题条目默认折叠 <details>，仅展开问题行（P1-3）
    - 总览下方保留"本次检查范围"小节，让用户知道覆盖面
      （P1-4 修订，2026-08-28 用户反馈恢复；footer 不再重复）
    """
    esc = lambda v: html_mod.escape("" if v is None else str(v))
    entries = data["entries"]
    results = data.get("verification", {})
    citations = data.get("citations", [])
    today = datetime.date.today().isoformat()
    stem = os.path.splitext(os.path.basename(data["paper"]["path"]))[0]

    checks = checks or {}
    complete = all(checks.get(k, {}).get("status") == "completed" for k in CHECK_TITLES)
    review_label = "检查记录完整" if complete else "部分检查 / 完成记录不全"
    # 多维结论取最高风险，不让后写入的 ok 覆盖警告。
    vmap = {}
    rank = {"ok": 0, "info": 1, "warn": 2}
    for v in verdicts:
        old = vmap.get(v["id"])
        if old is None or rank[v["final_status"]] > rank[old["final_status"]]:
            vmap[v["id"]] = v

    def cat(v):
        return v.get("category", "bibliography")

    must = [v for v in verdicts if v["final_status"] == "warn"
            and cat(v) in ("bibliography", "correspondence")]
    appro = [v for v in verdicts if cat(v) == "appropriateness"
             and v["final_status"] != "ok"]
    fmt = [v for v in verdicts if cat(v) == "format"
           and v["final_status"] != "ok"]
    check = [v for v in verdicts if v["final_status"] == "info"
             and cat(v) in ("bibliography", "correspondence")]

    def citation_item_html(v, level):
        """C 编号 verdict 的卡片：主体是正文引用，显示引用句原文与
        行动建议；复核入口用 作者+年份 构造 Scholar 搜索。"""
        c = _citation_by_cid(citations, v["id"])
        if c is None:
            return ""
        who = f"{', '.join(c.get('authors') or [])}, {c.get('year')}"
        context_label = (f"对应文献 {v.get('ref_id') or '待确认'}"
                         if cat(v) == "appropriateness" else ("文献列表中无对应条目" if any(
                             (item.get("cid") or _citation_cid(citations, item)) == v["id"]
                             for item in data.get("correspondence", {}).get("cited_but_missing_in_list", []))
                               else "对应关系待核对"))
        locator = c.get("locator") or c.get("section") or "正文位置未记录"
        scholar = ("https://scholar.google.com/scholar?q="
                   + urllib.parse.quote(who))
        return f"""<div class="item {level}">
<h3>{_badge(v['id'], level)} {esc(v.get('verdict') or '')}</h3>
<p class="muted">“{esc((c.get('sentence') or '')[:150])}”</p>
<p class="muted">→ 引用（{esc(who)}），{esc(context_label)} · {esc(locator)}</p>
<p><b>建议：</b>{esc(v.get('action') or '')}</p>
<p class="muted">依据：{esc(v.get('evidence') or '')}</p>
<p>复核：{_links_html({'google_scholar': scholar})}</p>
</div>"""

    def item_html(v, level):
        if _is_cid(v["id"]):
            return citation_item_html(v, level)
        e = _entry_by_id(entries, v["id"])
        r = results.get(v["id"], {})
        return f"""<div class="item {level}">
<h3>{_badge(v['id'], level)} {esc(v.get('verdict') or '')}</h3>
<p class="muted">{esc(e['raw'][:150])}</p>
<p><b>建议：</b>{esc(v.get('action') or '')}</p>
<p class="muted">依据：{esc(v.get('evidence') or '')}</p>
<p>复核：{_links_html(_final_links(r, e))}</p>
</div>"""

    sections = []
    if must:
        sections.append(("sec-must", "必须处理", len(must),
                         "".join(item_html(v, "warn") for v in must)))
    if appro or check:
        sections.append(("sec-check", "存疑 / 建议核对", len(appro) + len(check),
                         "".join(item_html(v, "warn" if v["final_status"] == "warn" else "info")
                                 for v in appro + check)))
    if fmt:
        sections.append(("sec-format", "格式调整", len(fmt),
                         "".join(item_html(v, "info") for v in fmt)))

    # 附录: 问题行显式列出，无问题条目折叠（P1-3）。
    # 不渲染任何"曾自动标记…误报"类过程备注（2026-08-28 用户反馈）
    problem_ids = {v["id"] for v in must + appro + fmt + check}
    # "其余确认"按条目计数——C 编号 verdict 的主体不在列表里，
    # 不应折减条目确认数（2026-08-28 缺口修复）
    rows_problem, rows_ok = [], []
    for e in entries:
        eid = e["id"]
        v = vmap.get(eid)
        r = results.get(eid, {})
        desc = esc(e["raw"][:90] + ("…" if len(e["raw"]) > 90 else ""))
        links = _links_html(_final_links(r, e))
        st = v["final_status"] if v else "info"
        badge = _badge(FINAL_ICON[st], st)
        row = (f'<tr><td>{esc(eid)}</td><td class="ref-text">{desc}</td>'
               f'<td>{badge}</td><td>{links}</td></tr>')
        (rows_problem if eid in problem_ids else rows_ok).append(row)

    n_a = sum(1 for c in citations if c.get("triage") == "A")
    n_confirmed = sum(1 for v in vmap.values() if not _is_cid(v["id"]) and v["final_status"] == "ok")
    caps = data.get("summary_stats", {}).get("source_capabilities", {})

    nav = ['<a href="#overview">总览</a>', '<a href="#scope">检查范围</a>']
    for sid, title, cnt, _body in sections:
        nav.append(f'<a href="#{sid}">{title}<span class="cnt hot-{"warn" if sid != "sec-check" else "info"}">{cnt}</span></a>')
    nav.append(f'<a href="#appendix">附录 · 全部 {len(entries)} 条</a>')

    H = [f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>参考文献检查报告（定稿）— {esc(stem)}</title>
<style>{REPORT_CSS}</style>
</head>
<body>
<div class="container">
<header class="paper">
  <div class="eyebrow">Reference integrity review · final</div>
  <h1>参考文献检查报告</h1>
  <div class="meta">{esc(stem)} ｜ 定稿 {today} ｜ {esc(review_label)}</div>
  <div class="context"><span>文献列表 {len(entries)} 条</span>
  <span>正文引用 {len(citations)} 处</span><span>A 类待覆盖 {n_a} 处</span></div>
</header>

<div class="report-layout">
<nav class="topnav" aria-label="报告目录">{''.join(nav)}</nav>
<main class="report-main">
<div id="overview" class="overview-heading"><h2>核验结论</h2></div>
<div class="summary-grid">"""]

    cards = [
        (len(must), "bad" if must else "zero", "🔴 必须修改"),
        (len(appro) + len(check), "warn" if (appro or check) else "zero", "⚠️ 存疑 / 建议核对"),
        (len(fmt), "info" if fmt else "zero", "🔵 格式调整"),
        (n_confirmed, "ok", "✅ 有通过记录（仅已检维度）"),
    ]
    for num, cls, label in cards:
        H.append(f'<div class="stat-card"><div class="num {cls}">{num}</div>'
                 f'<div class="label">{label}</div></div>')
    H.append('</div>')

    H.append('<section id="scope" class="scope review"><h2>本次检查范围</h2>'
             '<p class="scope-intro">完成表示已执行检查，不表示没有问题；未记录的检查不会推定通过。</p>'
             '<div class="scope-grid">')
    for key, title in CHECK_TITLES.items():
        state = checks.get(key, {})
        status = state.get('status', 'not_run')
        reviewed = state.get('reviewed', 0)
        total = _check_total(data, key)
        findings = [v for v in verdicts if verdict_check(v) == key and v['final_status'] != 'ok']
        labels = {'completed': '已完成检查', 'partial': '部分完成',
                  'not_run': '未记录检查', 'blocked': '检查受阻'}
        kind = 'ok' if status == 'completed' and not findings else 'warn' if findings else 'info'
        detail = f"{labels.get(status, '未记录检查')}：{reviewed}/{total}"
        if findings:
            detail += f"；{len(findings)} 项待处理，见下方详情"
        elif status == 'completed':
            detail += '；未发现问题'
        H.append(f'<div class="scope-item"><strong>{esc(title)}</strong>'
                 f'<span class="scope-{kind}">{esc(detail)}</span>'
                 f'<span>{esc(state.get("note", "尚无检查完成记录"))}</span></div>')
    H.append('</div></section>')

    for sid, title, cnt, body in sections:
        cls = {"sec-must": "critical", "sec-check": "attention",
               "sec-format": "notice"}[sid]
        H.append(f'<section id="{sid}" class="{cls}">'
                 f'<h2>{title}<span class="cnt2">{cnt}</span></h2>{body}</section>')

    H.append('<section id="appendix" class="review">'
             f'<h2>附录 · 全部文献核验清单</h2>')
    if rows_problem:
        H.append('<p class="muted">以下为需关注条目：'
                 '<b>⚠️</b> 确认需修改或存疑（对应上方"必须处理 / 存疑"卡片，'
                 '含修改建议）；<b>❓</b> 建议人工核对（证据不足以下定论，'
                 '需读原文或查证后自行判断）；其余条目默认折叠；徽章仅表示已有记录，不代表全部维度已核实。</p>'
                 '<table class="appendix"><thead><tr><th>ID</th><th>文献</th>'
                 '<th>结论</th><th>复核</th></tr></thead><tbody>'
                 + "".join(rows_problem) + '</tbody></table>')
    if rows_ok:
        H.append(f'<details><summary>显示其余 {len(rows_ok)} 条文献（仅显示已检维度）</summary>'
                 '<table class="appendix"><thead><tr><th>ID</th><th>文献</th>'
                 '<th>结论</th><th>复核</th></tr></thead><tbody>'
                 + "".join(rows_ok) + '</tbody></table></details>')
    H.append('</section>')

    H.append(f"""<footer>由 ob-reference-check 生成 · 自动初筛底稿，
<span class="foot-warn">建议经人工复核后再交付</span>——请勿仅依赖 AI 筛查结论</footer>
</main>
</div>
</div></body></html>""")
    return "\n".join(H)


def _file_sha256(path):
    """Hash source bytes without loading large documents entirely into memory."""
    try:
        digest = hashlib.sha256()
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(1024 * 1024), b""):
                digest.update(chunk)
        return digest.hexdigest()
    except (OSError, TypeError) as ex:
        sys.exit(f"[错误] 稿件无法读取: {path}（{ex}）。"
                 "若文件已移动，请用 --source 指定当前稿件。")


def review_binding(data):
    """Bind reviews to source bytes and the current, possibly corrected, screening."""
    paper = data.get("paper", {})
    source_hash = paper.get("sha256")
    if not isinstance(source_hash, str) or not re.fullmatch(r"[0-9a-f]{64}", source_hash):
        sys.exit("[错误] 初筛缺少有效稿件指纹，请使用当前版本重新初筛。")
    # Paths/timestamps may change when moving an audit; evidence must not change.
    content = dict(data)
    content["paper"] = {k: v for k, v in paper.items()
                        if k not in ("path", "report", "checked_at")}
    encoded = json.dumps(content, ensure_ascii=False, sort_keys=True,
                         separators=(",", ":")).encode("utf-8")
    return {"schema_version": 1, "source_sha256": source_hash,
            "screening_sha256": hashlib.sha256(encoded).hexdigest()}


def _validate_source(data, target=None, source_path=None):
    expected = review_binding(data)["source_sha256"]
    explicit_paper = target if target and not target.endswith(".json") else None
    selected = source_path or explicit_paper or data["paper"].get("path")
    # An explicit paper target must itself match, even when --source is also used.
    for path in dict.fromkeys([selected, explicit_paper] if explicit_paper else [selected]):
        if _file_sha256(path) != expected:
            sys.exit("[错误] 稿件内容已变化或选错稿件，请重新初筛并复核；"
                     "不能用旧稿结论生成当前稿件报告。")


def _validate_binding(data, payload, target=None, source_path=None):
    expected = review_binding(data)
    actual = payload.get("binding") if isinstance(payload, dict) else None
    if not isinstance(actual, dict) or actual.get("schema_version") != 1:
        sys.exit("[错误] 复核结论缺少有效版本绑定。请先核对原稿和初筛，"
                 "用 --prepare-final 生成新结论模板并重新确认适用结论。")
    if actual.get("source_sha256") != expected["source_sha256"]:
        sys.exit("[错误] 复核结论与初筛的稿件指纹不一致，请选择同一稿件版本的文件。")
    if actual.get("screening_sha256") != expected["screening_sha256"]:
        sys.exit("[错误] 复核结论与当前初筛内容不一致。若纠正了条目、引用对应或证据，"
                 "请重新准备结论模板并复核受影响结论。")
    _validate_source(data, target, source_path)


def run_prepare_final(target, final_path=None, source_path=None):
    """Create an empty, version-bound review template without replacing any work."""
    target = os.path.abspath(target)
    data_path, data = _load_refcheck_json(target)
    _validate_source(data, target, source_path)
    stem = os.path.splitext(os.path.basename(data["paper"]["path"]))[0]
    final_path = final_path or os.path.join(os.path.dirname(data_path), f"{stem}_final.json")
    payload = {"binding": review_binding(data), "verdicts": [], "checks": {}}
    try:
        with open(final_path, "x", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
    except FileExistsError:
        sys.exit(f"[错误] 结论文件已存在: {final_path}。请用 --final 指定新文件，保留已有复核。")
    except OSError as ex:
        sys.exit(f"[错误] 无法写入结论模板: {final_path}（{ex}）")
    print(f"✅ 待复核结论模板: {final_path}（尚无审读结论或完成记录）")


def _load_refcheck_json(target):
    target = os.path.abspath(target)
    if target.endswith(".json"):
        with open(target, encoding="utf-8") as f:
            data = json.load(f)
        if "entries" in data:
            return target, data
        # 传进来的是 *_final.json（只有 verdicts，2026-08-28 实测踩坑）:
        # 剥掉 _final 后缀，定位同目录的初筛数据文件
        stem = os.path.splitext(os.path.basename(target))[0]
        if stem.endswith("_final"):
            stem = stem[:-len("_final")]
        d = os.path.dirname(target)
        cands = [c for c in sorted(glob.glob(
            os.path.join(d, f"{stem}_refcheck_*.json")))
            if not c.endswith("_final.json")]
        if cands:
            with open(cands[-1], encoding="utf-8") as f:
                return cands[-1], json.load(f)
        sys.exit(f"[错误] {target} 是复核结论文件，同目录未找到 "
                 f"{stem}_refcheck_*.json（先跑一次初筛）")
    d = os.path.dirname(target)
    stem = os.path.splitext(os.path.basename(target))[0]
    cands = sorted(glob.glob(os.path.join(d, f"{stem}_refcheck_*.json")))
    cands = [c for c in cands if not c.endswith("_final.json")]
    if not cands:
        sys.exit(f"[错误] 未找到 {stem}_refcheck_*.json（先跑一次初筛）")
    with open(cands[-1], encoding="utf-8") as f:
        return cands[-1], json.load(f)


def run_finalize(target, final_path=None, source_path=None):
    target = os.path.abspath(target)
    # 直接传 _final.json 时，它本身就是复核结论文件
    if (not final_path and target.endswith("_final.json")
            and os.path.exists(target)):
        final_path = target
    data_path, data = _load_refcheck_json(target)
    stem = os.path.splitext(os.path.basename(data["paper"]["path"]))[0]
    outdir = os.path.dirname(data_path)
    if not final_path:
        # 候选: 论文原名_final.json（skill 约定）和数据文件同名_final.json
        for cand in (os.path.join(outdir, f"{stem}_final.json"),
                     os.path.join(outdir, os.path.splitext(os.path.basename(
                         data_path))[0] + "_final.json")):
            if os.path.exists(cand):
                final_path = cand
                break
    if not final_path or not os.path.exists(final_path):
        sys.exit(f"[错误] 未找到复核结论文件: "
                 f"{os.path.join(outdir, stem + '_final.json')}"
                 f"（可用 --final 指定路径）")
    with open(final_path, encoding="utf-8") as f:
        payload = json.load(f)
    _validate_binding(data, payload, target, source_path)
    verdicts = payload.get("verdicts", payload) if isinstance(payload, dict) else payload
    _validate_verdicts(data["entries"], data.get("verification", {}), verdicts,
                       data.get("citations", []), data.get("correspondence", {}))

    out = os.path.join(outdir, f"{stem}_refcheck_"
                       f"{datetime.date.today().strftime('%Y%m%d')}_final.html")
    checks = payload.get("checks", {}) if isinstance(payload, dict) else {}
    _validate_checks(data, verdicts, checks)
    report = build_final_report(data, verdicts, checks)
    _validate_source(data, target, source_path)
    _probe_writable(outdir)
    with open(out, "w", encoding="utf-8") as f:
        f.write(report)

    store = load_verdict_store()
    updated = updated_verdict_store(store, data, verdicts)
    n_saved = sum(1 for key, value in updated.items() if value != store.get(key))
    if n_saved:
        save_verdict_store(updated)

    n_bad = sum(1 for v in verdicts if v["final_status"] == "warn")
    n_info = sum(1 for v in verdicts if v["final_status"] == "info")
    print(f"✅ 最终报告: {out}")
    print(f"   ⚠️ 需处理 {n_bad} 项 · ❓ 建议核对 {n_info} 项 · "
          f"♻️ 复核结论已按 DOI 回流 {n_saved} 条")


def run_verify_doi(target, ids):
    """P2-2: 批量 Crossref DOI 直查，替代 Claude 逐条 WebFetch。"""
    if target.endswith(".json"):
        with open(target, encoding="utf-8") as f:
            entries = json.load(f)["entries"]
    else:
        paragraphs = parse_document(target)
        raw_entries, _ = split_references(paragraphs)
        if raw_entries is None:
            sys.exit("[错误] 未找到参考文献列表")
        entries = [parse_entry(r, i + 1) for i, r in enumerate(raw_entries)]
    by_id = {e["id"]: e for e in entries}
    want = [x.strip().upper() for x in ids.split(",") if x.strip()]
    unknown = [w for w in want if w not in by_id]
    verifier = Verifier()
    out = []
    for w in want:
        e = by_id.get(w)
        if e is None:
            continue
        doi = e.get("doi")
        if not doi:
            out.append({"id": w, "status": "no_doi"})
            print(f"{w}: 论文该条未写 DOI")
            continue
        rec = verifier._crossref_doi_lookup(doi)
        if rec is None:
            out.append({"id": w, "doi": doi, "status": "lookup_failed"})
            print(f"{w}: ❌ DOI 直查失败 {doi}")
        else:
            out.append({"id": w, "doi": doi, "status": "found",
                        "title": rec["title"], "year": rec["year"],
                        "venue": rec["venue"],
                        "url": f"https://doi.org/{doi}"})
            print(f"{w}: ✅ {doi} → {rec['title']} ({rec['year']}) {rec['venue']}")
    if unknown:
        print(f"⚠️ 未知条目: {','.join(unknown)}")
    print(json.dumps(out, ensure_ascii=False))


def _probe_writable(outdir):
    """C8: 输出目录可写性前置探测，避免整套检索跑完才在写报告时失败。"""
    probe = os.path.join(outdir, f".refcheck_probe_{os.getpid()}")
    try:
        with open(probe, "w") as f:
            f.write("")
        os.remove(probe)
    except OSError as ex:
        sys.exit(f"[错误] 输出目录不可写: {outdir}（{ex}）。"
                 f"请用 --outdir 指定可写目录")




# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

def _notify_update(offline=False):
    """Offer bounded update advice without affecting the reference-check workflow."""
    if offline:
        return
    try:
        import importlib.util
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "check_update.py")
        spec = importlib.util.spec_from_file_location("refcheck_update", path)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        result = module.check_update(__version__)
        if result["status"] == "available":
            print(f"[更新提示] ob-reference-check {__version__} → {result['latest_version']}；"
                  f"可请 AI 助手从 {module.REPOSITORY_URL} 更新。确认前继续当前检查。")
    except Exception:
        pass  # Optional advice cannot prevent parsing, review, or report generation.


def main():
    ap = argparse.ArgumentParser(description="论文参考文献机械检查")
    ap.add_argument("paper", help="论文文件 (.docx/.pdf/.md)，或 --finalize/"
                    "--verify-doi 模式下传 *_refcheck_*.json / 论文文件")
    ap.add_argument("--offline", action="store_true",
                    help="只用本地缓存，不访问网络")
    ap.add_argument("--delay", type=float, default=0.15,
                    help="API 请求间隔秒数（礼貌限速）")
    ap.add_argument("--open-draft", action="store_true",
                    help="显式在浏览器打开自动初筛底稿（默认不打开）")
    ap.add_argument("--outdir", help="报告/数据输出目录（默认论文所在目录）")
    modes = ap.add_mutually_exclusive_group()
    modes.add_argument("--prepare-final", action="store_true",
                       help="生成带版本绑定的空结论模板，不覆盖已有文件")
    modes.add_argument("--finalize", action="store_true",
                    help="读 <论文>_final.json 复核结论，重渲染最终报告 HTML")
    ap.add_argument("--final", metavar="FINAL_JSON",
                    help="final.json 路径（默认: 论文同目录 <论文名>_final.json）")
    ap.add_argument("--source", metavar="SOURCE_FILE",
                    help="定稿/准备模板时指定已移动的源稿件，内容须与初筛相同")
    modes.add_argument("--verify-doi", metavar="R_IDS",
                    help="批量 Crossref DOI 直查（如 R7,R16），输出 JSON 供 AI 层直接使用")
    args = ap.parse_args()

    if not os.path.exists(args.paper):
        sys.exit(f"[错误] 文件不存在: {args.paper}")

    _notify_update(args.offline)
    if args.source and not (args.finalize or args.prepare_final):
        ap.error("--source 仅用于 --prepare-final 或 --finalize")
    if args.prepare_final:
        run_prepare_final(args.paper, args.final, args.source)
        return
    if args.finalize:
        run_finalize(args.paper, args.final, args.source)
        return
    if args.verify_doi:
        run_verify_doi(args.paper, args.verify_doi)
        return

    print(f"[1/5] 解析文档: {args.paper}")
    source_hash = _file_sha256(args.paper)
    paragraphs = parse_document(args.paper)
    raw_entries, ref_idx = split_references(paragraphs)
    if raw_entries is None:
        sys.exit("[错误] 未找到参考文献列表（找不到 References 标题）。"
                 "请确认文档包含独立的参考文献部分。")

    print(f"[2/5] 解析 {len(raw_entries)} 条文献条目")
    entries = [parse_entry(r, i + 1) for i, r in enumerate(raw_entries)]
    n_bad = sum(1 for e in entries if not e["parse_ok"])
    if n_bad:
        print(f"      ⚠️ {n_bad} 条解析可疑（会在报告中标注，Claude 层兜底）")

    body = paragraphs[:ref_idx] if ref_idx else paragraphs
    print("[3/5] 提取正文引用标记")
    citations = extract_citations(body)
    n_a = sum(1 for c in citations if c["triage"] == "A")
    n_b = sum(1 for c in citations if c["triage"] == "B")
    n_c = sum(1 for c in citations if c["triage"] == "C")
    print(f"      共 {len(citations)} 处引用：A 类(承重) {n_a} ｜ B 类(顺带) {n_b} ｜ C 类(堆砌) {n_c}")

    verifier = Verifier(offline=args.offline)
    print(f"[4/5] 验证文献（三源轮换: OpenAlex / Crossref / Semantic Scholar，缓存 {CACHE_DIR}）")
    results = {}
    for i, e in enumerate(entries, 1):
        results[e["id"]] = verifier.verify(e)
        status = results[e["id"]]["status"]
        print(f"      [{i}/{len(entries)}] {e['id']} {STATUS_ICON.get(status, '?')} "
              f"{(e['title'] or e['raw'])[:60]}")
        if i < len(entries) and not args.offline:
            time.sleep(args.delay)

    store = load_verdict_store()
    n_prior = 0
    for e in entries:
        prior = prior_verdict_for(e, store)
        if prior:
            results[e["id"]]["prior_verdict"] = prior
            n_prior += 1
    if n_prior:
        print(f"      ♻️ {n_prior} 条命中历史复核结论（见各条 prior_verdict 字段）")

    print("[5/5] 机械检查 + 生成报告")
    corr = check_correspondence(entries, citations)
    dups = check_duplicates(entries)
    timeline = check_timeline(entries)
    preprints = check_preprint(entries)
    cross = {"doi_swaps": check_doi_swaps(entries, results),
             "ordering": check_ordering(entries),
             "title_artifacts": check_title_artifacts(entries)}

    today = datetime.date.today().strftime("%Y%m%d")
    stem = os.path.splitext(os.path.basename(args.paper))[0]
    outdir = args.outdir or os.path.dirname(os.path.abspath(args.paper))
    _probe_writable(outdir)
    report_path = os.path.join(outdir, f"{stem}_refcheck_{today}.html")
    data_path = os.path.join(outdir, f"{stem}_refcheck_{today}.json")

    if _file_sha256(args.paper) != source_hash:
        sys.exit("[错误] 初筛期间稿件内容已变化，请保存稿件后重新初筛。")
    report = build_report(args.paper, entries, citations, results, corr,
                          dups, timeline, preprints, verifier.stats, cross)
    with open(report_path, "w", encoding="utf-8") as f:
        f.write(report)

    data = {
        "paper": {"path": os.path.abspath(args.paper),
                  "sha256": source_hash,
                  "report": report_path,
                  "checked_at": datetime.datetime.now().isoformat()},
        "entries": entries,
        "citations": citations,
        "verification": results,
        "correspondence": {
            "cited_but_missing_in_list": corr["cited_but_missing_in_list"],
            "listed_but_never_cited": corr["listed_but_never_cited"],
            "ambiguous_citations": corr["ambiguous_citations"]},
        "duplicates": dups,
        "timeline": timeline,
        "preprints": preprints,
        "cross_checks": cross,
        "summary_stats": {
            "entries": len(entries), "citations": len(citations),
            "triage": {"A": n_a, "B": n_b, "C": n_c},
            "prior_verdicts": n_prior,
            "verifier": verifier.stats,
            "source_capabilities": verifier.source_capabilities},
    }
    with open(data_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)

    print()
    print(f"✅ 自动初筛底稿: {report_path}")
    print(f"✅ 数据文件: {data_path}（Claude 层深查用）")
    nf = sum(1 for r in results.values() if r["status"] == "not_found")
    if nf:
        print(f"⚠️ {nf} 条自动未匹配：必须人工复核，不能据此断言为编造或真实缺失")
    # 初筛底稿不应打断用户；只有显式要求时才打开。
    if args.open_draft and sys.platform == "darwin":
        try:
            subprocess.run(["open", report_path], check=False)
        except OSError:
            pass


if __name__ == "__main__":
    main()
