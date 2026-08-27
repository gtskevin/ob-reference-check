#!/usr/bin/env python3
"""
ob-reference-check — 论文参考文献系统检查（机械层）

用法:
    python refcheck.py <论文文件 .docx/.pdf/.md> [选项]

做什么（脚本层，零 LLM token）:
    1. 解析 Word / PDF / Markdown 三种格式
    2. 提取参考文献列表条目 + 正文引用标记
    3. OpenAlex / Crossref / Semantic Scholar 三源轮换验证 + 元数据比对（带全局缓存）
    4. 机械检查: 双向对应 / 重复条目 / 时间线异常 / preprint 版本
    5. A/B/C 分诊初筛（承重引用 / 顺带提及 / 引用堆砌）
    6. 生成自包含 HTML 报告（空问题模块不渲染）+ .json 数据文件（供 Claude 层继续深查）

不做什么:
    - 引用恰当性判断（层 3）→ 由 Claude 读 .json 中的句子+摘要完成
    - 引用格式一致性审查 → 由 Claude 完成
"""

import argparse
import datetime
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
        # 新条目开头: 大写字母开头 + 含 4 位年份（author-year style 通用特征）
        if _looks_like_entry_start(t) and current:
            entries.append(" ".join(current))
            current = [t]
        else:
            current.append(t)
    if current:
        entries.append(" ".join(current))
    return entries, ref_para_idx


_AUTHOR_START = re.compile(r"^[A-ZÀ-ÿ一-鿿][\w'’\-,. &]+")


def _looks_like_entry_start(t):
    return bool(_AUTHOR_START.match(t) and re.search(r"\d{4}", t) and len(t) > 20)


DOI_RE = re.compile(r"(10\.\d{4,9}/[^\s\"'\]\)]+)")


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
    ym = re.search(r"[ (.](\d{4})[a-z]?[).,]", raw + " ")
    if ym:
        e["year"] = int(ym.group(1))
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
                    if re.search(r"[a-zà-ÿ]{2}", a.strip())]

    # 标题: 年份括号之后到下一个句号分句
    rest = raw[ym.end():] if ym else raw[min(len(authors_str), len(raw)):]
    rest = rest.lstrip("). ").strip()
    title_m = re.match(r"^(.*?)(?:\.\s|\.$)", rest)
    e["title"] = (title_m.group(1).strip() if title_m else rest.split(".")[0]).strip()
    if len(e["title"]) < 8:
        e["title"] = rest[:120]
        e["parse_ok"] = False  # 标题解析可疑，让 Claude 兜底

    # 期刊/卷/期/页: 标题之后的剩余部分
    tail = rest[len(e["title"]):].lstrip(". ")
    vm = re.search(r"[, ](\d+)\s*\((\d+)\)", tail)
    if vm:
        e["volume"], e["issue"] = vm.group(1), vm.group(2)
        e["venue"] = tail[:vm.start()].strip(" ,.")
    else:
        vm2 = re.search(r"[, ](\d+):", tail)
        if vm2:
            e["volume"] = vm2.group(1)
            e["venue"] = tail[:vm2.start()].strip(" ,.")
        else:
            v = DOI_RE.sub("", tail).strip(" ,.")
            e["venue"] = v if 0 < len(v) < 120 else None
    pm = re.search(r"(\d+\s*[-–—]\s*\d+)", tail)
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
    r"([A-ZÀ-ÿ][\w'’\-]+(?:\s+(?:et al\.?|&|and)\s+[A-ZÀ-ÿ][\w'’\-]+)*)"
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
    for p in body_paragraphs:
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
                    im = INNER_CITE.match(part)
                    if not im:
                        continue
                    author_part, year_part = im.group(1).strip(), im.group(2)
                    years = re.findall(r"\d{4}[a-z]?", year_part)
                    surnames = _surnames_from_inline(author_part)
                    for y in years:
                        found.append({"authors": surnames, "year": y,
                                      "raw": part, "narrative": False})
            for m in NARR_CITE.finditer(sent):
                surnames = _surnames_from_inline(m.group(1))
                found.append({"authors": surnames, "year": m.group(2),
                              "raw": f"{m.group(1)} ({m.group(2)})", "narrative": True})
            for c in found:
                c["sentence"] = sent
                c["section"] = section
                citations.append(c)

    # 分诊: C 类=单括号内 ≥3 条并排; A 类=句子含假设/理论论证标记; B 类=其余
    for c in citations:
        sent = c["sentence"]
        pm = PAREN_CITE.search(sent)
        n_group = len(CITE_SPLIT.split(pm.group(1))) if pm else 1
        if n_group >= 3:
            c["triage"] = "C"
        elif HYPOTHESIS_MARKERS.search(sent):
            c["triage"] = "A"
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
        part = part.strip(" .")
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
        self.crossref_key = _env_any("CROSSREF_API_KEY")
        self.s2_key = _env_any("SEMANTIC_API_KEY", "SEMANTIC_SCHOLAR_API_KEY",
                               "S2_API_KEY")
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
        key = "v2:" + re.sub(r"\W+", " ", title.lower()).strip()[:200]
        cached = self._cache_get(key)
        if cached and self.offline:
            return cached["data"]
        if self.offline:
            return self._unverified(entry, "离线模式且无缓存")
        if cached and not self._stale(cached):
            self.stats["cache_hit"] += 1
            return cached["data"]

        result = self._verify_online(entry, title)
        if result.get("status") != "error":
            self._cache_put(key, result)
        return result

    @staticmethod
    def _stale(cached):
        age = (datetime.date.today()
               - datetime.date.fromisoformat(cached["cached_at"])).days
        return age > 180  # 半年后允许刷新（收录延迟 / 更正）

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
            return {"status": "not_found", "confidence": "high",
                    "source": "openalex+crossref+semanticscholar", "record": None,
                    "mismatches": [], "links": self._search_links(entry),
                    "abstract": None, "note": "各数据库均无匹配"}
        mismatches = _compare_metadata(entry, rec)
        return {"status": "found", "confidence": "medium" if mismatches else "high",
                "source": rec["_source"], "record": rec,
                "mismatches": mismatches, "links": self._record_links(rec, entry),
                "abstract": rec.get("abstract")}

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
            links["openalex_search"] = f"https://openalex.org/works?search={q}"
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


def _compare_metadata(entry, rec):
    """逐项比对，返回 mismatch 列表 [{field, paper, database}]"""
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
        if pl in dl or dl in pl:
            return  # 简称/缩写容差（如 "Human Factors" ⊂ 数据库全称）
        sim = difflib.SequenceMatcher(None, pl, dl).ratio()
        if sim < threshold:
            out.append({"field": field, "paper": p, "database": d})

    add("year", entry["year"], rec["year"])
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

def entry_key(e):
    # 年份统一为字符串，避免 int/str 元组不匹配
    return (re.sub(r"\W+", "", (e["authors"][0] if e["authors"] else "").lower()),
            str(e["year"]) if e["year"] else None)


def check_correspondence(entries, citations):
    entry_keys = {}
    for e in entries:
        entry_keys.setdefault(entry_key(e), []).append(e["id"])

    cited_keys = set()
    unmatched_citations = []
    for c in citations:
        matched = None
        cyear = c["year"].rstrip("abcdefghij")
        for surname in c["authors"]:
            if (re.sub(r"\W+", "", surname.lower()), cyear) in entry_keys:
                matched = (re.sub(r"\W+", "", surname.lower()), cyear)
                break
        if matched:
            cited_keys.add(matched)
        else:
            unmatched_citations.append(c)

    uncited = [e["id"] for e in entries if entry_key(e) not in cited_keys]
    return {"cited_but_missing_in_list": unmatched_citations,
            "listed_but_never_cited": uncited}


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
                           "issue": f"出版年份 {e['year']} 晚于当前年份，不可能存在"})
    return issues


def check_preprint(entries):
    issues = []
    for e in entries:
        venue = e["venue"] or ""
        if PREPRINT_MARKERS.search(venue):
            issues.append({"id": e["id"],
                           "issue": f"引用的是 preprint（{venue}）——正式版可能已发表，建议更新"})
    return issues

# ---------------------------------------------------------------------------
# 6. 报告生成（独立 HTML，自包含样式，浏览器直接打开）
# ---------------------------------------------------------------------------

STATUS_ICON = {"found": "✅", "not_found": "❌", "unverified": "❓"}

REPORT_CSS = """
:root {
  --ok: #047857; --bad: #b91c1c; --warn: #b45309; --info: #1d4ed8;
  --gray: #64748b; --ink: #1e293b; --border: #e2e8f0; --bg: #f8fafc;
  --card: #ffffff; --nav-h: 48px;
  --serif: Georgia, "Times New Roman", "Songti SC", "Noto Serif SC", serif;
  --mono: ui-monospace, "SF Mono", SFMono-Regular, Menlo, monospace;
}
* { box-sizing: border-box; margin: 0; padding: 0; }
@media (prefers-reduced-motion: no-preference) {
  html { scroll-behavior: smooth; }
}
body {
  font-family: -apple-system, "PingFang SC", "Microsoft YaHei", "Segoe UI", sans-serif;
  background: var(--bg); color: var(--ink); font-size: 14px; line-height: 1.75;
  padding: 24px 16px 60px;
}
.container { max-width: 1080px; margin: 0 auto; }
header.paper { background: #12283f; color: #f8fafc; border-radius: 12px;
  padding: 28px 32px 24px; margin-bottom: 14px;
  border-top: 3px solid #b08d3e; }
header.paper h1 { font-family: var(--serif); font-size: 24px; font-weight: 700;
  letter-spacing: .02em; }
header.paper .meta { font-size: 13px; opacity: .78; margin-top: 6px;
  font-variant-numeric: tabular-nums; }

/* 吸顶导航 */
.topnav { position: sticky; top: 0; z-index: 50; background: rgba(255,255,255,.94);
  backdrop-filter: blur(8px); border: 1px solid var(--border); border-radius: 10px;
  padding: 7px 12px; margin-bottom: 18px; display: flex; flex-wrap: wrap;
  gap: 4px; align-items: center; }
.topnav a { font-size: 13px; color: #334155; text-decoration: none;
  padding: 3px 10px; border-radius: 999px; white-space: nowrap; }
.topnav a:hover { background: #f1f5f9; text-decoration: none; color: var(--ink); }
.topnav a .cnt { font-size: 12px; font-weight: 700; margin-left: 3px;
  color: #94a3b8; font-variant-numeric: tabular-nums; }
.topnav a .cnt.hot-bad { color: var(--bad); }
.topnav a .cnt.hot-warn { color: var(--warn); }
.topnav a .cnt.hot-info { color: var(--info); }

/* 汇总卡片 */
.summary-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(155px, 1fr));
  gap: 10px; margin-bottom: 22px; }
.stat-card { background: var(--card); border: 1px solid var(--border);
  border-radius: 10px; padding: 13px 16px; text-align: center; }
.stat-card .num { font-size: 26px; font-weight: 700;
  font-variant-numeric: tabular-nums; }
.stat-card .num.zero { font-size: 19px; color: #94a3b8; }
.stat-card .label { font-size: 12px; color: var(--gray); margin-top: 3px; }
.num.ok { color: var(--ok); } .num.bad { color: var(--bad); }
.num.warn { color: var(--warn); } .num.info { color: var(--info); }

/* 区块 */
section { background: var(--card); border: 1px solid var(--border);
  border-radius: 12px; padding: 20px 26px 10px; margin-bottom: 16px; }
section h2 { font-family: var(--serif); font-size: 18px; font-weight: 700;
  padding-bottom: 10px; margin-bottom: 4px;
  border-bottom: 1px solid var(--border); }
section h2 .cnt2 { font-family: -apple-system, "PingFang SC", sans-serif;
  font-size: 13px; font-weight: 700; color: var(--gray);
  margin-left: 8px; font-variant-numeric: tabular-nums; }
section h3 { font-size: 14px; font-weight: 600; margin: 10px 0 2px; color: #334155; }

/* 详情行: 拍平的 hairline 行，不再套卡片 */
.item { padding: 13px 2px 13px 14px; border-bottom: 1px solid var(--border);
  border-left: 3px solid transparent; }
.item:last-child, .item:last-of-type { border-bottom: none; }
.item.crit { border-left-color: var(--bad); }
.item.warn { border-left-color: var(--warn); }
.item.info { border-left-color: var(--info); }
.item h3 { margin: 0 0 5px; font-size: 14px; font-weight: 600;
  word-break: break-all; }
.item p { font-size: 13.5px; }
.muted { color: var(--gray); font-size: 12.5px; }

/* 徽章: 描边样式，降饱和 */
.badge { display: inline-block; font-size: 12px; padding: 1px 9px;
  border-radius: 999px; font-weight: 600; margin-right: 6px;
  white-space: nowrap; background: #fff; border: 1px solid var(--border);
  color: var(--gray); }
.badge.ok { color: var(--ok); border-color: #a7d8c5; }
.badge.bad { color: var(--bad); border-color: #e8c4c4; }
.badge.warn { color: var(--warn); border-color: #e8d4b0; }
.badge.info { color: var(--info); border-color: #c3d3f0; }

/* 表格 */
table { width: 100%; border-collapse: collapse; font-size: 13.5px; margin: 8px 0; }
th { text-align: left; background: #f8fafc; padding: 8px 12px; font-size: 12.5px;
  color: #475569; border-bottom: 1px solid var(--border); }
td { padding: 8px 12px; border-bottom: 1px solid var(--border); vertical-align: top;
  font-variant-numeric: tabular-nums; }
td:first-child { font-family: var(--mono); font-size: 12.5px; }
td.val-paper { color: var(--bad); }
td.val-db { color: var(--ok); font-weight: 700; }

/* 附录长表: 表头吸顶 */
.appendix thead th { position: sticky; top: var(--nav-h); z-index: 10;
  background: #f8fafc; box-shadow: 0 1px 0 var(--border); }
.appendix td.ref-text { font-size: 12.5px; }

/* 链接按钮: 中性描边，DOI 用深墨实心 */
.lnk { display: inline-block; font-size: 12px; padding: 2px 10px;
  border-radius: 6px; border: 1px solid var(--border); color: #475569;
  background: #fff; text-decoration: none; margin: 1px 3px 1px 0;
  white-space: nowrap; }
.lnk:hover { border-color: #94a3b8; color: var(--ink); text-decoration: none; }
.lnk.doi { background: var(--ink); color: #fff; border-color: var(--ink); }
.lnk.doi:hover { background: #0f172a; }
a { color: var(--info); text-decoration: none; }
a:focus-visible { outline: 2px solid var(--info); outline-offset: 2px;
  border-radius: 4px; }

.todo { border: 1px dashed #cbd5e1; border-radius: 10px; padding: 14px 18px;
  color: var(--gray); font-size: 13.5px; margin: 10px 0 14px; background: #fafbfc; }
footer { text-align: center; color: #94a3b8; font-size: 12px; margin-top: 30px;
  line-height: 1.6; }
@media (max-width: 640px) {
  body { padding: 14px 8px 40px; }
  section { padding: 14px 14px 6px; border-radius: 9px; }
  table { display: block; overflow-x: auto; white-space: nowrap; }
  td { white-space: normal; }
}
@media print {
  body { background: #fff; padding: 0; }
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


def build_report(paper_path, entries, citations, results, corr, dups,
                 timeline, preprints, verifier_stats):
    today = datetime.date.today().isoformat()
    stem = os.path.splitext(os.path.basename(paper_path))[0]
    esc = html_mod.escape
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
        nf_rows.append(f"""<div class="item crit">
<h3>{_badge(eid, 'bad')} {esc(e['raw'][:150])}</h3>
<p class="muted">各数据库均无匹配（标题相似度 &lt; 0.75），建议人工确认是否为 AI 编造。</p>
<p>复核：{_links_html(r.get('links') or {})}</p>
</div>""")
    if nf_rows:
        sections.append(("sec-notfound", "❌ 疑似不存在的文献", "疑似编造",
                         n_notfound, "hot-bad", "".join(nf_rows)))

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
<p>复核：{_links_html(r.get('links') or {})}</p>
</div>""")
    if mm_rows:
        sections.append(("sec-meta", "⚠️ 元数据不一致", "元数据",
                         n_mm, "hot-warn", "".join(mm_rows)))

    corr_rows = []
    if corr["cited_but_missing_in_list"]:
        corr_rows.append("<h3>正文引了，但列表里没有</h3>")
        for c in corr["cited_but_missing_in_list"]:
            corr_rows.append(f"""<div class="item info">
<p>“{esc(c['sentence'][:180])}”</p>
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

    # 恰当性 / 格式两节占位（Claude 层填充）
    sections.append(("sec-appro", "⚠️ 引用恰当性（A 类深查）", "恰当性", None, "",
                     '<div class="todo">待 Claude 层补入：A 类承重引用逐条比对'
                     '（正文句子 vs 文献摘要），输出 支撑 / 存疑 / 摘要不足 三档判断。</div>'))
    sections.append(("sec-format", "📝 格式一致性", "格式", None, "",
                     '<div class="todo">待 Claude 层补入：同一 style 内部统一性审查'
                     '（et al. 规则、&amp; vs and、页码符号、期刊名缩写等）。</div>'))

    # ---- 组装 ----
    H = []
    nav_links = ['<a href="#overview">总览</a>']
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
  <h1>参考文献检查报告</h1>
  <div class="meta">{esc(stem)} ｜ 检查日期 {today} ｜
  文献列表 {len(entries)} 条 ｜ 正文引用 {len(citations)} 处</div>
</header>

<nav class="topnav">{''.join(nav_links)}</nav>

<div class="summary-grid" id="overview">""")

    n_ok = n_found - n_mm
    cards = [
        (n_ok, "ok" if n_ok else "zero", "✅ 确认存在且一致"),
        (n_notfound, "bad" if n_notfound else "zero", "❌ 疑似不存在"),
        (n_mm, "warn" if n_mm else "zero", "⚠️ 元数据不一致"),
        (n_unver, "info" if n_unver else "zero", "❓ 无法验证"),
        (n_corr_cited, "info" if n_corr_cited else "zero", "🔗 引用但列表缺失"),
        (n_corr_listed, "info" if n_corr_listed else "zero", "🔗 列表有但未引用"),
        (len(dups), "warn" if dups else "zero", "⧉ 重复条目"),
        (len(timeline) + len(preprints), "warn" if (timeline or preprints) else "zero", "⏰ 时间线/preprint"),
    ]
    for num, cls, label in cards:
        H.append(f'<div class="stat-card"><div class="num {cls}">{num}</div>'
                 f'<div class="label">{label}</div></div>')
    H.append('</div>')

    for sid, title, _short, cnt, _hot, body in sections:
        cnt_html = f'<span class="cnt2">{cnt}</span>' if cnt is not None else ""
        H.append(f'<section id="{sid}"><h2>{title}{cnt_html}</h2>{body}</section>')

    # 附录: 全部文献核验清单（问题条目排前面）
    H.append('<section id="appendix">'
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
        icon = {"found": "ok", "not_found": "bad", "unverified": "gray"}[status]
        label = STATUS_ICON.get(status, "❓")
        if r.get("mismatches"):
            icon, label = "warn", "⚠️"
        desc = esc(e["raw"][:90] + ("…" if len(e["raw"]) > 90 else ""))
        H.append(f'<tr><td>{esc(e["id"])}</td>'
                 f'<td class="ref-text">{desc}</td>'
                 f'<td>{_badge(label, icon)}</td>'
                 f'<td>{_links_html(r.get("links") or {})}</td></tr>')
    H.append("</tbody></table></section>")

    H.append(f"""<footer>脚本统计: {esc(json.dumps(verifier_stats, ensure_ascii=False))}
<br>ob-reference-check 机械层报告（恰当性 / 格式两节由 Claude 层补入）</footer>
</div></body></html>""")
    return "\n".join(H)




# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description="论文参考文献机械检查")
    ap.add_argument("paper", help="论文文件 (.docx/.pdf/.md)")
    ap.add_argument("--offline", action="store_true",
                    help="只用本地缓存，不访问网络")
    ap.add_argument("--delay", type=float, default=0.15,
                    help="API 请求间隔秒数（礼貌限速）")
    args = ap.parse_args()

    if not os.path.exists(args.paper):
        sys.exit(f"[错误] 文件不存在: {args.paper}")

    print(f"[1/5] 解析文档: {args.paper}")
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

    print("[5/5] 机械检查 + 生成报告")
    corr = check_correspondence(entries, citations)
    dups = check_duplicates(entries)
    timeline = check_timeline(entries)
    preprints = check_preprint(entries)

    today = datetime.date.today().strftime("%Y%m%d")
    stem = os.path.splitext(os.path.basename(args.paper))[0]
    outdir = os.path.dirname(os.path.abspath(args.paper))
    report_path = os.path.join(outdir, f"{stem}_refcheck_{today}.html")
    data_path = os.path.join(outdir, f"{stem}_refcheck_{today}.json")

    report = build_report(args.paper, entries, citations, results, corr,
                          dups, timeline, preprints, verifier.stats)
    with open(report_path, "w", encoding="utf-8") as f:
        f.write(report)

    data = {
        "paper": {"path": os.path.abspath(args.paper),
                  "report": report_path,
                  "checked_at": datetime.datetime.now().isoformat()},
        "entries": entries,
        "citations": citations,
        "verification": results,
        "correspondence": {
            "cited_but_missing_in_list": corr["cited_but_missing_in_list"],
            "listed_but_never_cited": corr["listed_but_never_cited"]},
        "duplicates": dups,
        "timeline": timeline,
        "preprints": preprints,
        "summary_stats": {
            "entries": len(entries), "citations": len(citations),
            "triage": {"A": n_a, "B": n_b, "C": n_c},
            "verifier": verifier.stats},
    }
    with open(data_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)

    print()
    print(f"✅ 检查报告: {report_path}")
    print(f"✅ 数据文件: {data_path}（Claude 层深查用）")
    nf = sum(1 for r in results.values() if r["status"] == "not_found")
    if nf:
        print(f"🚨 发现 {nf} 条疑似不存在的文献，见报告「疑似不存在的文献」一节")
    # macOS 下自动在浏览器打开（报告是自包含 HTML）
    if sys.platform == "darwin":
        try:
            subprocess.run(["open", report_path], check=False)
        except OSError:
            pass


if __name__ == "__main__":
    main()
