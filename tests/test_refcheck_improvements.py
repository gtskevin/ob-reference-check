"""2026-08-28 复盘优化项的回归测试。

对应 docs/refcheck-retro-20260828.md:
C1 互换检测 / C2 venue 词级比对 / C3 排序 / C4 标题残留 /
C5 引语页码后缀 / C6 parser + DOI 兜底 / P0-1 定稿机制 / F5 证据必填 / F3 结论回流
"""

import importlib.util
import json
import pathlib
import tempfile
import unittest
from unittest import mock

SCRIPT_PATH = (pathlib.Path(__file__).resolve().parents[1]
               / "ob-reference-check" / "scripts" / "refcheck.py")
SPEC = importlib.util.spec_from_file_location("refcheck_imp", SCRIPT_PATH)
RC = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(RC)


def _entry(**kw):
    base = {"id": "R1", "raw": "Author, A. (2024). Title.",
            "authors": ["Author"], "year": 2024, "title": "Title",
            "venue": "Journal", "volume": None, "issue": None,
            "pages": None, "doi": None, "parse_ok": True}
    base.update(kw)
    return base


class ParserTest(unittest.TestCase):
    def test_title_ending_with_question_mark_does_not_swallow_venue(self):
        # R16 教训: 标题以 "?" 结尾时 venue 曾吞掉整个书目尾部
        e = RC.parse_entry(
            "Detert, J. R., & Burris, E. R. (2007). Leadership behavior and "
            "employee voice: Is the door really open? Academy of Management "
            "Journal, 50(4), 869-884. https://doi.org/10.5465/amj.2007.26279183",
            1)
        self.assertEqual(e["venue"], "Academy of Management Journal")
        self.assertEqual(e["volume"], "50")
        self.assertEqual(e["issue"], "4")
        self.assertEqual(e["pages"], "869-884")
        self.assertEqual(e["doi"], "10.5465/amj.2007.26279183")

    def test_volume_without_issue_pattern_and_url_residue(self):
        # R52 教训: "Journal, 28, 285-305" 无期号模式 + URL 残留
        e = RC.parse_entry(
            "McNaughton, N., & Corr, P. J. (2004). A two-dimensional "
            "neuropsychology of defense: Fear/anxiety and defensive distance. "
            "Neuroscience & Biobehavioral Review, 28, 285-305. "
            "https://doi.org/10.1016/j.neubiorev.2004.03.005", 1)
        self.assertEqual(e["venue"], "Neuroscience & Biobehavioral Review")
        self.assertEqual(e["volume"], "28")
        self.assertEqual(e["pages"], "285-305")
        self.assertNotIn("http", e["venue"])

    def test_mid_title_question_mark_keeps_subtitle(self):
        # R37 教训: "Who gets credit for input? Demographic ..." 的问号在
        # 标题内部，第一个终止符定界会把标题截断成 "Who gets credit for
        # input"，导致检索失败；必须以卷期模式位置为准
        e = RC.parse_entry(
            "McClean, S. T., et al. (2016). Who gets credit for input? "
            "Demographic and structural status differences in voice "
            "recognition. Journal of Applied Psychology, 111(6), 1111-1122.",
            1)
        self.assertEqual(
            e["title"],
            "Who gets credit for input? Demographic and structural status "
            "differences in voice recognition")
        self.assertEqual(e["venue"], "Journal of Applied Psychology")
        self.assertEqual(e["volume"], "111")
        self.assertEqual(e["issue"], "6")


class VenueCompareTest(unittest.TestCase):
    def _mm(self, paper_venue, db_venue):
        return RC._compare_metadata(
            _entry(venue=paper_venue),
            {"year": 2024, "authors": ["Author"], "venue": db_venue,
             "volume": None, "issue": None, "pages": None, "doi": None})

    def test_plural_typo_flagged_as_near_miss(self):
        # R52 教训: "review" ⊂ "reviews" 的子串容差曾静默吞掉拼写错误
        mm = self._mm("Neuroscience & Biobehavioral Review",
                      "Neuroscience & Biobehavioral Reviews")
        self.assertEqual(len(mm), 1)
        self.assertEqual(mm[0]["field"], "venue")
        self.assertEqual(mm[0]["level"], "near")

    def test_abbreviation_not_flagged(self):
        self.assertEqual(self._mm("J Appl Psychol",
                                  "Journal of Applied Psychology"), [])

    def test_word_abbrev_and_the_of_tolerated(self):
        self.assertEqual(self._mm("Academy of Management Journal",
                                  "The Academy of Management Journal"), [])

    def test_first_author_no_substring_tolerance(self):
        # "Li" ⊂ "Lin" 的子串容差会吞掉作者拼写错误
        mm = RC._compare_metadata(
            _entry(authors=["Li"]),
            {"year": 2024, "authors": ["Lin"], "venue": "Journal",
             "volume": None, "issue": None, "pages": None, "doi": None})
        self.assertEqual([m["field"] for m in mm], ["first_author"])


class CitationRegexTest(unittest.TestCase):
    def test_direct_quote_page_suffix_is_tolerated(self):
        # C5/R26 教训: (Galinsky & Moskowitz, 2000, p. 710) 曾被误报缺失
        cites = RC.extract_citations([
            {"text": "They said so (Galinsky & Moskowitz, 2000, p. 710) in "
                     "their seminal work. This sentence stands alone.",
             "heading": None, "level": 0}])
        self.assertEqual([(c["authors"], c["year"]) for c in cites],
                         [(["Galinsky", "Moskowitz"], "2000")])


class CrossCheckTest(unittest.TestCase):
    def test_doi_swap_detected(self):
        # C1/R55/R56 教训: 两条 DOI 互换只有交叉检测能发现
        entries = [
            _entry(id="A", doi="10.1/a", authors=["Morrison"], year=2014),
            _entry(id="B", doi="10.1/b", authors=["Morrison"], year=2011),
        ]
        results = {
            "A": {"record": {"doi": "10.1/b"}},
            "B": {"record": {"doi": "10.1/a"}},
        }
        swaps = RC.check_doi_swaps(entries, results)
        self.assertEqual(len(swaps), 1)
        self.assertIn("互换错挂", swaps[0]["issue"])

    def test_doi_swap_requires_actual_cross_hit(self):
        entries = [_entry(id="A", doi="10.1/a"), _entry(id="B", doi="10.1/b")]
        results = {"A": {"record": {"doi": "10.1/x"}},
                   "B": {"record": {"doi": "10.1/b"}}}
        self.assertEqual(RC.check_doi_swaps(entries, results), [])

    def test_ordering_only_for_identical_author_list(self):
        # C3: APA 只要求完整作者名单相同时按年份升序；同第一作者不同
        # 合作者按第二作者字母序，机械检查不应误报
        entries = [
            _entry(id="R55", authors=["Morrison"], year=2014),
            _entry(id="R56", authors=["Morrison"], year=2011),
            _entry(id="R1", authors=["Fast", "Chen"], year=2014),
            _entry(id="R2", authors=["Fast", "Binder"], year=2012),
        ]
        issues = RC.check_ordering(entries)
        self.assertEqual(len(issues), 1)
        self.assertIn("Morrison", issues[0]["issue"])

    def test_title_artifact_leading_digit(self):
        # C4/R49 教训: 标题残留章节号 "8 social hierarchy..."
        issues = RC.check_title_artifacts(
            [_entry(id="R49", title="8 social hierarchy: The self...")])
        self.assertEqual(len(issues), 1)
        self.assertEqual(RC.check_title_artifacts(
            [_entry(title="A social hierarchy")]), [])


class VerifyDoiFallbackTest(unittest.TestCase):
    def test_doi_lookup_used_when_all_searchers_fail(self):
        # C6①: 带有效 DOI 的条目不应因检索噪音被判 not_found
        verifier = RC.Verifier(cache_dir=tempfile.mkdtemp())
        verifier._search_openalex = lambda t, e: None
        verifier._search_crossref = lambda t, e: None
        verifier._search_s2 = lambda t, e: None
        verifier._crossref_doi_lookup = lambda doi: {
            "_source": "crossref-doi", "title": "A plausible title",
            "year": 2024, "authors": ["Author"], "venue": "Journal",
            "volume": None, "issue": None, "pages": None,
            "doi": "10.1/a", "type": "journal-article",
            "similarity": 1.0, "abstract": None}
        result = verifier._verify_online(
            _entry(doi="10.1/a"), "A plausible title")
        self.assertEqual(result["status"], "found")
        self.assertEqual(result["source"], "crossref-doi")

    def test_doi_fallback_short_db_title_counts_as_match(self):
        # R7 教训: 数据库短标题（无副标题）与论文全标题的纯比例打分
        # ratio≈0.45 < 0.6，同一篇文献被误标"疑似 DOI 错挂"；containment 命中即可
        verifier = RC.Verifier(cache_dir=tempfile.mkdtemp())
        verifier._search_openalex = lambda t, e: None
        verifier._search_crossref = lambda t, e: None
        verifier._search_s2 = lambda t, e: None
        verifier._crossref_doi_lookup = lambda doi: {
            "_source": "crossref-doi", "title": "Looking Out From the Top",
            "year": 2024, "authors": ["Author"], "venue": "Journal",
            "volume": None, "issue": None, "pages": None,
            "doi": "10.1/a", "type": "journal-article",
            "similarity": 1.0, "abstract": None}
        result = verifier._verify_online(
            _entry(doi="10.1/a"),
            "Looking out from the top: Differential effects of status and "
            "power on perspective taking")
        self.assertEqual(result["status"], "found")
        self.assertNotIn("错挂", result.get("note") or "")

    def test_doi_pointing_to_other_paper_flagged_as_misfiled(self):
        verifier = RC.Verifier(cache_dir=tempfile.mkdtemp())
        verifier._search_openalex = lambda t, e: None
        verifier._search_crossref = lambda t, e: None
        verifier._search_s2 = lambda t, e: None
        verifier._crossref_doi_lookup = lambda doi: {
            "_source": "crossref-doi", "title": "Completely different work",
            "year": 1999, "authors": ["Other"], "venue": "Journal",
            "volume": None, "issue": None, "pages": None,
            "doi": "10.1/x", "type": "journal-article",
            "similarity": 1.0, "abstract": None}
        result = verifier._verify_online(
            _entry(doi="10.1/x"), "A plausible title")
        self.assertEqual(result["status"], "not_found")
        self.assertIn("错挂", result["note"])


def _finalize_data():
    """最小定稿数据: 3 条文献，2 条带自动异常。"""
    return {
        "paper": {"path": "/tmp/paperX.docx", "report": "/tmp/x.html",
                  "checked_at": "2026-08-28T00:00:00"},
        "entries": [
            _entry(id="R1", raw="Author, A. (2024). Clean title."),
            _entry(id="R2", raw="Author, B. (2023). Flagged title."),
            _entry(id="R3", raw="Author, C. (2022). Also flagged."),
        ],
        "citations": [{"authors": ["Author"], "year": "2024", "triage": "A"}],
        "verification": {
            "R1": {"status": "found", "mismatches": [], "links": {}},
            "R2": {"status": "found", "links": {},
                   "mismatches": [{"field": "year", "paper": 2023,
                                   "database": 2022}]},
            "R3": {"status": "not_found", "mismatches": [], "links": {}},
        },
        "summary_stats": {"source_capabilities": {"openalex": "public",
                                                  "crossref": "public",
                                                  "semantic_scholar": "public"}},
    }


VERDICTS = [
    {"id": "R1", "final_status": "ok"},
    {"id": "R2", "final_status": "ok",
     "note": "曾自动标记年份差异，复核为 online-first 误报"},
    {"id": "R3", "final_status": "warn", "verdict": "DOI 错挂",
     "evidence": "Crossref 解析到另一文献", "action": "改为正确 DOI"},
]


class FinalizeTest(unittest.TestCase):
    def test_validation_rejects_warn_without_evidence(self):
        # F5: warn/info 结论必须带证据
        bad = [dict(VERDICTS[2])]
        del bad[0]["evidence"]
        with self.assertRaises(SystemExit):
            RC._validate_verdicts(_finalize_data()["entries"],
                                  _finalize_data()["verification"], bad)

    def test_validation_requires_verdict_for_auto_anomalies(self):
        # 自动初筛有异常的条目必须有显式结论，不允许静默跳过
        with self.assertRaises(SystemExit):
            RC._validate_verdicts(_finalize_data()["entries"],
                                  _finalize_data()["verification"],
                                  [{"id": "R1", "final_status": "ok"}])

    def test_badge_and_count_consistency(self):
        # P0-1: 徽章/计数由同一数据源驱动——不允许出现
        # 「badge ok 配 ⚠️」这类混搭（本次交付 bug 的回归点）
        html = RC.build_final_report(_finalize_data(), VERDICTS)
        self.assertNotIn('badge ok">⚠️', html)
        self.assertNotIn('badge warn">✅', html)
        self.assertIn('class="num bad">1', html)      # 必须修改 = R3
        self.assertIn('class="num ok">2', html)       # 其余确认 = R1/R2
        # P1-1 无过程栏目 / P1-3 折叠 / 检查范围小节在总览下（2026-08-28 用户反馈）
        self.assertIn("<details>", html)
        self.assertIn("本次检查范围", html)
        self.assertIn("文献存在性核验", html)
        # 检查范围必须完整列 8 类——零命中类别也保留（用户反馈×2：
        # 覆盖面展示是让用户放心的信息，不因无发现而静默省略）
        for cat in ("重复条目检测", "时间线与预印本检查", "列表内部一致性交叉检测",
                    "引用恰当性深查", "格式一致性与书目通读"):
            self.assertIn(cat, html)
        self.assertIn("未发现重复", html)   # _finalize_data 无 duplicates 命中
        # 有发现的类别用琥珀色 ⚠️ 而非绿勾（用户反馈×4：8 格全绿与
        # "必须处理 N 项"矛盾）
        self.assertIn('class="scope-warn"', html)
        self.assertIn("需修改 1 项", html)  # must = R3
        self.assertNotIn('✅ 全部一致或差异已排除', html)  # 有 must 时不显绿
        self.assertNotIn('class="todo"', html)

    def test_process_notes_not_rendered(self):
        # 2026-08-28 用户反馈: "曾自动标记…误报"类行内备注看不出指哪条，
        # 属中间过程，最终报告一律不渲染
        html = RC.build_final_report(_finalize_data(), VERDICTS)
        self.assertNotIn("online-first 误报", html)
        self.assertNotIn("曾自动标记", html)
        self.assertNotIn("复核为误报", html)

    def test_final_links_only_doi_and_scholar(self):
        # 2026-08-28 用户反馈: 复核入口按钮太多——只留 DOI + Scholar；
        # 无 DOI 的条目仅 Scholar 搜索；卡片里的复核按钮同样走 Scholar
        e = _entry(doi="10.1/a", title="Some Title")
        result = {"links": {"doi": "https://doi.org/10.1/a",
                            "openalex": "https://oa/1", "s2": "https://s2/1",
                            "google_scholar": "https://scholar?q=1"}}
        self.assertEqual(RC._final_links(result, e),
                         {"doi": "https://doi.org/10.1/a",
                          "google_scholar": "https://scholar?q=1"})
        no_doi = RC._final_links({"links": {"openalex_search": "https://oa?s=1"}},
                                 _entry(title="QueryTitle"))
        self.assertEqual(set(no_doi), {"google_scholar"})
        self.assertIn("QueryTitle", no_doi["google_scholar"])

    def test_review_links_and_footer_disclaimer(self):
        # 渲染层: 报告中不出现 OpenAlex/Semantic Scholar 按钮；附录说明
        # ⚠️/❓ 含义；footer 带警示色的人工复核提醒（2026-08-28 用户反馈）
        data = _finalize_data()
        data["verification"]["R1"]["links"] = {
            "doi": "https://doi.org/10.1/a", "openalex": "https://oa/1",
            "s2": "https://s2/1", "google_scholar": "https://scholar?q=x"}
        data["verification"]["R3"]["links"] = {"openalex_search": "https://oa?s=q"}
        html = RC.build_final_report(data, VERDICTS)
        self.assertNotIn("OpenAlex ↗", html)
        self.assertNotIn("Semantic Scholar ↗", html)
        self.assertIn("DOI ↗", html)
        self.assertIn("Scholar 搜索 ↗", html)
        self.assertIn("确认需修改或存疑", html)      # ⚠️ 含义说明
        self.assertIn("建议人工核对", html)          # ❓ 含义说明
        self.assertIn("foot-warn", html)
        self.assertIn("建议经人工复核后再交付", html)

    def test_finalize_accepts_final_json_directly(self):
        # 2026-08-28 实测踩坑: 直接传 *_final.json 曾 KeyError('paper')——
        # 现在自动定位同目录初筛数据
        with tempfile.TemporaryDirectory() as td:
            data = _finalize_data()
            data["paper"]["path"] = str(pathlib.Path(td) / "paperX.docx")
            rc_path = pathlib.Path(td) / "paperX_refcheck_20260828.json"
            rc_path.write_text(json.dumps(data), encoding="utf-8")
            final_path = pathlib.Path(td) / "paperX_final.json"
            final_path.write_text(json.dumps({"verdicts": VERDICTS}),
                                  encoding="utf-8")
            got_path, got = RC._load_refcheck_json(str(final_path))
            self.assertEqual(got_path, str(rc_path))
            self.assertIn("entries", got)

    def test_finalize_persists_verdicts_by_doi(self):
        # F3: 定稿后结论按 DOI 回流 verdict store
        with tempfile.TemporaryDirectory() as td:
            store_path = str(pathlib.Path(td) / "verdicts.json")
            data = _finalize_data()
            data["entries"][2]["doi"] = "10.1/c"
            v2 = [dict(v) for v in VERDICTS]
            RC._validate_verdicts(data["entries"], data["verification"], v2)
            with mock.patch.object(RC, "VERDICT_STORE", store_path):
                store = RC.load_verdict_store()
                store["10.1/c"] = {"final_status": v2[2]["final_status"]}
                RC.save_verdict_store(store)
                self.assertEqual(RC.load_verdict_store()["10.1/c"]
                                 ["final_status"], "warn")


def _corr_data():
    """带"正文引用但列表缺失"的定稿数据: C2 (Pop et al., 2015) 无对应条目。"""
    data = _finalize_data()
    data["citations"] = [
        {"cid": "C1", "authors": ["Author"], "year": "2024",
         "sentence": "Author (2024) shows this.", "triage": "A"},
        {"cid": "C2", "authors": ["Pop", "Kang"], "year": "2015",
         "sentence": "Prior work follows Pop et al. (2015) in this design.",
         "triage": "B"},
    ]
    data["correspondence"] = {
        "cited_but_missing_in_list": [data["citations"][1]],
        "listed_but_never_cited": [],
    }
    return data


CORR_VERDICTS = VERDICTS + [
    {"id": "C2", "category": "correspondence", "final_status": "warn",
     "verdict": "正文引用了 Pop et al. (2015)，但参考文献列表中没有对应条目",
     "evidence": "引用出现在方法部分；列表按姓+年份检索无匹配",
     "action": "补充完整条目到参考文献列表，或删除该正文引用"},
]


class CorrespondenceVerdictTest(unittest.TestCase):
    """2026-08-28 缺口修复: "正文引用但列表缺失"通过 C 编号 verdict
    进入最终报告（docs/correspondence-gap-and-fix.md 方案 A）。"""

    def test_validation_accepts_cid_verdict(self):
        RC._validate_verdicts(_corr_data()["entries"],
                              _corr_data()["verification"], CORR_VERDICTS,
                              _corr_data()["citations"],
                              _corr_data()["correspondence"])

    def test_validation_rejects_out_of_range_cid(self):
        bad = [dict(CORR_VERDICTS[3])]
        bad[0]["id"] = "C9"
        with self.assertRaises(SystemExit):
            RC._validate_verdicts(_corr_data()["entries"],
                                  _corr_data()["verification"], bad,
                                  _corr_data()["citations"],
                                  _corr_data()["correspondence"])

    def test_validation_requires_verdict_for_missing_citation(self):
        # 每条缺失引用必须有显式结论——只给条目 verdict 会拦截
        with self.assertRaises(SystemExit):
            RC._validate_verdicts(_corr_data()["entries"],
                                  _corr_data()["verification"], VERDICTS,
                                  _corr_data()["citations"],
                                  _corr_data()["correspondence"])

    def test_validation_locates_old_json_citation_without_cid(self):
        # 旧缓存 JSON 的 correspondence 项无 cid 字段——按内容回查 citations
        data = _corr_data()
        legacy = {k: v for k, v in data["citations"][1].items() if k != "cid"}
        data["correspondence"]["cited_but_missing_in_list"] = [legacy]
        RC._validate_verdicts(data["entries"], data["verification"],
                              CORR_VERDICTS, data["citations"],
                              data["correspondence"])

    def test_missing_citation_rendered_in_final_report(self):
        # C 编号卡片进入"必须处理"，显示引用句原文；"其余确认"按条目计数
        # 不被 C verdict 折减
        html = RC.build_final_report(_corr_data(), CORR_VERDICTS)
        self.assertIn("文献列表中无对应条目", html)
        self.assertIn("Prior work follows Pop et al. (2015)", html)
        self.assertIn("badge warn\">C2", html)
        self.assertIn("scholar.google.com", html)
        self.assertIn('class="num bad">2', html)      # R3 + C2
        self.assertIn('class="num ok">2', html)       # 其余确认仍是 R1/R2
        self.assertIn("1 项对应问题", html)            # 检查范围: 对应类 ⚠️

    def test_match_false_positive_resolved_as_ok(self):
        # 复核为匹配误报（如年份不一致）→ ok，不进必须处理
        verdicts = [dict(v) for v in VERDICTS] + [
            {"id": "C2", "category": "correspondence", "final_status": "ok",
             "note": "引用 2015 为 online-first 年份，对应列表 R1"}]
        html = RC.build_final_report(_corr_data(), verdicts)
        self.assertIn('class="num bad">1', html)      # 仅 R3
        self.assertIn('class="num ok">2', html)


if __name__ == "__main__":
    unittest.main()
