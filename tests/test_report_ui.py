"""Regression checks for report navigation and AI follow-up guidance."""

import importlib.util
import os
import pathlib
import tempfile
import unittest
from html.parser import HTMLParser
from unittest import mock


SCRIPT_PATH = (pathlib.Path(__file__).resolve().parents[1]
               / "ob-reference-check" / "scripts" / "refcheck.py")
FIXTURE_PATH = (pathlib.Path(__file__).resolve().parent / "fixtures"
                / "test_paper_refcheck_20260828.html")
SPEC = importlib.util.spec_from_file_location("refcheck", SCRIPT_PATH)
REFCHECK = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(REFCHECK)


class NavigationParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self._in_navigation = False
        self.hrefs = []
        self.ids = set()

    def handle_starttag(self, tag, attrs):
        attributes = dict(attrs)
        if attributes.get("id"):
            self.ids.add(attributes["id"])
        if tag == "nav" and "topnav" in attributes.get("class", "").split():
            self._in_navigation = True
        if self._in_navigation and tag == "a" and attributes.get("href"):
            self.hrefs.append(attributes["href"])

    def handle_endtag(self, tag):
        if tag == "nav":
            self._in_navigation = False


class ReportUiTest(unittest.TestCase):
    def test_report_has_desktop_rail_mobile_fallback_and_ai_opt_in(self):
        report = REFCHECK.build_report(
            "paper.md",
            [{"id": "R1", "raw": "Author, A. (2024). Title.", "doi": None,
              "title": "Title", "year": "2024", "venue": "Journal"}],
            [],
            {"R1": {"status": "found", "mismatches": [], "links": {}}},
            {"cited_but_missing_in_list": [], "listed_but_never_cited": []},
            [], [], [], {},
        )

        self.assertIn('<div class="report-layout">', report)
        self.assertIn('<main class="report-main">', report)
        self.assertIn('</main>\n</div>', report)
        self.assertIn('grid-template-columns: 190px minmax(0, 1fr)', report)
        self.assertIn('top: 24px; display: block; margin: 0; padding: 15px 0;', report)
        self.assertIn('border-top: 2px solid #315e55;', report)
        self.assertIn('@media (max-width: 640px)', report)
        self.assertIn('.report-layout { display: block; }', report)
        self.assertIn('position: sticky; top: 0; display: flex; flex-wrap: wrap;', report)
        self.assertIn('.appendix thead th { position: static; }', report)
        self.assertIn('section, #overview { scroll-margin-top: 180px; }', report)
        self.assertIn('@media print {\n  body { background: #fff; padding: 0; }\n  .report-layout { display: block; }',
                      report)
        self.assertNotIn('class="return-nav"', report)
        self.assertIn('可继续由 AI 审读', report)
        self.assertEqual(report.count('可继续由 AI 审读。'), 2)
        self.assertIn('继续审读引用恰当性和格式一致性。', report)

        parser = NavigationParser()
        parser.feed(report)
        self.assertTrue(parser.hrefs)
        for href in parser.hrefs:
            self.assertTrue(href.startswith("#"))
            self.assertIn(href[1:], parser.ids)

    def test_found_record_uses_google_scholar_search_when_no_direct_record_link(self):
        links = REFCHECK.Verifier._record_links(
            {"_source": "crossref", "title": "A verified reference", "doi": None},
            {"raw": "Author, A. (2024). A verified reference."},
        )

        self.assertIn("google_scholar", links)
        self.assertNotIn("openalex_search", links)
        self.assertIn("scholar.google.com/scholar?q=", links["google_scholar"])

    def test_found_record_replaces_cached_openalex_search_in_report(self):
        report = REFCHECK.build_report(
            "cached.md",
            [{"id": "R1", "raw": "Author, A. (2024). Cached title.", "doi": None,
              "title": "Cached title", "year": "2024", "venue": "Journal"}],
            [],
            {"R1": {"status": "found", "mismatches": [],
                    "record": {"title": "Cached title"},
                    "links": {"openalex_search": "https://openalex.org/works?search=Cached"}}},
            {"cited_but_missing_in_list": [], "listed_but_never_cited": []},
            [], [], [], {},
        )

        self.assertIn("Scholar 搜索 ↗", report)
        self.assertNotIn("OpenAlex 搜索 ↗", report)

    def test_fixture_retains_navigation_and_ai_follow_up_guidance(self):
        report = FIXTURE_PATH.read_text(encoding="utf-8")

        self.assertIn('<div class="report-layout">', report)
        self.assertIn('<main class="report-main">', report)
        self.assertIn('可继续由 AI 审读。', report)
        self.assertIn('继续审读引用恰当性和格式一致性。', report)

    def test_severity_preserves_unverified_and_warning_issues(self):
        report = REFCHECK.build_report(
            "severity.md",
            [
                {"id": "R1", "raw": "Author, A. (2024). Unverified.", "doi": None,
                 "title": "Unverified", "year": "2024", "venue": "Journal"},
                {"id": "R2", "raw": "Author, B. (2024). Duplicate.", "doi": None,
                 "title": "Duplicate", "year": "2024", "venue": "Journal"},
            ],
            [],
            {
                "R1": {"status": "unverified", "mismatches": [], "links": {}},
                "R2": {"status": "found", "mismatches": [], "links": {}},
            },
            {"cited_but_missing_in_list": [], "listed_but_never_cited": []},
            [{"ids": ["R1", "R2"], "by": "title"}], [], [], {},
        )

        self.assertIn(
            '<div class="num warn">2</div><div class="label">🔗 对应、版本或待复核问题</div>',
            report,
        )
        self.assertIn('class="item warn"', report)
        self.assertIn('class="badge warn"', report)
        self.assertIn('<span class="scope-info">无法确认 1 项：见下方详情</span>', report)
        self.assertNotIn('<span class="scope-warn">发现 1 项：见下方详情</span>'
                         '<span>通过学术数据库核对每一条参考文献。</span>', report)

    def test_correspondence_scope_uses_info_severity(self):
        report = REFCHECK.build_report(
            "correspondence.md",
            [{"id": "R1", "raw": "Author, A. (2024). Title.", "doi": None,
              "title": "Title", "year": "2024", "venue": "Journal"}],
            [],
            {"R1": {"status": "found", "mismatches": [], "links": {}}},
            {"cited_but_missing_in_list": [
                {"sentence": "Claim (Missing, 2024).", "authors": ["Missing"], "year": "2024"},
            ], "listed_but_never_cited": []},
            [], [], [], {},
        )

        self.assertIn('<span class="scope-info">发现 1 项：见下方详情</span>', report)
        self.assertNotIn('<span class="scope-warn">发现 1 项：见下方详情</span>', report)

    def test_automatic_unmatched_result_is_not_presented_as_nonexistent(self):
        verifier = REFCHECK.Verifier(cache_dir=tempfile.mkdtemp())
        verifier._search_openalex = lambda title, entry: None
        verifier._search_crossref = lambda title, entry: None
        verifier._search_s2 = lambda title, entry: None

        result = verifier._verify_online(
            {"raw": "Author, A. (2024). A plausible title.",
             "title": "A plausible title."},
            "A plausible title.",
        )

        self.assertEqual("not_found", result["status"])
        self.assertEqual("low", result["confidence"])
        self.assertIn("人工复核", result["note"])
        self.assertNotIn("不存在", result["note"])

    def test_draft_report_labels_automatic_unmatched_as_review_required(self):
        report = REFCHECK.build_report(
            "draft.md",
            [{"id": "R1", "raw": "Author, A. (2024). Unmatched title.",
              "doi": None, "title": "Unmatched title", "year": "2024",
              "venue": "Journal"}],
            [],
            {"R1": {"status": "not_found", "mismatches": [], "links": {},
                    "note": "自动检索未匹配，必须人工复核"}},
            {"cited_but_missing_in_list": [], "listed_but_never_cited": []},
            [], [], [], {},
        )

        self.assertIn("自动未匹配（需复核）", report)
        self.assertNotIn("疑似不存在的文献", report)
        self.assertIn("自动初筛结果", report)

    def test_source_capabilities_are_explicit_without_exposing_keys(self):
        with mock.patch.dict(os.environ, {}, clear=True):
            verifier = REFCHECK.Verifier(cache_dir=tempfile.mkdtemp())

        self.assertEqual("public", verifier.source_capabilities["openalex"])
        self.assertEqual("public", verifier.source_capabilities["crossref"])
        self.assertEqual("not_configured", verifier.source_capabilities["semantic_scholar"])


if __name__ == "__main__":
    unittest.main()
