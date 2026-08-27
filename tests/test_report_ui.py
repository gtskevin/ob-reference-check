"""Regression checks for report navigation and AI follow-up guidance."""

import importlib.util
import pathlib
import unittest


SCRIPT_PATH = (pathlib.Path(__file__).resolve().parents[1]
               / "ob-reference-check" / "scripts" / "refcheck.py")
SPEC = importlib.util.spec_from_file_location("refcheck", SCRIPT_PATH)
REFCHECK = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(REFCHECK)


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
        self.assertIn('可继续由 AI 审读', report)
        self.assertEqual(report.count('可继续由 AI 审读。'), 2)
        self.assertIn('继续审读引用恰当性和格式一致性。', report)

        for anchor in ("overview", "scope", "sec-appro", "sec-format", "appendix"):
            self.assertIn(f'id="{anchor}"', report)


if __name__ == "__main__":
    unittest.main()
