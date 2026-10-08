import importlib.util
from pathlib import Path
import unittest

spec=importlib.util.spec_from_file_location('release_rc',Path(__file__).parents[1]/'ob-reference-check/scripts/refcheck.py')
rc=importlib.util.module_from_spec(spec);spec.loader.exec_module(rc)

class ReleaseIntegrityTest(unittest.TestCase):
    def test_unknown_or_single_dimension_entry_not_called_confirmed(self):
        data=dict(paper={'path':'paper.md'},entries=[dict(id='R1',raw='Smith 2020',doi=None)],citations=[],verification={},correspondence={},summary_stats={})
        for verdicts in ([],[dict(id='R1',category='format',final_status='ok',note='Style checked')]):
            with self.subTest(verdicts=verdicts):
                report=rc.build_final_report(data,verdicts)
                self.assertNotIn('已确认条目',report)
                self.assertNotIn('条目已确认无误',report)
                self.assertNotIn('其余 ✅ 条目',report)
