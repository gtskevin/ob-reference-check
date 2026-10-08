"""End-to-end protections against stale manuscripts and review-file mixups."""
import contextlib
import datetime
import copy
import hashlib
import importlib.util
import io
import json
from pathlib import Path
import sys
import tempfile
import unittest
from unittest import mock

spec = importlib.util.spec_from_file_location('binding_rc', Path(__file__).parents[1] / 'ob-reference-check/scripts/refcheck.py')
rc = importlib.util.module_from_spec(spec)
spec.loader.exec_module(rc)


def binding(data):
    content = copy.deepcopy(data)
    for key in ('path', 'report', 'checked_at'):
        content['paper'].pop(key, None)
    digest = hashlib.sha256(json.dumps(content, ensure_ascii=False, sort_keys=True, separators=(',', ':')).encode()).hexdigest()
    return {'schema_version': 1, 'source_sha256': data['paper']['sha256'], 'screening_sha256': digest}


class VersionBindingTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        self.source = self.root / 'paper.md'
        self.source.write_text('Source v1', encoding='utf-8')
        self.data = dict(paper=dict(path=str(self.source), sha256=hashlib.sha256(self.source.read_bytes()).hexdigest()), entries=[], citations=[], verification={}, correspondence={}, duplicates=[], timeline=[], preprints=[], cross_checks={})
        self.screening = self.root / 'paper_refcheck_20261008.json'
        self.final = self.root / 'paper_final.json'
        self.payload = dict(binding=binding(self.data), verdicts=[], checks={})
        self.report = self.root / f'paper_refcheck_{datetime.date.today():%Y%m%d}_final.html'
        self.store = self.root / 'verdicts.json'
        self.patch = mock.patch.object(rc, 'VERDICT_STORE', str(self.store))
        self.patch.start()
        self.addCleanup(self.patch.stop)

    def write(self):
        self.screening.write_text(json.dumps(self.data), encoding='utf-8')
        self.final.write_text(json.dumps(self.payload), encoding='utf-8')

    def finalize(self, target=None, **kwargs):
        self.write()
        with contextlib.redirect_stdout(io.StringIO()):
            rc.run_finalize(str(target or self.screening), str(self.final), **kwargs)

    def reject(self, pattern, **kwargs):
        # Preserve prior deliverables and prove no report/history writes on failure.
        self.report.write_text('Existing report', encoding='utf-8')
        self.store.write_text('{}', encoding='utf-8')
        with self.assertRaisesRegex(SystemExit, pattern):
            self.finalize(**kwargs)
        self.assertEqual(self.report.read_text(), 'Existing report')
        self.assertEqual(self.store.read_text(), '{}')

    def test_changed_manuscript_rejected(self):
        self.source.write_text('Source v2', encoding='utf-8')
        self.reject('稿件.*变化')

    def test_other_screening_rejected(self):
        self.payload['binding']['screening_sha256'] = '0' * 64
        self.reject('初筛.*不一致')

    def test_review_source_mismatch_rejected(self):
        self.payload['binding']['source_sha256'] = '0' * 64
        self.reject('稿件.*不一致')

    def test_corrected_citation_invalidates_old_review(self):
        self.data['citations'] = [dict(cid='C1', sentence='Corrected claim')]
        self.reject('初筛.*不一致')

    def test_changed_database_evidence_invalidates_old_review(self):
        self.data['verification']['R1'] = dict(abstract='New evidence')
        self.reject('初筛.*不一致')

    def test_legacy_final_must_not_silently_pass(self):
        self.payload.pop('binding')
        self.reject('版本绑定')

    def test_legacy_screening_must_be_rescreened(self):
        self.data['paper'].pop('sha256')
        self.reject('重新初筛')

    def test_unavailable_source_is_not_assumed_current(self):
        self.source.unlink()
        self.reject('稿件.*读取')

    def test_explicit_paper_target_cannot_hide_behind_stored_path(self):
        other = self.root / 'other' / 'paper.md'
        other.parent.mkdir()
        other.write_text('Another manuscript', encoding='utf-8')
        shutil_path = other.parent / self.screening.name
        shutil_path.write_text(json.dumps(self.data), encoding='utf-8')
        self.reject('稿件.*变化', target=other)

    def test_valid_binding_generates_report(self):
        self.finalize()
        self.assertTrue(list(self.root.glob('*_final.html')))

    def test_same_bytes_at_new_source_location_allowed(self):
        moved = self.root / 'moved.md'
        self.source.rename(moved)
        self.finalize(source_path=str(moved))
        self.assertTrue(list(self.root.glob('*_final.html')))

    def test_explicit_source_with_wrong_content_rejected(self):
        other = self.root / 'other.md'
        other.write_text('Wrong bytes')
        self.reject('稿件.*变化', source_path=str(other))

    def test_final_direct_target_still_works(self):
        self.finalize(target=self.final)
        self.assertTrue(list(self.root.glob('*_final.html')))

    def test_json_formatting_and_paths_do_not_invalidate_content(self):
        self.data['paper'].update(report='new-report.html', checked_at='later')
        self.finalize()
        self.assertTrue(list(self.root.glob('*_final.html')))

    def test_prepare_template_has_binding_and_no_fake_reviews(self):
        self.write()
        self.final.unlink()
        with contextlib.redirect_stdout(io.StringIO()):
            with mock.patch.object(sys, 'argv', ['refcheck', str(self.screening), '--prepare-final', '--offline', '--final', str(self.final)]):
                rc.main()
        payload = json.loads(self.final.read_text())
        self.assertEqual(payload['binding'], binding(self.data))
        self.assertEqual(payload['verdicts'], [])
        self.assertEqual(payload['checks'], {})

    def test_prepare_does_not_overwrite_existing_verdicts(self):
        self.write()
        original = self.final.read_bytes()
        with contextlib.redirect_stdout(io.StringIO()):
            with mock.patch.object(sys, 'argv', ['refcheck', str(self.screening), '--prepare-final', '--offline', '--final', str(self.final)]):
                with self.assertRaisesRegex(SystemExit, '已存在'):
                    rc.main()
        self.assertEqual(self.final.read_bytes(), original)

class ScreeningBindingTest(unittest.TestCase):
    def run_screening(self, root, mutate=False):
        source = root / 'sample.md'
        source.write_text('# Introduction\nPrior work is related (Smith, 2020).\n\n# References\nSmith, A. (2020). Sample research title. Journal, 1(1), 1-10.\n')
        original = source.read_bytes()
        parse = rc.parse_document
        def parse_and_change(path):
            paragraphs = parse(path)
            if mutate:
                source.write_text('Changed during screening')
            return paragraphs
        # Isolate external retrieval and persistent user history.
        with mock.patch.object(rc, 'parse_document', side_effect=parse_and_change), mock.patch.object(rc, 'Verifier') as verifier, mock.patch.object(rc, 'load_verdict_store', return_value={}):
            verifier.return_value.verify.return_value = dict(status='found', mismatches=[], links={})
            verifier.return_value.stats = {}
            verifier.return_value.source_capabilities = {}
            with contextlib.redirect_stdout(io.StringIO()), mock.patch.object(sys, 'argv', ['refcheck', str(source), '--offline', '--outdir', str(root)]):
                rc.main()
        return source, original, list(root.glob('*_refcheck_*.json'))[0]

    def test_screening_records_actual_source_bytes(self):
        with tempfile.TemporaryDirectory() as td:
            source, original, screening = self.run_screening(Path(td))
            data = json.loads(screening.read_text())
            self.assertEqual(data['paper'].get('sha256'), hashlib.sha256(original).hexdigest())
            final = Path(td) / 'sample_final.json'
            with contextlib.redirect_stdout(io.StringIO()):
                rc.run_prepare_final(str(screening), str(final))
                self.assertEqual(json.loads(final.read_text())['binding'], binding(data))
                # Normal parser/API replay produces a usable partially reviewed report.
                with mock.patch.object(rc, 'VERDICT_STORE', str(Path(td) / 'history.json')):
                    rc.run_finalize(str(screening), str(final))
            self.assertTrue(list(Path(td).glob('*_final.html')))

    def test_file_changed_during_screening_does_not_publish_outputs(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            with self.assertRaisesRegex(SystemExit, '初筛期间.*变化'):
                self.run_screening(root, mutate=True)
            self.assertFalse(list(root.glob('*_refcheck_*')))


if __name__ == '__main__':
    unittest.main()
