"""Isolated regression cases for review integrity; no network or user cache."""
import copy
import datetime
import importlib.util
import json
from pathlib import Path
import tempfile
import unittest
from unittest import mock

s = importlib.util.spec_from_file_location('rc', Path(__file__).parents[1] / 'ob-reference-check/scripts/refcheck.py')
rc = importlib.util.module_from_spec(s)
s.loader.exec_module(rc)


def entry(**kw):
    e = dict(id='R1', raw='Smith (2020). Sample research title.', authors=['Smith'], year=2020,
             title='Sample research title', venue='Journal', volume='1', issue='1', pages='1-10',
             doi='10.1234/a', parse_ok=True)
    e.update(kw)
    return e


def data():
    return dict(paper={'path': 'paper.md'}, entries=[entry()],
                citations=[dict(cid='C1', authors=['Smith'], year='2020', sentence='We hypothesize an effect (Smith, 2020).', triage='A')],
                verification={'R1': dict(status='found', mismatches=[], links={}, abstract=None)},
                correspondence={'cited_but_missing_in_list': [], 'listed_but_never_cited': []},
                duplicates=[], timeline=[], preprints=[], cross_checks={})


class ReliabilityTest(unittest.TestCase):
    def test_empty_reviews_do_not_claim_completion(self):
        h = rc.build_final_report(data(), [])
        self.assertNotIn('已完成 AI 复核', h)
        self.assertNotIn('已核，未发现存疑', h)
        self.assertNotIn('其余确认', h)
        self.assertIn('未记录检查', h)

    def test_volume_and_issue_not_substrings(self):
        rec = dict(entry(), volume='10', issue='12')
        fields = {x['field'] for x in rc._compare_metadata(entry(), rec)}
        self.assertTrue({'volume', 'issue'} <= fields)

    def test_title_and_coauthor_compared(self):
        fields = {x['field'] for x in rc._compare_metadata(entry(authors=['Smith', 'Wrong']), dict(entry(), title='Sample research titles', authors=['Smith', 'Correct']))}
        self.assertTrue({'title', 'authors'} <= fields)

    def test_year_suffix_preserved(self):
        e = rc.parse_entry('Smith, A. (2020a). Sample research title. Journal, 1(1), 1-10.', 1)
        self.assertEqual(e.get('year_suffix'), 'a')

    def test_one_suffix_citation_does_not_mark_both_entries_cited(self):
        es = [entry(year_suffix='a'), entry(id='R2', year_suffix='b')]
        cs = [dict(authors=['Smith'], year='2020a')]
        corr = rc.check_correspondence(es, cs)
        self.assertEqual(corr['listed_but_never_cited'], ['R2'])
        self.assertEqual(cs[0].get('ref_ids'), ['R1'])

    def test_ambiguous_same_year_does_not_silently_match(self):
        cs = [dict(authors=['Smith'], year='2020')]
        corr = rc.check_correspondence([entry(), entry(id='R2')], cs)
        self.assertEqual(len(corr.get('ambiguous_citations', [])), 1)
        self.assertEqual(cs[0].get('ref_ids'), [])

    def test_coauthor_not_used_as_first_author(self):
        corr = rc.check_correspondence([entry(authors=['Jones'])], [dict(authors=['Smith', 'Jones'], year='2020')])
        self.assertEqual(len(corr['cited_but_missing_in_list']), 1)

    def test_hypothesis_cluster_stays_A(self):
        cs = rc.extract_citations([dict(text='We hypothesize an effect (Smith, 2020; Jones, 2021; Brown, 2022).', heading=None)])
        self.assertEqual([c['triage'] for c in cs], ['A'] * 3)

    def test_group_size_is_per_parenthesis(self):
        cs = rc.extract_citations([dict(text='Past work (Smith, 2020; Jones, 2021; Brown, 2022) relates to work (White, 2023).', heading=None)])
        self.assertEqual(cs[-1]['triage'], 'B')
        self.assertEqual(cs[-1].get('group_size'), 1)
        self.assertTrue(cs[-1].get('locator'))

    def test_cache_recompares_current_entry(self):
        with tempfile.TemporaryDirectory() as td:
            v = rc.Verifier(cache_dir=td)
            rec = dict(entry(), _source='test', abstract='Abstract')
            with mock.patch.object(v, '_verify_online', return_value=dict(status='found', record=rec, mismatches=[])):
                v.verify(entry())
            got = v.verify(entry(year=2030))
            self.assertIn('year', [x['field'] for x in got['mismatches']])

    def test_changed_doi_does_not_hit_old_record(self):
        with tempfile.TemporaryDirectory() as td:
            v = rc.Verifier(cache_dir=td)
            with mock.patch.object(v, '_verify_online', return_value=dict(status='found', record=dict(entry(), _source='test'), mismatches=[])) as lookup:
                v.verify(entry())
                v.verify(entry(doi='10.1234/b'))
                self.assertEqual(lookup.call_count, 2)

    def test_negative_cache_has_short_ttl(self):
        cached = dict(cached_at=(datetime.date.today()-datetime.timedelta(days=2)).isoformat(), data={'status':'not_found'})
        self.assertTrue(rc.Verifier._stale(cached))

    def test_completed_appropriateness_requires_citation_verdict(self):
        checks = {'appropriateness': {'status':'completed', 'reviewed':1, 'note':'Reviewed'}}
        with self.assertRaises(SystemExit):
            rc._validate_checks(data(), [], checks)

    def test_two_categories_do_not_overwrite_badge(self):
        vs = [dict(id='R1', final_status='warn', verdict='Bad DOI', evidence='source', action='fix'),
              dict(id='R1', category='format', final_status='ok')]
        h = rc.build_final_report(data(), vs)
        self.assertNotIn('badge ok">✅', h)

    def test_citation_appropriateness_does_not_claim_missing_reference(self):
        vs = [dict(id='C1', category='appropriateness', ref_id='R1', final_status='info', verdict='Insufficient abstract', evidence='Abstract unavailable', action='Read full text')]
        h = rc.build_final_report(data(), vs)
        self.assertNotIn('文献列表中无对应条目', h)

    def test_unknown_category_rejected(self):
        with self.assertRaises(SystemExit):
            rc._validate_verdicts(data()['entries'], {}, [dict(id='R1', category='typo', final_status='ok')])

    def test_legacy_doi_verdict_not_reused(self):
        self.assertIsNone(rc.prior_verdict_for(entry(), {'10.1234/a': {'final_status':'ok'}}))

    def test_saved_verdict_bound_to_entry_and_preserves_evidence(self):
        vs = [dict(id='R1', final_status='warn', verdict='year typo', evidence='Publisher', action='correct'),
              dict(id='R1', category='format', final_status='ok')]
        store = rc.updated_verdict_store({}, data(), vs)
        saved = rc.prior_verdict_for(entry(), store)
        self.assertEqual(len(saved['verdicts']), 2)
        self.assertEqual(saved['verdicts'][0]['evidence'], 'Publisher')
        self.assertIsNone(rc.prior_verdict_for(entry(year=2030), store))

    def test_unresolved_duplicates_cannot_be_completed(self):
        d = data()
        d['duplicates'] = [{'ids':['R1', 'R2']}]
        with self.assertRaises(SystemExit):
            rc._validate_checks(d, [], {'duplicates': {'status':'completed', 'reviewed':1, 'note':'checked'}})

    def test_cleared_duplicate_does_not_remain_a_report_warning(self):
        d = data()
        d['duplicates'] = [{'ids':['R1']}]
        vs = [dict(id='R1', check='duplicates', final_status='ok', note='Distinct works')]
        checks = {'duplicates': {'status':'completed', 'reviewed':1, 'note':'Compared both records'}}
        rc._validate_checks(d, vs, checks)
        h = rc.build_final_report(d, vs, checks)
        self.assertIn('已完成检查：1/1；未发现问题', h)
        self.assertNotIn('发现 1 项重复', h)

    def test_support_without_abstract_rejected(self):
        vs = [dict(id='C1', category='appropriateness', ref_id='R1', final_status='ok', evidence='Supports it', evidence_level='abstract')]
        with self.assertRaises(SystemExit):
            rc._validate_verdicts(data()['entries'], data()['verification'], vs, data()['citations'])

    def test_format_ok_cannot_clear_metadata_anomaly(self):
        d = data()
        d['verification']['R1']['mismatches'] = [{'field':'year'}]
        with self.assertRaises(SystemExit):
            rc._validate_verdicts(d['entries'], d['verification'], [dict(id='R1', category='format', final_status='ok')])

    def test_appropriateness_cannot_clear_missing_correspondence(self):
        d = data()
        vs = [dict(id='C1', category='appropriateness', final_status='info', verdict='Unknown', evidence='No match', action='Verify')]
        with self.assertRaises(SystemExit):
            rc._validate_verdicts(d['entries'], d['verification'], vs, d['citations'], {'cited_but_missing_in_list':d['citations']})

    def test_stale_cache_requires_explicit_verdict(self):
        d = data()
        d['verification']['R1']['cache_stale'] = True
        with self.assertRaises(SystemExit):
            rc._validate_verdicts(d['entries'], d['verification'], [])

    def test_parse_failure_requires_explicit_verdict(self):
        d = data()
        d['entries'][0]['parse_ok'] = False
        with self.assertRaises(SystemExit):
            rc._validate_verdicts(d['entries'], d['verification'], [])

    def test_wrong_dimension_cannot_clear_metadata(self):
        d = data()
        d['verification']['R1']['mismatches'] = [{'field':'year'}]
        with self.assertRaises(SystemExit):
            rc._validate_verdicts(d['entries'], d['verification'], [dict(id='R1', check='duplicates', final_status='ok')])

    def test_ambiguous_reference_requires_explicit_verdict(self):
        d = data()
        with self.assertRaises(SystemExit):
            rc._validate_verdicts(d['entries'], d['verification'], [], d['citations'], {'ambiguous_citations':d['citations']})

    def test_completed_support_with_source_excerpt(self):
        d = data()
        vs = [dict(id='C1', category='appropriateness', ref_id='R1', final_status='ok', evidence='Publisher abstract', evidence_level='abstract', source_excerpt='Reports the same relationship')]
        rc._validate_verdicts(d['entries'], d['verification'], vs, d['citations'])
        checks = {'appropriateness': {'status':'completed', 'reviewed':1, 'note':'Compared source abstract'}}
        rc._validate_checks(d, vs, checks)
        self.assertIn('已完成检查：1/1；未发现问题', rc.build_final_report(d, vs, checks))

    def test_negative_cache_preserves_search_links(self):
        with tempfile.TemporaryDirectory() as td:
            v = rc.Verifier(cache_dir=td)
            with mock.patch.object(v, '_verify_online', return_value=dict(status='not_found', record=None, mismatches=[], links={'google_scholar':'example'})):
                v.verify(entry())
            self.assertIn('google_scholar', v.verify(entry())['links'])

    def test_finalize_does_not_reuse_appropriateness_across_manuscripts(self):
        vs = [dict(id='C1', category='appropriateness', ref_id='R1', final_status='ok', evidence='source')]
        self.assertEqual(rc.updated_verdict_store({}, data(), vs), {})

    def test_support_must_match_citation_reference(self):
        d = data()
        d['entries'].append(entry(id='R2'))
        d['citations'][0]['ref_ids'] = ['R1']
        vs = [dict(id='C1', category='appropriateness', ref_id='R2', final_status='ok', evidence='Full text', evidence_level='full_text')]
        with self.assertRaises(SystemExit):
            rc._validate_verdicts(d['entries'], d['verification'], vs, d['citations'])

    def test_format_pass_not_presented_as_full_entry_confirmation(self):
        d = data()
        vs = [dict(id='R1', category='format', final_status='ok')]
        h = rc.build_final_report(d, vs)
        self.assertNotIn('✅ 有明确通过结论的条目', h)

    def test_future_issue_is_not_impossible_existence(self):
        future = entry(year=datetime.date.today().year+1)
        self.assertNotIn('不可能存在', str(rc.check_timeline([future])))


if __name__ == '__main__':
    unittest.main()
