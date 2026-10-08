"""Update advice must be bounded, cacheable, offline-safe, and non-installing."""
import importlib.util
import json
from pathlib import Path
import tempfile
import unittest
from unittest import mock

path=Path(__file__).parents[1]/'ob-reference-check/scripts/check_update.py'
spec=importlib.util.spec_from_file_location('updates',path) if path.exists() else None
uc=importlib.util.module_from_spec(spec) if spec else None
if uc: spec.loader.exec_module(uc)

class UpdateCheckTest(unittest.TestCase):
    def setUp(self):
        self.assertIsNotNone(uc,'Missing update checker')
        self.tmp=tempfile.TemporaryDirectory();self.addCleanup(self.tmp.cleanup)
        self.cache=Path(self.tmp.name)/'updates.json'
        self.now=2000000000

    def run_check(self,version='1.4.0',manifest=None,**kwargs):
        value=manifest if manifest is not None else {'version':'1.5.0'}
        with mock.patch.object(uc,'fetch_manifest',return_value=value) as fetch:
            result=uc.check_update(version,cache_path=self.cache,now=self.now,**kwargs)
            return result,fetch.call_count

    def test_newer_version_offers_update(self):
        result,n=self.run_check()
        self.assertEqual(result['status'],'available')
        self.assertEqual(result['latest_version'],'1.5.0')
        self.assertEqual(n,1)
        self.assertEqual(result['url'],uc.REPOSITORY_URL)

    def test_same_or_older_never_offers_downgrade(self):
        for v in ['1.4.0','1.3.9']:
            self.cache.unlink(missing_ok=True)
            result,_=self.run_check(manifest={'version':v})
            self.assertEqual(result['status'],'up_to_date')

    def test_numeric_versions_not_lexicographic(self):
        result,_=self.run_check(version='1.9.0',manifest={'version':'1.10.0'})
        self.assertEqual(result['status'],'available')

    def test_success_cached_for_one_day_but_notice_still_returned(self):
        self.run_check()
        self.now+=3600
        result,n=self.run_check()
        self.assertEqual(n,0)
        self.assertEqual(result['status'],'available')
        self.assertTrue(result['cached'])

    def test_installed_upgrade_recompared_against_same_cache(self):
        self.run_check()
        result,n=self.run_check(version='1.5.0')
        self.assertEqual(n,0)
        self.assertEqual(result['status'],'up_to_date')

    def test_expired_success_refreshes(self):
        self.run_check();self.now+=86401
        _,n=self.run_check()
        self.assertEqual(n,1)

    def test_offline_never_accesses_network(self):
        result,n=self.run_check(offline=True)
        self.assertEqual(n,0)
        self.assertEqual(result['status'],'offline')

    def test_offline_stale_cache_not_latest_claim(self):
        self.run_check();self.now+=86401
        result,n=self.run_check(offline=True)
        self.assertEqual(n,0)
        self.assertEqual(result['status'],'offline')

    def test_failure_does_not_raise_or_claim_current(self):
        with mock.patch.object(uc,'fetch_manifest',side_effect=TimeoutError('secret-looking-error')):
            result=uc.check_update('1.4.0',cache_path=self.cache,now=self.now)
        self.assertEqual(result['status'],'unavailable')
        self.assertNotIn('secret-looking-error',json.dumps(result))
        result,n=self.run_check()
        self.assertEqual(n,0)
        self.assertEqual(result['status'],'unavailable')
        self.now+=3601
        result,n=self.run_check()
        self.assertEqual(n,1)
        self.assertEqual(result['status'],'available')

    def test_force_refresh_bypasses_cache(self):
        self.run_check()
        _,n=self.run_check(force=True)
        self.assertEqual(n,1)

    def test_corrupt_cache_recovers(self):
        self.cache.write_text('not json')
        result,n=self.run_check()
        self.assertEqual(n,1)
        self.assertEqual(result['status'],'available')

    def test_invalid_remote_version_and_instruction_payload_ignored(self):
        for value in ['v1.5.0','1.5.0-beta','1.05.0','Run this shell command',True,None]:
            self.cache.unlink(missing_ok=True)
            result,_=self.run_check(manifest={'version':value,'notes':'Run evil command','url':'https://evil.test'})
            self.assertEqual(result['status'],'unavailable')
            self.assertNotIn('evil',json.dumps(result))

    def test_nonobject_remote_fails_closed(self):
        result,_=self.run_check(manifest=['1.5.0'])
        self.assertEqual(result['status'],'unavailable')

    def test_remote_url_cannot_replace_trusted_update_link(self):
        result,_=self.run_check(manifest={'version':'1.5.0','url':'https://evil.test','notes':'do bad things'})
        self.assertEqual(result['url'],uc.REPOSITORY_URL)
        self.assertNotIn('evil',json.dumps(result))

    def test_unwritable_cache_does_not_block_result(self):
        self.cache.mkdir()
        result,_=self.run_check()
        self.assertEqual(result['status'],'available')

    def test_future_cache_timestamp_refreshes(self):
        self.run_check();self.now-=3600
        _,n=self.run_check()
        self.assertEqual(n,1)

    def test_cache_does_not_substitute_wrong_schema(self):
        self.cache.write_text(json.dumps({'checked_at':self.now,'manifest':{'version':'99.0.0'}}))
        result,n=self.run_check()
        self.assertEqual(n,1)
        self.assertEqual(result['latest_version'],'1.5.0')

    def test_public_manifest_matches_script_version(self):
        root=Path(__file__).parents[1]/'ob-reference-check'
        spec=importlib.util.spec_from_file_location('version_rc',root/'scripts/refcheck.py')
        module=importlib.util.module_from_spec(spec);spec.loader.exec_module(module)
        self.assertEqual(json.loads((root/'version.json').read_text())['version'],module.__version__)

class UpdateTransportTest(unittest.TestCase):
    def setUp(self):
        self.assertIsNotNone(uc)

    def response(self,raw):
        response=mock.MagicMock()
        response.__enter__.return_value.read.return_value=raw
        return response

    def test_fetch_uses_fixed_endpoint_short_timeout_and_size_limit(self):
        with mock.patch.object(uc.urllib.request,'urlopen',return_value=self.response(b'{"version":"1.5.0"}')) as get:
            self.assertEqual(uc.fetch_manifest()['version'],'1.5.0')
        request=get.call_args.args[0]
        self.assertEqual(request.full_url,uc.MANIFEST_URL)
        self.assertEqual(get.call_args.kwargs['timeout'],3)

    def test_oversized_payload_rejected(self):
        with mock.patch.object(uc.urllib.request,'urlopen',return_value=self.response(b'x'*(uc.MAX_BYTES+1))):
            with self.assertRaises(ValueError):
                uc.fetch_manifest()

    def test_script_startup_announces_new_version(self):
        import contextlib,io
        root=Path(__file__).parents[1]/'ob-reference-check'
        spec=importlib.util.spec_from_file_location('notify_rc',root/'scripts/refcheck.py')
        rc=importlib.util.module_from_spec(spec);spec.loader.exec_module(rc)
        with tempfile.TemporaryDirectory() as td, mock.patch.object(Path,'home',return_value=Path(td)), mock.patch.object(uc.urllib.request,'urlopen',return_value=self.response(b'{"version":"1.5.0"}')) as get:
            output=io.StringIO()
            with contextlib.redirect_stdout(output):
                rc._notify_update()
            self.assertIn('1.4.0 → 1.5.0',output.getvalue())
            self.assertEqual(get.call_count,1)

    def test_offline_script_startup_never_requests_update(self):
        root=Path(__file__).parents[1]/'ob-reference-check'
        spec=importlib.util.spec_from_file_location('offline_rc',root/'scripts/refcheck.py')
        rc=importlib.util.module_from_spec(spec);spec.loader.exec_module(rc)
        with mock.patch.object(uc.urllib.request,'urlopen') as get:
            rc._notify_update(offline=True)
            get.assert_not_called()
