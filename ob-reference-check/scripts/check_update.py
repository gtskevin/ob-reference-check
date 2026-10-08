#!/usr/bin/env python3
"""Check a fixed public version manifest; never download or execute updates."""
import argparse
import json
import math
import os
from pathlib import Path
import re
import time
import urllib.request

REPOSITORY_URL = 'https://github.com/gtskevin/ob-reference-check'
MANIFEST_URL = 'https://raw.githubusercontent.com/gtskevin/ob-reference-check/main/ob-reference-check/version.json'
CACHE_PATH = Path.home() / '.reference_check' / 'update-check.json'
SUCCESS_TTL = 86400
FAILURE_TTL = 3600
MAX_BYTES = 16384


def _version(value):
    if not isinstance(value, str) or not re.fullmatch(r'(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)', value):
        raise ValueError('Invalid stable version')
    return tuple(int(part) for part in value.split('.'))


def fetch_manifest():
    """Read a small JSON document from the canonical repository with a short timeout."""
    request = urllib.request.Request(MANIFEST_URL, headers={'User-Agent': 'ob-reference-check-update', 'Accept': 'application/json'})
    with urllib.request.urlopen(request, timeout=3) as response:
        raw = response.read(MAX_BYTES + 1)
    if len(raw) > MAX_BYTES:
        raise ValueError('Oversized manifest')
    return json.loads(raw)


def _load_cache(path):
    try:
        with open(path, 'rb') as f:
            raw = f.read(MAX_BYTES + 1)
        value = json.loads(raw) if len(raw) <= MAX_BYTES else None
        return value if isinstance(value, dict) and value.get('schema_version') == 1 else {}
    except (OSError, ValueError):
        return {}


def _save_cache(path, value):
    temporary = path.with_name(path.name + f'.{os.getpid()}.tmp')
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        temporary.write_text(json.dumps(value), encoding='utf-8')
        os.replace(temporary, path)
    except OSError:
        try:
            temporary.unlink()
        except OSError:
            pass


def check_update(current_version, cache_path=None, now=None, offline=False, force=False):
    """Return advice only. Failures and unavailable cache storage never block use."""
    now = time.time() if now is None else now
    path = Path(cache_path) if cache_path is not None else CACHE_PATH
    cached = _load_cache(path)
    base = {'current_version': current_version, 'url': REPOSITORY_URL, 'cached': False}
    try:
        current = _version(current_version)
    except ValueError:
        return dict(base, status='unavailable')
    timestamp = cached.get('checked_at')
    age = now - timestamp if isinstance(timestamp, (int, float)) and not isinstance(timestamp, bool) and math.isfinite(timestamp) else -1
    ttl = SUCCESS_TTL if cached.get('status') == 'success' else FAILURE_TTL
    fresh = 0 <= age < ttl
    latest = None
    if fresh and (not force or offline):
        if cached.get('status') == 'success':
            try:
                latest = cached['latest_version']
                _version(latest)
            except (KeyError, ValueError):
                latest = None
        elif cached.get('status') == 'failure':
            return dict(base, status='offline' if offline else 'unavailable', cached=True)
        if latest is not None:
            base['cached'] = True
    if latest is None:
        if offline:
            return dict(base, status='offline')
        try:
            manifest = fetch_manifest()
            if not isinstance(manifest, dict):
                raise ValueError('Invalid manifest')
            latest = manifest.get('version')
            _version(latest)
        except Exception:
            # Advice must never interrupt the academic task or expose network errors.
            _save_cache(path, {'schema_version': 1, 'checked_at': now, 'status': 'failure'})
            return dict(base, status='unavailable')
        _save_cache(path, {'schema_version': 1, 'checked_at': now, 'status': 'success', 'latest_version': latest})
    return dict(base, latest_version=latest, status='available' if _version(latest) > current else 'up_to_date')


def main():
    parser = argparse.ArgumentParser(description='检查 skill 更新，仅提示，不自动安装')
    parser.add_argument('--offline', action='store_true')
    parser.add_argument('--force', action='store_true', help='忽略检查缓存重新联网')
    parser.add_argument('--json', action='store_true', help='结构化结果供 AI 助手读取')
    parser.add_argument('--current-version', help='用于验证旧版本提示；默认已安装版本')
    parser.add_argument('--cache-file', help='指定更新检查缓存文件')
    args = parser.parse_args()
    current = args.current_version
    if current is None:
        try:
            current = json.loads((Path(__file__).resolve().parents[1] / 'version.json').read_text())['version']
        except (OSError, ValueError, KeyError, TypeError):
            current = 'unknown'
    result = check_update(current, cache_path=args.cache_file, offline=args.offline, force=args.force)
    if args.json:
        print(json.dumps(result, ensure_ascii=False))
    elif result['status'] == 'available':
        print(f"[更新提示] ob-reference-check {current} → {result['latest_version']}；"
              f"可请 AI 助手从 {REPOSITORY_URL} 更新。确认前继续使用当前版本。")
    elif result['status'] in ('unavailable', 'offline'):
        print('[更新检查] 暂时无法确认新版；继续当前任务。')


if __name__ == '__main__':
    main()
