"""发布 NCC 收款结算单工具到 GitHub：强制同时上传“轻量更新包 + 完整离线包”并同步已签名 latest.json；需显式 --publish。"""
import argparse
import base64
import json
import os
from pathlib import Path
import subprocess
import sys
from urllib.error import HTTPError
from urllib.parse import quote, urlsplit
from urllib.request import Request, HTTPRedirectHandler, ProxyHandler, build_opener

from receipt_updater import sha256_file, verify_envelope
import receipt_release_identity as identity


def git_credential():
    env = dict(os.environ, GIT_TERMINAL_PROMPT='0', GCM_INTERACTIVE='Never')
    result = subprocess.run(
        ['git', 'credential', 'fill'],
        input=b'protocol=https\nhost=github.com\n\n',
        capture_output=True,
        env=env,
        timeout=20,
    )
    fields = {}
    for line in result.stdout.decode('utf-8', errors='replace').splitlines():
        if '=' in line:
            key, value = line.split('=', 1)
            fields[key] = value
    token = fields.get('password')
    if result.returncode or not token:
        raise RuntimeError('本机未找到已登录的 GitHub 凭据，请先完成 GitHub 登录。')
    return token


class NoRedirect(HTTPRedirectHandler):
    def redirect_request(self, *args, **kwargs):
        raise ValueError('发布接口发生跳转，未继续发送登录凭据。')


def api(token, path, method='GET', data=None, content_type='application/json'):
    url = path if path.startswith('https://') else 'https://api.github.com' + path
    parsed = urlsplit(url)
    if parsed.scheme != 'https' or parsed.netloc not in ('api.github.com', 'uploads.github.com'):
        raise ValueError('拒绝将登录凭据发送至其他站点。')
    body = json.dumps(data).encode() if isinstance(data, (dict, list)) else data
    request = Request(
        url,
        data=body,
        method=method,
        headers={
            'Authorization': 'Bearer ' + token,
            'User-Agent': 'NCC-Receipt-Release',
            'Accept': 'application/vnd.github+json',
            'Content-Type': content_type,
            'X-GitHub-Api-Version': '2022-11-28',
        },
    )
    with build_opener(ProxyHandler({}), NoRedirect()).open(
        request, timeout=15 if method == 'GET' else 600
    ) as response:
        raw = response.read()
    return json.loads(raw) if raw else {}


def validate_release_folder(folder):
    """发布前强校验：必须同时具备已校验的 update 与 full 两个 ZIP 及已签名 latest.json。"""
    folder = Path(folder).resolve()
    plan = json.loads((folder / 'release-plan.json').read_text(encoding='utf-8'))
    expected_repo = identity.GITHUB_OWNER + '/' + identity.GITHUB_REPO
    if plan.get('repository') != expected_repo or plan.get('tag') != identity.TAG_PREFIX + identity.APP_VERSION:
        raise ValueError('发行目标仓库或版本号不符。')

    manifest_bytes = (folder / 'latest.json').read_bytes()
    verified = verify_envelope(json.loads(manifest_bytes))
    if verified['version'] != identity.APP_VERSION:
        raise ValueError('已签名清单版本与当前源码版本不符。')

    assets = plan.get('assets')
    if not isinstance(assets, list):
        raise ValueError('release-plan.json 缺少 assets 列表。')
    roles = {item.get('role') for item in assets if isinstance(item, dict)}
    if roles != {'update', 'full'}:
        raise ValueError('每次发布必须同时上传轻量更新包（update）与完整离线包（full）。')

    resolved_assets = []
    for item in assets:
        name = item['name']
        asset_path = folder / name
        if asset_path.name != name or asset_path.suffix != '.zip' or asset_path.parent != folder:
            raise ValueError('发行文件路径非法：' + str(name))
        if not asset_path.is_file():
            raise ValueError('缺少待上传压缩包：' + str(asset_path))
        actual_sha = sha256_file(asset_path)
        actual_size = asset_path.stat().st_size
        if actual_sha != item['sha256'] or actual_size != item['size_bytes']:
            raise ValueError('压缩包校验不通过：' + name)
        if item['role'] == 'update':
            if actual_sha != verified['package']['sha256'] or actual_size != verified['package']['size_bytes']:
                raise ValueError('轻量更新包与已签名 latest.json 不一致。')
        resolved_assets.append((item['role'], asset_path))
    return plan, verified, manifest_bytes, resolved_assets


def publish(folder):
    plan, verified, manifest_bytes, resolved_assets = validate_release_folder(folder)
    token = git_credential()
    root = '/repos/' + plan['repository']

    release = api(
        token,
        root + '/releases',
        'POST',
        dict(
            tag_name=plan['tag'],
            name=plan['title'],
            body=verified['notes'],
            draft=True,
            prerelease=False,
            make_latest='true',
            target_commitish=identity.GITHUB_BRANCH,
        ),
    )
    release_id = release['id']
    try:
        upload_base = release['upload_url'].split('{', 1)[0]
        for role, asset_path in resolved_assets:
            print(f'正在上传 {role} 包：{asset_path.name} ({asset_path.stat().st_size} 字节)...', flush=True)
            upload_url = upload_base + '?name=' + quote(asset_path.name)
            api(token, upload_url, 'POST', asset_path.read_bytes(), 'application/zip')

        published = api(
            token,
            f'{root}/releases/{release_id}',
            'PATCH',
            dict(draft=False, prerelease=False, make_latest='true'),
        )

        target = root + '/contents/' + identity.MANIFEST_PATH
        previous = None
        try:
            previous = api(token, target)
        except HTTPError as exc:
            if exc.code != 404:
                raise
        change = dict(
            message='Publish NCC receipt ' + identity.APP_VERSION + ' signed update manifest',
            content=base64.b64encode(manifest_bytes).decode(),
            branch=identity.GITHUB_BRANCH,
        )
        if previous:
            change['sha'] = previous['sha']
        api(token, target, 'PUT', change)
        print('已发布（含轻量更新包 + 完整离线包）：' + published['html_url'])
        print('NCC-Receipt 仓库的已签名 latest.json 已同步更新。')
    except Exception:
        print('发布未完成，请核对 GitHub Release 草稿与 latest.json。', file=sys.stderr)
        raise


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('folder', help='prepare_update_release.py 生成的发行目录')
    parser.add_argument('--publish', action='store_true', help='确认同时上传更新包与完整包并更新 latest.json')
    args = parser.parse_args()
    if not args.publish:
        parser.error('请核对双包目录后，使用 --publish 执行正式发布。')
    try:
        publish(args.folder)
    except Exception as exc:
        print(str(exc) if isinstance(exc, (ValueError, RuntimeError)) else '发布失败：' + type(exc).__name__, file=sys.stderr)
        sys.exit(1)
