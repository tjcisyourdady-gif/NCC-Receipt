"""准备已签名的 NCC 收款结算单工具更新清单（latest.json）与发布计划；本脚本不直接上传。"""
import argparse
from datetime import datetime, timezone
import json
from pathlib import Path
import shutil
import subprocess
import zipfile

from receipt_updater import sha256_file, verify_envelope
import receipt_release_identity as identity


def prepare(archive, destination=None, update_only=True):
    project = Path(__file__).resolve().parent
    archive = Path(archive).resolve()
    destination = Path(destination or project / 'release_upload' / identity.APP_VERSION).resolve()
    if destination.exists():
        shutil.rmtree(destination)

    with zipfile.ZipFile(str(archive)) as package:
        names = package.namelist()
        if any('/.updates/' in n or '/.release-secrets/' in n for n in names):
            raise ValueError('发行包包含内部更新或密钥目录，已拒绝。')
        if any(
            n.endswith('/收款结算单填写顺序设置.json') or n.endswith('/收款结算单时间设置.json')
            for n in names
        ):
            raise ValueError('公开更新包不得包含个人配置文件。')
        if update_only and any('/ms-playwright（勿动）/' in n for n in names):
            raise ValueError('轻量更新包不得包含内置浏览器目录。')
        for exe_name in identity.EXE_NAMES:
            if not any(n.endswith('/' + exe_name) or n == exe_name for n in names):
                raise ValueError('更新包缺少双版本主程序：' + exe_name)

    destination.mkdir(parents=True, exist_ok=True)
    asset = destination / (
        'NCC_receipt_' + identity.APP_VERSION + ('_update.zip' if update_only else '_full.zip')
    )
    shutil.copyfile(str(archive), str(asset))

    payload = dict(
        schema=1,
        app_id=identity.APP_ID,
        version=identity.APP_VERSION,
        package_type='application_update' if update_only else 'offline_full',
        requires_existing_browser=bool(update_only),
        released_at=datetime.now(timezone.utc).isoformat(),
        notes=identity.RELEASE_NOTES,
        full=dict(
            url='https://github.com/{}/{}/releases/download/{}{}/{}'.format(
                identity.GITHUB_OWNER,
                identity.GITHUB_REPO,
                identity.TAG_PREFIX,
                identity.APP_VERSION,
                asset.name,
            ),
            sha256=sha256_file(asset),
            size_bytes=asset.stat().st_size,
        ),
    )
    raw = destination / 'payload.json'
    raw.write_text(
        json.dumps(payload, ensure_ascii=True, sort_keys=True, separators=(',', ':')),
        encoding='utf-8',
    )
    signed = destination / 'latest.json'
    result = subprocess.run(
        [
            'powershell.exe', '-NoProfile', '-ExecutionPolicy', 'Bypass',
            '-File', str(project / 'update_keys.ps1'),
            '-PayloadPath', str(raw), '-OutputPath', str(signed),
        ],
        capture_output=True,
        timeout=30,
    )
    if result.returncode:
        err = result.stderr.decode('utf-8', errors='replace') or result.stdout.decode('utf-8', errors='replace')
        raise RuntimeError('本机更新签名失败，请使用创建私钥的 Windows 账号运行：' + err)

    verify_envelope(json.loads(signed.read_text(encoding='utf-8-sig')))
    plan = dict(
        repository=identity.GITHUB_OWNER + '/' + identity.GITHUB_REPO,
        tag=identity.TAG_PREFIX + identity.APP_VERSION,
        title='NCC 收款结算单工具 ' + identity.DISPLAY_VERSION,
        asset=asset.name,
        manifest=identity.MANIFEST_PATH,
        update_only=update_only,
        sha256=payload['full']['sha256'],
        size_bytes=asset.stat().st_size,
    )
    (destination / 'release-plan.json').write_text(
        json.dumps(plan, ensure_ascii=False, indent=2),
        encoding='utf-8',
    )
    print('发行文件已准备：' + str(destination))
    print('已签名更新清单：' + str(signed))
    return destination


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('archive')
    parser.add_argument('--destination')
    parser.add_argument('--full', action='store_true', help='打包完整离线包而非轻量更新包')
    args = parser.parse_args()
    prepare(args.archive, args.destination, update_only=not args.full)
