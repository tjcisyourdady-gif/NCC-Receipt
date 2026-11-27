"""准备已签名的 NCC 收款结算单工具发布目录：强制同时校验并准备“轻量更新包 + 完整离线包”与签名 latest.json；本脚本不直接上传。"""
import argparse
from datetime import datetime, timezone
import io
import json
from pathlib import Path
import shutil
import subprocess
import zipfile

from openpyxl import load_workbook

from receipt_updater import sha256_file, verify_envelope
import receipt_release_identity as identity


def _check_common_safety(names, label):
    if any('/.updates/' in n or '/.release-secrets/' in n for n in names):
        raise ValueError(f'{label}包含内部更新或密钥目录，已拒绝。')
    if any(
        n.endswith('/收款结算单填写顺序设置.json') or n.endswith('/收款结算单时间设置.json')
        for n in names
    ):
        raise ValueError(f'{label}不得包含个人配置文件。')
    for exe_name in identity.EXE_NAMES:
        if not any(n.endswith('/' + exe_name) or n == exe_name for n in names):
            raise ValueError(f'{label}缺少双版本主程序：{exe_name}')


def _verify_blank_template_bytes(raw_bytes):
    wb = load_workbook(io.BytesIO(raw_bytes), read_only=True, data_only=True)
    try:
        for sheet_name in ('收款结算明细', '例子（在此工作表填写无效）'):
            ws = wb[sheet_name]
            for row in ws.iter_rows(min_row=3, values_only=True):
                if any(v is not None and str(v).strip() != '' for v in row):
                    raise ValueError(f'完整包内的底稿模板“{sheet_name}”含业务数据，严禁公开上传。')
    finally:
        wb.close()


def prepare(update_archive, full_archive, destination=None):
    """同时校验轻量更新包与完整离线包，生成已签名 latest.json 与双包 release-plan.json。"""
    project = Path(__file__).resolve().parent
    update_archive = Path(update_archive).resolve()
    full_archive = Path(full_archive).resolve()
    if not update_archive.is_file():
        raise ValueError('缺少轻量更新包：' + str(update_archive))
    if not full_archive.is_file():
        raise ValueError('缺少完整离线包：' + str(full_archive))

    destination = Path(destination or project / 'release_upload' / identity.APP_VERSION).resolve()
    if destination.exists():
        shutil.rmtree(destination)

    # 1. 校验轻量更新包（不含内置浏览器与个人配置）
    with zipfile.ZipFile(str(update_archive)) as pkg:
        names = pkg.namelist()
        _check_common_safety(names, '轻量更新包')
        if any('/ms-playwright（勿动）/' in n for n in names):
            raise ValueError('轻量更新包不得包含内置浏览器目录。')

    # 2. 校验完整离线包（必须含内置 Chromium、空白 33 列宏模板与对照表方案，且模板 0 业务数据）
    with zipfile.ZipFile(str(full_archive)) as pkg:
        names = pkg.namelist()
        _check_common_safety(names, '完整离线包')
        if not any(n.endswith('/ms-playwright（勿动）/chromium-1161/chrome-win/chrome.exe') for n in names):
            raise ValueError('完整离线包缺少内置 Chromium 浏览器。')
        template_entries = [n for n in names if n.endswith('/收款结算单底稿导入模板.xlsm')]
        if len(template_entries) != 1:
            raise ValueError('完整离线包缺少唯一标准空白模板“收款结算单底稿导入模板.xlsm”。')
        if not any(n.endswith('/对照表方案.xlsx') for n in names):
            raise ValueError('完整离线包缺少“对照表方案.xlsx”。')
        _verify_blank_template_bytes(pkg.read(template_entries[0]))

    destination.mkdir(parents=True, exist_ok=True)
    update_asset = destination / f'NCC_receipt_{identity.APP_VERSION}_update.zip'
    full_asset = destination / f'NCC_receipt_{identity.APP_VERSION}_full.zip'
    shutil.copyfile(str(update_archive), str(update_asset))
    shutil.copyfile(str(full_archive), str(full_asset))

    update_sha = sha256_file(update_asset)
    full_sha = sha256_file(full_asset)

    # latest.json 继续指向轻量更新包，保证客户端秒级增量升级并保留浏览器与用户模板
    payload = dict(
        schema=1,
        app_id=identity.APP_ID,
        version=identity.APP_VERSION,
        package_type='application_update',
        requires_existing_browser=True,
        released_at=datetime.now(timezone.utc).isoformat(),
        notes=identity.RELEASE_NOTES,
        full=dict(
            url='https://github.com/{}/{}/releases/download/{}{}/{}'.format(
                identity.GITHUB_OWNER,
                identity.GITHUB_REPO,
                identity.TAG_PREFIX,
                identity.APP_VERSION,
                update_asset.name,
            ),
            sha256=update_sha,
            size_bytes=update_asset.stat().st_size,
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
        manifest=identity.MANIFEST_PATH,
        assets=[
            dict(
                role='update',
                name=update_asset.name,
                sha256=update_sha,
                size_bytes=update_asset.stat().st_size,
            ),
            dict(
                role='full',
                name=full_asset.name,
                sha256=full_sha,
                size_bytes=full_asset.stat().st_size,
            ),
        ],
    )
    (destination / 'release-plan.json').write_text(
        json.dumps(plan, ensure_ascii=False, indent=2),
        encoding='utf-8',
    )
    print('双包发行文件已准备：' + str(destination))
    print('  轻量更新包：' + update_asset.name)
    print('  完整离线包：' + full_asset.name)
    print('  已签名清单：' + str(signed))
    return destination


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--update-archive', required=True, help='轻量更新包 ZIP 路径（不含浏览器）')
    parser.add_argument('--full-archive', required=True, help='完整离线包 ZIP 路径（含浏览器与空白模板）')
    parser.add_argument('--destination', help='可选输出目录')
    args = parser.parse_args()
    prepare(args.update_archive, args.full_archive, args.destination)
