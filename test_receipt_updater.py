"""收款结算单工具自动检测更新模块离线单元测试（不连接真实网络，不修改真实签名密钥）。"""
import ast
import base64
import io
import json
from pathlib import Path
import subprocess
import tempfile
import unittest
from unittest.mock import patch
import zipfile

import receipt_release_identity as identity
import receipt_updater as updater
import prepare_update_release as prep
import publish_receipt_release as publisher

REPO = Path(__file__).resolve().parent


def make_zip(path: Path, entries: dict[str, bytes]):
    with zipfile.ZipFile(path, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
        for arcname, content in entries.items():
            zf.writestr(arcname, content)


class FakeResponse(io.BytesIO):
    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        self.close()
        return False


class ReceiptUpdaterTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.tmp = Path(self.temp.name)

    def test_version_tuple_and_trusted_asset_url(self):
        self.assertLess(updater.version_tuple('7.0.0'), updater.version_tuple('7.1.0'))
        self.assertLess(updater.version_tuple('7.1.0'), updater.version_tuple('7.1.1'))
        self.assertEqual(updater.version_tuple('V7.1'), updater.version_tuple('7.1.0.0'))
        with self.assertRaises(updater.UpdateError):
            updater.version_tuple('7.1-beta')

        good_url = (
            f'https://github.com/{identity.GITHUB_OWNER}/{identity.GITHUB_REPO}'
            f'/releases/download/{identity.TAG_PREFIX}7.2.0/NCC_receipt_7.2.0_update.zip'
        )
        self.assertTrue(updater.trusted_asset_url(good_url, '7.2.0'))
        self.assertFalse(updater.trusted_asset_url(good_url, '7.1.0'))
        self.assertFalse(updater.trusted_asset_url('https://evil.example.com/a.zip', '7.2.0'))
        self.assertFalse(updater.trusted_asset_url(good_url + '?token=1', '7.2.0'))
        self.assertFalse(updater.trusted_asset_url(good_url.replace('.zip', '.exe'), '7.2.0'))

    def test_rsa_sign_and_verify_envelope_with_isolated_temp_keys(self):
        key_dir = self.tmp / 'keys'
        pub_path = self.tmp / 'pub.json'
        ps1 = REPO / 'update_keys.ps1'

        subprocess.run(
            [
                'powershell.exe', '-NoProfile', '-ExecutionPolicy', 'Bypass',
                '-File', str(ps1), '-Initialize',
                '-KeyDirectory', str(key_dir), '-PublicKeyPath', str(pub_path),
            ],
            check=True, capture_output=True, text=True,
        )
        pub = json.loads(pub_path.read_text(encoding='utf-8-sig'))
        payload = {
            'schema': 1,
            'app_id': identity.APP_ID,
            'version': '7.9.0',
            'notes': '测试新版本',
            'released_at': '2026-11-27T01:00:00+00:00',
            'full': {
                'url': (
                    f'https://github.com/{identity.GITHUB_OWNER}/{identity.GITHUB_REPO}'
                    f'/releases/download/{identity.TAG_PREFIX}7.9.0/NCC_receipt_7.9.0_update.zip'
                ),
                'sha256': 'a' * 64,
                'size_bytes': 123456,
            },
        }
        payload_path = self.tmp / 'payload.json'
        signed_path = self.tmp / 'latest.json'
        payload_path.write_text(
            json.dumps(payload, ensure_ascii=True, sort_keys=True, separators=(',', ':')),
            encoding='utf-8',
        )
        subprocess.run(
            [
                'powershell.exe', '-NoProfile', '-ExecutionPolicy', 'Bypass',
                '-File', str(ps1),
                '-PayloadPath', str(payload_path), '-OutputPath', str(signed_path),
                '-KeyDirectory', str(key_dir), '-PublicKeyPath', str(pub_path),
            ],
            check=True, capture_output=True, text=True,
        )
        envelope = json.loads(signed_path.read_text(encoding='utf-8-sig'))

        with patch.object(identity, 'SIGNING_KEY_ID', pub['key_id']), \
                patch.object(identity, 'SIGNING_MODULUS', pub['modulus']):
            verified = updater.verify_envelope(envelope)
            self.assertEqual(verified['version'], '7.9.0')
            self.assertEqual(verified['notes'], '测试新版本')

            # 篡改 payload 必须验签失败
            tampered = dict(envelope)
            raw_bytes = bytearray(base64.b64decode(envelope['payload']))
            raw_bytes[0] ^= 0x01
            tampered['payload'] = base64.b64encode(bytes(raw_bytes)).decode('ascii')
            with self.assertRaisesRegex(updater.UpdateError, '签名校验失败'):
                updater.verify_envelope(tampered)

            # 错误 key_id 必须拒绝
            wrong_key = dict(envelope, key_id='0' * 64)
            with self.assertRaisesRegex(updater.UpdateError, '签名校验失败'):
                updater.verify_envelope(wrong_key)

    def test_stage_package_preserves_user_files_and_rejects_bad_archives(self):
        good_zip = self.tmp / 'good.zip'
        make_zip(good_zip, {
            f'收款结算填写工具/{identity.EXE_NAMES[0]}': b'exe-normal',
            f'收款结算填写工具/{identity.EXE_NAMES[1]}': b'exe-time',
            '收款结算填写工具/V7双版本及GUI功能说明.txt': '说明'.encode('utf-8'),
            # 用户数据/配置即使出现在包内也必须跳过，严禁覆盖
            '收款结算填写工具/收款结算单底稿导入模板.xlsm': b'user-template',
            '收款结算填写工具/对照表方案.xlsx': b'user-mapping',
            '收款结算填写工具/收款结算单填写顺序设置.json': b'{}',
            '收款结算填写工具/收款结算单时间设置.json': b'{}',
        })
        staged = updater.stage_package(good_zip, self.tmp / 'stage_good')
        staged_paths = {item['path'] for item in staged}
        self.assertEqual(staged_paths, {
            identity.EXE_NAMES[0],
            identity.EXE_NAMES[1],
            'V7双版本及GUI功能说明.txt',
        })

        # 缺少其中一个版本 EXE 必须拒绝
        missing_exe = self.tmp / 'missing_exe.zip'
        make_zip(missing_exe, {f'收款结算填写工具/{identity.EXE_NAMES[0]}': b'only-one'})
        with self.assertRaisesRegex(updater.UpdateError, '缺少双版本'):
            updater.stage_package(missing_exe, self.tmp / 'stage_missing')

        # 路径穿越必须拒绝
        traversal = self.tmp / 'traversal.zip'
        make_zip(traversal, {'../evil.exe': b'bad'})
        with self.assertRaisesRegex(updater.UpdateError, '非法路径'):
            updater.stage_package(traversal, self.tmp / 'stage_trav')

        # 未知顶层文件必须拒绝
        unknown_file = self.tmp / 'unknown.zip'
        make_zip(unknown_file, {
            f'收款结算填写工具/{identity.EXE_NAMES[0]}': b'1',
            f'收款结算填写工具/{identity.EXE_NAMES[1]}': b'2',
            '收款结算填写工具/evil.dll': b'3',
        })
        with self.assertRaisesRegex(updater.UpdateError, '不允许覆盖'):
            updater.stage_package(unknown_file, self.tmp / 'stage_unk')

    def test_update_manager_check_download_and_plan(self):
        app_dir = self.tmp / 'app'
        app_dir.mkdir()
        pkg_zip = self.tmp / 'pkg.zip'
        make_zip(pkg_zip, {
            f'收款结算填写工具/{identity.EXE_NAMES[0]}': b'new-normal-exe',
            f'收款结算填写工具/{identity.EXE_NAMES[1]}': b'new-time-exe',
            '收款结算填写工具/V7双版本及GUI功能说明.txt': '新说明'.encode('utf-8'),
        })
        pkg_bytes = pkg_zip.read_bytes()
        manifest = {
            'version': '9.9.9',
            'notes': '测试升级到 9.9.9',
            'released_at': '2026-11-27T01:00:00+00:00',
            'source': '国内直连通道',
            'package': {
                'url': (
                    f'https://github.com/{identity.GITHUB_OWNER}/{identity.GITHUB_REPO}'
                    f'/releases/download/{identity.TAG_PREFIX}9.9.9/NCC_receipt_9.9.9_update.zip'
                ),
                'sha256': updater.sha256_file(pkg_zip),
                'size_bytes': len(pkg_bytes),
            },
        }
        mgr = updater.UpdateManager(app_dir)
        with patch.object(updater, 'fetch_manifest', return_value=manifest):
            state = mgr.check()
        self.assertEqual(state['status'], 'available')
        self.assertEqual(state['latest_version'], '9.9.9')

        # 模拟第一条镜像失败、第二条镜像成功下载
        calls = {'n': 0}

        def fake_open(url, timeout=15):
            calls['n'] += 1
            if calls['n'] == 1:
                raise OSError('mirror 1 timeout')
            return FakeResponse(pkg_bytes)

        with patch.object(updater, 'open_url', side_effect=fake_open):
            dl_state = mgr.download()
        self.assertEqual(dl_state['status'], 'ready')
        self.assertEqual(dl_state['progress'], 100)
        plan = mgr.installation_plan()
        self.assertEqual(plan['app_id'], identity.APP_ID)
        self.assertEqual(plan['version'], '9.9.9')
        self.assertEqual(plan['exe_name'], identity.EXE_NAMES[0])
        self.assertEqual(len(plan['files']), 3)

        # 版本相同则提示已是最新
        same_manifest = dict(manifest, version=identity.APP_VERSION)
        with patch.object(updater, 'fetch_manifest', return_value=same_manifest):
            up_state = mgr.check()
        self.assertEqual(up_state['status'], 'up_to_date')

    def test_prepare_and_publish_enforce_both_update_and_full_packages(self):
        clean_update = self.tmp / 'clean_update.zip'
        make_zip(clean_update, {
            f'收款结算填写工具/{identity.EXE_NAMES[0]}': b'exe1',
            f'收款结算填写工具/{identity.EXE_NAMES[1]}': b'exe2',
            '收款结算填写工具/V7双版本及GUI功能说明.txt': '说明'.encode('utf-8'),
        })
        clean_full = self.tmp / 'clean_full.zip'
        make_zip(clean_full, {
            f'收款结算填写工具/{identity.EXE_NAMES[0]}': b'exe1',
            f'收款结算填写工具/{identity.EXE_NAMES[1]}': b'exe2',
            '收款结算填写工具/V7双版本及GUI功能说明.txt': '说明'.encode('utf-8'),
            '收款结算填写工具/收款结算单底稿导入模板.xlsm': (REPO / '收款结算单底稿导入模板.xlsm').read_bytes(),
            '收款结算填写工具/对照表方案.xlsx': (REPO / '对照表方案.xlsx').read_bytes(),
            '收款结算填写工具/ms-playwright（勿动）/chromium-1161/chrome-win/chrome.exe': b'chrome',
        })

        # 1. 轻量包包含浏览器必须拒绝
        bad_browser = self.tmp / 'with_browser.zip'
        make_zip(bad_browser, {
            f'收款结算填写工具/{identity.EXE_NAMES[0]}': b'1',
            f'收款结算填写工具/{identity.EXE_NAMES[1]}': b'2',
            '收款结算填写工具/ms-playwright（勿动）/chromium-1161/chrome-win/chrome.exe': b'3',
        })
        with self.assertRaisesRegex(ValueError, '不得包含内置浏览器'):
            prep.prepare(bad_browser, clean_full, destination=self.tmp / 'out1')

        # 2. 轻量包包含个人配置必须拒绝
        bad_cfg = self.tmp / 'with_cfg.zip'
        make_zip(bad_cfg, {
            f'收款结算填写工具/{identity.EXE_NAMES[0]}': b'1',
            f'收款结算填写工具/{identity.EXE_NAMES[1]}': b'2',
            '收款结算填写工具/收款结算单填写顺序设置.json': b'{}',
        })
        with self.assertRaisesRegex(ValueError, '不得包含个人配置'):
            prep.prepare(bad_cfg, clean_full, destination=self.tmp / 'out2')

        # 3. 完整包缺少浏览器必须拒绝
        with self.assertRaisesRegex(ValueError, '缺少内置 Chromium'):
            prep.prepare(clean_update, clean_update, destination=self.tmp / 'out3')

        # 4. 完整包 + 轻量包齐全时，生成的 release-plan.json 同时记录 update 与 full 且通过发布前校验
        if (REPO / '.release-secrets' / 'ncc-receipt-signing.dpapi').is_file():
            out_ok = prep.prepare(clean_update, clean_full, destination=self.tmp / 'out_ok')
            plan, verified, _manifest_bytes, resolved = publisher.validate_release_folder(out_ok)
            self.assertEqual([role for role, _ in resolved], ['update', 'full'])
            self.assertEqual(len(plan['assets']), 2)
            self.assertEqual(verified['version'], identity.APP_VERSION)

            # 若人为删掉 full 记录，发布脚本必须拒绝
            broken_plan = dict(plan, assets=[plan['assets'][0]])
            (out_ok / 'release-plan.json').write_text(
                json.dumps(broken_plan, ensure_ascii=False, indent=2),
                encoding='utf-8',
            )
            with self.assertRaisesRegex(ValueError, '必须同时上传轻量更新包'):
                publisher.validate_release_folder(out_ok)

    def test_gui_and_build_script_wire_auto_updater(self):
        main_src = (REPO / '收款结算单工具.py').read_text(encoding='utf-8-sig')
        ast.parse(main_src)
        self.assertIn('ReceiptUpdateManager', main_src)
        self.assertIn('ReceiptUpdateController', main_src)
        self.assertIn("text='检查/自动更新...'", main_src)
        self.assertIn('update_controller.schedule()', main_src)
        self.assertIn('update_controller.cancel_timers()', main_src)

        build_src = (REPO / 'build_scripts' / 'build_release.py').read_text(encoding='utf-8-sig')
        self.assertIn('update_public_key.json', build_src)
        self.assertIn('update_install.ps1', build_src)


if __name__ == '__main__':
    unittest.main()
