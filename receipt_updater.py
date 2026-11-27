"""NCC 收款结算单工具独立自动更新客户端（仅使用标准库，支持双版本 EXE 与签名校验）。"""
import base64
from contextlib import contextmanager
import hashlib
import hmac
import json
import ntpath
import os
import re
import shutil
import stat
import subprocess
import sys
import threading
import time
import uuid
import zipfile
from pathlib import Path
from urllib.error import HTTPError
from urllib.parse import unquote, urlsplit
from urllib.request import Request, HTTPRedirectHandler, ProxyHandler, build_opener

import receipt_release_identity as identity

MAX_ARCHIVE = 1024 * 1024 * 1024
MAX_EXPANDED = 2 * 1024 * 1024 * 1024

INSTALL_ROOT_FILES = frozenset((
    *identity.EXE_NAMES,
    'V7双版本及GUI功能说明.txt',
    '运行调试（选择版本并保留日志）.bat',
    'README.md',
))
# 严禁自动更新覆盖用户个人配置、填写好的底稿模板或已配置的对照表方案。
SKIP_ROOTS = frozenset((
    '收款结算单填写顺序设置.json',
    '收款结算单时间设置.json',
    '收款结算单底稿导入模板.xlsm',
    '对照表方案.xlsx',
))
_RESERVED = re.compile(r'^(con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\.|$)', re.I)


class UpdateError(RuntimeError):
    pass


def clean_child_environment():
    env = dict(os.environ)
    for name in list(env):
        if name == '_MEIPASS2' or name.startswith('_PYI_') or name in ('TCL_LIBRARY', 'TK_LIBRARY', 'PSModulePath'):
            env.pop(name, None)
    env['PYINSTALLER_RESET_ENVIRONMENT'] = '1'
    bundled = getattr(sys, '_MEIPASS', None)
    if bundled:
        root = os.path.normcase(os.path.abspath(str(bundled)))
        env['PATH'] = os.pathsep.join(
            part for part in env.get('PATH', '').split(os.pathsep)
            if part and not (
                os.path.normcase(os.path.abspath(part)) == root
                or os.path.normcase(os.path.abspath(part)).startswith(root + os.sep)
            )
        )
    return env


@contextmanager
def system_dll_search():
    if os.name != 'nt' or not getattr(sys, 'frozen', False):
        yield
        return
    import ctypes
    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    kernel.GetDllDirectoryW.argtypes = [ctypes.c_uint32, ctypes.c_wchar_p]
    kernel.SetDllDirectoryW.argtypes = [ctypes.c_wchar_p]
    buffer = ctypes.create_unicode_buffer(32768)
    kernel.GetDllDirectoryW(len(buffer), buffer)
    previous = buffer.value
    if not kernel.SetDllDirectoryW(None):
        raise UpdateError('无法准备系统安装助手环境。')
    try:
        yield
    finally:
        kernel.SetDllDirectoryW(previous or None)


def await_installer_ready(child, ready_path, plan_hash, log_path, timeout=20):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if ready_path.exists():
            try:
                marker = json_bytes(ready_path.read_bytes())
                if (
                    marker.get('app_id') == identity.APP_ID
                    and marker.get('plan_sha256') == plan_hash
                    and child.poll() is None
                ):
                    return
            except (ValueError, OSError):
                pass
        code = child.poll()
        if code is not None:
            detail = '退出码 ' + str(code)
            try:
                result = json_bytes((ready_path.parents[1] / 'last-result.json').read_bytes())
                if result.get('plan_sha256') == plan_hash:
                    detail = result.get('message', detail)
            except (ValueError, OSError):
                pass
            raise UpdateError(f'安装助手未能启动（{detail}），原程序未关闭。日志：{log_path}')
        time.sleep(0.1)
    child.terminate()
    child.wait(timeout=5)
    raise UpdateError('安装助手启动超时，原程序未关闭。日志：' + str(log_path))


def version_tuple(value):
    match = re.fullmatch(r'[vV]?(\d+(?:\.\d+){1,3})', str(value))
    if match is None:
        raise UpdateError('更新版本号格式无效。')
    parts = tuple(int(x) for x in match.group(1).split('.'))
    return parts + (0,) * (4 - len(parts))


def sha256_file(path):
    digest = hashlib.sha256()
    with Path(path).open('rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            digest.update(block)
    return digest.hexdigest()


def trusted_asset_url(url, version):
    try:
        parsed = urlsplit(url)
        prefix = '/{}/{}/releases/download/{}{}/'.format(
            identity.GITHUB_OWNER, identity.GITHUB_REPO, identity.TAG_PREFIX, version
        )
        filename = unquote(parsed.path[len(prefix):])
        return (
            parsed.scheme == 'https'
            and parsed.netloc == 'github.com'
            and not parsed.query
            and not parsed.fragment
            and parsed.path.startswith(prefix)
            and '/' not in parsed.path[len(prefix):]
            and re.fullmatch(r'[A-Za-z0-9_.-]+\.zip', filename) is not None
        )
    except ValueError:
        return False


def validate_manifest(raw):
    if not isinstance(raw, dict) or raw.get('schema') != 1 or raw.get('app_id') != identity.APP_ID:
        raise UpdateError('更新清单不属于 NCC 收款结算单工具，已拒绝。')
    version = str(raw.get('version', ''))
    version_tuple(version)
    package = raw.get('full')
    if not isinstance(package, dict) or not trusted_asset_url(str(package.get('url', '')), version):
        raise UpdateError('更新包地址不属于 NCC 收款结算单独立发布通道。')
    if re.fullmatch('[0-9a-fA-F]{64}', str(package.get('sha256', ''))) is None:
        raise UpdateError('更新包缺少有效的 SHA-256 校验值。')
    size = package.get('size_bytes')
    if isinstance(size, bool) or not isinstance(size, int) or not 0 < size <= MAX_ARCHIVE:
        raise UpdateError('更新包大小超出允许范围。')
    return dict(
        version=version,
        notes=str(raw.get('notes', ''))[:20000],
        released_at=str(raw.get('released_at', '')),
        package=dict(
            url=package['url'],
            sha256=package['sha256'].lower(),
            size_bytes=size,
        ),
    )


class SafeRedirect(HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        parsed = urlsplit(newurl)
        allowed = {
            'github.com', 'api.github.com', 'raw.githubusercontent.com',
            'release-assets.githubusercontent.com', 'objects.githubusercontent.com',
            'ghfast.top', 'gh-proxy.com', 'gcore.jsdelivr.net', 'cdn.jsdelivr.net',
        }
        if (
            parsed.scheme != 'https'
            or parsed.hostname not in allowed
            or parsed.username
            or parsed.password
            or parsed.port
        ):
            raise UpdateError('更新下载发生未允许的跳转。')
        return super().redirect_request(req, fp, code, msg, headers, newurl)


def open_url(url, timeout=15):
    # 独立直连，不依赖用户电脑上正在运行的代理软件。
    return build_opener(ProxyHandler({}), SafeRedirect()).open(
        Request(
            url,
            headers={
                'User-Agent': 'NCC-Receipt/' + identity.APP_VERSION,
                'Cache-Control': 'no-cache',
            },
        ),
        timeout=timeout,
    )


def json_bytes(data):
    def pairs(items):
        result = {}
        for key, value in items:
            if key in result:
                raise UpdateError('更新清单包含重复字段。')
            result[key] = value
        return result
    return json.loads(data.decode('utf-8-sig'), object_pairs_hook=pairs)


def verify_envelope(envelope):
    """严格验证完整 PKCS#1 v1.5 SHA-256 RSA 签名；镜像无法自行替换更新清单。"""
    try:
        if not isinstance(envelope, dict) or envelope.get('key_id') != identity.SIGNING_KEY_ID:
            raise ValueError('key')
        payload = base64.b64decode(envelope['payload'], validate=True)
        signature = base64.b64decode(envelope['signature'], validate=True)
        modulus = int.from_bytes(base64.b64decode(identity.SIGNING_MODULUS, validate=True), 'big')
        length = (modulus.bit_length() + 7) // 8
        if len(signature) != length or int.from_bytes(signature, 'big') >= modulus or modulus.bit_length() < 2048:
            raise ValueError('size')
        decoded = pow(int.from_bytes(signature, 'big'), 65537, modulus).to_bytes(length, 'big')
        digest_info = bytes.fromhex('3031300d060960864801650304020105000420') + hashlib.sha256(payload).digest()
        expected = b'\x00\x01' + b'\xff' * (length - len(digest_info) - 3) + b'\x00' + digest_info
        if not hmac.compare_digest(decoded, expected):
            raise ValueError('signature')
        return validate_manifest(json_bytes(payload))
    except UpdateError:
        raise
    except (ValueError, KeyError, TypeError, OverflowError):
        raise UpdateError('更新清单签名校验失败，已停止更新。') from None


def fetch_manifest():
    root = '{}/{}'.format(identity.GITHUB_OWNER, identity.GITHUB_REPO)
    raw_url = 'https://raw.githubusercontent.com/{}/{}/{}'.format(
        root, identity.GITHUB_BRANCH, identity.MANIFEST_PATH
    )
    urls = [mirror + raw_url for mirror in identity.DOWNLOAD_MIRRORS]
    urls += [
        'https://gcore.jsdelivr.net/gh/{}@{}/{}'.format(root, identity.GITHUB_BRANCH, identity.MANIFEST_PATH),
        'https://cdn.jsdelivr.net/gh/{}@{}/{}'.format(root, identity.GITHUB_BRANCH, identity.MANIFEST_PATH),
        'https://api.github.com/repos/{}/contents/{}?ref={}'.format(root, identity.MANIFEST_PATH, identity.GITHUB_BRANCH),
        raw_url,
    ]
    missing = 0
    rejected = False
    for url in urls:
        try:
            with open_url(url, 8) as response:
                body = response.read(512 * 1024 + 1)
            if len(body) > 512 * 1024:
                raise UpdateError('更新清单过大。')
            raw = json_bytes(body)
            if url.startswith('https://api.github.com/'):
                if raw.get('encoding') != 'base64':
                    raise UpdateError('官方更新清单格式无效。')
                raw = json_bytes(base64.b64decode(raw['content'], validate=False))
            manifest = verify_envelope(raw)
            manifest['source'] = '国内直连通道' if url not in urls[-2:] else 'GitHub 官方备用'
            return manifest
        except HTTPError as exc:
            if exc.code == 404:
                missing += 1
        except UpdateError:
            rejected = True
        except Exception:
            pass
    if missing == len(urls):
        raise UpdateError('收款工具自动更新清单尚未发布，当前程序仍可正常使用。')
    if rejected:
        raise UpdateError('未取得签名校验通过的收款工具更新清单，请稍后重试。')
    raise UpdateError('暂时无法连接更新通道，请稍后重试；当前程序可继续使用。')


def _parts(name):
    raw = name.replace('\\', '/')
    parts = raw.rstrip('/').split('/')
    if (
        not raw
        or raw.startswith('/')
        or ntpath.splitdrive(raw)[0]
        or any(not p or p in ('.', '..') for p in parts)
        or any(
            any(c in p for c in '<>:"|?*')
            or p.endswith((' ', '.'))
            or _RESERVED.match(p)
            or any(ord(c) < 32 for c in p)
            for p in parts
        )
    ):
        raise UpdateError('更新压缩包包含非法路径。')
    return parts


def stage_package(zip_path, stage_dir):
    stage_dir = Path(stage_dir)
    if stage_dir.exists():
        raise UpdateError('更新暂存目录已存在，未覆盖。')
    with zipfile.ZipFile(str(zip_path)) as archive:
        infos = archive.infolist()
        if not infos or len(infos) > 20000 or sum(i.file_size for i in infos) > MAX_EXPANDED:
            raise UpdateError('更新包为空或解压大小超限。')
        parsed = [(info, _parts(info.filename)) for info in infos]
        roots = {parts[0] for _, parts in parsed}
        known_roots = INSTALL_ROOT_FILES | SKIP_ROOTS | {'ms-playwright（勿动）'}
        strip = len(roots) == 1 and next(iter(roots)) not in known_roots
        seen, selected = set(), []
        for info, parts in parsed:
            mode = info.external_attr >> 16
            if stat.S_ISLNK(mode) or info.flag_bits & 1:
                raise UpdateError('更新包包含链接或加密文件，已拒绝。')
            if strip:
                parts = parts[1:]
            if not parts:
                continue
            lower = '/'.join(parts).casefold()
            if lower in seen and not info.is_dir():
                raise UpdateError('更新包存在重名文件。')
            seen.add(lower)
            top = parts[0]
            if top in SKIP_ROOTS:
                continue
            if top not in INSTALL_ROOT_FILES and top != 'ms-playwright（勿动）':
                raise UpdateError('更新包包含不允许覆盖的内容：' + top)
            if top in INSTALL_ROOT_FILES and len(parts) != 1:
                raise UpdateError('更新包文件结构异常。')
            if not info.is_dir():
                selected.append((info, parts))
        present_exes = {parts[0] for _, parts in selected if len(parts) == 1}
        if not set(identity.EXE_NAMES).issubset(present_exes):
            raise UpdateError('更新包缺少双版本收款结算单主程序。')
        if shutil.disk_usage(str(stage_dir.parent)).free < MAX_EXPANDED // 4 + sum(i.file_size for i, _ in selected) * 2:
            raise UpdateError('可用磁盘空间不足以暂存和备份更新。')
        stage_dir.mkdir()
        files = []
        for info, parts in selected:
            target = stage_dir.joinpath(*parts)
            target.parent.mkdir(parents=True, exist_ok=True)
            with archive.open(info) as source, target.open('wb') as dest:
                shutil.copyfileobj(source, dest, 1024 * 1024)
            files.append({'path': '/'.join(parts), 'sha256': sha256_file(target)})
        return files


def validate_target(app_dir, relative):
    parts = _parts(relative)
    if not (
        (len(parts) == 1 and parts[0] in INSTALL_ROOT_FILES)
        or (len(parts) > 1 and parts[0] == 'ms-playwright（勿动）')
    ):
        raise UpdateError('不允许更新此文件。')
    root = Path(app_dir).resolve()
    target = root.joinpath(*parts)
    try:
        target.resolve().relative_to(root)
    except ValueError:
        raise UpdateError('更新目标路径越界。') from None
    for part in [target] + list(target.parents):
        if part == root:
            break
        if part.exists():
            attributes = getattr(part.stat(), 'st_file_attributes', 0)
            if part.is_symlink() or attributes & 0x400:
                raise UpdateError('更新目标存在链接或目录联接。')
    return target


class UpdateManager:
    def __init__(self, base_dir):
        self.base_dir = Path(base_dir).resolve()
        self.root = self.base_dir / '.updates'
        self.lock = threading.RLock()
        self.operation = threading.Lock()
        self.manifest = self.ready = None
        self.state = {
            'status': 'idle',
            'message': '启动后会自动后台检查更新。',
            'progress': 0,
            'current_version': identity.APP_VERSION,
            'latest_version': '',
            'notes': '',
            'source': '',
        }

    def status(self):
        with self.lock:
            return dict(self.state)

    def _state(self, **values):
        with self.lock:
            self.state.update(values)
            return dict(self.state)

    def check(self):
        if not self.operation.acquire(False):
            return self.status()
        try:
            self.manifest = self.ready = None
            self._state(
                status='checking',
                message='正在后台检查新版本…',
                progress=0,
                latest_version='',
                notes='',
                source='',
            )
            manifest = fetch_manifest()
            self.manifest = manifest
            available = version_tuple(manifest['version']) > version_tuple(identity.APP_VERSION)
            return self._state(
                status='available' if available else 'up_to_date',
                latest_version=manifest['version'],
                notes=manifest['notes'],
                source=manifest.get('source', ''),
                message='发现新版本 V' + manifest['version'] if available else '当前已是最新版本。',
            )
        except Exception as exc:
            return self._state(
                status='error',
                message=str(exc) if isinstance(exc, UpdateError) else '检查更新失败，请稍后重试。',
            )
        finally:
            self.operation.release()

    def download(self):
        if not self.operation.acquire(False):
            return self.status()
        try:
            if self.manifest is None or version_tuple(self.manifest['version']) <= version_tuple(identity.APP_VERSION):
                raise UpdateError('请先检查是否有新版本。')
            self.ready = None
            self.root.mkdir(parents=True, exist_ok=True)
            if self.root.is_symlink() or getattr(self.root.stat(), 'st_file_attributes', 0) & 0x400:
                raise UpdateError('更新目录存在链接，未继续。')
            job = self.root / uuid.uuid4().hex
            job.mkdir()
            package = self.manifest['package']
            part = job / 'package.part'
            urls = [
                ('国内直连通道 ' + str(i + 1), mirror + package['url'])
                for i, mirror in enumerate(identity.DOWNLOAD_MIRRORS)
            ] + [('GitHub 官方备用', package['url'])]
            verified = False
            for label, url in urls:
                self._state(status='downloading', message='正在下载更新包…', progress=0, source=label)
                try:
                    hasher, done = hashlib.sha256(), 0
                    with open_url(url, 15) as response, part.open('wb') as stream:
                        while True:
                            block = response.read(128 * 1024)
                            if not block:
                                break
                            done += len(block)
                            if done > package['size_bytes']:
                                raise UpdateError('下载大小超过清单声明。')
                            stream.write(block)
                            hasher.update(block)
                            self._state(progress=min(99, int(done * 100 / package['size_bytes'])))
                    if done != package['size_bytes'] or hasher.hexdigest() != package['sha256']:
                        raise UpdateError('下载完整性校验失败。')
                    verified = True
                    break
                except Exception:
                    if part.exists():
                        part.unlink()
            if not verified:
                raise UpdateError('官方下载和国内镜像通道均未取得校验通过的更新包。')
            self._state(status='staging', message='校验通过，正在准备更新文件…', progress=99)
            stage = job / 'stage'
            files = stage_package(part, stage)
            present_exes = {f['path'] for f in files}
            if not set(identity.EXE_NAMES).issubset(present_exes):
                raise UpdateError('更新包缺少双版本收款结算单主程序。')
            for file_entry in files:
                validate_target(self.base_dir, file_entry['path'])
            part.unlink()
            self.ready = {'job': job, 'stage': stage, 'files': files, 'version': self.manifest['version']}
            return self._state(
                status='ready',
                message='更新已校验完成。停止任务后，可安装并重启。',
                progress=100,
            )
        except Exception as exc:
            self.ready = None
            return self._state(
                status='error',
                message=str(exc) if isinstance(exc, UpdateError) else '下载或准备更新失败，当前版本未修改。',
            )
        finally:
            self.operation.release()

    def installation_plan(self):
        if self.ready is None or self.status()['status'] not in ('ready', 'starting'):
            raise UpdateError('尚无已校验的更新。')
        ready = self.ready
        stage = ready['stage'].resolve()
        stage.relative_to(self.root.resolve())
        for file_entry in ready['files']:
            validate_target(self.base_dir, file_entry['path'])
            source = stage.joinpath(*_parts(file_entry['path']))
            if source.is_symlink() or sha256_file(source) != file_entry['sha256']:
                raise UpdateError('暂存文件发生变化，请重新下载更新。')
        return {
            'app_id': identity.APP_ID,
            'app_dir': str(self.base_dir),
            'stage_dir': str(stage),
            'backup_dir': str(ready['job'] / 'backup'),
            'exe_name': identity.current_exe_name(),
            'exe_names': list(identity.EXE_NAMES),
            'pid': os.getpid(),
            'version': ready['version'],
            'files': ready['files'],
        }

    def start_installer(self, is_busy):
        if is_busy():
            raise UpdateError('请先停止当前填报或转换任务后再安装更新。')
        if not getattr(sys, 'frozen', False):
            raise UpdateError('源码模式仅支持检查和下载；请在打包版 EXE 中安装更新。')
        if os.name != 'nt':
            raise UpdateError('自动安装目前只支持 Windows。')
        if not self.operation.acquire(False):
            raise UpdateError('更新操作正在进行中。')
        try:
            self._state(status='starting', message='正在校验文件并启动安装助手，请稍候…')
            if is_busy():
                raise UpdateError('任务状态已变化，未开始安装。')
            plan = self.installation_plan()
            plan['parent_pid'] = os.getppid()
            ready = self.ready
            plan_path = ready['job'] / 'install-plan.json'
            plan_path.write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding='utf-8-sig')
            resource = Path(getattr(sys, '_MEIPASS', Path(__file__).parent)) / 'update_install.ps1'
            helper = ready['job'] / 'install.ps1'
            shutil.copyfile(resource, helper)
            ready_path = ready['job'] / ('helper-ready-' + uuid.uuid4().hex + '.json')
            log_path = ready['job'] / 'installer.log'
            plan_hash = sha256_file(plan_path)
            powershell = (
                Path(os.environ.get('SystemRoot', r'C:\Windows'))
                / 'System32' / 'WindowsPowerShell' / 'v1.0' / 'powershell.exe'
            )
            with log_path.open('ab') as log, system_dll_search():
                child = subprocess.Popen(
                    [
                        str(powershell), '-NoProfile', '-ExecutionPolicy', 'Bypass',
                        '-WindowStyle', 'Hidden', '-File', str(helper),
                        '-PlanPath', str(plan_path), '-PlanSha256', plan_hash,
                        '-ReadyPath', str(ready_path),
                    ],
                    creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0),
                    close_fds=True,
                    stdin=subprocess.DEVNULL,
                    stdout=log,
                    stderr=subprocess.STDOUT,
                    env=clean_child_environment(),
                )
            await_installer_ready(child, ready_path, plan_hash, log_path)
            return self._state(status='installing', message='正在关闭程序并安装更新…')
        except Exception as exc:
            self._state(status='ready' if self.ready is not None else 'error', message=str(exc))
            raise
        finally:
            self.operation.release()
