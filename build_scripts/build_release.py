"""从源码构建 V7 双版本 EXE（正常版 / 可修改时间版-卡顿使用）。"""
import argparse
import ast
import os
from pathlib import Path
import shutil
import subprocess
import sys

SCRIPT_DIR = Path(__file__).resolve().parent
REPO_ROOT = SCRIPT_DIR.parent


def build_variants(python_exe: str, out_dir: Path, tcl_dir: Path | None = None):
    source = (REPO_ROOT / '收款结算单工具.py').read_text(encoding='utf-8-sig')
    dist_dir = out_dir / 'dist'
    if dist_dir.exists():
        shutil.rmtree(dist_dir)

    extra_datas = [
        (str(REPO_ROOT / 'update_public_key.json'), '.'),
        (str(REPO_ROOT / 'update_install.ps1'), '.'),
    ]
    extra_binaries = []
    extra_hooks = []
    extra_rthooks = []
    if tcl_dir is not None and tcl_dir.is_dir():
        extra_datas += [
            (str(tcl_dir / '_tcl_data'), '_tcl_data'),
            (str(tcl_dir / '_tk_data'), '_tk_data'),
        ]
        extra_binaries = [
            (str(tcl_dir / 'tcl86t.dll'), '.'),
            (str(tcl_dir / 'tk86t.dll'), '.'),
        ]
        zlib_dll = Path(sys.base_prefix) / 'DLLs' / 'zlib1.dll'
        if zlib_dll.is_file():
            extra_binaries.append((str(zlib_dll), '.'))
        extra_hooks = [str(SCRIPT_DIR / 'custom_hooks')]
        extra_rthooks = [str(SCRIPT_DIR / 'custom_runtime_hooks' / 'pyinstaller_tk_runtime_hook.py')]

    for normal in (False, True):
        name = '收款结算单填写工具V7-' + ('正常版' if normal else '可修改时间版-卡顿使用')
        srcdir = out_dir / ('normal_source' if normal else 'adjustable_source')
        if srcdir.exists():
            shutil.rmtree(srcdir)
        srcdir.mkdir(parents=True, exist_ok=True)
        content = (
            source.replace('ENABLE_TIME_SETTINGS = True  #', 'ENABLE_TIME_SETTINGS = False  #')
            if normal else source
        )
        assert ('ENABLE_TIME_SETTINGS = False' in content) == normal
        ast.parse(content)
        entry = srcdir / '收款结算单工具.py'
        entry.write_text(content, encoding='utf-8')

        spec = rf'''# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all
datas, binaries, hiddenimports = collect_all('playwright')
datas += {extra_datas!r}
binaries += {extra_binaries!r}
a = Analysis(
    [{str(entry)!r}],
    pathex=[{str(REPO_ROOT)!r}],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath={extra_hooks!r},
    runtime_hooks={extra_rthooks!r},
    excludes=['pandas', 'numpy', 'xlrd', 'xlsxwriter'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz, a.scripts, a.binaries, a.datas, [],
    name={name!r},
    debug=False, bootloader_ignore_signals=False,
    strip=False, upx=False, console=True,
)
'''
        specpath = srcdir / 'release.spec'
        specpath.write_text(spec, encoding='utf-8')
        print(f'=== Building {name} ===', flush=True)
        subprocess.run(
            [
                python_exe, '-m', 'PyInstaller', '--noconfirm',
                '--distpath', str(dist_dir),
                '--workpath', str(srcdir / 'work'),
                str(specpath),
            ],
            check=True,
            env=dict(os.environ),
        )
    print(f'=== Build complete: {dist_dir} ===', flush=True)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--python', default=sys.executable, help='Python executable with playwright/PIL/openpyxl/PyInstaller')
    parser.add_argument('--out', default=str(REPO_ROOT / 'build'), help='Build output directory')
    parser.add_argument('--tcl-dir', default='', help='Optional directory containing _tcl_data/_tk_data/tcl86t.dll/tk86t.dll')
    args = parser.parse_args()
    tcl_dir = Path(args.tcl_dir) if args.tcl_dir else None
    build_variants(args.python, Path(args.out), tcl_dir)


if __name__ == '__main__':
    main()
