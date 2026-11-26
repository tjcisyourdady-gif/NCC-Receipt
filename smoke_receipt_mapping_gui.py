"""隔离Tk对话框与后台转换烟测；不打开Excel/NCC，不修改正式模板。"""
import shutil
import ctypes
import os
from pathlib import Path
import sys
import time
import tkinter as tk
from tkinter import ttk
from unittest.mock import patch

import receipt_excel_converter as converter
from test_receipt_mapping_table import MappingTableTests, TEMPLATE


case = MappingTableTests()
case.setUp()
root = None
try:
    template = case.directory / TEMPLATE.name
    shutil.copy2(TEMPLATE, template)
    original = template.read_bytes()
    # 与正式程序的Tcl初始化一致；只影响本次隔离测试进程。
    if os.name == 'nt':
        runtime = Path(__file__).resolve().parent / '付款结算单工具开发/tcl_from_receipt'
        # 某些本机Tcl构建不能从含中文目录加载初始化脚本，用ASCII临时副本。
        for name in ('_tcl_data', '_tk_data'):
            shutil.copytree(runtime / name, case.directory / name)
        os.environ['TCL_LIBRARY'] = (case.directory / '_tcl_data').as_posix()
        os.environ['TK_LIBRARY'] = (case.directory / '_tk_data').as_posix()
        tcl = ctypes.WinDLL(str(Path(sys.base_prefix) / 'DLLs' / 'tcl86t.dll'))
        tcl.Tcl_FindExecutable.argtypes = (ctypes.c_char_p,)
        tcl.Tcl_FindExecutable(sys.executable.encode('utf-8'))
    root = tk.Tk()
    root.withdraw()
    output = tk.StringVar(root)
    observed = {'choices': None, 'success': [], 'errors': [], 'confirm': []}

    def descendants(widget):
        for child in widget.winfo_children():
            yield child
            yield from descendants(child)

    def select_scheme():
        widgets = list(descendants(root))
        choices = [w for w in widgets if isinstance(w, ttk.Combobox)]
        if not choices:
            root.after(20, select_scheme)
            return
        observed['choices'] = tuple(choices[0]['values'])
        choices[0].set('工资专用')
        for widget in widgets:
            if isinstance(widget, ttk.Button) and widget.cget('text') == '下一步':
                widget.invoke()
                return
        raise AssertionError('No next button')

    def confirm(*args, **kwargs):
        observed['confirm'].append(args)
        return True

    root.after(20, select_scheme)
    with patch.object(converter.filedialog, 'askopenfilename', return_value=str(case.source)), \
            patch.object(converter.messagebox, 'askokcancel', side_effect=confirm), \
            patch.object(converter.messagebox, 'showinfo', side_effect=lambda *a, **k: observed['success'].append(a)), \
            patch.object(converter.messagebox, 'showerror', side_effect=lambda *a, **k: observed['errors'].append(a)):
        converter.start_conversion_wizard(root, output, str(template))
        timeout = time.monotonic() + 20
        while getattr(root, '_receipt_conversion_busy', False) and time.monotonic() < timeout:
            root.update()
            time.sleep(0.02)
    assert not observed['errors'], observed['errors']
    assert observed['choices'] == ('工资专用',), observed
    assert observed['confirm'] and observed['success'], observed
    assert output.get() == str(template), output.get()
    assert not root._receipt_conversion_busy
    backups = list((case.directory / '底稿备份').glob('*.xlsm'))
    assert len(backups) == 1 and backups[0].read_bytes() == original
    assert converter.read_receipt_macro(template) == converter.read_receipt_macro(TEMPLATE)
    print('PASS native Tk: source picker -> renamed scheme -> confirmation -> worker -> backup -> auto-select macro output')
finally:
    if root is not None:
        root.destroy()
    case.doCleanups()
