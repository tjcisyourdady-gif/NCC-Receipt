"""打包版在导入 tkinter 前初始化随包携带的 Tcl/Tk 运行库。"""
import os
import sys


if getattr(sys, 'frozen', False):
    runtime_root = sys._MEIPASS
    try:
        import ctypes

        tcl_dll_path = os.path.join(runtime_root, 'tcl86t.dll')
        tcl_dll = ctypes.WinDLL(tcl_dll_path)
        tcl_dll.Tcl_FindExecutable.argtypes = (ctypes.c_char_p,)
        tcl_dll.Tcl_FindExecutable(
            tcl_dll_path.encode('mbcs', errors='replace')
        )
    except Exception:
        pass
    os.environ.setdefault('TCL_LIBRARY', os.path.join(runtime_root, '_tcl_data'))
    os.environ.setdefault('TK_LIBRARY', os.path.join(runtime_root, '_tk_data'))
