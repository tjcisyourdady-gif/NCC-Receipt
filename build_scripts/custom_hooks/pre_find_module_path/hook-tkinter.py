"""打包时保留 tkinter 源码；Tcl/Tk 资源由构建脚本显式随包提供。"""


def pre_find_module_path(hook_api):
    """不要因构建机的 Tcl 安装异常而清空 tkinter 的搜索路径。"""
    return None
