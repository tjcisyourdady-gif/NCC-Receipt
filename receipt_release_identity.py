"""NCC 收款结算单工具发布身份与自动更新公钥配置。"""
import json
import sys
from pathlib import Path

APP_ID = 'ncc-receipt'
APP_VERSION = '7.1.0'
DISPLAY_VERSION = 'V7.1.0（自动检测更新版）'
RELEASE_NOTES = (
    '新增启动自动检测更新包与一键安全更新功能：每次打开程序自动在后台检查 '
    'NCC-Receipt 独立仓库的已签名更新清单（优先国内直连镜像，备用 GitHub 官方通道）。'
    '发现新版本时弹出“更新 / 忽略”提示，也支持随时在主界面点击“检查/自动更新...”手动检查；'
    '更新过程执行 RSA-3072 签名校验与 SHA-256 完整性校验，安装前自动备份并在失败时回滚，'
    '自动保留已有的底稿模板、对照表方案、填写顺序设置、时间设置与内置浏览器。'
)
EXE_NAMES = (
    '收款结算单填写工具V7-正常版.exe',
    '收款结算单填写工具V7-可修改时间版-卡顿使用.exe',
)
GITHUB_OWNER = 'tjcisyourdady-gif'
GITHUB_REPO = 'NCC-Receipt'
GITHUB_BRANCH = 'main'
MANIFEST_PATH = 'latest.json'
TAG_PREFIX = 'ncc-receipt-v'
DOWNLOAD_MIRRORS = ('https://ghfast.top/', 'https://gh-proxy.com/')

_key_path = Path(getattr(sys, '_MEIPASS', Path(__file__).resolve().parent)) / 'update_public_key.json'
_key = json.loads(_key_path.read_text(encoding='utf-8-sig'))
SIGNING_KEY_ID = _key['key_id']
SIGNING_MODULUS = _key['modulus']


def current_exe_name():
    """返回当前运行的 EXE 文件名（区分正常版与可修改时间版），源码模式下默认正常版。"""
    if getattr(sys, 'frozen', False):
        name = Path(sys.executable).name
        if name in EXE_NAMES:
            return name
    return EXE_NAMES[0]
