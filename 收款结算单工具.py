# -*- coding: utf-8 -*-
"""
==================== 收款结算单自动化制单工具 ====================
基于 ClearDealings3.4（往来清理工具）同架构重构。
运行方式：
    1. 双击运行“收款结算单工具.exe”（打包后），或用 Python 运行本脚本
    2. 在界面中选择填好的《收款结算单底稿导入模板.xlsm》
    3. 点击“开始运行程序”，在打开的浏览器中登录 NCC
    4. 程序自动进入“收款结算单”节点按底稿逐笔制单，并按实际布局回写填写状态和异常原因
================================================================
"""
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import threading
import queue
import traceback
import json
import math
import tempfile
from decimal import Decimal, InvalidOperation
from playwright.sync_api import sync_playwright
from playwright.sync_api import expect
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from PIL import Image, ImageTk
import time
import os
import re
import sys
from receipt_data_rules import (
    load_receipt_workbook, read_receipt_macro,
    normalize_receipt_value, normalize_receipt_row, receipt_sheet_layout,
)

try:
    from receipt_excel_converter import start_conversion_wizard
except ModuleNotFoundError as import_error:
    if import_error.name != 'receipt_excel_converter':
        raise
    # 兼容早期源码目录：转换模块也可随内部运行文件一同存放。
    internal_dependency_dir = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), 'ms-playwright（勿动）'
    )
    if internal_dependency_dir not in sys.path:
        sys.path.insert(0, internal_dependency_dir)
    from receipt_excel_converter import start_conversion_wizard


def resource_path(rel):
    """
    定位资源文件，兼容两种运行方式：
    - 源码调试运行：资源在 项目目录\\dist\\ 下
    - 打包 exe 运行：资源在 exe 同目录下（和旧工具一致）
    """
    rel = rel.lstrip('./')
    if getattr(sys, 'frozen', False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    for cand in (os.path.join(base, rel), os.path.join(base, 'dist', rel)):
        if os.path.exists(cand):
            return cand
    return os.path.join(base, rel)   # 都找不到时返回默认路径，让报错信息可读


# ╔══════════════════════════════════════════════════════════════╗
# ║  【配置区】新表所有“数据位置 / 输入框位置”只改这里，不改下面代码  ║
# ╚══════════════════════════════════════════════════════════════╝

# ---------- 1. 旧版底稿基准列；新增C/D部门列由 receipt_sheet_layout 同步顺延 ----------
EXCEL_CONFIG = {
    "sheet_name": "收款结算明细",      # TODO 底稿工作表名称（按新模板改）
    "header_row": 2,                  # 新版模板表头所在行
    "start_row": 3,                   # TODO 第一条数据所在行（1=注意事项 2=表头）
    "status_col": 28,                 # AB列：仅允许空值/“异常”/“填报成功”
    "reason_col": 29,                 # AC列：回写异常原因
    "bill_number_col": 8,             # H 列：仅保存成功后回写 NCC 单据编号
    # 字段 → 绝对列号（E=5, H=8, ...），后续步骤用到新列时在这里加
    "fields": {
        "company_code":             1,   # A 列：公司编码（填入“结算财务组织”）
        "company_name":             2,   # B 列：公司名称（用于分组及运行日志）
        "bank_account":             3,   # C 列：银行账户（填入“收款银行账户”）
        "bill_date":                4,   # D 列：单据日期（填入“单据日期”）
        "merge_key_i":              5,   # E 列：并表识别键
        "expense_period":           6,   # F 列：费用所属期间（明细仅填写到月份）
        "summary":                  7,   # G 列：摘要
        "bill_number":              8,   # H 列：单据号（保存后回写）
        "customer_code":            9,   # I 列：客商编码（填入明细“客户”）
        "customer_name":           10,   # J 列：客户（判断交易对象类型）
        "accounting_subject":      11,   # K 列：会计科目
        "auxiliary_accounting":    12,   # L 列：辅助核算（编号/名称，仅填写编号）
        "represented_party":       13,   # M 列：被代付账套/被代收账套名称
        "represented_project":     14,   # N 列：被代收方（被收回方）项目档案
        "business_category":       15,   # O 列：经营类别名称
        "internal_collection_flag": 16,  # P 列：代收代付情况（是/否）
        "internal_collection_type": 17,  # Q 列：代收代付类型
        "receipt_method":          18,   # R 列：收款方式名称
        "settlement_method":       19,   # S 列：结算方式名称
        "actual_charge_type":      20,   # T 列：收费类型名称
        "new_income_expense_item": 21,   # U 列：NCC收支项目名称
        "project_archive_name":    22,   # V 列：项目档案名称
        "receipt_original_amount": 23,   # W 列：原币收入
        "tax_method":              24,   # X 列：计税方式
        "tax_rate":                25,   # Y 列：税率
        "split_income_type":       26,   # Z 列：分灶收入类型
        "merge_flag":              27,   # AA列：是否并表
    },
}

EXPECTED_FIELD_HEADERS = {
    "company_code": "公司编码",
    "company_name": "公司名称",
    "bank_account": ("银行账号", "收款银行账号"),
    "bill_date": "单据日期",
    "merge_key_i": "结算日期",
    "expense_period": ("费用所属期间", "费用所属期"),
    "summary": "摘要",
    "bill_number": ("单据号", "单据编号"),
    "customer_code": "客商编码",
    "customer_name": "客商名称",
    "accounting_subject": "会计科目",
    "auxiliary_accounting": "辅助核算",
    "represented_party": (
        "被代付账套/被代付账套名称",
        "被代付账套/被代收账套名称",
        "被代收方（被收回方）",
    ),
    "represented_project": "被代收方（被收回方）项目档案",
    "business_category": "经营类别名称",
    "internal_collection_flag": "代收代付情况",
    "internal_collection_type": "代收代付类型",
    "receipt_method": "收款方式名称",
    "settlement_method": "结算方式名称",
    "actual_charge_type": (
        "收费类型名称",
        "代收（收回）业务实际业务收费类型",
    ),
    "new_income_expense_item": "NCC收支项目名称",
    "project_archive_name": "项目档案名称",
    "receipt_original_amount": "原币收入",
    "tax_method": "计税方式",
    "tax_rate": ("税率（非增值税收益等）", "税率"),
    "split_income_type": "分灶收入类型",
    "merge_flag": "是否并表",
}
EXPECTED_OUTPUT_HEADERS = {
    EXCEL_CONFIG["status_col"]: "填写状态",
    EXCEL_CONFIG["reason_col"]: "异常原因",
}

# ---------- 2. NCC 页面配置（★ 新节点 / 新地址在这里改 ★） ----------
NCC_CONFIG = {
    "login_url": "https://ncc.bgyfw.com/nccloud/resources/uap/rbac/login/main/index.html",
    "node_name": "收款结算",             # NCC 菜单搜索框输入的节点名，选第一条结果（已按截图确认）
    "chrome_path": "./ms-playwright（勿动）/chromium-1161/chrome-win/chrome.exe",
    "main_iframe": "#mainiframe",
}

# ---------- 3. 页面元素选择器（★ 输入框位置更新在这里改 ★） ----------
# 说明：以下为旧工具“收款单（内转）”的实测选择器，新节点若布局一致可直接用；
#       不一致的项已标 TODO，按“使用说明.md 第 4 节”的方法重新取选择器后替换。
SELECTORS = {
    # —— 登录与节点打开 ——
    "login_user":        "#username",
    "home_search_icon":  'xpath=//i[@class="iconfont icon-sousuo1"]',   # 首页加载完成标志
    # 右上角搜索入口（点击后弹出搜索框）；按截图，放大镜在页面右上角
    "search_open_icon":  'xpath=//i[@class="iconfont icon-sousuo1"]',   # TODO 若点击无反应，F12 取右上角放大镜图标的选择器替换
    "node_search_input": "xpath=//input[@class='ant-input ant-select-search__field']",  # TODO 若定位不到，改用 placeholder 模糊匹配（见 operate_page 备用逻辑）
    "node_search_first": "xpath=//ul[@class='ant-select-dropdown-menu  ant-select-dropdown-menu-root ant-select-dropdown-menu-vertical']/li[1]",  # TODO 新版下拉若为其他结构，见 operate_page 备用逻辑

    # —— 新增按钮 / 返回按钮 ——
    "add_btn":   "xpath=//div[@popcontainer='[object HTMLDivElement]']/button[@tabindex='-1']",  # TODO 新节点“新增”按钮
    "back_btn":  "span.back-btn.iconfont.icon.icon-danjufanhui",
    "back_btn_xpath": "xpath=//span[@class=' back-btn iconfont icon icon-danjufanhui']",

    # —— 卡片头字段 ——
    "save_only_btn":    "span.button-app-wrapper.card_head-button-app-wrapper button.u-button.null.nc-button-wrapper.button-weishabushezhiyigemoren",           # 仅保存
    "save_submit_btn":  "span.button-app-wrapper.card_head-button-app-wrapper button.u-button.button-primary.nc-button-wrapper.button-weishabushezhiyigemoren",   # 保存并提交
    "income_org_input": "div.pk_org_v.js-type-refer.form-item input.nc-input.refer-input.u-form-control.md",      # TODO 收方单位输入框
    "income_org_item":  "ul.refer-cascading-item > li[title*='{text}']",
    "income_org_match": "div.def4.js-type-refer.form-item div[title*='公司']",
    "pay_org_input":    "div.customer.js-type-refer.form-item input.nc-input.refer-input.u-form-control.md",      # TODO 付方单位输入框
    "pay_org_item":     "ul.refer-cascading-item > li[title*='{text}']",
    "pay_org_match":    "div.def21.js-type-refer.form-item div[title*='公司']",
    "biz_type_input":   "div.def13.js-type-refer.form-item input.nc-input.refer-input.u-form-control.md",          # TODO 业务类型输入框
    "biz_type_item":    "li[title='{text}']",
    "biz_type_value":   "input[value='{text}']",
    "head_remark":      "div.def1.js-type-input.form-item input.nc-input.nc-input.undefined.u-form-control.md",    # TODO 卡片头备注输入框

    # —— 内部账户参照弹窗 ——
    "refer_icon":        "div.u-col-xs-12.refer.clearfix.refer-input-common span.icon-refer",
    "income_acct_row":   "div.refer-table.selectable-recaccount div[title*='{text}']",
    "pay_acct_row":      "div.refer-table.selectable-payaccount div[title*='{text}']",
    "refer_sure_btn":    "div.refer-bottom.refer-resize-icon button.u-button.sure-btn.nc-button-wrapper.button-weishabushezhiyigemoren",

    # —— 表体（行编辑弹窗）——
    "body_table":        "div.selectable-bodys div.u-table-content div.u-table-body",
    "row_edit_btn":      "div.u-table-fixed-right div.u-table-body-outer a.row-edit-option.row-btn-open_edit",
    "modal_inputs":      "div.card-table-modal-form-wrap-body.cf div.card-table-modal-form-item.fl input.nc-input.refer-input.u-form-control.md",
    "modal_save_btn":    "span.button-app-wrapper.card_body-button-app-wrapper button.null.nc-button-wrapper.button-weishabushezhiyigemoren",
    "modal_title":       "h2.card-table-modal-form-wrap-header",
    # TODO 以下 nth-child 序号 = 行编辑弹窗中各字段的位置，新表单需重新核对
    "modal_money":       "div.card-table-modal-form-wrap-body.cf div.card-table-modal-form-item.fl input[label='收款原币金额']",
    "modal_remark_nth":  2,     # TODO 摘要输入框 nth-child 序号
    "modal_wl_type_nth": 6,     # TODO 收支类型 nth-child 序号
    "modal_s_file_nth":  9,     # TODO 被收回方项目档案 nth-child 序号
    "modal_file_nth":   11,     # TODO 项目档案 nth-child 序号
    "modal_ncc_nth":    14,     # TODO 收方NCC收支项目 nth-child 序号
    "modal_income_nth": 15,     # TODO 收方分灶收入类型 nth-child 序号
    "modal_operate_nth": 16,    # TODO 经营类别 nth-child 序号
    "modal_dept_nth":   13,     # TODO 部门 nth-child 序号
    "modal_project_nth": 26,    # TODO 项目档案（收回被代收）nth-child 序号
    "modal_cascade_item": 'body > div:last-of-type ul.refer-cascading-item li[title*="{text}"]',
    "modal_dropdown_item": 'div.u-dropdown.nc-dropdown.refer-container.u-dropdown-placement-bottomLeft ul.refer-cascading-item li[title*="{text}"]',
    "modal_row_by_title": "tr.u-table-row.row-key-nokey.u-table-row-level-0 div[title='{text}']",

    # —— 收款结算：列表页 ——
    "list_add_btn":      "xpath=//button[normalize-space(.)='新增']",

    # —— 收款结算：制单卡片（★ 第2~4步已按截图配置，取不到时按使用说明4节重新取 ★） ——
    "settle_org_input":  "xpath=(//div[contains(@class,'form-item') and contains(.,'结算财务组织') and not(contains(.,'结算财务组织编码'))]//input[not(contains(@class,'hidden-input')) and not(@readonly)])[1]",
    "settle_org_first":  "xpath=(//ul[contains(@class,'refer-cascading-item')]/li[not(contains(.,'没有更多了'))])[1]",
    "bill_date_input":   "xpath=(//input[contains(@label,'单据日期') and not(contains(@class,'hidden-input'))])[1]",
    "card_cancel_btn":   "xpath=(//button[contains(normalize-space(.),'取消')])[1]",   # 卡片页“取消”按钮（异常时退出卡片用）

    # —— 报错与等待 ——
    "error_toast":  "div[class=toast-content]",
    "error_group":  "ul[class=groupOperationMsg]",
    "error_modal":  "div.fl.u-modal-title",
    "bill_info":    "span.bill-info-title",
    "loading":      'xpath=//div[@class="u-loading u-loading-rotate nc-loading"]',
}

# ---------- 4. 运行参数 ----------
RETRY_MAX = 3                      # 单个输入框操作失败后的自动重试次数
SETTLE_ORG_CARD_RELOAD_MAX = 3     # 输入框不出现时取消并重新新增的次数
BILL_DATE_CARD_RELOAD_MAX = 3      # 日期框不出现时取消并重新新增的次数

# 所有可影响填报节奏的时间统一使用“秒”。常用项保持业务语义，高级项
# 归并同类 Playwright 操作，避免同一个 10 秒等待在界面中重复几十次。
# min/max 既用于 GUI 校验，也防止误填 0 或极大数值导致程序忙循环/长期卡住。
TIME_SETTING_SPECS = (
    # key, 显示名称, 分组, 默认值, 最小值, 最大值
    ('login_wait', '等待人工登录最长时间', '常用等待', 120.0, 1.0, 3600.0),
    ('reference_wait', '普通参照/下拉每轮最长时间', '常用等待', 30.0, 1.0, 600.0),
    ('settle_org_option_wait', '结算财务组织候选最长时间', '常用等待', 60.0, 1.0, 600.0),
    ('settle_org_input_wait', '结算财务组织输入框最长时间', '常用等待', 15.0, 1.0, 600.0),
    ('bill_date_input_wait', '单据日期输入框最长时间', '常用等待', 15.0, 1.0, 600.0),
    ('action_result_wait', '保存/提交结果最长时间', '常用等待', 60.0, 1.0, 1200.0),
    ('tax_rate_bind_wait', '税率选中后绑定等待', '常用等待', 2.5, 0.0, 60.0),

    ('auxiliary_modal_wait', '辅助核算弹窗出现最长时间', '辅助核算', 30.0, 1.0, 600.0),
    ('auxiliary_modal_open_wait', '弹窗打开后再输入', '辅助核算', 3.0, 0.0, 60.0),
    ('auxiliary_option_ready_wait', '候选出现后再点击', '辅助核算', 2.0, 0.0, 60.0),
    ('auxiliary_selection_settle_wait', '候选提交稳定最长时间', '辅助核算', 8.0, 0.0, 600.0),
    ('auxiliary_popup_dismiss_wait', '候选层残留后收起时间', '辅助核算', 3.0, 0.0, 60.0),
    ('auxiliary_selection_commit_wait', '点击候选后再确定', '辅助核算', 3.0, 0.0, 60.0),
    ('auxiliary_confirm_commit_wait', '点击确定后再继续', '辅助核算', 3.0, 0.0, 60.0),
    ('auxiliary_type_interval', '辅助核算逐字输入间隔', '辅助核算', 0.15, 0.0, 2.0),
    ('auxiliary_input_timeout', '辅助核算输入最长时间', '辅助核算', 15.0, 1.0, 600.0),
    ('auxiliary_input_verify_wait', '辅助核算输入值核验最长时间', '辅助核算', 3.0, 1.0, 600.0),
    ('auxiliary_input_stable_wait', '辅助核算输入值稳定时间', '辅助核算', 1.0, 0.0, 60.0),
    ('auxiliary_modal_close_stable_wait', '辅助核算弹窗关闭稳定时间', '辅助核算', 1.5, 0.0, 60.0),
    ('auxiliary_retry_interval', '辅助核算输入失败重试间隔', '辅助核算', 0.5, 0.0, 60.0),

    ('element_action_timeout', '普通控件操作最长时间', '页面与控件', 10.0, 1.0, 600.0),
    ('short_element_timeout', '短控件查找最长时间', '页面与控件', 5.0, 1.0, 600.0),
    ('tiny_element_timeout', '快速控件查找最长时间', '页面与控件', 3.0, 0.5, 600.0),
    ('long_element_timeout', '页面/侧栏/按钮最长时间', '页面与控件', 30.0, 1.0, 1200.0),
    ('page_load_timeout', '页面加载圈最长时间', '页面与控件', 60.0, 1.0, 1200.0),
    ('work_frame_wait', 'NCC业务页面识别最长时间', '页面与控件', 30.0, 1.0, 1200.0),
    ('bill_number_wait', '保存后读取单据编号最长时间', '页面与控件', 30.0, 1.0, 600.0),
    ('list_return_wait', '取消/保存后返回列表最长时间', '页面与控件', 30.0, 1.0, 600.0),
    ('notification_clear_wait', '保存前旧提示清理最长时间', '页面与控件', 5.0, 0.0, 60.0),

    ('menu_open_wait', '菜单搜索框打开等待', '动作稳定间隔', 1.0, 0.0, 60.0),
    ('menu_result_render_wait', '菜单搜索结果渲染等待', '动作稳定间隔', 3.0, 0.0, 60.0),
    ('page_transition_wait', '页面切换后稳定等待', '动作稳定间隔', 1.0, 0.0, 60.0),
    ('reference_option_stable_wait', '参照候选列表稳定时间', '动作稳定间隔', 0.5, 0.0, 60.0),
    ('reference_post_select_wait', '普通参照选中后等待', '动作稳定间隔', 0.8, 0.0, 60.0),
    ('reference_retry_interval', '普通参照失败重试间隔', '动作稳定间隔', 1.0, 0.0, 60.0),
    ('reference_type_interval', '普通参照逐字输入间隔', '动作稳定间隔', 0.06, 0.0, 2.0),
    ('reference_input_timeout', '普通参照输入最长时间', '动作稳定间隔', 15.0, 1.0, 600.0),
    ('settle_org_commit_wait', '结算财务组织选中后等待', '动作稳定间隔', 1.5, 0.0, 60.0),
    ('plain_input_commit_wait', '普通输入框失焦后等待', '动作稳定间隔', 0.8, 0.0, 60.0),
    ('checkbox_commit_wait', '复选框状态确认后等待', '动作稳定间隔', 1.5, 0.0, 60.0),
    ('internal_collection_reset_wait', '代收付联动字段清理前等待', '动作稳定间隔', 1.0, 0.0, 60.0),
    ('field_poll_interval', '字段/页面状态轮询间隔', '动作稳定间隔', 0.2, 0.02, 5.0),
    ('candidate_poll_interval', '候选/结果状态轮询间隔', '动作稳定间隔', 0.25, 0.02, 5.0),
    ('short_stable_wait', '清空/收起控件后等待', '动作稳定间隔', 0.5, 0.0, 60.0),

    ('date_calendar_open_wait', '日期面板打开最长时间', '日期与明细', 5.0, 1.0, 600.0),
    ('date_field_stable_wait', '日期输入框稳定时间', '日期与明细', 0.8, 0.0, 60.0),
    ('date_year_step_wait', '日期逐年切换间隔', '日期与明细', 0.05, 0.0, 5.0),
    ('date_month_switch_wait', '日期月份切换等待', '日期与明细', 0.3, 0.0, 60.0),
    ('date_input_commit_wait', '日期文本输入后等待', '日期与明细', 0.5, 0.0, 60.0),
    ('date_selection_commit_wait', '日期选中后等待', '日期与明细', 1.5, 0.0, 60.0),
    ('date_confirm_commit_wait', '日期确定后等待', '日期与明细', 1.0, 0.0, 60.0),
    ('detail_row_detect_wait', '明细增行识别最长时间', '日期与明细', 30.0, 1.0, 1200.0),
    ('detail_virtual_fallback_wait', '虚拟表格增行兜底稳定时间', '日期与明细', 5.0, 0.0, 60.0),
    ('detail_signature_stable_wait', '明细行结构稳定时间', '日期与明细', 0.6, 0.0, 60.0),

    ('cancel_confirm_wait', '取消后确认框最长时间', '保存与回写', 10.0, 1.0, 600.0),
    ('save_error_close_wait', '保存报错关闭按钮最长时间', '保存与回写', 10.0, 1.0, 600.0),
    ('workbook_unlock_wait', 'Excel占用释放最长时间', '保存与回写', 10.0, 0.0, 600.0),
    ('workbook_retry_interval', 'Excel占用重试间隔', '保存与回写', 1.0, 0.05, 60.0),
    ('browser_close_poll_interval', '批次完成后浏览器关闭检测间隔', '保存与回写', 1.0, 0.05, 60.0),

    ('pause_stop_poll_interval', '暂停/停止状态检测间隔', '系统轮询', 0.1, 0.02, 5.0),
    ('gui_poll_interval', 'GUI消息及任务结束检测间隔', '系统轮询', 0.2, 0.05, 5.0),
)

TIME_SETTINGS_VERSION = 1
TIME_SETTINGS_FILENAME = '收款结算单时间设置.json'


def default_time_settings():
    """返回一份可独立修改的默认时间设置。"""
    return {spec[0]: float(spec[3]) for spec in TIME_SETTING_SPECS}


def validate_time_settings(values):
    """补齐缺项并严格校验秒数；未知旧字段会被安全忽略。"""
    if values is None:
        values = {}
    if not isinstance(values, dict):
        raise ValueError('时间设置必须是键值对象')
    validated = {}
    for key, label, _group, default, minimum, maximum in TIME_SETTING_SPECS:
        raw = values.get(key, default)
        if isinstance(raw, bool):
            raise ValueError(f'“{label}”必须填写数字')
        try:
            number = float(str(raw).strip())
        except (TypeError, ValueError) as error:
            raise ValueError(f'“{label}”必须填写数字') from error
        if not math.isfinite(number):
            raise ValueError(f'“{label}”必须是有限数字')
        if number < minimum or number > maximum:
            raise ValueError(
                f'“{label}”应在 {minimum:g}～{maximum:g} 秒之间'
            )
        validated[key] = number
    if (
        validated['auxiliary_input_stable_wait']
        > validated['auxiliary_input_verify_wait']
    ):
        raise ValueError(
            '“辅助核算输入值稳定时间”不能大于“辅助核算输入值核验最长时间”'
        )
    if (
        validated['auxiliary_modal_close_stable_wait']
        > validated['auxiliary_modal_wait']
    ):
        raise ValueError(
            '“辅助核算弹窗关闭稳定时间”不能大于“辅助核算弹窗出现最长时间”'
        )
    return validated


def application_directory():
    """时间配置始终写在源码/EXE同目录，不写入 PyInstaller 临时目录。"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def timing_config_path():
    return os.path.join(application_directory(), TIME_SETTINGS_FILENAME)


def load_time_settings(path=None):
    """读取最后一次保存的设置；损坏或版本不兼容时回退默认值。"""
    config_path = path or timing_config_path()
    if not os.path.isfile(config_path):
        return default_time_settings()
    try:
        with open(config_path, 'r', encoding='utf-8') as config_file:
            payload = json.load(config_file)
        if not isinstance(payload, dict):
            raise ValueError('配置根节点不是对象')
        if payload.get('version') != TIME_SETTINGS_VERSION:
            raise ValueError('配置版本不兼容')
        return validate_time_settings(payload.get('values'))
    except Exception as error:
        print(f'时间设置读取失败，已使用默认值：{error}')
        return default_time_settings()


def save_time_settings(values, path=None):
    """原子覆盖最新时间设置；不会累积历史配置。"""
    validated = validate_time_settings(values)
    config_path = path or timing_config_path()
    config_dir = os.path.dirname(os.path.abspath(config_path))
    os.makedirs(config_dir, exist_ok=True)
    file_descriptor, temporary_path = tempfile.mkstemp(
        prefix='.receipt_time_settings_', suffix='.tmp', dir=config_dir
    )
    try:
        with os.fdopen(file_descriptor, 'w', encoding='utf-8') as config_file:
            json.dump(
                {'version': TIME_SETTINGS_VERSION, 'values': validated},
                config_file, ensure_ascii=False, indent=2
            )
            config_file.flush()
            os.fsync(config_file.fileno())
        os.replace(temporary_path, config_path)
    finally:
        if os.path.exists(temporary_path):
            os.unlink(temporary_path)
    return validated


ENABLE_TIME_SETTINGS = True  # 正常版构建时固定为 False，不读取用户调时配置。
current_time_settings = load_time_settings() if ENABLE_TIME_SETTINGS else default_time_settings()


def get_time_settings_snapshot():
    """开始任务时复制一份，保证本批次运行过程中参数不变化。"""
    return dict(current_time_settings)


def timing_seconds(key):
    """优先读取本次运行快照，GUI空闲时读取当前已保存设置。"""
    context_local = globals().get('run_context_local')
    run_context = (
        getattr(context_local, 'value', None)
        if context_local is not None else None
    )
    settings = (
        run_context.get('timing_settings')
        if isinstance(run_context, dict) else None
    )
    if not isinstance(settings, dict):
        settings = current_time_settings
    if key not in settings:
        settings = default_time_settings()
    return float(settings[key])


def timing_ms(key):
    """把 GUI 中的秒转换为 Playwright 使用的整数毫秒。"""
    return max(1, int(round(timing_seconds(key) * 1000)))


def format_timing_seconds(key):
    """日志和错误提示使用实际设置值，整数不显示小数点。"""
    value = timing_seconds(key)
    return f'{value:g}'


# ---------- 4.1 填写顺序 / 停用设置（仅表头内部、明细内部各自互换） ----------
# “结算财务组织”“单据日期”承担卡片加载与单公司权限切换，固定最前；
# “付款银行账号”依赖明细数据，固定在所有数据（含全部明细）填妥后、点击保存前最后填入；
# 表头与表头、明细与明细各自独立调整顺序，严格禁止表头与明细之间互换顺序。
HEADER_FILL_DEFAULT_ORDER = (
    "settlement_method",
    "transaction_object",
    "receive_bank_account",
    "remark",
)
HEADER_FILL_STEP_LABELS = {
    "settlement_method": "结算方式",
    "transaction_object": "交易对象类型",
    "receive_bank_account": "收款银行账户",
    "remark": "备注",
}
HEADER_FILL_OPTIONAL_STEPS = (
    "remark",
)
HEADER_FILL_ORDER_DEPENDENCIES = {}

DETAIL_FILL_DEFAULT_ORDER = (
    "summary",
    "internal_collection",
    "customer",
    "receipt_method",
    "expense_period",
    "tax_method",
    "accounting_subject",
    "project_archive",
    "receipt_original_amount",
    "internal_collection_references",
    "new_income_expense_item",
    "department",
    "split_income_type",
    "business_category",
    "auxiliary_accounting",
    "tax_rate",
)
DETAIL_FILL_STEP_LABELS = {
    "summary": "摘要",
    "internal_collection": "代收代付情况",
    "customer": "客户",
    "receipt_method": "收款方式",
    "expense_period": "费用所属期",
    "tax_method": "计税方式",
    "accounting_subject": "会计科目",
    "project_archive": "项目档案",
    "receipt_original_amount": "收款原币金额",
    "internal_collection_references": "代收代付类型及被代收信息",
    "new_income_expense_item": "新收支项目",
    "department": "部门",
    "split_income_type": "分灶收入类型",
    "business_category": "经营类别",
    "auxiliary_accounting": "辅助核算",
    "tax_rate": "税率",
}
DETAIL_FILL_OPTIONAL_STEPS = (
    "project_archive",
    "department",
    "auxiliary_accounting",
)
DETAIL_FILL_ORDER_DEPENDENCIES = {
    "project_archive": ("internal_collection",),
    "internal_collection_references": ("internal_collection",),
    "split_income_type": ("new_income_expense_item",),
    "business_category": ("new_income_expense_item",),
}

POST_DETAIL_FILL_STEPS = (
    "payer_bank_account",
)
POST_DETAIL_FILL_STEP_LABELS = {
    "payer_bank_account": "付款银行账号（保存前最后填入）",
}
POST_DETAIL_FILL_OPTIONAL_STEPS = (
    "payer_bank_account",
)

ALL_FILL_STEP_LABELS = {
    **HEADER_FILL_STEP_LABELS,
    **DETAIL_FILL_STEP_LABELS,
    **POST_DETAIL_FILL_STEP_LABELS,
}
ALL_FILL_OPTIONAL_STEPS = (
    HEADER_FILL_OPTIONAL_STEPS
    + DETAIL_FILL_OPTIONAL_STEPS
    + POST_DETAIL_FILL_OPTIONAL_STEPS
)

FILL_ORDER_SETTINGS_VERSION = 1
FILL_ORDER_SETTINGS_FILENAME = '收款结算单填写顺序设置.json'


def fill_order_settings_path():
    """填写顺序设置始终保存在源码/EXE同目录，跨次运行复用。"""
    return os.path.join(application_directory(), FILL_ORDER_SETTINGS_FILENAME)


def default_fill_order_settings():
    """返回一份可独立修改的默认填写顺序与停用设置。"""
    return {
        "header_order": list(HEADER_FILL_DEFAULT_ORDER),
        "detail_order": list(DETAIL_FILL_DEFAULT_ORDER),
        "disabled": [],
    }


def normalize_fill_order_settings(raw_settings):
    """校验并标准化填写顺序设置：严格禁止表头与明细互换及破坏字段联动依赖。"""
    if not isinstance(raw_settings, dict):
        raise ValueError("填写顺序设置不是有效对象")
    if (
        "version" in raw_settings
        and raw_settings.get("version") != FILL_ORDER_SETTINGS_VERSION
    ):
        raise ValueError("填写顺序设置版本不兼容")

    header_order = raw_settings.get("header_order")
    detail_order = raw_settings.get("detail_order")
    disabled = raw_settings.get("disabled", [])

    if not isinstance(header_order, list) or not all(
        isinstance(item, str) for item in header_order
    ):
        raise ValueError("表头填写顺序必须是字符串列表")
    if not isinstance(detail_order, list) or not all(
        isinstance(item, str) for item in detail_order
    ):
        raise ValueError("明细填写顺序必须是字符串列表")
    if not isinstance(disabled, list) or not all(
        isinstance(item, str) for item in disabled
    ):
        raise ValueError("停用字段必须是字符串列表")

    header_set = set(HEADER_FILL_DEFAULT_ORDER)
    detail_set = set(DETAIL_FILL_DEFAULT_ORDER)
    post_detail_set = set(POST_DETAIL_FILL_STEPS)

    if any(item in post_detail_set for item in header_order) or any(
        item in post_detail_set for item in detail_order
    ):
        raise ValueError(
            "“付款银行账号”固定在全部数据（含所有明细）填妥后、点击保存前最后填入，"
            "不允许放入表头或明细调序列表"
        )

    header_with_detail = [item for item in header_order if item in detail_set]
    if header_with_detail:
        raise ValueError(
            "不允许表头与明细之间互换顺序：表头顺序中包含了明细字段："
            + "、".join(
                DETAIL_FILL_STEP_LABELS.get(key, key)
                for key in dict.fromkeys(header_with_detail)
            )
        )
    detail_with_header = [item for item in detail_order if item in header_set]
    if detail_with_header:
        raise ValueError(
            "不允许表头与明细之间互换顺序：明细顺序中包含了表头字段："
            + "、".join(
                HEADER_FILL_STEP_LABELS.get(key, key)
                for key in dict.fromkeys(detail_with_header)
            )
        )

    if len(header_order) != len(set(header_order)):
        raise ValueError("表头填写顺序存在重复字段")
    if set(header_order) != header_set:
        raise ValueError("表头填写顺序缺少字段或包含未知字段")

    if len(detail_order) != len(set(detail_order)):
        raise ValueError("明细填写顺序存在重复字段")
    if set(detail_order) != detail_set:
        raise ValueError("明细填写顺序缺少字段或包含未知字段")

    disabled_set = set(disabled)
    if len(disabled_set) != len(disabled):
        raise ValueError("停用字段存在重复项")
    unknown_disabled = disabled_set.difference(ALL_FILL_STEP_LABELS)
    if unknown_disabled:
        raise ValueError(
            "停用字段包含未知步骤：" + "、".join(sorted(unknown_disabled))
        )
    invalid_disabled = disabled_set.difference(ALL_FILL_OPTIONAL_STEPS)
    if invalid_disabled:
        raise ValueError(
            "必填或联动字段不允许停用："
            + "、".join(
                ALL_FILL_STEP_LABELS.get(key, key)
                for key in sorted(invalid_disabled)
            )
        )

    header_positions = {key: index for index, key in enumerate(header_order)}
    for step_key, prerequisites in HEADER_FILL_ORDER_DEPENDENCIES.items():
        for prerequisite in prerequisites:
            if header_positions[prerequisite] > header_positions[step_key]:
                raise ValueError(
                    f"表头中“{HEADER_FILL_STEP_LABELS[prerequisite]}”必须排在"
                    f"“{HEADER_FILL_STEP_LABELS[step_key]}”之前"
                )

    detail_positions = {key: index for index, key in enumerate(detail_order)}
    for step_key, prerequisites in DETAIL_FILL_ORDER_DEPENDENCIES.items():
        for prerequisite in prerequisites:
            if detail_positions[prerequisite] > detail_positions[step_key]:
                raise ValueError(
                    f"明细中“{DETAIL_FILL_STEP_LABELS[prerequisite]}”必须排在"
                    f"“{DETAIL_FILL_STEP_LABELS[step_key]}”之前"
                )

    return {
        "header_order": list(header_order),
        "detail_order": list(detail_order),
        "disabled": sorted(disabled_set),
    }


def load_fill_order_settings(path=None):
    """读取上次保存的填写顺序；读不到或损坏时回退默认顺序。"""
    config_path = path or fill_order_settings_path()
    if not os.path.isfile(config_path):
        return default_fill_order_settings()
    try:
        with open(config_path, 'r', encoding='utf-8') as config_file:
            payload = json.load(config_file)
        return normalize_fill_order_settings(payload)
    except Exception as error:
        print(f'填写顺序设置无效，已恢复默认设置：{error}')
        return default_fill_order_settings()


def save_fill_order_settings(settings, path=None):
    """原子保存填写顺序设置，避免异常退出导致 JSON 损坏。"""
    normalized = normalize_fill_order_settings(settings)
    config_path = path or fill_order_settings_path()
    config_dir = os.path.dirname(os.path.abspath(config_path))
    os.makedirs(config_dir, exist_ok=True)
    file_descriptor, temporary_path = tempfile.mkstemp(
        prefix='.receipt_fill_order_', suffix='.tmp', dir=config_dir
    )
    try:
        with os.fdopen(file_descriptor, 'w', encoding='utf-8') as config_file:
            json.dump(
                {'version': FILL_ORDER_SETTINGS_VERSION, **normalized},
                config_file, ensure_ascii=False, indent=2
            )
            config_file.flush()
            os.fsync(config_file.fileno())
        os.replace(temporary_path, config_path)
    finally:
        if os.path.exists(temporary_path):
            os.unlink(temporary_path)
    return normalized


current_fill_order_settings = load_fill_order_settings()


def get_fill_order_settings_snapshot():
    """开始任务时复制一份，保证本批次运行过程中顺序和停用状态不变。"""
    return {
        "header_order": list(current_fill_order_settings["header_order"]),
        "detail_order": list(current_fill_order_settings["detail_order"]),
        "disabled": list(current_fill_order_settings["disabled"]),
    }


def active_fill_order_settings():
    """优先读取本次运行快照，GUI空闲或单元测试时读取当前/默认设置。"""
    context_local = globals().get('run_context_local')
    run_context = (
        getattr(context_local, 'value', None)
        if context_local is not None else None
    )
    settings = (
        run_context.get('fill_order_settings')
        if isinstance(run_context, dict) else None
    )
    if isinstance(settings, dict):
        return settings
    current = globals().get('current_fill_order_settings')
    if isinstance(current, dict):
        return current
    return default_fill_order_settings()


def enabled_header_fill_order():
    """获取当前设置下实际执行的表头填写步骤（不含固定最前项和保存前付款银行账号）。"""
    settings = active_fill_order_settings()
    disabled = set(settings.get("disabled", ()))
    return [
        step_key
        for step_key in settings.get("header_order", HEADER_FILL_DEFAULT_ORDER)
        if step_key not in disabled
    ]


def enabled_detail_fill_order():
    """获取当前设置下实际执行的明细填写步骤，供每一条明细共用。"""
    settings = active_fill_order_settings()
    disabled = set(settings.get("disabled", ()))
    return [
        step_key
        for step_key in settings.get("detail_order", DETAIL_FILL_DEFAULT_ORDER)
        if step_key not in disabled
    ]


def is_fill_step_enabled(step_key):
    """判断指定填写步骤当前是否处于启用状态。"""
    settings = active_fill_order_settings()
    return step_key not in set(settings.get("disabled", ()))


RUN_TIME_SHEET_NAME = '运行时间'
RUN_TIME_HEADERS = ('运行次数', '运行数据', '运行时间', '运行结果')

current_thread = None
pause_event = threading.Event()
stop_event = threading.Event()
pause_clock_lock = threading.Lock()
run_context_local = threading.local()
pause_started_at = None
total_paused_seconds = 0.0
ui_message_queue = queue.Queue()
active_browser = None
active_context = None
active_cdp_sessions = []
last_cdp_activity_refresh_at = 0.0
automation_priority_raised = False
closing_requested = False

STATUS_SUCCESS = '填报成功'
STATUS_ABNORMAL = '异常'
LEGACY_SUCCESS_STATUS = '制单成功'
LEGACY_REVIEW_STATUS = 'NCC结果待核对'
LEGACY_PROCESSING_STATUSES = ('保存处理中', '保存提交处理中')
# ╚══════════════════════════ 配置区结束 ══════════════════════════╝


class FieldFillError(Exception):
    """携带具体字段名的填写异常，供底稿回写精确标注。"""

    def __init__(self, field_name, reason):
        self.field_name = field_name
        self.reason = str(reason).strip() or '未知错误'
        super().__init__(f'{field_name}：{self.reason}')


class ReferenceMatchAmbiguousError(FieldFillError):
    """参照下拉存在多个匹配项；供客商名称模式立即终止重试。"""


class SettleOrgInputUnavailableError(FieldFillError):
    """新增卡片已经打开，但结算财务组织输入框没有按时渲染。"""


class BillDateInputUnavailableError(FieldFillError):
    """新增卡片已经打开，但单据日期输入框没有按时渲染。"""


class TaskCancelledError(BaseException):
    """用户主动停止任务；继承 BaseException，避免被字段重试误吞。"""


class NccCommitUncertainError(FieldFillError):
    """保存/提交已经点击，但无法确认 NCC 最终结果。"""


class NccSaveRejectedError(NccCommitUncertainError):
    """NCC 已明确返回保存/提交失败；可关闭提示并跳过当前单据。"""


class ResultWorkbookRedirectedError(FieldFillError):
    """防重复标记只能写入备用文件，禁止继续点击 NCC 保存。"""


def count_elements(nested_dict):
    """统计底稿中的总笔数（合并类按子行计数）"""
    total = 0
    for value in nested_dict.values():
        if isinstance(value, dict):
            total += len(value)
        else:
            total += 1
    return total


def clear_before_fill(locator_input, field_name='输入框', force=False):
    """统一填写前置：聚焦、全选并清空；只读日期框可强制解锁。"""
    pause()
    locator_input.wait_for(
        state='visible', timeout=timing_ms('element_action_timeout')
    )
    locator_input.click(
        force=force, timeout=timing_ms('element_action_timeout')
    )
    if force:
        locator_input.evaluate(
            "el => { el.removeAttribute('readonly'); el.readOnly = false; "
            "el.removeAttribute('disabled'); el.disabled = false; el.focus(); }"
        )

    try:
        locator_input.press('Control+a')
        locator_input.press('Backspace')
        locator_input.press('Delete')
        locator_input.fill(
            '', force=force, timeout=timing_ms('short_element_timeout')
        )
    except Exception:
        pass

    # NCC 受控组件可能立即恢复旧值；最后只清 DOM，不触发回滚事件，
    # 紧接着由 Playwright 写入完整的新值。
    actual = locator_input.input_value(
        timeout=timing_ms('short_element_timeout')
    )
    if actual:
        locator_input.evaluate(
            """el => {
                el.removeAttribute('readonly');
                el.readOnly = false;
                const proto = el.tagName === 'TEXTAREA'
                    ? window.HTMLTextAreaElement.prototype
                    : window.HTMLInputElement.prototype;
                const setter = Object.getOwnPropertyDescriptor(
                    proto, 'value'
                ).set;
                setter.call(el, '');
            }"""
        )
        actual = locator_input.input_value(
            timeout=timing_ms('short_element_timeout')
        )

    if actual:
        raise Exception(f'填写“{field_name}”前清空失败，当前值={actual}')
    print(f'     [前置] 已清空{field_name}')


def try_except_click(iframe, locator_input, text, locator_choose):
    """输入文本并点选下拉参照项，失败自动重试"""
    retry_count = 0
    while retry_count < RETRY_MAX:
        try:
            clear_before_fill(locator_input)
            locator_input.fill(text)
            locator_choose.wait_for(
                state='visible', timeout=timing_ms('reference_wait')
            )
            locator_choose.click()
            return
        except Exception:
            print(f'     操作失败，正在重新调试(第{retry_count + 1}次)......')
            try:
                iframe.locator(SELECTORS["bill_info"]).click()
                current = locator_input.nth(0)
                if current.get_attribute('value'):
                    return
            except Exception:
                pass
            retry_count += 1


def expand_try_except_click(iframe, locator_input, text, locator_choose, n):
    """表体第 n 行：输入文本并点选下拉参照项，失败自动重试"""
    retry_count = 0
    while retry_count < RETRY_MAX:
        try:
            clear_before_fill(locator_input)
            locator_input.fill(text)
            locator_choose.wait_for(
                state='visible', timeout=timing_ms('reference_wait')
            )
            locator_choose.click()
            return
        except Exception:
            print(f'     操作失败，正在重新调试(第{retry_count + 1}次)......')
            try:
                iframe.locator(SELECTORS["error_modal"]).click()
                current = locator_input.nth(n - 1)
                if current.get_attribute('value'):
                    return
            except Exception:
                pass
            retry_count += 1


def expand_try_except(iframe, locator_input, text, locator_wait, n):
    """表体第 n 行：输入文本并等待值生效，失败自动重试"""
    retry_count = 0
    while retry_count < RETRY_MAX:
        try:
            clear_before_fill(locator_input)
            locator_input.fill(text)
            locator_wait.wait_for(
                state='visible', timeout=timing_ms('reference_wait')
            )
            return
        except Exception:
            print(f'     操作失败，正在重新调试(第{retry_count + 1}次)......')
            try:
                iframe.locator(SELECTORS["error_modal"]).click()
                current = locator_input.nth(n - 1)
                if current.get_attribute('value'):
                    return
            except Exception:
                pass
            retry_count += 1


def task_now():
    """返回剔除暂停时长后的单调时钟，供业务超时与可暂停等待使用。"""
    now = time.monotonic()
    with pause_clock_lock:
        paused_seconds = total_paused_seconds
        if pause_started_at is not None:
            paused_seconds += now - pause_started_at
    return now - paused_seconds


def reset_pause_clock():
    """开始新任务时重置暂停时钟。"""
    global pause_started_at, total_paused_seconds
    with pause_clock_lock:
        pause_event.clear()
        stop_event.clear()
        pause_started_at = None
        total_paused_seconds = 0.0


def set_task_paused(paused):
    """原子地切换暂停状态，并累计本次任务的暂停时长。"""
    global pause_started_at, total_paused_seconds
    now = time.monotonic()
    with pause_clock_lock:
        if paused:
            if not pause_event.is_set():
                pause_started_at = now
                pause_event.set()
            return
        if pause_event.is_set() and pause_started_at is not None:
            total_paused_seconds += now - pause_started_at
        pause_started_at = None
        pause_event.clear()


def pause():
    """响应 GUI 暂停按钮；业务超时不会在暂停期间继续倒计时。"""
    # Chromium 在 Windows 中即使带有启动参数，窗口最小化一段时间后仍可能
    # 再次把页面标记为后台。这里在自动化工作线程内定期重申活动状态，避免
    # 跨线程调用 Playwright。
    _refresh_page_activity_when_minimized()
    if stop_event.is_set():
        raise TaskCancelledError('用户已停止任务')
    while pause_event.is_set():
        if stop_event.wait(timing_seconds('pause_stop_poll_interval')):
            raise TaskCancelledError('用户已停止任务')
    if stop_event.is_set():
        raise TaskCancelledError('用户已停止任务')


def task_sleep(seconds):
    """可响应暂停的短间隔等待。"""
    deadline = task_now() + max(0, seconds)
    while task_now() < deadline:
        pause()
        time.sleep(min(
            timing_seconds('field_poll_interval'),
            max(0, deadline - task_now())
        ))


def wait_end(iframe, timeout=None):
    """等待 NCC 页面加载圈消失；超时返回 False，绝不当作成功证据。"""
    if timeout is None:
        timeout = timing_ms('page_load_timeout')
    try:
        iframe.locator(SELECTORS["loading"]).wait_for(
            state='hidden', timeout=timeout
        )
        return True
    except Exception:
        return False


def keep_page_active_when_minimized(context, page):
    """让 Chromium 在窗口最小化时仍把 NCC 页面视为已聚焦、活动状态。"""
    global active_cdp_sessions
    try:
        session = context.new_cdp_session(page)
        active_cdp_sessions.append(session)
        _set_cdp_page_active(session)
        print('     已启用最小化页面活动状态保护')
    except Exception as error:
        print(f'     最小化页面活动状态保护未启用，将继续运行：{error}')


def _set_cdp_page_active(session):
    """在当前 Playwright 工作线程中强制维持页面为前台活动状态。"""
    session.send('Emulation.setFocusEmulationEnabled', {'enabled': True})
    try:
        session.send('Page.setWebLifecycleState', {'state': 'active'})
    except Exception:
        # 部分内置 Chromium 没有该实验协议，但焦点模拟仍有效。
        pass


def _refresh_page_activity_when_minimized():
    """最小化期间定时重申 CDP 活动状态，防止 Chromium 再次后台化。"""
    global last_cdp_activity_refresh_at, active_cdp_sessions
    now = time.monotonic()
    if now - last_cdp_activity_refresh_at < 2.0:
        return
    last_cdp_activity_refresh_at = now
    alive_sessions = []
    for session in active_cdp_sessions:
        try:
            _set_cdp_page_active(session)
            alive_sessions.append(session)
        except Exception:
            # 页面关闭或 CDP 已断开时自然丢弃该会话。
            continue
    active_cdp_sessions = alive_sessions


def raise_automation_process_priority():
    """将自动化进程提高到“高于普通”优先级，浏览器子进程会继承该优先级。"""
    global automation_priority_raised
    if automation_priority_raised or os.name != 'nt':
        return
    try:
        import ctypes
        kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
        # ABOVE_NORMAL 比 HIGH 更适合办公电脑：提高 NCC 渲染优先级，同时不让
        # 浏览器在页面卡顿时挤占系统输入、Excel 等前台程序。
        above_normal_priority_class = 0x00008000
        if kernel32.SetPriorityClass(
            kernel32.GetCurrentProcess(), above_normal_priority_class
        ):
            automation_priority_raised = True
            print('     已提高自动化与浏览器子进程优先级（高于普通）')
        else:
            print('     未能提高自动化进程优先级，将按普通优先级继续运行')
    except Exception as error:
        print(f'     提高自动化进程优先级失败，将按普通优先级继续运行：{error}')


def safe_get_text(page, selector, timeout=None):
    """安全获取元素文本，不存在时返回None"""
    if timeout is None:
        timeout = timing_ms('tiny_element_timeout')
    try:
        return page.locator(selector).inner_text(timeout=timeout)
    except Exception:
        return None


def wait_for_come(xpath):
    """等待元素出现（用于登录后主界面加载）"""
    retry_count = 0
    last_error = None
    while retry_count < RETRY_MAX:
        try:
            xpath.wait_for(
                state='visible', timeout=timing_ms('login_wait')
            )
            return
        except Exception as error:
            last_error = error
            retry_count += 1
    raise last_error or Exception('等待页面元素超时')


def _normalize_template_header(value):
    """比较表头时忽略末尾必填标记，保留其余文字的严格校验。"""
    text = str(value or '').strip()
    return re.sub(r'[\s*＊]+$', '', text).strip()


def _normalize_existing_fill_status(value):
    """把程序旧版状态迁移为当前允许的“异常”或“填报成功”。"""
    text = str(value or '').strip()
    if text == LEGACY_SUCCESS_STATUS:
        return STATUS_SUCCESS
    if text == LEGACY_REVIEW_STATUS or text in LEGACY_PROCESSING_STATUSES:
        return STATUS_ABNORMAL
    return text


def _simplify_legacy_review_reason(value):
    """从旧版冗长待核对原因中保留用户需要的保存错误正文。"""
    text = str(value or '').strip()
    marker = '保存结果需人工核对：'
    marker_index = text.find(marker)
    if marker_index < 0:
        return text
    text = text[marker_index:]
    stop_marker = '。程序已停止本批次'
    if stop_marker in text:
        text = text.split(stop_marker, 1)[0]
    return text.strip()


def _clean_input_text(value):
    """
    清除 Excel 文本首尾的可见/隐藏空白。

    除普通空格外，还处理从网页、微信或其他系统复制时常见的全角空格、
    不换行空格、零宽空格和 BOM。只清理首尾，字段正文中的空格保持不变。
    """
    if value is None:
        return ''
    edge_whitespace = (
        ' \t\r\n\v\f\u00a0\u1680\u180e\u2000\u2001\u2002\u2003'
        '\u2004\u2005\u2006\u2007\u2008\u2009\u200a\u200b\u200c'
        '\u200d\u2028\u2029\u202f\u205f\u2060\u3000\ufeff'
    )
    return str(value).strip(edge_whitespace)


def get_data(file_path, sheet_name):
    """
    读取底稿 Excel（新表：按 EXCEL_CONFIG 里的【绝对列号】取值）
    返回 [{'_row': Excel行号, 字段: 值, ...}, ...]：
    - 自动跳过所有字段都为空的行；
    - 仅保留 “填写状态”为空或“异常”的行，避免重复填写已完成数据。
    """
    # openpyxl 不会计算公式。为避免读取陈旧缓存，待处理输入列一旦含公式就明确拒绝。
    from receipt_excel_converter import identifier_source_issue, IDENTIFIER_HEADERS, normalize_header
    if os.path.splitext(file_path)[1].lower() == '.xlsm':
        read_receipt_macro(file_path)
    wb = load_receipt_workbook(file_path, data_only=False)
    ws = wb[sheet_name]
    config = receipt_sheet_layout(ws, EXCEL_CONFIG)
    F = config['fields']
    header_row = EXCEL_CONFIG["header_row"]
    header_errors = []
    field_headers = dict(EXPECTED_FIELD_HEADERS)
    if 'department_code' in F:
        field_headers['department_code'] = '部门编码'
    if 'department_name' in F:
        field_headers['department_name'] = ('部门', '部门名称')
    if 'payer_bank_account' in F:
        field_headers['payer_bank_account'] = (
            '付款银行账号', '付款银行账户', '付款账号'
        )
    if 'remark' in F:
        field_headers['remark'] = '备注'
    for name, expected_header in field_headers.items():
        col = F[name]
        actual_header_raw = ws.cell(row=header_row, column=col).value
        actual_header_raw = str(actual_header_raw or '').strip()
        actual_header = _normalize_template_header(actual_header_raw)
        allowed_headers = (
            expected_header
            if isinstance(expected_header, (tuple, list))
            else (expected_header,)
        )
        normalized_allowed_headers = tuple(
            _normalize_template_header(header) for header in allowed_headers
        )
        if actual_header not in normalized_allowed_headers:
            col_name = get_column_letter(col)
            expected_text = ' / '.join(allowed_headers)
            header_errors.append(
                f'{col_name}列应为“{expected_text}”，实际为“{actual_header_raw}”'
            )
    for col, expected_header in {config['status_col']: '填写状态', config['reason_col']: '异常原因'}.items():
        actual_header_raw = ws.cell(row=header_row, column=col).value
        actual_header_raw = str(actual_header_raw or '').strip()
        actual_header = _normalize_template_header(actual_header_raw)
        expected_header_normalized = _normalize_template_header(expected_header)
        if actual_header != expected_header_normalized:
            col_name = get_column_letter(col)
            header_errors.append(
                f'{col_name}列应为“{expected_header}”，实际为“{actual_header_raw}”'
            )
    if header_errors:
        wb.close()
        raise Exception(
            '底稿列结构与新版模板不一致：' + '；'.join(header_errors)
        )

    # 再打开一份 data_only 工作簿，仅用于判断公式缓存后的显示值。
    # 这样模板空白预留行中的“返回空字符串公式”不会误拦整个批次；
    # 一旦同一行存在真实业务数据，该行任何输入公式仍会被明确拒绝。
    value_wb = load_workbook(file_path, data_only=True)
    value_ws = value_wb[sheet_name]
    rows = []
    skipped_status_rows = []
    normalized_status_rows = []
    formula_errors = []
    cleaned_input_cells = []
    format_errors = []
    source_names = tuple(name for name in F if name != 'bill_number')
    formula_row_anchors = {
        'company_code', 'company_name', 'bill_date', 'summary',
        'accounting_subject',
        'internal_collection_flag', 'internal_collection_type',
        'receipt_method', 'settlement_method',
        'receipt_original_amount', 'tax_method', 'tax_rate', 'merge_flag',
    }
    for i in range(EXCEL_CONFIG["start_row"], ws.max_row + 1):
        original_fill_status = excel_text(
            ws.cell(row=i, column=config["status_col"]).value
        )
        fill_status = _normalize_existing_fill_status(original_fill_status)
        row = {'_row': i}
        row_formula_cells = []
        row_formula_names = set()
        row_cleaned_cells = []
        has_business_data = False
        for name, col in F.items():
            cell = ws.cell(row=i, column=col)
            v = cell.value
            cleaned_value = normalize_receipt_value(_clean_input_text(v), name) if isinstance(v, str) else v
            if isinstance(v, str) and cleaned_value != v:
                row_cleaned_cells.append(f'第{i}行{get_column_letter(col)}列')
            if name in source_names:
                displayed_value = (
                    value_ws.cell(row=i, column=col).value
                    if cell.data_type == 'f' else v
                )
                if excel_text(displayed_value) != '':
                    has_business_data = True
                if cell.data_type == 'f':
                    row_formula_names.add(name)
                    row_formula_cells.append(
                        f'第{i}行{get_column_letter(col)}列'
                    )
            row[name] = cleaned_value if cleaned_value is not None else ''
        if not has_business_data:
            if (
                fill_status in ('', STATUS_ABNORMAL)
                and row_formula_names.intersection(formula_row_anchors)
            ):
                # 关键业务列只有公式且无缓存时，无法区分“空白预留行”和
                # “未重算的真实业务行”。财务制单宁可中止，也不能静默漏单。
                formula_errors.extend(row_formula_cells)
            continue
        if fill_status != original_fill_status:
            ws.cell(
                row=i, column=config["status_col"],
                value=fill_status
            )
            if original_fill_status == LEGACY_REVIEW_STATUS:
                reason_cell = ws.cell(
                    row=i, column=config["reason_col"]
                )
                reason_cell.value = _simplify_legacy_review_reason(
                    reason_cell.value
                )
            normalized_status_rows.append(
                (i, original_fill_status, fill_status)
            )
        if fill_status not in ('', STATUS_ABNORMAL):
            skipped_status_rows.append((i, fill_status))
            continue
        formula_errors.extend(row_formula_cells)
        for name, col in F.items():
            cell = ws.cell(row=i, column=col)
            header = normalize_header(ws.cell(row=header_row, column=col).value)
            if header in IDENTIFIER_HEADERS and cell.data_type != 'f':
                issue = identifier_source_issue(cell.value, cell.number_format)
                if issue:
                    format_errors.append(f'{get_column_letter(col)}{i}（{header}）：{issue}')
        rows.append(row)
        cleaned_input_cells.extend(row_cleaned_cells)

    value_wb.close()
    if normalized_status_rows:
        try:
            wb.save(file_path)
        except PermissionError as error:
            wb.close()
            raise Exception(
                '需要把底稿中的旧填写状态改为“异常”，但Excel文件正被占用；'
                '请关闭Excel后重新运行'
            ) from error
        preview = '；'.join(
            f'第{row_no}行={old_status or "空值"}→{new_status or "空值"}'
            for row_no, old_status, new_status in normalized_status_rows[:10]
        )
        if len(normalized_status_rows) > 10:
            preview += f'；另有{len(normalized_status_rows) - 10}行'
        print(f'填写状态规范化完成：{preview}。')
    wb.close()
    if format_errors:
        raise FieldFillError('底稿格式校验', '；'.join(format_errors[:20]) +
                             '。请从完整原始数据重新粘贴，程序尚未开始填报。')
    if formula_errors:
        preview = '、'.join(formula_errors[:20])
        if len(formula_errors) > 20:
            preview += f'，另有{len(formula_errors) - 20}处'
        raise Exception(
            '待处理输入区域含公式，程序无法保证公式结果为最新值：'
            f'{preview}。请在Excel中复制这些公式并“粘贴为值”后重试。'
        )

    if skipped_status_rows:
        preview = '；'.join(
            f'第{row_no}行={status}'
            for row_no, status in skipped_status_rows[:10]
        )
        if len(skipped_status_rows) > 10:
            preview += f'；另有{len(skipped_status_rows) - 10}行'
        print(
            '填写状态检查完成：仅处理“填写状态”为空或“异常”的数据；'
            f'本次待处理 {len(rows)} 行，已跳过 '
            f'{len(skipped_status_rows)} 行（{preview}）。'
        )
    else:
        print(
            '填写状态检查完成：所有有效数据的“填写状态”均为空或“异常”，'
            f'本次待处理 {len(rows)} 行。'
        )
    if cleaned_input_cells:
        preview = '、'.join(cleaned_input_cells[:10])
        if len(cleaned_input_cells) > 10:
            preview += f'，另有{len(cleaned_input_cells) - 10}处'
        print(
            '数据清洗完成：已清除首尾普通/隐藏空格并统一文本中英文标点，'
            f'共 {len(cleaned_input_cells)} 处（{preview}）；'
            '填入NCC时将使用清洗后的数据。'
        )
    return rows


def _format_calendar_date(value, field_name, source_name):
    """把 Excel 日期或常用日期文本统一为 YYYY-MM-DD。"""
    import datetime as _dt

    if isinstance(value, (_dt.datetime, _dt.date)):
        return value.strftime('%Y-%m-%d')
    text = excel_text(value)
    if text == '':
        raise FieldFillError(field_name, f'{source_name}为空')
    for sep in ('/', '.', '年', '月'):
        text = text.replace(sep, '-')
    text = text.replace('日', '')
    parts = [p for p in text.split('-') if p != '']
    if len(parts) == 3 and all(p.isdigit() for p in parts):
        y, m, d = parts
        try:
            parsed = _dt.date(int(y), int(m), int(d))
        except ValueError as error:
            raise FieldFillError(
                field_name, f'{source_name}日期无效：{value}（{error}）'
            ) from error
        return parsed.strftime('%Y-%m-%d')
    raise FieldFillError(
        field_name,
        f'{source_name}格式无法识别：{value}；'
        '请使用YYYY-MM-DD或YYYY/MM/DD'
    )


def format_bill_date(value):
    """
    单据日期标准化为 YYYY-MM-DD（自动补零）。
    明确支持 YYYY-MM-DD、YYYY/MM/DD 及 Excel 日期单元格。
    """
    return _format_calendar_date(value, '单据日期', '底稿“单据日期”')


def format_settlement_date(value):
    """
    结算日期标准化为 YYYY-MM-DD（自动补零）。
    明确支持 YYYY-MM-DD、YYYY/MM/DD 及 Excel 日期单元格。
    """
    return _format_calendar_date(value, '结算日期', '底稿“结算日期”')


def login(main_page):
    """等待人工完成 NCC 登录，以首页搜索图标出现作为唯一成功证据。"""
    login_wait_text = format_timing_seconds('login_wait')
    print(f'正在等待NCC网站登录（最长{login_wait_text}秒）......')
    deadline = task_now() + timing_seconds('login_wait')
    finder = main_page.locator(SELECTORS["home_search_icon"]).first
    while task_now() < deadline:
        pause()
        if main_page.is_closed():
            raise FieldFillError('NCC登录', '登录页面已被关闭')
        try:
            if finder.count() > 0 and finder.is_visible():
                print('NCC网站登录成功')
                return
        except Exception:
            pass
        task_sleep(timing_seconds('candidate_poll_interval'))
    raise FieldFillError(
        'NCC登录',
        f'{login_wait_text}秒内未检测到NCC首页，登录未完成'
    )


def operate_page(main_page, url, context):
    """
    登录成功后：点击右上角搜索图标 → 输入节点名 → 点选第一条结果
    返回节点页面对象（新开标签页或当前页跳转均兼容）
    """
    # ---------- 1. login() 已确认首页加载完成 ----------
    finder = main_page.locator(SELECTORS["home_search_icon"]).first
    finder.wait_for(
        state='visible', timeout=timing_ms('element_action_timeout')
    )
    task_sleep(timing_seconds('menu_open_wait'))
    pause()

    # ---------- 2. 点击右上角搜索图标，唤出搜索框 ----------
    print(f'登录NCC网站成功，正在搜索并打开{NCC_CONFIG["node_name"]}节点......')
    try:
        pause()
        main_page.locator(SELECTORS["search_open_icon"]).first.click()
        task_sleep(timing_seconds('menu_open_wait'))
    except Exception:
        # 搜索框可能默认已展开，点击失败则直接尝试输入
        pass

    # ---------- 3. 在搜索框输入节点名（主选择器失败时尝试备用定位） ----------
    input_candidates = [
        SELECTORS["node_search_input"],
        "xpath=//input[contains(@class,'ant-select-search__field')]",
        "xpath=//input[contains(@placeholder,'搜索') or contains(@placeholder,'Search')]",
        "xpath=(//input[@type='text' and contains(@class,'ant-input')])[last()]",
    ]
    url_input = None
    for css in input_candidates:
        try:
            cand = main_page.locator(css).last
            cand.wait_for(
                state='visible', timeout=timing_ms('short_element_timeout')
            )
            url_input = cand
            break
        except Exception:
            continue
    if url_input is None:
        raise Exception('未找到菜单搜索框，请更新 SELECTORS["node_search_input"]')
    clear_before_fill(url_input, '菜单搜索框')
    url_input.fill(url)
    task_sleep(timing_seconds('menu_result_render_wait'))

    # ---------- 4. 点选第一条搜索结果 ----------
    first_candidates = [
        SELECTORS["node_search_first"],
        "xpath=(//ul[contains(@class,'ant-select-dropdown-menu')]/li)[1]",
        f"xpath=(//li[contains(normalize-space(.),'{url}')])[1]",
        f"xpath=(//div[contains(@class,'dropdown') or contains(@class,'menu')]//*[normalize-space(text())='{url}'])[1]",
    ]
    new_page = None
    last_click_error = None
    for css in first_candidates:
        try:
            url_choose = main_page.locator(css).first
            url_choose.wait_for(
                state='visible', timeout=timing_ms('short_element_timeout')
            )
        except Exception:
            continue
        click_returned = False
        try:
            pause()
            try:
                with context.expect_page(
                    timeout=timing_ms('element_action_timeout')
                ) as new_page_info:
                    url_choose.click()
                    click_returned = True
                new_page = new_page_info.value
            except PlaywrightTimeoutError as error:
                if not click_returned:
                    raise error
                new_page = main_page
            break
        except Exception as error:
            last_click_error = error
            continue
    if new_page is None:
        if last_click_error is not None:
            raise FieldFillError(
                '打开NCC节点',
                f'已找到“{url}”搜索结果，但点击失败：{last_click_error}'
            ) from last_click_error
        raise FieldFillError(
            '打开NCC节点',
            f'未找到“{url}”的搜索结果，请更新节点搜索选择器'
        )
    pause()
    new_page.wait_for_load_state(
        timeout=timing_ms('page_load_timeout')
    )
    task_sleep(timing_seconds('page_transition_wait'))
    return new_page


def add(iframe):
    """点击“新增”按钮，进入制单卡片"""
    finder_btn = iframe.locator(SELECTORS["add_btn"]).first
    finder_btn.wait_for(
        state='visible', timeout=timing_ms('long_element_timeout')
    )
    finder_btn.click(timeout=timing_ms('long_element_timeout'))
    task_sleep(timing_seconds('page_transition_wait'))
    expect(iframe.locator(SELECTORS["back_btn_xpath"])).not_to_be_attached(
        timeout=timing_ms('page_load_timeout')
    )
    pause()


def choose_collection_status(iframe, group):
    """选择收方单位组织"""
    if group == '':
        return
    locator_input = iframe.locator(SELECTORS["income_org_input"])
    locator_choose = iframe.locator(SELECTORS["income_org_item"].format(text=group)).first
    try_except_click(iframe, locator_input, group, locator_choose)
    try:
        iframe.locator(SELECTORS["income_org_match"]).wait_for(
            state='visible', timeout=timing_ms('reference_wait')
        )
    except Exception:
        pass
    pause()


def choose_payment_status(iframe, group):
    """选择付方单位组织"""
    if group == '':
        return
    locator_input = iframe.locator(SELECTORS["pay_org_input"])
    locator_choose = iframe.locator(SELECTORS["pay_org_item"].format(text=group)).first
    try_except_click(iframe, locator_input, group, locator_choose)
    try:
        iframe.locator(SELECTORS["pay_org_match"]).wait_for(
            state='visible', timeout=timing_ms('reference_wait')
        )
    except Exception:
        pass
    pause()


def choose_income_account(iframe, account):
    """选择收款单位内部账户（参照弹窗，选填）"""
    if account == '':
        return
    iframe.locator(SELECTORS["refer_icon"]).first.click()
    row = iframe.locator(SELECTORS["income_acct_row"].format(text=account)).first
    try:
        row.wait_for(
            state='visible', timeout=timing_ms('element_action_timeout')
        )
        row.click(timeout=timing_ms('element_action_timeout'))
        iframe.locator(SELECTORS["refer_sure_btn"]).click(
            timeout=timing_ms('element_action_timeout')
        )
    except Exception:
        print('      未找到指定的收款单位内部账户,请核实底稿是否正确.......')
        raise Exception('报错')
    pause()


def choose_pay_account(iframe, account):
    """选择付款单位内部账户（参照弹窗，选填）"""
    if account == '':
        return
    iframe.locator(SELECTORS["refer_icon"]).first.click()
    row = iframe.locator(SELECTORS["pay_acct_row"].format(text=account)).first
    try:
        row.wait_for(
            state='visible', timeout=timing_ms('element_action_timeout')
        )
        row.click(timeout=timing_ms('element_action_timeout'))
        iframe.locator(SELECTORS["refer_sure_btn"]).click(
            timeout=timing_ms('element_action_timeout')
        )
    except Exception:
        print('      未找到指定的付款单位内部账户,请核实底稿是否正确.......')
        raise Exception('报错')
    pause()


def business_type(iframe, typer):
    """选择业务类型（如 10/非交易性往来清理）"""
    if typer == '':
        return
    locator_input = iframe.locator(SELECTORS["biz_type_input"])
    locator_choose = iframe.locator(SELECTORS["biz_type_item"].format(text=typer))
    try_except_click(iframe, locator_input, typer, locator_choose)
    try:
        iframe.locator(
            SELECTORS["biz_type_value"].format(text=typer)
        ).wait_for(
            state='visible', timeout=timing_ms('element_action_timeout')
        )
    except Exception:
        pass
    pause()


def remark(iframe, remake):
    """填写卡片头备注"""
    if remake == '':
        return
    locator_input = iframe.locator(SELECTORS["head_remark"])
    try:
        clear_before_fill(locator_input, '备注')
        locator_input.fill(remake)
    except Exception:
        pass
    pause()


def expand(iframe, n):
    """打开表体第 n 行的行编辑弹窗"""
    body = iframe.locator(SELECTORS["body_table"])
    body.evaluate('element => element.scrollTop = element.scrollHeight')
    edit_btn = iframe.locator(SELECTORS["row_edit_btn"]).nth(n - 1)
    edit_btn.wait_for(
        state='visible', timeout=timing_ms('long_element_timeout')
    )
    edit_btn.click(timeout=timing_ms('long_element_timeout'))
    iframe.locator(SELECTORS["modal_title"]).wait_for(
        state='visible', timeout=timing_ms('long_element_timeout')
    )
    pause()


def add_expand(iframe):
    """表体新增一行（合并类第 2 行起）"""
    # TODO 新表单若“增行”按钮选择器不同，请修改此处（旧版通过表体工具栏增行）
    body = iframe.locator(SELECTORS["body_table"])
    body.evaluate('element => element.scrollTop = element.scrollHeight')
    task_sleep(timing_seconds('page_transition_wait'))
    pause()


def _modal_input(iframe, nth):
    """行编辑弹窗中第 nth 个表单项的输入框"""
    return iframe.locator(
        f'div.card-table-modal-form-wrap-body.cf > div:nth-child({nth}) input.nc-input.refer-input.u-form-control.md')


def expand_remark(iframe, remake, n):
    """表体第 n 行：填写摘要"""
    if remake == '':
        return
    locator_input = _modal_input(iframe, SELECTORS["modal_remark_nth"])
    locator_wait = iframe.locator(
        f'div.card-table-modal-form-wrap-body.cf > div:nth-child({SELECTORS["modal_remark_nth"]}) input[value=\'{remake}\']')
    expand_try_except(iframe, locator_input, remake, locator_wait, n)
    pause()


def expand_money(iframe, money, n):
    """表体第 n 行：填写收款原币金额"""
    locator_input = iframe.locator(SELECTORS["modal_money"])
    try:
        clear_before_fill(locator_input, '收款原币金额')
        locator_input.fill(f'{money:.2f}')
        iframe.locator(
            f'div.card-table-modal-form-wrap-body.cf > div:nth-child({SELECTORS["modal_remark_nth"]}) input[value]').first.click()
    except Exception:
        pass
    pause()


def expand_type(iframe, type, s_file, file, n):
    """表体第 n 行：选择收支类型（收回代付 / 收回被代收）"""
    if type == '':
        return
    locator_input = _modal_input(iframe, SELECTORS["modal_wl_type_nth"])
    locator_choose = iframe.locator(SELECTORS["modal_cascade_item"].format(text=type))
    expand_try_except_click(iframe, locator_input, type, locator_choose, n)
    if type == '收回代付':
        expand_file(iframe, s_file, n)
    elif type == '收回被代收':
        expand_project(iframe, file, n)
    pause()


def income_type(iframe, type, text):
    """表体：收方分灶收入类型（仅“收回被代收”时）"""
    if type != '收回被代收' or text == '':
        return
    locator_input = _modal_input(iframe, SELECTORS["modal_income_nth"])
    locator_choose = iframe.locator(SELECTORS["modal_cascade_item"].format(text=text))
    try_except_click(iframe, locator_input, text, locator_choose)
    pause()


def expand_file(iframe, text, n):
    """表体第 n 行：被收回方项目档案"""
    if text == '':
        return
    locator_input = _modal_input(iframe, SELECTORS["modal_s_file_nth"])
    locator_choose = iframe.locator(SELECTORS["modal_cascade_item"].format(text=text))
    expand_try_except_click(iframe, locator_input, text, locator_choose, n)
    pause()


def expand_department(iframe, text, n):
    """表体第 n 行：部门"""
    if text == '':
        return
    locator_input = _modal_input(iframe, SELECTORS["modal_dept_nth"])
    locator_choose = iframe.locator(SELECTORS["modal_dropdown_item"].format(text=text))
    expand_try_except_click(iframe, locator_input, text, locator_choose, n)
    pause()


def expand_project(iframe, project, n):
    """表体第 n 行：项目档案（收回被代收）"""
    if project == '':
        return
    locator_input = _modal_input(iframe, SELECTORS["modal_project_nth"])
    locator_choose = iframe.locator(SELECTORS["modal_row_by_title"].format(text=project))
    expand_try_except_click(iframe, locator_input, project, locator_choose, n)
    iframe.locator(SELECTORS["modal_save_btn"]).first.click()
    pause()


def expand_operate_type(iframe, abc, n):
    """表体第 n 行：经营类别（收支项为社群时必填）"""
    if abc == '':
        return
    try:
        locator_input = _modal_input(iframe, SELECTORS["modal_operate_nth"])
        locator_choose = iframe.locator(SELECTORS["modal_dropdown_item"].format(text=abc))
        expand_try_except_click(iframe, locator_input, abc, locator_choose, n)
    except Exception:
        print('     请核对是否无法选择经营类别，跳过该单据！！！')
        raise
    pause()


def expand_exit(iframe):
    """关闭行编辑弹窗（收起侧拉）"""
    try:
        iframe.locator('span.icon.iconfont.icon-you.icon-celashouqi').click()
    except Exception:
        pass
    pause()


def _normalize_ncc_error_text(value):
    """压缩 NCC 提示文本，并去掉通知组件附带的“关闭”按钮文案。"""
    text = ' '.join(str(value or '').split()).strip()
    text = re.sub(r'(关闭|Close)$', '', text, flags=re.IGNORECASE).strip()
    return text


def visible_ncc_messages(frame):
    """读取当前可见的 NCC 通知文本。"""
    selectors = (
        SELECTORS["error_toast"],
        "css=.toast-content",
        "css=.u-message-notice-content",
        "css=.u-notification-notice-description",
        "css=.u-notification-notice-message",
        SELECTORS["error_group"] + " li",
    )
    messages = []
    for selector in selectors:
        elements = frame.locator(selector)
        try:
            count = min(elements.count(), 20)
        except Exception:
            continue
        for index in range(count):
            element = elements.nth(index)
            try:
                if not element.is_visible():
                    continue
                text = _normalize_ncc_error_text(
                    element.inner_text(
                        timeout=timing_ms('tiny_element_timeout')
                    )
                )
                if text == '':
                    continue
                if text not in messages:
                    messages.append(text)
            except Exception:
                continue
    return messages


def _is_ncc_success_message(text, action_text=''):
    negative_words = ('失败', '错误', '异常', '未成功', '不成功', '未能')
    if any(word in text for word in negative_words):
        return False
    negative_success = re.compile(
        r'(?:未|尚未|无法|不能|没有|未能|不是|并非)'
        r'[^，,。；;！？!?\n]{0,30}(?:保存|提交|操作)'
        r'[^，,。；;！？!?\n]{0,10}成功'
        r'|(?:保存|提交|操作)[^，,。；;！？!?\n]{0,12}'
        r'(?:未|尚未|无法|不能|没有|未能|不是|并非)'
        r'[^，,。；;！？!?\n]{0,10}成功'
    )
    if negative_success.search(text):
        return False
    success_end = r'(?:[！!。；;，,：:\s]|$)'
    if action_text == '保存提交':
        return bool(re.search(
            rf'(?:保存提交|提交)(?:操作)?成功{success_end}', text
        ))
    if action_text == '保存':
        return bool(re.search(
            rf'(?:保存(?:操作)?成功|操作成功){success_end}', text
        ))
    return bool(re.search(
        rf'(?:保存(?:操作)?成功|提交(?:操作)?成功|操作成功)'
        rf'{success_end}',
        text,
    ))


def visible_ncc_success_messages(frame, action_text='', messages=None):
    source_messages = (
        visible_ncc_messages(frame) if messages is None else messages
    )
    return [
        message for message in source_messages
        if _is_ncc_success_message(message, action_text)
    ]


def visible_ncc_system_errors(frame, messages=None):
    """读取当前错误；“未成功/不成功”绝不会再被当作成功过滤。"""
    progress_words = ('正在保存', '正在提交', '处理中', '请稍候')
    negative_words = ('失败', '错误', '异常', '未成功', '不成功', '未能')
    errors = []
    source_messages = (
        visible_ncc_messages(frame) if messages is None else messages
    )
    for message in source_messages:
        if _is_ncc_success_message(message):
            continue
        if (
            any(word in message for word in progress_words)
            and not any(word in message for word in negative_words)
        ):
            continue
        errors.append(message)
    return errors


def _wait_for_old_notifications_to_clear(frame, timeout=None):
    """保存前清掉旧提示，避免旧 toast 被误当作本次结果。"""
    if timeout is None:
        timeout = timing_ms('notification_clear_wait')
    deadline = task_now() + timeout / 1000
    while task_now() < deadline:
        pause()
        if not visible_ncc_messages(frame):
            return True
        task_sleep(timing_seconds('field_poll_interval'))
    return not visible_ncc_messages(frame)


def _submitted_status_visible(frame):
    selectors = (
        "xpath=//*[contains(@class,'status') and normalize-space(.)='已提交']",
        "xpath=//*[contains(@class,'bill') and normalize-space(.)='已提交']",
    )
    for selector in selectors:
        elements = frame.locator(selector)
        try:
            for index in range(min(elements.count(), 20)):
                if elements.nth(index).is_visible():
                    return True
        except Exception:
            continue
    return False


def _loading_visible(frame):
    elements = frame.locator(SELECTORS["loading"])
    try:
        return any(
            elements.nth(index).is_visible()
            for index in range(min(elements.count(), 20))
        )
    except Exception:
        return False


def _wait_for_action_result(
    frame, go, before_bill_number='', before_list_ready=False,
    before_submitted_status=False,
    before_notifications=None,
    timeout=None
):
    """保存/提交后必须取得明确正向证据，否则结果按异常处理。"""
    if timeout is None:
        timeout = timing_ms('action_result_wait')
    button_text = '保存提交' if go else '保存'
    persisting_old_notifications = set(before_notifications or ())
    deadline = task_now() + timeout / 1000
    while task_now() < deadline:
        pause()
        current_notifications = list(visible_ncc_messages(frame))
        current_notification_set = set(current_notifications)
        persisting_old_notifications.intersection_update(
            current_notification_set
        )
        new_notifications = [
            message for message in current_notifications
            if message not in persisting_old_notifications
        ]
        errors = visible_ncc_system_errors(
            frame, messages=new_notifications
        )
        if errors:
            raise NccSaveRejectedError(
                button_text,
                f'{button_text}结果需人工核对：{"；".join(errors)}'
            )

        bill_number = _find_bill_number_once(frame)
        successes = visible_ncc_success_messages(
            frame, button_text, messages=new_notifications
        )
        if successes:
            print(f'     已检测到明确成功提示: {"；".join(successes)}')
            return bill_number
        if (
            not go and bill_number
            and (not before_bill_number or bill_number != before_bill_number)
        ):
            print(f'     已检测到保存后生成的新单据编号: {bill_number}')
            return bill_number
        if go and not before_submitted_status and _submitted_status_visible(frame):
            print('     已检测到单据状态“已提交”')
            return bill_number
        if not go and not before_list_ready and _list_page_ready(frame):
            print('     保存操作后已返回单据列表')
            return bill_number
        task_sleep(timing_seconds('candidate_poll_interval'))

    loading_note = '，且加载状态仍未结束' if _loading_visible(frame) else ''
    raise NccCommitUncertainError(
        button_text,
        f'{button_text}结果需人工核对：'
        f'{int(timeout / 1000)}秒内未检测到明确成功或失败证据{loading_note}；'
        '请在NCC中人工核对，禁止直接重跑'
    )


def save(iframe, go):
    """点击保存/提交并等待明确成功证据；返回当前单据编号（如可读取）。"""
    button_text = '保存提交' if go else '保存'
    buttons = iframe.locator(
        f"xpath=//button[normalize-space(.)='{button_text}']"
    )
    save_choose = None
    for index in range(buttons.count()):
        button = buttons.nth(index)
        try:
            if button.is_visible() and button.is_enabled():
                save_choose = button
                break
        except Exception:
            continue
    if save_choose is None:
        raise FieldFillError(button_text, f'未找到可点击的“{button_text}”按钮')
    _wait_for_old_notifications_to_clear(iframe)
    before_notifications = set(visible_ncc_messages(iframe))
    if before_notifications:
        print(
            '     保存前仍有旧提示，程序会继续点击保存，并在旧提示消失前'
            '忽略这些既有文案，避免把填写阶段提示误判为保存结果'
        )
    before_bill_number = _find_bill_number_once(iframe)
    before_list_ready = _list_page_ready(iframe)
    before_submitted_status = _submitted_status_visible(iframe)
    pause()
    try:
        save_choose.click(timeout=timing_ms('long_element_timeout'))
    except Exception as error:
        raise NccCommitUncertainError(
            button_text,
            f'点击“{button_text}”时发生异常，无法确认请求是否已发出：{error}'
        ) from error
    print(f'     已点击“{button_text}”，正在等待明确结果......')
    try:
        bill_number = _wait_for_action_result(
            iframe, go, before_bill_number, before_list_ready,
            before_submitted_status,
            before_notifications=before_notifications
        )
    except TaskCancelledError as error:
        raise NccCommitUncertainError(
            button_text,
            '按钮已经点击后用户停止了任务，无法确认NCC最终结果；'
            '必须人工核对，禁止直接重跑'
        ) from error
    except NccCommitUncertainError:
        raise
    except Exception as error:
        raise NccCommitUncertainError(
            button_text,
            f'按钮已点击，但等待结果时发生异常：{error}'
        ) from error
    return bill_number


def _parse_bill_number(value):
    """从保存后的标题、字段文本或输入框值中提取 NCC 单据编号。"""
    text = ' '.join(str(value or '').split())
    if text == '':
        return ''
    patterns = (
        r'单据编号\s*[：:]?\s*([A-Za-z0-9][A-Za-z0-9_-]{5,})',
        r'收款结算\s*[：:]?\s*([A-Za-z0-9][A-Za-z0-9_-]{5,})',
        r'\b(F[A-Za-z0-9_-]{8,})\b',
    )
    for pattern in patterns:
        match = re.search(pattern, text)
        if match is not None:
            return match.group(1)
    if re.fullmatch(r'[A-Za-z0-9][A-Za-z0-9_-]{5,}', text):
        return text
    return ''


def _find_bill_number_once(frame):
    """单次扫描当前卡片中的单据编号；未找到时返回空字符串。"""
    input_selectors = (
        "css=input[label*='单据编号']",
        "css=input[attrcode*='billno']",
        "css=input[attrcode*='vbillno']",
    )
    text_selectors = (
        SELECTORS["bill_info"],
        "xpath=//*[contains(@class,'bill-info') and contains(normalize-space(.),'单据编号')]",
        "xpath=//*[contains(@class,'bill-info') and contains(normalize-space(.),'收款结算') and "
        "(contains(normalize-space(.),'：') or contains(normalize-space(.),':'))]",
        "xpath=//*[contains(normalize-space(.),'单据编号') and "
        "string-length(normalize-space(.)) <= 100]",
    )
    for selector in input_selectors:
        inputs = frame.locator(selector)
        try:
            count = min(inputs.count(), 20)
        except Exception:
            continue
        for index in range(count):
            locator_input = inputs.nth(index)
            try:
                if not locator_input.is_visible():
                    continue
                for candidate in (
                    locator_input.input_value(
                        timeout=timing_ms('tiny_element_timeout')
                    ),
                    locator_input.get_attribute('title'),
                ):
                    bill_number = _parse_bill_number(candidate)
                    if bill_number:
                        return bill_number
            except Exception:
                continue

    for selector in text_selectors:
        elements = frame.locator(selector)
        try:
            count = min(elements.count(), 80)
        except Exception:
            continue
        for index in range(count):
            element = elements.nth(index)
            try:
                if not element.is_visible():
                    continue
                for candidate in (
                    element.inner_text(
                        timeout=timing_ms('tiny_element_timeout')
                    ),
                    element.text_content(
                        timeout=timing_ms('tiny_element_timeout')
                    ),
                    element.get_attribute('title'),
                ):
                    bill_number = _parse_bill_number(candidate)
                    if bill_number:
                        return bill_number
            except Exception:
                continue

    # 部分 NCC 版本把“单据编号：”与编号拆成相邻的普通文本节点，
    # 没有 bill-info/input 等稳定属性。最后扫描当前卡片正文，按
    # “单据编号 + 编号”的上下文提取，避免只查旧版标题节点而漏读。
    try:
        body = frame.locator('body')
        for candidate in (
            body.inner_text(timeout=timing_ms('tiny_element_timeout')),
            body.text_content(timeout=timing_ms('tiny_element_timeout')),
        ):
            bill_number = _parse_bill_number(candidate)
            if bill_number:
                return bill_number
    except Exception:
        pass
    return ''


def read_saved_bill_number(frame, timeout=None):
    """仅保存后读取页面生成的单据编号，供回写底稿 L 列。"""
    if timeout is None:
        timeout = timing_ms('bill_number_wait')
    deadline = task_now() + timeout / 1000
    while task_now() < deadline:
        pause()
        bill_number = _find_bill_number_once(frame)
        if bill_number:
            print(f'     已读取保存后的单据编号: {bill_number}')
            return bill_number
        task_sleep(timing_seconds('candidate_poll_interval'))
    raise FieldFillError(
        '单据编号',
        f'保存成功后{int(timeout / 1000)}秒内未能从页面读取单据编号'
    )


def _visible_list_add_button(iframe):
    """返回列表页中真正可见、可用的精确“新增”按钮。"""
    try:
        buttons = iframe.locator(SELECTORS["list_add_btn"])
        for index in range(buttons.count() - 1, -1, -1):
            button = buttons.nth(index)
            if (
                button.is_visible()
                and button.is_enabled()
                and button.inner_text(
                    timeout=timing_ms('tiny_element_timeout')
                ).strip() == '新增'
            ):
                return button
    except Exception:
        pass
    return None


def _list_page_ready(iframe):
    """判断是否已经回到包含可见“新增”按钮的单据列表页。"""
    if _visible_list_add_button(iframe) is None:
        return False
    # NCC 可能把列表保留在卡片背后；只要卡片操作按钮仍可见，就不算返回列表。
    card_buttons = iframe.locator(
        "xpath=//button[normalize-space(.)='保存' or "
        "normalize-space(.)='保存提交' or normalize-space(.)='取消']"
    )
    try:
        for index in range(card_buttons.count()):
            if card_buttons.nth(index).is_visible():
                return False
    except Exception:
        return False
    return True


def _wait_for_list_page(iframe, timeout=None):
    if timeout is None:
        timeout = timing_ms('list_return_wait')
    deadline = task_now() + timeout / 1000
    while task_now() < deadline:
        pause()
        if _list_page_ready(iframe):
            return True
        task_sleep(timing_seconds('field_poll_interval'))
    return False


def back(iframe):
    """保存后返回单据列表，为下一张单据做准备。"""
    if _wait_for_list_page(
        iframe, timeout=timing_ms('short_element_timeout')
    ):
        print('     已检测到“新增”按钮，将继续新增下一张单据')
        return

    back_candidates = [
        iframe.locator(SELECTORS["back_btn"]),
        iframe.locator(SELECTORS["back_btn_xpath"]),
        iframe.locator("css=span.back-btn"),
    ]
    for candidates in back_candidates:
        for index in range(candidates.count() - 1, -1, -1):
            button = candidates.nth(index)
            try:
                if button.is_visible():
                    button.click(
                        timeout=timing_ms('element_action_timeout')
                    )
                    if _wait_for_list_page(iframe):
                        print('     已返回单据列表')
                        pause()
                        return
            except Exception:
                continue

    # 部分 NCC 版本保存后仍停留卡片，可用“取消”退出已保存卡片。
    cancel_buttons = iframe.locator(
        "xpath=//button[normalize-space(.)='取消']"
    )
    for index in range(cancel_buttons.count() - 1, -1, -1):
        button = cancel_buttons.nth(index)
        try:
            if button.is_visible() and button.is_enabled():
                button.click(timeout=timing_ms('element_action_timeout'))
                task_sleep(timing_seconds('reference_post_select_wait'))
                confirms = iframe.locator(
                    "xpath=//button[normalize-space(.)='确定']"
                )
                for confirm_index in range(
                    confirms.count() - 1, -1, -1
                ):
                    confirm = confirms.nth(confirm_index)
                    if confirm.is_visible() and confirm.is_enabled():
                        confirm.click(
                            timeout=timing_ms('element_action_timeout')
                        )
                        break
                if _wait_for_list_page(iframe):
                    print('     已退出已保存卡片并返回单据列表')
                    pause()
                    return
        except Exception:
            continue
    raise Exception('单据保存后未能返回列表，无法继续处理下一张单据')


def error(iframe):
    """兼容旧调用；填写阶段不再读取 NCC 业务报错。"""
    return '请检查底稿表格信息是否正确，或者NCC运行状况是否正常......'


# ==================== 收款结算 · 制单卡片操作（第1~3步） ====================

def _node_marker_visible(root):
    """识别节点标题/页签，允许“收款结算 - 查询”等后缀文案。"""
    node_name = NCC_CONFIG["node_name"]
    markers = root.locator(
        "xpath=//*[contains(normalize-space(.), "
        f'\"{node_name}\") and ('
        "self::h1 or self::h2 or self::h3 or "
        "contains(translate(@class,'ABCDEFGHIJKLMNOPQRSTUVWXYZ',"
        "'abcdefghijklmnopqrstuvwxyz'),'title') or "
        "contains(translate(@class,'ABCDEFGHIJKLMNOPQRSTUVWXYZ',"
        "'abcdefghijklmnopqrstuvwxyz'),'tab') or "
        "contains(translate(@class,'ABCDEFGHIJKLMNOPQRSTUVWXYZ',"
        "'abcdefghijklmnopqrstuvwxyz'),'breadcrumb') or "
        "contains(translate(@class,'ABCDEFGHIJKLMNOPQRSTUVWXYZ',"
        "'abcdefghijklmnopqrstuvwxyz'),'header') or "
        "contains(translate(@class,'ABCDEFGHIJKLMNOPQRSTUVWXYZ',"
        "'abcdefghijklmnopqrstuvwxyz'),'caption'))]"
    )
    try:
        for index in range(min(markers.count(), 50)):
            if markers.nth(index).is_visible():
                return True
    except Exception:
        pass
    return False


def _work_root_ready(root, marker_root=None):
    """业务控件与节点标识可分别位于 iframe 和外层壳。"""
    settle_inputs = root.locator(SELECTORS["settle_org_input"])
    try:
        for index in range(min(settle_inputs.count(), 20)):
            if settle_inputs.nth(index).is_visible():
                return True
    except Exception:
        pass

    add_visible = False
    add_buttons = root.locator(SELECTORS["list_add_btn"])
    try:
        for index in range(min(add_buttons.count(), 20)):
            if add_buttons.nth(index).is_visible():
                add_visible = True
                break
    except Exception:
        return False
    if not add_visible:
        return False
    if _node_marker_visible(root):
        return True
    return marker_root is not None and _node_marker_visible(marker_root)


def get_work_frame(page, timeout=None):
    """等待真正就绪的 iframe 或直出业务页面，避免 SPA 注入竞态。"""
    if timeout is None:
        timeout = timing_ms('work_frame_wait')
    deadline = task_now() + timeout / 1000
    while task_now() < deadline:
        pause()
        if page.is_closed():
            raise FieldFillError('NCC节点页', '节点页面已被关闭')
        try:
            iframe_elements = page.locator(NCC_CONFIG["main_iframe"])
            frame_count = iframe_elements.count()
            for index in range(frame_count):
                try:
                    candidate = page.frame_locator(
                        NCC_CONFIG["main_iframe"]
                    ).nth(index)
                    if _work_root_ready(candidate, marker_root=page):
                        return candidate
                except Exception:
                    continue
            if _work_root_ready(page):
                return page
        except Exception:
            pass
        task_sleep(timing_seconds('candidate_poll_interval'))
    raise FieldFillError(
        'NCC节点页',
        f'{timeout / 1000:g}秒内未检测到收款结算节点标识与'
        '“新增/结算财务组织”业务控件'
    )


def click_add_bill(frame):
    """第1步：点击列表“新增”；卡片字段是否就绪由后续步骤判断。"""
    btn = None
    deadline = task_now() + timing_seconds('long_element_timeout')
    while task_now() < deadline:
        pause()
        btn = _visible_list_add_button(frame)
        if btn is not None:
            break
        task_sleep(timing_seconds('field_poll_interval'))
    if btn is None:
        raise Exception('未找到可点击的列表“新增”按钮')
    pause()
    btn.click(timeout=timing_ms('long_element_timeout'))
    print('     已点击“新增”，开始等待新增卡片字段加载')
    pause()


def settlement_org_query_value(value):
    """结算财务组织为“编号/名称”时只返回编号；纯名称保持原样。"""
    source_text = excel_text(value)
    if source_text == '':
        raise FieldFillError('结算财务组织', '底稿“公司编码”为空')
    parts = re.split(r'[/／]', source_text, maxsplit=1)
    if len(parts) == 1:
        return source_text
    code = parts[0].strip()
    name = parts[1].strip()
    if code == '' or name == '':
        raise FieldFillError(
            '结算财务组织',
            f'底稿A列值“{source_text}”不是有效的“编号/名称”格式'
        )
    return code


def project_archive_query_value(value):
    """项目档案为“编码/名称”或“编码\\名称”时只查询编码；纯名称保持原样。"""
    source_text = excel_text(value)
    if source_text == '':
        return ''
    parts = re.split(r'[/\\／＼]', source_text, maxsplit=1)
    if len(parts) == 1:
        return source_text
    code = parts[0].strip()
    name = parts[1].strip()
    if code == '' or name == '':
        raise FieldFillError(
            '项目档案',
            f'底稿“项目档案名称”值“{source_text}”不是有效的“编码/名称”格式'
        )
    return code


def _find_settle_org_input(frame, timeout=None):
    """用字段语义定位组织输入框，兼容 NCC 调整 form-item 层级或类名。"""
    if timeout is None:
        timeout = timing_ms('settle_org_input_wait')
    body = frame.locator('body')
    try:
        return _find_field_input(body, '结算财务组织', timeout=timeout)
    except Exception as semantic_error:
        selectors = (
            "css=input[label='结算财务组织']:not(.hidden-input)",
            "xpath=//input[contains(@label,'结算财务组织') "
            "and not(contains(@label,'编码')) "
            "and not(contains(@class,'hidden-input'))]",
            SELECTORS["settle_org_input"],
        )
        for selector in selectors:
            try:
                inputs = frame.locator(selector)
                for index in range(min(inputs.count(), 30)):
                    locator_input = inputs.nth(index)
                    if (
                        locator_input.is_visible()
                        and locator_input.is_editable()
                    ):
                        print(
                            '     [调试] 结算财务组织输入框通过备用规则命中: '
                            f'{selector}'
                        )
                        return locator_input
            except Exception:
                continue
        raise SettleOrgInputUnavailableError(
            '结算财务组织',
            f'{int(timeout / 1000)}秒内未找到可编辑输入框：{semantic_error}'
        ) from semantic_error


def fill_settle_org(frame, company_code):
    """
    第2步：输入 D 列结算财务组织查询值，等待并点击唯一匹配项。
    NCC 选中后会重绘该字段，因此不再回头查找原输入框。
    """
    company_code = excel_text(company_code)
    org_query = settlement_org_query_value(company_code)
    if org_query != company_code:
        print(
            f'     结算财务组织按“编号/名称”解析：{company_code} -> {org_query}'
        )
    last_error = None
    for attempt in range(1, RETRY_MAX + 1):
        pause()
        org_input = None
        try:
            org_input = _find_settle_org_input(
                frame, timeout=timing_ms('settle_org_input_wait')
            )
            _clear_reference_before_fill(org_input, '结算财务组织')
            org_input.fill(
                org_query, timeout=timing_ms('element_action_timeout')
            )
            option_wait_text = format_timing_seconds(
                'settle_org_option_wait'
            )
            print(
                f'     第{attempt}/{RETRY_MAX}次已填入结算财务组织查询值: '
                f'{org_query}，等待下拉匹配项（最长{option_wait_text}秒）......'
            )
            _click_matching_reference_option(
                frame, '结算财务组织', org_query,
                timeout=timing_ms('settle_org_option_wait'),
                dismiss_input=org_input
            )
            task_sleep(timing_seconds('settle_org_commit_wait'))
            pause()
            return True
        except TaskCancelledError:
            raise
        except SettleOrgInputUnavailableError:
            # 输入框没有渲染属于卡片加载故障，由外层执行“取消→确定→新增”。
            raise
        except Exception as error:
            last_error = error
        if attempt < RETRY_MAX:
            try:
                if org_input is not None:
                    org_input.press(
                        'Escape', timeout=timing_ms('tiny_element_timeout')
                    )
            except Exception:
                pass
            print(
                f'     结算财务组织第{attempt}/{RETRY_MAX}次匹配失败，'
                f'将清空后重新填入：{last_error}'
            )
            task_sleep(timing_seconds('reference_retry_interval'))

    reason = getattr(last_error, 'reason', str(last_error))
    raise FieldFillError(
        '结算财务组织',
        f'连续{RETRY_MAX}次填入A列查询值“{org_query}”后'
        f'仍未匹配到下拉数据：{reason}'
    )


def fill_settle_org_with_card_recovery(
    frame, company_code, card_already_open=False
):
    """
    结算财务组织输入框未在设定时间内出现时，取消当前卡片并重新新增。
    初次打开之外最多重新加载3次；普通下拉无数据仍由字段自身重试。
    """
    last_error = None
    for reload_count in range(SETTLE_ORG_CARD_RELOAD_MAX + 1):
        try:
            if not (reload_count == 0 and card_already_open):
                click_add_bill(frame)
            return fill_settle_org(frame, company_code)
        except TaskCancelledError:
            raise
        except SettleOrgInputUnavailableError as error:
            last_error = error
            if reload_count >= SETTLE_ORG_CARD_RELOAD_MAX:
                break
            print(
                '     结算财务组织输入框'
                f'{format_timing_seconds("settle_org_input_wait")}秒内未出现，'
                f'开始第{reload_count + 1}/{SETTLE_ORG_CARD_RELOAD_MAX}次'
                '重新加载（取消→确定→新增）......'
            )
            if not cancel_card(frame):
                raise FieldFillError(
                    '结算财务组织',
                    '输入框未出现，且取消当前卡片后未能安全返回列表；'
                    '为避免误操作，本单已停止'
                ) from error

    reason = getattr(last_error, 'reason', str(last_error))
    raise FieldFillError(
        '结算财务组织',
        f'初次打开及{SETTLE_ORG_CARD_RELOAD_MAX}次重新加载后，'
        f'输入框仍未出现：{reason}'
    )


def _resolve_input(frame, candidates, field_name):
    """按候选顺序定位输入框，命中时把元素信息打印到控制台（便于排查定位问题）"""
    for css in candidates:
        try:
            cand = frame.locator(css).first
            cand.wait_for(
                state='visible', timeout=timing_ms('element_action_timeout')
            )
            try:
                info = cand.evaluate("el => el.outerHTML.slice(0, 220)")
                print(f'     [调试] {field_name} 命中: {css}')
                print(f'     [调试] 元素: {info}')
            except Exception:
                pass
            return cand
        except Exception:
            continue
    return None


def _dump_card_inputs(frame):
    """排查用：列出当前卡片上所有可见输入框的 class / 是否只读 / 所属字段名"""
    print('     [调试] 未命中任何候选，开始转储卡片上的输入框清单：')
    try:
        inputs = frame.locator("xpath=//input[not(contains(@class,'hidden-input'))]")
        n = min(inputs.count(), 60)
        for i in range(n):
            try:
                info = inputs.nth(i).evaluate(
                    "el => ({cls: el.className, ro: el.readOnly, vis: !!(el.offsetWidth||el.offsetHeight), "
                    "val: el.value, field: el.closest('.form-item') ? el.closest('.form-item').innerText.slice(0,20) : '?'})")
                print(f'     [调试] #{i}: {info}')
            except Exception:
                continue
    except Exception as e:
        print(f'     [调试] 转储失败: {e}')


def _access_input_by_field(frame, field_name, value=None):
    """按输入框最近的 form-item 精确找字段；value 非空时通过原生事件写入只读控件。"""
    return frame.locator("body").evaluate(
        """(root, args) => {
            const wanted = String(args.fieldName || '').replace(/\\s+/g, '');
            const visible = el => !!(el.offsetWidth || el.offsetHeight || el.getClientRects().length);
            const normalize = text => String(text || '').replace(/\\s+/g, '');
            const matches = Array.from(root.querySelectorAll('input:not(.hidden-input)'))
                .map((el, index) => {
                    const item = el.closest('.form-item');
                    const label = normalize(el.getAttribute('label'));
                    const field = normalize(item ? item.innerText : '');
                    let score = 0;
                    if (label.includes(wanted)) score += 100;
                    if (field === wanted || field === `*${wanted}`) score += 80;
                    else if (field.includes(wanted)) score += 40;
                    return { el, index, label, field, score };
                })
                .filter(x => visible(x.el) && x.score > 0)
                .sort((a, b) => b.score - a.score || a.index - b.index);

            if (!matches.length) return null;
            const picked = matches[0];
            const el = picked.el;
            const before = el.value;

            if (args.write) {
                el.removeAttribute('readonly');
                el.readOnly = false;
                el.removeAttribute('disabled');
                el.disabled = false;
                el.focus();
                const descriptor = Object.getOwnPropertyDescriptor(
                    window.HTMLInputElement.prototype, 'value'
                );
                if (descriptor && descriptor.set) {
                    descriptor.set.call(el, args.value);
                } else {
                    el.value = args.value;
                }
                el.dispatchEvent(new Event('input', { bubbles: true }));
                el.dispatchEvent(new Event('change', { bubbles: true }));
                el.dispatchEvent(new KeyboardEvent(
                    'keydown', { key: 'Enter', code: 'Enter', bubbles: true }
                ));
                el.dispatchEvent(new KeyboardEvent(
                    'keyup', { key: 'Enter', code: 'Enter', bubbles: true }
                ));
                el.blur();
            }

            return {
                index: picked.index,
                label: picked.label,
                field: picked.field.slice(0, 40),
                before,
                val: el.value,
                readonly: el.readOnly,
                disabled: el.disabled
            };
        }""",
        {"fieldName": field_name, "value": value, "write": value is not None}
    )


def _visible_date_calendar(frame):
    """返回当前可见的 bee-datepicker 日历面板。"""
    calendars = frame.locator("css=.rc-calendar")
    try:
        for i in range(calendars.count() - 1, -1, -1):
            calendar = calendars.nth(i)
            if calendar.is_visible():
                return calendar
    except Exception:
        pass
    return None


def _open_date_calendar(frame, date_input):
    """打开 bee-datepicker 面板；优先使用组件原生支持的向下方向键。"""
    calendar = _visible_date_calendar(frame)
    if calendar is not None:
        return calendar

    for action in ('ArrowDown', 'click'):
        try:
            if action == 'ArrowDown':
                date_input.press('ArrowDown')
            else:
                date_input.click(
                    force=True, timeout=timing_ms('element_action_timeout')
                )
        except Exception:
            continue

        deadline = task_now() + timing_seconds('date_calendar_open_wait')
        while task_now() < deadline:
            pause()
            calendar = _visible_date_calendar(frame)
            if calendar is not None:
                return calendar
            task_sleep(timing_seconds('field_poll_interval'))

    raise Exception('已清空“单据日期”，但未能打开日期选择面板')


def _read_bill_date(frame):
    info = _access_input_by_field(frame, '单据日期')
    return info["val"] if info else ''


def _wait_for_bill_date_input(frame, timeout=None):
    """等待 NCC 新增卡片稳定渲染出“单据日期”输入框。"""
    if timeout is None:
        timeout = timing_ms('bill_date_input_wait')
    deadline = task_now() + timeout / 1000
    stable_signature = None
    stable_since = None
    while task_now() < deadline:
        pause()
        info = _access_input_by_field(frame, '单据日期')
        if info is not None:
            signature = (
                info.get("index"), info.get("label"), info.get("field")
            )
            if signature != stable_signature:
                stable_signature = signature
                stable_since = task_now()
            elif (
                stable_since is not None
                and task_now() - stable_since
                >= timing_seconds('date_field_stable_wait')
            ):
                return info
        else:
            stable_signature = None
            stable_since = None
        task_sleep(timing_seconds('field_poll_interval'))

    _dump_card_inputs(frame)
    seconds = max(1, int(timeout / 1000))
    raise BillDateInputUnavailableError(
        '单据日期',
        f'{seconds}秒内未找到“单据日期”输入框，新增卡片可能尚未完成加载'
    )


def _select_date_by_calendar_cells(frame, date_input, text):
    """日历输入提交失败时，按目标年、月、日点击面板。"""
    year, month, day = [int(part) for part in text.split('-')]
    calendar = _open_date_calendar(frame, date_input)

    year_select = calendar.locator("css=.rc-calendar-year-select").first
    year_select.wait_for(
        state='visible', timeout=timing_ms('element_action_timeout')
    )
    year_text = year_select.inner_text().strip()
    year_digits = ''.join(ch for ch in year_text if ch.isdigit())
    if not year_digits:
        raise Exception(f'无法识别日期面板年份: {year_text}')
    current_year = int(year_digits)
    year_diff = year - current_year
    if abs(year_diff) > 200:
        raise Exception(f'日期面板年份跨度异常: {current_year} → {year}')

    year_button = (
        ".rc-calendar-next-year-btn"
        if year_diff > 0 else
        ".rc-calendar-prev-year-btn"
    )
    for _ in range(abs(year_diff)):
        calendar.locator(f"css={year_button}").first.click()
        task_sleep(timing_seconds('date_year_step_wait'))

    calendar.locator("css=.rc-calendar-month-select").first.click()
    month_panel = calendar.locator("css=.rc-calendar-month-panel").first
    month_panel.wait_for(
        state='visible', timeout=timing_ms('element_action_timeout')
    )
    month_items = month_panel.locator("css=.rc-calendar-month-panel-month")
    if month_items.count() < 12:
        raise Exception(f'日期面板月份数量异常: {month_items.count()}')
    month_items.nth(month - 1).click()
    task_sleep(timing_seconds('date_month_switch_wait'))

    title = f'{year}年{month}月{day}日'
    day_cell = calendar.locator(
        f"css=td.rc-calendar-cell[title='{title}']"
        ":not(.rc-calendar-disabled-cell)"
    ).first
    if day_cell.count() == 0:
        day_cell = calendar.locator(
            "xpath=.//td[contains(@class,'rc-calendar-cell') "
            "and not(contains(@class,'rc-calendar-last-month-cell')) "
            "and not(contains(@class,'rc-calendar-next-month-btn-day')) "
            "and not(contains(@class,'rc-calendar-disabled-cell'))]"
            f"[.//*[contains(@class,'rc-calendar-date') and normalize-space(.)='{day}']]"
        ).first
    if day_cell.count() == 0:
        raise Exception(f'日期面板中未找到可选择日期: {text}')

    day_cell.click()
    task_sleep(timing_seconds('date_selection_commit_wait'))
    ok_button = calendar.locator("css=.rc-calendar-ok-btn").first
    if ok_button.count() > 0 and ok_button.is_visible():
        ok_button.click()
        task_sleep(timing_seconds('date_confirm_commit_wait'))


def fill_bill_date(frame, value):
    """第3步：清空外层日期框，在 bee-datepicker 面板内提交目标日期。"""
    text = format_bill_date(value)
    info = _wait_for_bill_date_input(frame)
    print(f'     [调试] 单据日期精确命中: {info}')

    inputs = frame.locator("css=input:not(.hidden-input)")
    date_input = inputs.nth(info["index"])
    clear_before_fill(date_input, '单据日期', force=True)

    # 外层框 keyboardInput=false，只负责显示；真正可输入的是弹窗内的 rc-calendar-input。
    calendar = _open_date_calendar(frame, date_input)
    panel_input = calendar.locator("css=input.rc-calendar-input").first
    try:
        panel_input.wait_for(
            state='visible', timeout=timing_ms('element_action_timeout')
        )
        clear_before_fill(panel_input, '日历面板日期')
        panel_input.fill(
            text, timeout=timing_ms('element_action_timeout')
        )
        task_sleep(timing_seconds('date_input_commit_wait'))
        panel_input.press('Enter')
        task_sleep(timing_seconds('date_selection_commit_wait'))
    except Exception as e:
        print(f'     [调试] 日历面板输入未提交，改用年月日点击: {e}')

    actual = _read_bill_date(frame)
    if actual != text:
        print(f'     [调试] 日历面板输入后页面值: {actual}，改用年月日点击')
        _select_date_by_calendar_cells(frame, date_input, text)
        actual = _read_bill_date(frame)

    if actual != text:
        raise Exception(f'“单据日期”写入失败：目标值={text}，页面值={actual}')
    print(f'     已填入单据日期: {text}（页面显示: {actual}）')
    pause()


def fill_bill_date_with_card_recovery(
    frame, value, card_already_open=False
):
    """
    单家公司权限模式没有结算财务组织步骤，因此单独等待日期框。
    日期框未在设定时间内出现时执行“取消→确定→新增”，最多重新加载3次。
    """
    last_error = None
    for reload_count in range(BILL_DATE_CARD_RELOAD_MAX + 1):
        try:
            if not (reload_count == 0 and card_already_open):
                click_add_bill(frame)
            return fill_bill_date(frame, value)
        except TaskCancelledError:
            raise
        except BillDateInputUnavailableError as error:
            last_error = error
            if reload_count >= BILL_DATE_CARD_RELOAD_MAX:
                break
            print(
                '     单据日期输入框'
                f'{format_timing_seconds("bill_date_input_wait")}秒内未出现，'
                f'开始第{reload_count + 1}/{BILL_DATE_CARD_RELOAD_MAX}次'
                '重新加载（取消→确定→新增）......'
            )
            if not cancel_card(frame):
                raise FieldFillError(
                    '单据日期',
                    '输入框未出现，且取消当前卡片后未能安全返回列表；'
                    '为避免误操作，本单已停止'
                ) from error

    reason = getattr(last_error, 'reason', str(last_error))
    raise FieldFillError(
        '单据日期',
        f'初次打开及{BILL_DATE_CARD_RELOAD_MAX}次重新加载后，'
        f'输入框仍未出现：{reason}'
    )


def excel_text(value):
    """把 Excel 单元格值转成适合页面输入和判空的文本。"""
    if value is None:
        return ''
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return _clean_input_text(value)


def format_expense_period(value):
    """
    把底稿 J 列所属期间转换为 NCC 要求的 YYYY-MM。
    明确支持 YYYY-MM、YYYY/MM 及 Excel 日期单元格。
    """
    import datetime as _dt
    from openpyxl.utils.datetime import from_excel

    if isinstance(value, (_dt.datetime, _dt.date)):
        return value.strftime('%Y-%m')

    text = excel_text(value)
    if text == '':
        raise FieldFillError('费用所属期', '底稿“费用所属期间”为空')
    match = re.fullmatch(r'(\d{4})\s*[-/]\s*(\d{1,2})', text)
    if match is None:
        # 保留既有兼容：完整日期、点号/中文年月及紧凑数字。
        match = re.match(r'^(\d{4})\s*[-/.年]\s*(\d{1,2})', text)
    if match is None:
        match = re.fullmatch(r'(\d{4})(\d{2})(?:\d{2})?', text)
    if match is not None:
        year = int(match.group(1))
        month = int(match.group(2))
        if not 1 <= month <= 12:
            raise FieldFillError(
                '费用所属期',
                f'底稿“费用所属期间”月份无效，当前值为“{text}”'
            )
        return f'{year:04d}-{month:02d}'

    # 只有不符合 YYYYMM/YYYYMMDD 的裸数字才按 Excel 日期序号兼容。
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            return from_excel(value).strftime('%Y-%m')
        except Exception:
            pass
    raise FieldFillError(
        '费用所属期',
        f'底稿“费用所属期间”无法识别为年月，当前值为“{text}”'
    )


def format_tax_rate(value):
    """保留文本税率；Excel 百分比数值转换为下拉框常见的百分号文本。"""
    if (
        isinstance(value, (int, float))
        and not isinstance(value, bool)
        and abs(value) <= 1
    ):
        percentage = value * 100
        number = f'{percentage:.10f}'.rstrip('0').rstrip('.')
        return f'{number}%'
    return excel_text(value)


def accounting_subject_code(value):
    """从 O 列会计科目文本开头提取连续数字编码。"""
    text = excel_text(value)
    if text == '':
        raise FieldFillError('会计科目', '底稿“会计科目”为空')
    match = re.match(r'^([0-9]+)', text)
    if match is None:
        raise FieldFillError(
            '会计科目',
            f'底稿“会计科目”开头未找到数字编码，当前值为“{text}”'
        )
    code = match.group(1)
    if code != text:
        print(f'     会计科目仅取数字编码: {text} -> {code}')
    return code


def auxiliary_accounting_code(value):
    """从 “编号/名称”中取编号，用于辅助核算弹窗的核算内容。"""
    text = excel_text(value)
    if text == '':
        return ''
    code = re.split(r'[/\\／＼]', text, maxsplit=1)[0].strip()
    if code == '':
        raise FieldFillError(
            '辅助核算',
            f'底稿“辅助核算”无法识别编号，当前值为“{text}”'
        )
    if code != text:
        print(f'     辅助核算仅取编号: {text} -> {code}')
    return code


def customer_grouping_key(row):
    """并表增加客商维度时，优先使用编码，缺少编码才使用名称。"""
    customer_code = excel_text(row["customer_code"])
    if customer_code != '':
        return 'code', customer_code
    return 'name', excel_text(row["customer_name"])


def merge_basis_label(merge_basis):
    labels = {
        'company_account_date': '同一公司、同一账号、同一日期',
        'company_account_date_customer': (
            '同一公司、同一账号、同一日期、同一客商'
        ),
    }
    try:
        return labels[merge_basis]
    except KeyError as error:
        raise FieldFillError(
            '并表依据选择', f'无法识别的选项：{merge_basis}'
        ) from error


def build_bill_groups(rows, merge_basis='company_account_date'):
    """
    根据 GUI 选择的并表依据和 AS 列生成制单分组：
    - 基础模式：E/G/I 相同且 AS 为“是”的行合并；
    - 客商模式：E/G/I/客商相同且 AS 为“是”的行合并；
    - AS 非“是”：每行单独生成一张单据。
    分组顺序按各组在底稿中的首次出现位置保留。
    """
    merge_basis_label(merge_basis)  # 启动分组前先校验 GUI 传值
    include_customer = merge_basis == 'company_account_date_customer'
    groups = []
    merged_group_indexes = {}
    for row in rows:
        if excel_text(row["merge_flag"]) == '是':
            key_parts = [
                excel_text(row["company_name"]),
                excel_text(row["bank_account"]),
                format_settlement_date(row["merge_key_i"]),
            ]
            if include_customer:
                key_parts.append(customer_grouping_key(row))
            key = tuple(key_parts)
            group_index = merged_group_indexes.get(key)
            if group_index is None:
                group_index = len(groups)
                merged_group_indexes[key] = group_index
                groups.append([])
            groups[group_index].append(row)
        else:
            groups.append([row])
    return groups


def validate_grouping_fields(row):
    """在分组前严格校验 AT 及并表键，避免错字造成残缺或错误并表。"""
    if excel_text(row["company_code"]) == '':
        return '“公司编码”为空，无法填写结算财务组织'
    merge_flag = excel_text(row["merge_flag"])
    if merge_flag not in ('是', '否'):
        return (
            '“是否并表”只能填写“是”或“否”，'
            f'当前值为“{merge_flag}”'
        )
    if merge_flag == '是':
        required_keys = (
            ('company_name', '“公司名称”为空，无法并表'),
            ('bank_account', '“银行账号”为空，无法并表'),
            ('merge_key_i', '“结算日期”为空，无法并表'),
        )
        for field_name, reason in required_keys:
            if excel_text(row[field_name]) == '':
                return reason
    settlement_date = excel_text(row["merge_key_i"])
    if settlement_date != '':
        try:
            format_settlement_date(row["merge_key_i"])
        except Exception as error:
            return str(error)
    return ''


def validate_source_row(row):
    """返回底稿行异常原因；空字符串表示校验通过。"""
    row = normalize_receipt_row(row)
    grouping_error = validate_grouping_fields(row)
    if grouping_error:
        return grouping_error
    internal_flag = excel_text(row["internal_collection_flag"])
    if internal_flag not in ('是', '否'):
        return '“代收代付情况”只能填写“是”或“否”'
    if internal_flag == '是':
        required_internal_fields = (
            (
                "represented_party",
                '“被代付账套/被代收账套名称”为空'
            ),
            (
                "represented_project",
                '“被代收方（被收回方）项目档案”为空'
            ),
            (
                "internal_collection_type",
                '“代收代付类型”为空'
            ),
            (
                "actual_charge_type",
                '“收费类型名称”为空'
            ),
        )
        for field_name, reason in required_internal_fields:
            if excel_text(row[field_name]) == '':
                return reason
    return ''


def transaction_object_of(row):
    has_customer = (
        excel_text(row["customer_code"]) != ''
        or excel_text(row["customer_name"]) != ''
    )
    return '客户' if has_customer else '人员'


def customer_reference_value(row):
    """客商优先按 M 列编码匹配；无编码时才使用 N 列名称。"""
    customer_code = excel_text(row["customer_code"])
    if customer_code != '':
        return customer_code, '“客商编码”', True
    customer_name = excel_text(row["customer_name"])
    if customer_name != '':
        return customer_name, '“客商名称”', False
    return '', '“客商编码”或“客商名称”', False


def validate_group_header(group):
    """确认同一并表组所有单据头字段一致，避免静默套用首行值。"""
    if len(group) <= 1:
        return None, ''
    first = group[0]
    specs = (
        ('“公司编码”', lambda item: excel_text(item["company_code"])),
        ('“单据日期”', lambda item: format_bill_date(item["bill_date"])),
        ('“结算方式名称”', lambda item: excel_text(item["settlement_method"])),
        ('交易对象类型', transaction_object_of),
        ('“付款银行账号”', lambda item: excel_text(item.get("payer_bank_account", ""))),
        (
            '“备注”',
            lambda item: excel_text(
                normalize_receipt_value(item.get("remark", ""), "remark")
            ),
        ),
    )
    for label, getter in specs:
        try:
            expected = getter(first)
        except Exception as error:
            return first, f'{label}无法识别：{error}'
        for row in group[1:]:
            try:
                actual = getter(row)
            except Exception as error:
                return row, f'{label}无法识别：{error}'
            if actual != expected:
                return row, (
                    f'并表数据的{label}必须一致：'
                    f'第{first["_row"]}行为“{expected}”，'
                    f'第{row["_row"]}行为“{actual}”'
                )
    return None, ''


def _field_item_info(root, field_name, require_input=False):
    """在指定区域内按字段首行标签匹配 form-item。"""
    return root.evaluate(
        """(root, args) => {
            const clean = value => String(value || '')
                .replace(/[\\s*＊\\uE000-\\uF8FF]/g, '');
            const wanted = clean(args.fieldName);
            const visible = el => !!(
                el.offsetWidth || el.offsetHeight || el.getClientRects().length
            );
            const items = Array.from(root.querySelectorAll(
                '.form-item, .card-table-modal-form-item'
            ));
            const candidates = items.map((el, index) => {
                const lines = String(el.innerText || '')
                    .split(/\\r?\\n/)
                    .map(clean)
                    .filter(Boolean);
                const first = lines[0] || '';
                const labels = Array.from(el.querySelectorAll('input, textarea'))
                    .map(input => clean(input.getAttribute('label')))
                    .filter(Boolean);
                const editableInputs = Array.from(el.querySelectorAll(
                    'input:not(.hidden-input), textarea:not(.hidden-input)'
                )).filter(input => (
                    visible(input) && !input.disabled && !input.readOnly
                )).length;
                const allText = clean(el.innerText);
                const exact = first === wanted || labels.includes(wanted);
                const wrongCustomerCode = (
                    wanted === '客户'
                    && (
                        first === '客户编码'
                        || labels.includes('客户编码')
                    )
                );
                let score = 0;
                if (first === wanted) score += 400;
                if (labels.includes(wanted)) score += 300;
                if (first.startsWith(wanted)) score += 120;
                if (labels.some(label => label.includes(wanted))) score += 100;
                if (allText.includes(wanted)) score += 30;
                if (args.requireInput && editableInputs > 0) score += 1000;
                if (wrongCustomerCode) score = 0;
                return {
                    index,
                    first,
                    text: allText.slice(0, 80),
                    score,
                    exact,
                    editableInputs,
                    visible: visible(el)
                };
            }).filter(item => (
                item.visible
                && item.score > 0
                && (!args.requireInput || item.editableInputs > 0)
            ));
            const exactCandidates = candidates.filter(item => item.exact);
            const pool = args.requireInput
                ? exactCandidates
                : (exactCandidates.length ? exactCandidates : candidates);
            const matches = pool
                .sort((a, b) => b.score - a.score || a.index - b.index);
            return matches.length ? matches[0] : null;
        }""",
        {"fieldName": field_name, "requireInput": require_input}
    )


def _exact_input_info(root, field_name):
    """优先按 input 自身的 label/attrcode 精确定位可编辑输入框。"""
    return root.evaluate(
        """(root, fieldName) => {
            const clean = value => String(value || '')
                .replace(/[\\s*＊\\uE000-\\uF8FF]/g, '');
            const wanted = clean(fieldName);
            const visible = el => !!(
                el.offsetWidth || el.offsetHeight || el.getClientRects().length
            );
            const inputs = Array.from(root.querySelectorAll(
                'input:not(.hidden-input), textarea:not(.hidden-input)'
            ));
            const matches = inputs.map((el, index) => {
                const label = clean(el.getAttribute('label'));
                const attrcode = String(el.getAttribute('attrcode') || '');
                const customerAttr = (
                    wanted === '客户'
                    && ['pk_customer', 'customer'].includes(attrcode)
                );
                const wrongCustomerCode = (
                    wanted === '客户'
                    && (
                        label === '客户编码'
                        || attrcode === 'pk_customer.code'
                        || attrcode.endsWith('.code')
                    )
                );
                let score = 0;
                if (label === wanted) score += 1000;
                if (customerAttr) score += 900;
                if (wrongCustomerCode) score = 0;
                return {
                    index,
                    label,
                    attrcode,
                    score,
                    usable: visible(el) && !el.disabled && !el.readOnly
                };
            }).filter(item => item.usable && item.score > 0)
              .sort((a, b) => b.score - a.score || a.index - b.index);
            return matches.length ? matches[0] : null;
        }""",
        field_name
    )


def _dump_region_inputs(root, field_name):
    """定位失败时打印侧拉区域输入框属性，便于精确补充选择规则。"""
    print(f'     [调试] 未找到可编辑的“{field_name}”，输入框清单如下：')
    inputs = root.locator(
        "css=input:not(.hidden-input), textarea:not(.hidden-input)"
    )
    try:
        count = min(inputs.count(), 80)
    except Exception:
        return
    for index in range(count):
        try:
            info = inputs.nth(index).evaluate(
                """el => {
                    const item = el.closest(
                        '.card-table-modal-form-item, .form-item'
                    );
                    return {
                        label: el.getAttribute('label') || '',
                        attrcode: el.getAttribute('attrcode') || '',
                        disabled: el.disabled,
                        readonly: el.readOnly,
                        visible: !!(
                            el.offsetWidth
                            || el.offsetHeight
                            || el.getClientRects().length
                        ),
                        field: item
                            ? String(item.innerText || '').slice(0, 30)
                            : ''
                    };
                }"""
            )
            print(f'     [调试] 明细输入框#{index}: {info}')
        except Exception:
            continue


def _find_field_item(
    root, field_name, timeout=None, require_input=False
):
    """返回指定区域内与字段名最匹配的可见表单项。"""
    if timeout is None:
        timeout = timing_ms('element_action_timeout')
    deadline = task_now() + timeout / 1000
    while task_now() < deadline:
        pause()
        try:
            info = _field_item_info(root, field_name, require_input)
            if info:
                items = root.locator(
                    "css=.form-item, .card-table-modal-form-item"
                )
                item = items.nth(info["index"])
                if item.is_visible():
                    return item, info
        except Exception:
            pass
        task_sleep(timing_seconds('field_poll_interval'))
    raise Exception(f'未找到“{field_name}”字段')


def _find_field_input(root, field_name, timeout=None):
    """在指定字段内取得真实可见输入框，避免同名或邻近字段误命中。"""
    if timeout is None:
        timeout = timing_ms('element_action_timeout')
    exact = _exact_input_info(root, field_name)
    if exact:
        locator_input = root.locator(
            "css=input:not(.hidden-input), textarea:not(.hidden-input)"
        ).nth(exact["index"])
        print(
            f'     [调试] {field_name}输入框精确命中: '
            f'input.label={exact["label"]}，attrcode={exact["attrcode"]}'
        )
        return locator_input

    try:
        item, info = _find_field_item(
            root, field_name, timeout, require_input=True
        )
    except Exception:
        _dump_region_inputs(root, field_name)
        raise Exception(f'未找到可编辑的“{field_name}”输入框')

    inputs = item.locator(
        "css=input:not(.hidden-input), textarea:not(.hidden-input)"
    )
    for index in range(inputs.count()):
        locator_input = inputs.nth(index)
        try:
            if locator_input.is_visible() and locator_input.is_editable():
                label = locator_input.get_attribute('label') or ''
                attrcode = locator_input.get_attribute('attrcode') or ''
                print(
                    f'     [调试] {field_name}字段命中: '
                    f'标签={info["first"]}，区域索引={info["index"]}，'
                    f'可编辑输入框={info["editableInputs"]}，'
                    f'input.label={label}，attrcode={attrcode}'
                )
                return locator_input
        except Exception:
            continue
    raise Exception(f'已找到“{field_name}”字段，但其中没有可见输入框')


def _find_visible_field_input_state(
    root, field_name, timeout=None
):
    """取得字段内的可见输入框及编辑状态，允许命中被禁用的联动字段。"""
    if timeout is None:
        timeout = timing_ms('element_action_timeout')
    item, info = _find_field_item(
        root, field_name, timeout, require_input=False
    )
    inputs = item.locator(
        "css=input:not(.hidden-input), textarea:not(.hidden-input)"
    )
    for index in range(inputs.count()):
        locator_input = inputs.nth(index)
        try:
            if not locator_input.is_visible():
                continue
            state = locator_input.evaluate(
                """el => ({
                    disabled: !!el.disabled,
                    readonly: !!el.readOnly,
                    value: String(el.value || ''),
                    label: el.getAttribute('label') || '',
                    attrcode: el.getAttribute('attrcode') || ''
                })"""
            )
            state["editable"] = not (
                state["disabled"] or state["readonly"]
            )
            print(
                f'     [调试] {field_name}字段状态: '
                f'标签={info["first"]}，区域索引={info["index"]}，'
                f'可编辑={state["editable"]}，'
                f'禁用={state["disabled"]}，只读={state["readonly"]}'
            )
            return locator_input, state
        except Exception:
            continue
    raise Exception(f'已找到“{field_name}”字段，但其中没有可见输入框')


def _normalize_reference_text(value):
    return re.sub(r'\s+', '', normalize_receipt_value(excel_text(value), '__reference__')).casefold()


def _reference_value_parts(value):
    """把参照文本拆成完整列值/token，禁止用无边界子串冒充命中。"""
    raw = excel_text(value).strip()
    if raw == '':
        return []
    parts = re.split(
        r'[\s/\\／＼|｜,，;；:：=＝()（）\[\]【】]+', raw
    )
    normalized = []
    label_prefixes = (
        '公司名称', '客商名称', '客商编码', '组织编码',
        '编码', '编号', '名称',
    )
    for part in parts:
        compact = _normalize_reference_text(part)
        if compact:
            normalized.append(compact)
        for prefix in label_prefixes:
            prefix_norm = _normalize_reference_text(prefix)
            if compact.startswith(prefix_norm) and compact != prefix_norm:
                normalized.append(compact[len(prefix_norm):])
                break
    return list(dict.fromkeys(normalized))


def _reference_text_matches(expected, candidate):
    """按完整文本、完整列值或带边界编码匹配参照候选。"""
    expected_raw = excel_text(expected).strip()
    candidate_raw = excel_text(candidate).strip()
    expected_norm = _normalize_reference_text(expected_raw)
    candidate_norm = _normalize_reference_text(candidate_raw)
    if not expected_norm or not candidate_norm:
        return False
    if expected_norm == candidate_norm:
        return True

    expected_parts = _reference_value_parts(expected_raw)
    candidate_parts = set(_reference_value_parts(candidate_raw))
    if expected_parts and all(part in candidate_parts for part in expected_parts):
        return True

    # 编码与中文名称有时在 NCC 中紧挨显示，如“1001甲公司”。
    compact_match = re.fullmatch(
        r'([A-Za-z0-9_.-]+)([\u3400-\u9fff].*)',
        re.sub(r'\s+', '', candidate_raw)
    )
    if compact_match and expected_norm in {
        _normalize_reference_text(compact_match.group(1)),
        _normalize_reference_text(compact_match.group(2)),
    }:
        return True

    # 纯代码允许与中文相邻，但不得命中更长代码（1001 不能命中 10010）。
    if re.fullmatch(r'[A-Za-z0-9_.-]+', expected_raw):
        return bool(re.search(
            rf'(?<![A-Za-z0-9_.-]){re.escape(expected_raw)}'
            rf'(?![A-Za-z0-9_.-])',
            candidate_raw,
            flags=re.IGNORECASE,
        ))
    return False


def _customer_name_from_reference_option(value):
    """从 NCC 客商候选“编号/名称”中只取名称部分。"""
    raw = excel_text(value).strip()
    if raw == '':
        return ''
    parts = re.split(r'[/／]', raw, maxsplit=1)
    return (parts[1] if len(parts) == 2 else parts[0]).strip()


def _customer_name_option_matches(expected_name, candidate):
    """名称查询只比较候选斜杠后的客商名称，不比较编号。"""
    expected_norm = _normalize_reference_text(expected_name)
    candidate_name = _customer_name_from_reference_option(candidate)
    return bool(
        expected_norm
        and expected_norm == _normalize_reference_text(candidate_name)
    )


def _reference_selection_matches(expected, actual, selected_text=''):
    """核验页面回显值确实来自唯一命中的目标候选。"""
    if not excel_text(actual).strip() or not excel_text(selected_text).strip():
        return False
    if not _reference_text_matches(expected, selected_text):
        return False
    return (
        _reference_text_matches(expected, actual)
        or _reference_text_matches(actual, selected_text)
    )


def _click_matching_reference_option(
    frame, field_name, expected, timeout=None,
    dismiss_input=None, allow_multiple_matches=False,
    customer_name_only=False
):
    """
    点击目标匹配项。默认要求唯一；客商编码模式允许多个匹配项并选第一个。
    """
    if timeout is None:
        timeout = timing_ms('reference_wait')
    expected_text = excel_text(expected)
    expected_norm = _normalize_reference_text(expected_text)
    if expected_norm == '':
        raise FieldFillError(field_name, '用于匹配下拉项的目标值为空')
    selectors = [
        "css=div.u-dropdown.nc-dropdown.refer-container "
        "ul.refer-cascading-item li",
        "css=ul.refer-cascading-item li",
        "css=div.refer-container tr.u-table-row",
        "css=div.refer-table tr.u-table-row",
        # 辅助核算弹窗中的候选有时不使用上述级联列表结构，而是挂在
        # 普通 dropdown、select dropdown 或弹窗内的表格中。
        "css=div.u-dropdown [role='option']",
        "css=div.u-dropdown li",
        "css=div.u-select-dropdown li",
        "css=div.u-modal div.refer-container li",
        "css=div[role='dialog'] div.refer-container li",
        "css=div.u-modal tr.u-table-row",
        "css=div[role='dialog'] tr.u-table-row",
    ]
    ignored_text = ('历史记录', '没有更多', '暂无数据', '无数据')
    loading_text = ('加载中', '正在加载', '请稍候')
    deadline = task_now() + timeout / 1000
    ambiguous_options = []
    last_visible_options = []
    stable_signature = None
    stable_since = None
    while task_now() < deadline:
        pause()
        visible_options = None
        loading_visible = False
        for selector in selectors:
            options = frame.locator(selector)
            try:
                count = min(options.count(), 200)
            except Exception:
                continue
            current_options = []
            current_loading = False
            for index in range(count):
                option = options.nth(index)
                try:
                    if not option.is_visible():
                        continue
                    classes = option.get_attribute('class') or ''
                    option_text = (
                        option.get_attribute('title')
                        or option.inner_text(
                            timeout=timing_ms('tiny_element_timeout')
                        )
                        or ''
                    ).strip()
                    if any(word in option_text for word in loading_text):
                        current_loading = True
                        continue
                    if 'disabled' in classes:
                        continue
                    if (
                        not option_text
                        or any(word in option_text for word in ignored_text)
                    ):
                        continue
                    current_options.append((option, option_text))
                except Exception:
                    continue
            if current_options or current_loading:
                visible_options = current_options
                loading_visible = current_loading
                break

        if visible_options is None or loading_visible:
            stable_signature = None
            stable_since = None
            task_sleep(timing_seconds('candidate_poll_interval'))
            continue

        last_visible_options = [text for _, text in visible_options]
        if customer_name_only:
            matches = [
                item for item in visible_options
                if _customer_name_option_matches(expected_text, item[1])
            ]
        else:
            matches = [
                item for item in visible_options
                if _reference_text_matches(expected_text, item[1])
            ]
        if len(matches) > 1:
            ambiguous_options = [text for _, text in matches]

        signature = (
            tuple(sorted(
                _normalize_reference_text(text)
                for _, text in visible_options
            )),
            tuple(sorted(
                _normalize_reference_text(text)
                for _, text in matches
            )),
        )
        now = task_now()
        if signature != stable_signature:
            stable_signature = signature
            stable_since = now
            task_sleep(timing_seconds('candidate_poll_interval'))
            continue
        if (
            stable_since is None
            or now - stable_since
            < timing_seconds('reference_option_stable_wait')
        ):
            task_sleep(timing_seconds('candidate_poll_interval'))
            continue

        if len(matches) > 1 and not allow_multiple_matches:
            preview = '；'.join(ambiguous_options[:10])
            match_description = (
                f'按“编号/名称”中的名称部分发现多个同名客商'
                f'（目标名称：{expected_text}）'
                if customer_name_only
                else f'下拉数据匹配不唯一（目标：{expected_text}）'
            )
            raise ReferenceMatchAmbiguousError(
                field_name,
                f'{match_description}：{preview}'
            )
        if matches:
            option, option_text = matches[0]
            pause()
            option.click(timeout=timing_ms('element_action_timeout'))
            if dismiss_input is not None:
                try:
                    dismiss_input.press(
                        'Escape', timeout=timing_ms('tiny_element_timeout')
                    )
                    dismiss_input.press(
                        'Tab', timeout=timing_ms('tiny_element_timeout')
                    )
                    task_sleep(timing_seconds('short_stable_wait'))
                except Exception:
                    pass
            if len(matches) > 1:
                print(
                    f'     {field_name}按编码匹配到{len(matches)}项，'
                    f'已选择第一项: {option_text}'
                )
            else:
                print(f'     {field_name}已选择唯一匹配项: {option_text}')
            return option_text
        task_sleep(timing_seconds('candidate_poll_interval'))
    seconds = max(1, int(timeout / 1000))
    if ambiguous_options and not allow_multiple_matches:
        preview = '；'.join(ambiguous_options[:10])
        match_description = (
            f'按“编号/名称”中的名称部分发现多个同名客商'
            f'（目标名称：{expected_text}）'
            if customer_name_only
            else f'下拉数据匹配不唯一（目标：{expected_text}）'
        )
        raise ReferenceMatchAmbiguousError(
            field_name,
            f'{match_description}：{preview}'
        )
    preview = '；'.join(last_visible_options[:10])
    option_note = f'；当前可见候选：{preview}' if preview else ''
    raise FieldFillError(
        field_name,
        f'等待{seconds}秒后未找到与“{expected_text}”匹配的下拉数据{option_note}'
    )


def _clear_reference_before_fill(locator_input, field_name, force=False):
    """优先点击参照组件清除图标，确保显示值和内部 refpk 同时清空。"""
    scope = locator_input
    cleared_by_icon = False
    clear_selectors = (
        ".uf-close-c, .uf-close, .refer-input-clear, "
        ".input-clear, .clear-icon, [title='清空']"
    )
    for _ in range(3):
        try:
            scope = scope.locator("xpath=..")
            buttons = scope.locator(f"css={clear_selectors}")
            for index in range(buttons.count()):
                button = buttons.nth(index)
                if button.is_visible():
                    pause()
                    button.click(
                        force=True,
                        timeout=timing_ms('short_element_timeout')
                    )
                    cleared_by_icon = True
                    task_sleep(timing_seconds('short_stable_wait'))
                    break
            if cleared_by_icon:
                break
        except Exception:
            continue
    if cleared_by_icon:
        print(f'     [前置] 已通过参照清除按钮重置{field_name}')
    clear_before_fill(locator_input, field_name, force=force)


def _fill_reference_field(
    frame, root, field_name, value, source_name, previous_value=None,
    retry_ambiguous=True, timeout=None,
    allow_multiple_matches=False, customer_name_only=False
):
    """完整执行“清空—输入—匹配”，失败时总共尝试三次。"""
    if timeout is None:
        timeout = timing_ms('reference_wait')
    text = normalize_receipt_value(excel_text(value), field_name)
    if text == '':
        raise FieldFillError(field_name, f'底稿{source_name}为空')

    try:
        locator_input = _find_field_input(root, field_name)
        inherited_value = locator_input.input_value(
            timeout=timing_ms('short_element_timeout')
        ).strip()
    except FieldFillError:
        raise
    except Exception as e:
        raise FieldFillError(field_name, e) from e
    if (
        previous_value is not None
        and excel_text(previous_value) == text
        and _reference_selection_matches(text, inherited_value, inherited_value)
    ):
        print(
            f'     {field_name}与上一条明细相同，'
            f'复用新增行继承值: {inherited_value}'
        )
        pause()
        return

    last_error = None
    for attempt in range(1, RETRY_MAX + 1):
        pause()
        try:
            locator_input = _find_field_input(root, field_name)
            _clear_reference_before_fill(locator_input, field_name)
            try:
                locator_input.click(
                    timeout=timing_ms('element_action_timeout')
                )
                locator_input.type(
                    text,
                    delay=max(
                        0, int(round(
                            timing_seconds('reference_type_interval') * 1000
                        ))
                    ),
                    timeout=timing_ms('reference_input_timeout')
                )
            except Exception:
                locator_input.fill(
                    text, timeout=timing_ms('element_action_timeout')
                )
            seconds = max(1, int(timeout / 1000))
            print(
                f'     第{attempt}/{RETRY_MAX}次已填入{field_name}: {text}，'
                f'等待下拉匹配项（最长{seconds}秒）......'
            )
            _click_matching_reference_option(
                frame, field_name, expected=text, timeout=timeout,
                allow_multiple_matches=allow_multiple_matches,
                customer_name_only=customer_name_only
            )
            task_sleep(timing_seconds('reference_post_select_wait'))
            pause()
            return
        except TaskCancelledError:
            raise
        except ReferenceMatchAmbiguousError as error:
            if not retry_ambiguous:
                raise
            last_error = error
        except Exception as error:
            last_error = error

        if attempt < RETRY_MAX:
            try:
                locator_input.press(
                    'Escape', timeout=timing_ms('tiny_element_timeout')
                )
            except Exception:
                pass
            print(
                f'     {field_name}第{attempt}/{RETRY_MAX}次输入或匹配失败，'
                f'将清空后重新填入：{last_error}'
            )
            task_sleep(timing_seconds('reference_retry_interval'))

    reason = getattr(last_error, 'reason', str(last_error))
    raise FieldFillError(
        field_name,
        f'连续{RETRY_MAX}次清空并重新填入底稿值“{text}”后'
        f'仍未匹配到下拉数据：{reason}'
    )


def _fill_reference_field_if_editable(
    frame, root, field_name, value, source_name, previous_value=None
):
    """填写由 NCC 联动控制的参照字段；页面禁用时按业务规则跳过。"""
    try:
        _, state = _find_visible_field_input_state(root, field_name)
    except Exception as e:
        raise FieldFillError(field_name, e) from e
    if not state["editable"]:
        reason = '禁用' if state["disabled"] else '只读'
        page_value = state["value"].strip()
        if not _reference_selection_matches(value, page_value, page_value):
            raise FieldFillError(
                field_name,
                f'字段当前为{reason}状态，且NCC联动值与底稿不一致：'
                f'底稿={excel_text(value)}，页面={page_value or "空"}'
            )
        print(
            f'     {field_name}当前为{reason}状态，'
            f'已核验NCC联动值: {page_value}'
        )
        pause()
        return
    _fill_reference_field(
        frame, root, field_name, value, source_name,
        previous_value=previous_value
    )


def _clear_optional_reference_field(
    root, field_name, source_name, preserve_linked_readonly=False
):
    """底稿可选字段为空时，确认 NCC 没有保留新增行继承值。"""
    try:
        locator_input, state = _find_visible_field_input_state(
            root, field_name, timeout=timing_ms('tiny_element_timeout')
        )
    except Exception:
        print(f'     底稿{source_name}为空，页面未渲染{field_name}，无需清除')
        return
    current = state["value"].strip()
    if current == '':
        print(f'     底稿{source_name}为空，{field_name}页面值也为空')
        return
    if not state["editable"]:
        if preserve_linked_readonly:
            reason = '禁用' if state["disabled"] else '只读'
            print(
                f'     底稿{source_name}为空，但{field_name}为{reason}状态，'
                f'保留NCC自动联动值: {current}'
            )
            return
        raise FieldFillError(
            field_name,
            f'底稿{source_name}为空，但页面继承值“{current}”不可编辑，无法安全清除'
        )
    _clear_reference_before_fill(locator_input, field_name)
    _, final_state = _find_visible_field_input_state(
        root, field_name, timeout=timing_ms('short_element_timeout')
    )
    if final_state["value"].strip() != '':
        raise FieldFillError(
            field_name,
            f'底稿{source_name}为空，但页面继承值清除失败：{final_state["value"]}'
        )
    print(f'     底稿{source_name}为空，已确认清除新增行继承的{field_name}')


def fill_split_income_type_after_income_item(
    frame, drawer, row, previous_row=None
):
    """新收支项目联动后处理分灶收入类型；禁用即视为 NCC 自动取值。"""
    task_sleep(timing_seconds('reference_post_select_wait'))
    try:
        _, state = _find_visible_field_input_state(
            drawer, '分灶收入类型',
            timeout=timing_ms('element_action_timeout')
        )
    except Exception as error:
        raise FieldFillError('分灶收入类型', error) from error

    if state["disabled"]:
        page_value = state["value"].strip()
        print(
            '     分灶收入类型已被NCC禁用，表明系统已自动取值，'
            f'跳过底稿填写（页面值：{page_value or "空"}）'
        )
        pause()
        return

    split_income_type = excel_text(row["split_income_type"])
    if split_income_type != '':
        _fill_reference_field_if_editable(
            frame, drawer, '分灶收入类型', split_income_type,
            '“分灶收入类型”',
            previous_value=(
                previous_row["split_income_type"]
                if previous_row is not None else None
            )
        )
    else:
        _clear_optional_reference_field(
            drawer, '分灶收入类型', '“分灶收入类型”'
        )


def _find_visible_auxiliary_modal_once(frame):
    """单次查找当前真正可见的“辅助核算”弹窗。"""
    selectors = (
        "css=div.u-modal",
        "css=div[role='dialog']",
        "xpath=//div[contains(@class,'modal') "
        "and not(contains(@class,'card-table-modal'))]",
    )
    for selector in selectors:
        modals = frame.locator(selector)
        try:
            count = modals.count()
        except Exception:
            continue
        for index in range(count - 1, -1, -1):
            modal = modals.nth(index)
            try:
                if (
                    modal.is_visible()
                    and '辅助核算' in modal.inner_text(
                        timeout=timing_ms('tiny_element_timeout')
                    )
                ):
                    return modal
            except Exception:
                continue
    return None


def _visible_auxiliary_modal(frame, timeout=None):
    """等待并返回当前可见的“辅助核算”独立弹窗。"""
    if timeout is None:
        timeout = timing_ms('auxiliary_modal_wait')
    deadline = task_now() + timeout / 1000
    while task_now() < deadline:
        pause()
        modal = _find_visible_auxiliary_modal_once(frame)
        if modal is not None:
            return modal
        task_sleep(timing_seconds('field_poll_interval'))
    raise FieldFillError(
        '辅助核算',
        f'点击后{timeout / 1000:g}秒内未出现“辅助核算”窗口'
    )


def _open_auxiliary_modal(frame, drawer):
    """点击明细侧栏的“辅助核算”，并等待其独立弹窗。"""
    try:
        item, _ = _find_field_item(
            drawer, '辅助核算', require_input=False
        )
    except Exception as e:
        raise FieldFillError('辅助核算', e) from e

    # NCC 将辅助核算表格内容挂载在弹窗标题之外。打开前标记当前
    # 已显示的参照框，打开后即可精确识别新出现的“核算内容”输入框。
    frame.locator('body').evaluate(
        """() => {
            document.querySelectorAll('[data-codex-aux-before]')
                .forEach(el => el.removeAttribute('data-codex-aux-before'));
            document.querySelectorAll('input.refer-input:not(.hidden-input)')
                .forEach(el => {
                    const visible = !!(
                        el.offsetWidth || el.offsetHeight || el.getClientRects().length
                    );
                    if (visible) el.setAttribute('data-codex-aux-before', '1');
                });
        }"""
    )

    targets = (
        item.locator("css=input:not(.hidden-input)").first,
        item.locator("css=.icon-refer, span[class*='icon-refer']").first,
        item,
    )
    last_error = None
    for target in targets:
        try:
            if target.count() == 0 or not target.is_visible():
                continue
            _dispatch_auxiliary_dom_click(target)
            return _visible_auxiliary_modal(
                frame, timeout=timing_ms('auxiliary_modal_wait')
            )
        except Exception as e:
            last_error = e
            continue
    raise FieldFillError(
        '辅助核算',
        f'无法点击“辅助核算”打开窗口：{last_error or "未找到可点击控件"}'
    )


def _dump_auxiliary_modal_inputs(frame, modal):
    """辅助核算窗口定位失败时输出输入控件属性，便于兼容 NCC 页面差异。"""
    print('     [调试] 辅助核算窗口输入控件清单：')
    for scope_name, scope in (('窗口', modal), ('页面', frame)):
        inputs = scope.locator("css=input, textarea")
        try:
            count = min(inputs.count(), 200)
        except Exception:
            continue
        for index in range(count):
            locator_input = inputs.nth(index)
            try:
                info = locator_input.evaluate(
                    """el => ({
                        tag: el.tagName,
                        type: el.type || '',
                        label: el.getAttribute('label') || '',
                        placeholder: el.getAttribute('placeholder') || '',
                        cls: el.className || '',
                        disabled: !!el.disabled,
                        readonly: !!el.readOnly,
                        visible: !!(el.offsetWidth || el.offsetHeight),
                        value: String(el.value || '')
                    })"""
                )
                print(
                    f'     [调试] 辅助核算{scope_name}输入框'
                    f'#{index}: {info}'
                )
            except Exception:
                continue


def _find_auxiliary_content_input(frame, modal):
    """取得打开辅助核算弹窗后新出现的可编辑参照输入框。"""
    base_selectors = (
        "css=input.nc-input.refer-input:not(.hidden-input)",
        "css=input.refer-input:not(.hidden-input)",
    )
    # 先限定在“辅助核算”弹窗内，避免页面重渲染后全局 nth 定位漂移；
    # 只有弹窗结构不标准时，才回退到打开弹窗后新出现的全局输入框。
    search_scopes = (
        ('弹窗内', modal, False),
        ('页面新出现', frame, True),
    )
    for scope_name, scope, only_new in search_scopes:
        for base_selector in base_selectors:
            selector = base_selector
            if only_new:
                selector += ":not([data-codex-aux-before])"
            candidates = scope.locator(selector)
            try:
                count = candidates.count()
            except Exception:
                continue
            for index in range(count - 1, -1, -1):
                candidate = candidates.nth(index)
                try:
                    info = candidate.evaluate(
                        """el => ({
                            type: String(el.type || '').toLowerCase(),
                            cls: el.className || '',
                            disabled: !!el.disabled,
                            readonly: !!el.readOnly,
                            visible: !!(
                                el.offsetWidth || el.offsetHeight
                                || el.getClientRects().length
                            ),
                            value: String(el.value || '')
                        })"""
                    )
                    if (
                        info['visible'] and not info['disabled']
                        and not info['readonly']
                        and info['type'] not in (
                            'checkbox', 'radio', 'hidden'
                        )
                    ):
                        if scope_name == '弹窗内':
                            hit_message = '辅助核算弹窗内输入框命中'
                        else:
                            hit_message = '辅助核算新出现输入框命中'
                        print(f'     [调试] {hit_message}: {info}')
                        return candidate
                except Exception:
                    continue

    # 极少数版本会重建整张明细表，导致旧输入框的标记一并丢失。
    # 弹窗内容在 DOM 末尾，取最后一个可见、可编辑参照框作为兜底。
    candidates = frame.locator(
        "css=input.nc-input.refer-input:not(.hidden-input)"
    )
    for index in range(candidates.count() - 1, -1, -1):
        candidate = candidates.nth(index)
        try:
            if candidate.is_visible() and candidate.is_enabled():
                input_type = (candidate.get_attribute('type') or 'text').lower()
                if input_type not in ('checkbox', 'radio', 'hidden'):
                    info = candidate.evaluate(
                        "el => ({cls: el.className || '', value: el.value || ''})"
                    )
                    print(f'     [调试] 辅助核算末尾参照输入框命中: {info}')
                    return candidate
        except Exception:
            continue

    _dump_auxiliary_modal_inputs(frame, modal)
    raise FieldFillError('辅助核算', '弹窗打开后未检测到新出现的核算内容输入框')


def _mark_auxiliary_reference_popup(option):
    """标记本次辅助核算候选所在弹层，避免误判弹窗内的固定表格。"""
    try:
        option.evaluate(
            """element => {
                document.querySelectorAll('[data-codex-aux-popup]')
                    .forEach(node => node.removeAttribute('data-codex-aux-popup'));
                let node = element;
                while (node && node !== document.body) {
                    if (node.matches(
                        'div.u-dropdown, div.refer-container, '
                        + 'div.refer-table, ul.refer-cascading-item'
                    )) {
                        node.setAttribute('data-codex-aux-popup', '1');
                        return;
                    }
                    node = node.parentElement;
                }
                element.setAttribute('data-codex-aux-popup', '1');
            }"""
        )
    except Exception:
        # 标记失败不阻止候选点击；后续会把元素重建视为弹层已收起。
        pass


def _settle_auxiliary_reference_selection(frame):
    """等待本次点击的候选弹层收起，不误读辅助核算弹窗内固定表格。"""
    hidden_since = None
    visible_since = None
    deadline = (
        task_now() + timing_seconds('auxiliary_selection_settle_wait')
    )
    while task_now() < deadline:
        pause()
        marked_popups = frame.locator('[data-codex-aux-popup="1"]')
        popup_visible = False
        try:
            for index in range(marked_popups.count() - 1, -1, -1):
                if marked_popups.nth(index).is_visible():
                    popup_visible = True
                    break
        except Exception:
            # NCC 候选点击后经常直接重建DOM，找不到标记即等同于已收起。
            popup_visible = False
        if not popup_visible:
            visible_since = None
            if hidden_since is None:
                hidden_since = task_now()
            elif task_now() - hidden_since >= 1.0:
                return
        else:
            hidden_since = None
            if visible_since is None:
                visible_since = task_now()
            elif (
                task_now() - visible_since
                >= timing_seconds('auxiliary_popup_dismiss_wait')
            ):
                # 部分 NCC 版本选中后仍保留候选表格DOM。此时继续等待只会
                # 误判失败；先尝试收起，再交给后续“确定”前的稳定等待处理。
                try:
                    marked_popups.nth(0).press(
                        'Escape', timeout=timing_ms('tiny_element_timeout')
                    )
                except Exception:
                    try:
                        frame.locator('body').press(
                            'Escape', timeout=timing_ms('tiny_element_timeout')
                        )
                    except Exception:
                        pass
                task_sleep(timing_seconds('reference_post_select_wait'))
                print(
                    '     辅助核算候选层持续显示，已尝试收起并继续确认'
                )
                return
        task_sleep(timing_seconds('field_poll_interval'))
    raise FieldFillError(
        '辅助核算',
        '候选点击后本次参照下拉框仍未收起，可能尚未完成选择；'
        '为避免触发NCC确认错误，本次未点击“确定”'
    )


def _dispatch_auxiliary_dom_click(locator):
    """在页面内部触发一次完整点击，不依赖最小化窗口的鼠标命中测试。"""
    locator.evaluate(
        """element => {
            element.scrollIntoView({block: 'nearest', inline: 'nearest'});
            try {
                element.focus({preventScroll: true});
            } catch (_) {
                element.focus();
            }
            const common = {
                bubbles: true,
                cancelable: true,
                composed: true,
                button: 0,
            };
            const PointerCtor = window.PointerEvent || window.MouseEvent;
            element.dispatchEvent(new PointerCtor('pointerdown', {
                ...common, buttons: 1, pointerId: 1, pointerType: 'mouse',
                isPrimary: true
            }));
            element.dispatchEvent(new MouseEvent('mousedown', {
                ...common, buttons: 1
            }));
            element.dispatchEvent(new PointerCtor('pointerup', {
                ...common, buttons: 0, pointerId: 1, pointerType: 'mouse',
                isPrimary: true
            }));
            element.dispatchEvent(new MouseEvent('mouseup', {
                ...common, buttons: 0
            }));
            element.click();
        }"""
    )


def _click_auxiliary_reference_option(frame, modal, expected, timeout):
    """在辅助核算窗口及其页面下拉层中选择核算内容候选。

    NCC 会把候选层挂在弹窗内或 body 末尾，两种结构都可能出现；这里不
    复用输入框清空逻辑，候选识别失败时保留已输入的编号后继续重试。
    """
    expected_text = excel_text(expected)
    selectors = (
        # NCC 辅助核算弹层中的候选本身是带 title 的 refer-li，
        # 其 title 即“编号/名称”（例如 08/能源转售收入）。先精确命中该节点，
        # 不去点击“已选择”标题、历史记录等同层内容。
        "css=div.u-dropdown.nc-dropdown.refer-container "
        "ul.refer-cascading-item li.refer-li[title]",
        "css=div.refer-cascading-list ul.refer-cascading-item "
        "li.refer-li[title]",
        "css=div.u-dropdown.nc-dropdown.refer-container "
        "ul.refer-cascading-item li",
        "css=ul.refer-cascading-item li",
        "css=div.refer-container tr.u-table-row",
        "css=div.refer-table tr.u-table-row",
        "css=div.u-dropdown [role='option']",
        "css=div.u-dropdown li",
        "css=div.u-select-dropdown li",
        "css=div.u-modal div.refer-container li",
        "css=div.u-modal div.refer-container tr.u-table-row",
        "css=div[role='dialog'] div.refer-container li",
        "css=div[role='dialog'] div.refer-container tr.u-table-row",
    )
    ignored_text = ('历史记录', '没有更多', '暂无数据', '无数据')
    loading_text = ('加载中', '正在加载', '请稍候')
    deadline = task_now() + timeout / 1000
    last_visible_options = []
    while task_now() < deadline:
        pause()
        # 优先查辅助核算弹窗本身；候选被 portal 到页面时再查全页。
        for scope in (modal, frame):
            for selector in selectors:
                options = scope.locator(selector)
                try:
                    count = min(options.count(), 200)
                except Exception:
                    continue
                for index in range(count):
                    option = options.nth(index)
                    try:
                        if not option.is_visible():
                            continue
                        classes = option.get_attribute('class') or ''
                        option_text = (
                            option.get_attribute('title')
                            or option.inner_text(
                                timeout=timing_ms('tiny_element_timeout')
                            )
                            or ''
                        ).strip()
                        if (
                            not option_text
                            or 'disabled' in classes
                            or any(word in option_text for word in ignored_text)
                            or any(word in option_text for word in loading_text)
                        ):
                            continue
                        last_visible_options.append(option_text)
                        if not _reference_text_matches(expected_text, option_text):
                            continue
                        print(
                            '     辅助核算候选已出现，等待'
                            f'{format_timing_seconds("auxiliary_option_ready_wait")}秒后点击......'
                        )
                        task_sleep(timing_seconds('auxiliary_option_ready_wait'))
                        # Playwright 的普通 click 需要对最小化窗口执行坐标命中测试，
                        # NCC 后台渲染稍慢时会出现“调用成功但选项未提交”。页面内
                        # 事件直接交给 React 组件处理，不依赖窗口是否处于前台。
                        _dispatch_auxiliary_dom_click(option)
                        print(
                            '     辅助核算已选择匹配候选项: '
                            f'{option_text}'
                        )
                        return option_text
                    except TaskCancelledError:
                        raise
                    except Exception:
                        continue
        task_sleep(timing_seconds('candidate_poll_interval'))
    seconds = max(1, int(timeout / 1000))
    preview = '；'.join(list(dict.fromkeys(last_visible_options))[:10])
    option_note = f'；当前可见候选：{preview}' if preview else ''
    raise FieldFillError(
        '辅助核算',
        f'等待{seconds}秒后未找到与“{expected_text}”匹配的辅助核算候选'
        f'{option_note}'
    )


def _confirm_auxiliary_modal(frame, modal):
    """等待参照选择稳定后确认辅助核算，只判断窗口是否真正关闭。"""
    # 候选被点击后 NCC 会异步更新受控组件；固定等待后再点“确定”。
    task_sleep(timing_seconds('auxiliary_selection_commit_wait'))
    pause()
    current_modal = _find_visible_auxiliary_modal_once(frame)
    if current_modal is None:
        raise FieldFillError('辅助核算', '点击“确定”前辅助核算窗口已消失')

    button_selectors = (
        "xpath=.//*[contains(@class,'modal-footer')]"
        "//button[starts-with(normalize-space(.),'确定') "
        "or normalize-space(.)='确认']",
        "xpath=.//button[starts-with(normalize-space(.),'确定') "
        "or normalize-space(.)='确认']",
    )
    confirm_button = None
    for selector in button_selectors:
        buttons = current_modal.locator(selector)
        try:
            count = buttons.count()
        except Exception:
            continue
        for index in range(count - 1, -1, -1):
            button = buttons.nth(index)
            try:
                if button.is_visible() and button.is_enabled():
                    confirm_button = button
                    break
            except Exception:
                continue
        if confirm_button is not None:
            break
    if confirm_button is None:
        raise FieldFillError('辅助核算', '未找到可点击的“确定”按钮')

    pause()
    try:
        _dispatch_auxiliary_dom_click(confirm_button)
    except Exception as error:
        raise FieldFillError(
            '辅助核算', f'点击“确定”失败：{error}'
        ) from error
    print(
        '     已点击辅助核算“确定”，等待'
        f'{format_timing_seconds("auxiliary_confirm_commit_wait")}秒后继续......'
    )
    task_sleep(timing_seconds('auxiliary_confirm_commit_wait'))

    # 旧 modal locator 可能因 React 重渲染而漂移；每轮重新按标题查找。
    # 填写阶段不读取 NCC 业务提示，业务报错只在点击保存/提交后判定。
    hidden_since = None
    deadline = task_now() + timing_seconds('auxiliary_modal_wait')
    while task_now() < deadline:
        pause()
        if _find_visible_auxiliary_modal_once(frame) is None:
            if hidden_since is None:
                hidden_since = task_now()
            elif (
                task_now() - hidden_since
                >= timing_seconds('auxiliary_modal_close_stable_wait')
            ):
                print('     辅助核算窗口已确认')
                return
        else:
            hidden_since = None
        task_sleep(timing_seconds('field_poll_interval'))

    raise FieldFillError(
        '辅助核算',
        '点击“确定”后'
        f'{format_timing_seconds("auxiliary_modal_wait")}秒内窗口未关闭，'
        '页面可能仍在处理'
    )


def _fill_auxiliary_code_input(frame, modal, code):
    """把辅助核算编号真正写入 NCC 受控输入框，并确认不会立即回滚为空。"""
    last_error = None
    for attempt in range(1, RETRY_MAX + 1):
        pause()
        try:
            # NCC 可能在搜索时重建输入框，因此每次重试都重新定位当前控件。
            locator_input = _find_auxiliary_content_input(frame, modal)
            # 仅首次输入前清空。若后续候选识别失败，绝不能再次清掉已输入
            # 的编号，否则会把本程序的重试误判成 NCC 自动清空。
            if attempt == 1:
                _clear_reference_before_fill(
                    locator_input, '辅助核算核算内容', force=True
                )
            locator_input.wait_for(
                state='visible', timeout=timing_ms('element_action_timeout')
            )
            _dispatch_auxiliary_dom_click(locator_input)

            if attempt == 1:
                # 优先走真实键盘事件，最符合 NCC 参照组件的 onChange 逻辑。
                locator_input.type(
                    code,
                    delay=max(
                        0, int(round(
                            timing_seconds('auxiliary_type_interval') * 1000
                        ))
                    ),
                    timeout=timing_ms('auxiliary_input_timeout')
                )
            else:
                # React 受控组件偶尔会吞掉键入；后续只覆盖为完整编号，
                # 不先执行 Playwright.fill()，从而不产生额外的清空动作。
                locator_input.evaluate(
                    """(el, value) => {
                        el.removeAttribute('readonly');
                        el.readOnly = false;
                        el.removeAttribute('disabled');
                        el.disabled = false;
                        el.focus();
                        const setter = Object.getOwnPropertyDescriptor(
                            window.HTMLInputElement.prototype, 'value'
                        ).set;
                        setter.call(el, value);
                        el.dispatchEvent(new Event('input', {
                            bubbles: true
                        }));
                        el.dispatchEvent(new Event('change', {
                            bubbles: true
                        }));
                    }""",
                    code,
                )

            # 不能在 type()/fill() 返回后就声称成功；按设置等待连续稳定。
            stable_since = None
            actual = ''
            deadline = (
                task_now() + timing_seconds('auxiliary_input_verify_wait')
            )
            while task_now() < deadline:
                pause()
                actual = locator_input.input_value(
                    timeout=timing_ms('tiny_element_timeout')
                ).strip()
                if actual == code:
                    if stable_since is None:
                        stable_since = task_now()
                    elif (
                        task_now() - stable_since
                        >= timing_seconds('auxiliary_input_stable_wait')
                    ):
                        print(f'     已确认辅助核算输入框当前值: {actual}')
                        return locator_input
                else:
                    stable_since = None
                task_sleep(timing_seconds('field_poll_interval'))
            last_error = f'NCC将输入值回滚为“{actual}”'
        except Exception as error:
            last_error = error
        print(
            f'     辅助核算编号第{attempt}/{RETRY_MAX}次输入未生效，'
            f'准备重试：{last_error}'
        )
        task_sleep(timing_seconds('auxiliary_retry_interval'))

    raise FieldFillError(
        '辅助核算',
        f'连续{RETRY_MAX}次未能把编号“{code}”写入“自定义档案”输入框：{last_error}'
    )


def _fill_auxiliary_accounting_once(frame, drawer, code):
    """执行一次辅助核算弹窗填写、候选选择和确认。"""
    modal = _open_auxiliary_modal(frame, drawer)
    print(
        '     辅助核算窗口已打开，等待'
        f'{format_timing_seconds("auxiliary_modal_open_wait")}秒后输入......'
    )
    task_sleep(timing_seconds('auxiliary_modal_open_wait'))
    # 编号只写入一次；后续下拉候选重试必须保留该值，不能重新清空输入框。
    auxiliary_input = _fill_auxiliary_code_input(frame, modal, code)
    last_error = None
    for attempt in range(1, RETRY_MAX + 1):
        try:
            seconds = format_timing_seconds('reference_wait')
            print(
                f'     第{attempt}/{RETRY_MAX}次已填入辅助核算编号: {code}，'
                f'等待下拉匹配项（最长{seconds}秒）......'
            )
            _click_auxiliary_reference_option(
                frame, modal, expected=code,
                timeout=timing_ms('reference_wait')
            )
            break
        except TaskCancelledError:
            raise
        except Exception as error:
            last_error = error
            if attempt < RETRY_MAX:
                print(
                    f'     辅助核算第{attempt}/{RETRY_MAX}次匹配失败，'
                    f'保留已输入编号并重新打开下拉框：{error}'
                )
                try:
                    current_value = auxiliary_input.input_value(
                        timeout=timing_ms('tiny_element_timeout')
                    ).strip()
                    if current_value != code:
                        raise FieldFillError(
                            '辅助核算',
                            f'候选重试前输入值意外变化：期望“{code}”，'
                            f'当前“{current_value or "空"}”'
                        )
                    _dispatch_auxiliary_dom_click(auxiliary_input)
                    auxiliary_input.press(
                        'ArrowDown', timeout=timing_ms('tiny_element_timeout')
                    )
                except TaskCancelledError:
                    raise
                except Exception as reopen_error:
                    last_error = reopen_error
                task_sleep(timing_seconds('reference_retry_interval'))
    else:
        reason = getattr(last_error, 'reason', str(last_error))
        raise FieldFillError(
            '辅助核算',
            f'连续{RETRY_MAX}次清空并重新填入编号“{code}”后'
            f'仍未匹配到下拉数据：{reason}'
        )
    # 候选点击后直接执行窗口确认；不再额外等待、按 Escape 收起候选层，
    # 也不检查是否已回显，避免在最小化或页面较慢时把有效选择撤销。
    _confirm_auxiliary_modal(frame, modal)
    pause()
    return _visible_detail_drawer(
        frame, timeout=timing_ms('element_action_timeout')
    )


def fill_auxiliary_accounting(frame, drawer, value, previous_value=None):
    """L 列有值时填写辅助核算；弹窗确认关闭即视为填写完成。"""
    code = auxiliary_accounting_code(value)
    if code == '':
        print('     底稿“辅助核算”为空，跳过辅助核算')
        return drawer
    return _fill_auxiliary_accounting_once(frame, drawer, code)


def _field_display_text(root, field_name):
    """读取字段区域的显示文本和所有输入值，用于识别 NCC 自动带出的只读值。"""
    try:
        item, _ = _find_field_item(
            root, field_name,
            timeout=timing_ms('tiny_element_timeout'),
            require_input=False
        )
    except Exception:
        return ''
    try:
        return item.evaluate(
            """el => {
                const values = Array.from(el.querySelectorAll('input, textarea'))
                    .flatMap(input => [
                        input.value || '',
                        input.getAttribute('title') || '',
                        input.getAttribute('display') || ''
                    ]);
                values.push(el.innerText || '');
                return values.filter(Boolean).join(' / ');
            }"""
        ).strip()
    except Exception:
        return ''


def _normalize_linked_yes_no(value, field_name=''):
    """把 NCC 联动字段的显示值统一识别为“是”或“否”。"""
    text = excel_text(value).strip()
    if field_name:
        text = text.replace(field_name, '')
    option_hint = re.sub(r'[\s()（）*＊]+', '', text).casefold()
    if option_hint in {'是/否', '是否', '是或否'}:
        return ''
    tokens = [text]
    tokens.extend(re.split(
        r'[\s/\\／＼|｜,，;；:：=＝()（）\[\]【】*＊]+', text
    ))
    yes_values = {'是', 'yes', 'y', 'true', '1'}
    no_values = {'否', 'no', 'n', 'false', '0'}
    for token in tokens:
        normalized = re.sub(r'\s+', '', token).casefold()
        if normalized in yes_values:
            return '是'
        if normalized in no_values:
            return '否'
    return ''


def _read_community_income_flag(root, timeout=None):
    """在新收支项目联动完成后读取“是否社群收入类收支”的是/否值。"""
    if timeout is None:
        timeout = timing_ms('element_action_timeout')
    field_name = '是否社群收入类收支'
    deadline = task_now() + timeout / 1000
    last_snapshot = ''
    stable_flag = ''
    stable_since = None
    while task_now() < deadline:
        pause()
        try:
            item, _ = _find_field_item(
                root, field_name,
                timeout=timing_ms('tiny_element_timeout'),
                require_input=False
            )
            snapshot = item.evaluate(
                """el => {
                    const visible = node => !!(
                        node.offsetWidth || node.offsetHeight
                        || node.getClientRects().length
                    );
                    const values = [];
                    for (const input of el.querySelectorAll('input')) {
                        if (!visible(input)) continue;
                        const type = String(input.type || '').toLowerCase();
                        if (type === 'checkbox') {
                            const itemType = String(
                                input.getAttribute('itemtype') || ''
                            ).toLowerCase();
                            const switchValue = String(
                                input.value
                                || input.getAttribute('value')
                                || ''
                            ).trim().toLowerCase();
                            if (
                                itemType.includes('checkbox_switch')
                                && ['true', 'false', '1', '0'].includes(
                                    switchValue
                                )
                            ) {
                                values.push(
                                    ['true', '1'].includes(switchValue)
                                        ? '是' : '否'
                                );
                            } else {
                                values.push(input.checked ? '是' : '否');
                            }
                            continue;
                        }
                        if (type === 'radio') {
                            if (input.checked) {
                                values.push(
                                    input.getAttribute('aria-label') || '',
                                    input.getAttribute('title') || '',
                                    input.getAttribute('label') || '',
                                    ...Array.from(input.labels || []).map(
                                        label => label.textContent || ''
                                    ),
                                    input.getAttribute('data-value') || '',
                                    input.getAttribute('value') || ''
                                );
                            }
                            continue;
                        }
                        values.push(
                            input.value || '',
                            input.getAttribute('title') || '',
                            input.getAttribute('display') || '',
                            input.getAttribute('data-value') || ''
                        );
                    }
                    for (const node of el.querySelectorAll(
                        '.u-select-selection-selected-value, '
                        + '.u-select-selection-rendered, [aria-checked]'
                    )) {
                        if (!visible(node)) continue;
                        const checked = node.getAttribute('aria-checked');
                        const role = String(
                            node.getAttribute('role') || ''
                        ).toLowerCase();
                        const classes = String(
                            node.className || ''
                        ).toLowerCase();
                        if (checked !== null) {
                            const booleanControl = (
                                role === 'switch'
                                || role === 'checkbox'
                                || classes.includes('switch')
                                || classes.includes('checkbox')
                            );
                            if (booleanControl) {
                                values.push(checked === 'true' ? '是' : '否');
                                continue;
                            }
                            const radioControl = (
                                role === 'radio' || classes.includes('radio')
                            );
                            if (radioControl && checked !== 'true') {
                                continue;
                            }
                            if (radioControl) {
                                const label = node.closest('label');
                                values.push(
                                    node.getAttribute('aria-label') || '',
                                    node.getAttribute('title') || '',
                                    node.textContent || '',
                                    label ? (label.textContent || '') : '',
                                    node.getAttribute('data-value') || '',
                                    node.getAttribute('value') || ''
                                );
                                continue;
                            }
                        }
                        values.push(
                            node.getAttribute('data-value') || '',
                            node.getAttribute('value') || '',
                            node.getAttribute('aria-label') || '',
                            node.getAttribute('title') || '',
                            node.textContent || ''
                        );
                    }
                    return {
                        values: values.filter(Boolean),
                        text: String(el.innerText || '')
                    };
                }"""
            )
            candidates = list(snapshot.get('values') or [])
            candidates.append(snapshot.get('text') or '')
            last_snapshot = ' / '.join(
                excel_text(candidate).strip()
                for candidate in candidates
                if excel_text(candidate).strip()
            )
            observed_flag = ''
            for candidate in candidates:
                observed_flag = _normalize_linked_yes_no(
                    candidate, field_name
                )
                if observed_flag:
                    break
            if observed_flag:
                if observed_flag != stable_flag:
                    stable_flag = observed_flag
                    stable_since = task_now()
                elif (
                    stable_since is not None
                    and task_now() - stable_since
                    >= timing_seconds('reference_option_stable_wait')
                ):
                    print(f'     {field_name}已读取: {observed_flag}')
                    return observed_flag
            else:
                stable_flag = ''
                stable_since = None
        except Exception as error:
            last_snapshot = str(error)
            stable_flag = ''
            stable_since = None
        task_sleep(timing_seconds('candidate_poll_interval'))

    raise FieldFillError(
        field_name,
        f'新收支项目选择后{timeout / 1000:g}秒内未读取到明确的“是/否”值'
        + (f'，当前页面内容：{last_snapshot}' if last_snapshot else '')
    )


def validate_community_business_category(root, business_category):
    """社群收入类收支必须在底稿 W 列提供经营类别。"""
    flag = _read_community_income_flag(root)
    if flag == '是' and excel_text(business_category).strip() == '':
        raise FieldFillError(
            '经营类别',
            'NCC“是否社群收入类收支”为“是”，'
            '底稿“经营类别名称”不能为空'
        )
    return flag


def fill_accounting_subject(
    frame, drawer, value, previous_value=None
):
    """填写会计科目；字段被 NCC 联动临时卸载时刷新当前明细后重试。"""
    code = accounting_subject_code(value)
    previous_code = (
        accounting_subject_code(previous_value)
        if previous_value is not None else None
    )

    def already_filled(root):
        display_text = _field_display_text(root, '会计科目')
        if code in display_text:
            print(
                f'     会计科目已由NCC自动带出: {display_text}，'
                '无需重复填写'
            )
            return True
        return False

    roots = [drawer]
    try:
        active_drawer = _visible_detail_drawer(
            frame, timeout=timing_ms('tiny_element_timeout')
        )
        roots.append(active_drawer)
    except Exception:
        active_drawer = drawer
    roots.append(frame.locator('body'))

    last_error = None
    for root in roots:
        if already_filled(root):
            return active_drawer
        try:
            _fill_reference_field(
                frame, root, '会计科目', code,
                '“会计科目”', previous_value=previous_code
            )
            return _visible_detail_drawer(
                frame, timeout=timing_ms('tiny_element_timeout')
            )
        except FieldFillError as error:
            last_error = error
            if '未找到' not in error.reason:
                raise

    print('     [恢复] 会计科目字段暂未渲染，重新展开当前明细后重试')
    close_detail_drawer(frame, active_drawer)
    refreshed_drawer = click_latest_expand(frame)
    if already_filled(refreshed_drawer):
        return refreshed_drawer
    try:
        _fill_reference_field(
            frame, refreshed_drawer, '会计科目', code,
            '“会计科目”', previous_value=previous_code
        )
    except FieldFillError as error:
        if '未找到' not in error.reason:
            raise
        _dump_region_inputs(refreshed_drawer, '会计科目')
        raise FieldFillError(
            '会计科目',
            '重新展开当前明细后仍未找到字段；'
            f'NCC可能尚未完成字段联动（首次错误：{last_error}）'
        ) from error
    return refreshed_drawer


def _plain_values_equivalent(expected, actual):
    expected_text = excel_text(expected).replace('\r\n', '\n').replace('\r', '\n')
    actual_text = excel_text(actual).replace('\r\n', '\n').replace('\r', '\n')
    if expected_text == actual_text:
        return True
    try:
        expected_number = Decimal(expected_text.replace(',', ''))
        actual_number = Decimal(actual_text.replace(',', ''))
        return expected_number == actual_number
    except (InvalidOperation, ValueError):
        return False


def _fill_plain_field(root, field_name, value, source_name):
    """清空并填写不需要下拉选择的普通输入字段。"""
    text = excel_text(value).replace('\r\n', '\n').replace('\r', '\n')
    if text == '':
        raise FieldFillError(field_name, f'底稿{source_name}为空')

    try:
        locator_input = _find_field_input(root, field_name)
        clear_before_fill(locator_input, field_name)
        locator_input.fill(
            text, timeout=timing_ms('element_action_timeout')
        )
        locator_input.press('Tab')
        task_sleep(timing_seconds('plain_input_commit_wait'))

        actual_input = _find_field_input(root, field_name)
        actual = actual_input.input_value(
            timeout=timing_ms('short_element_timeout')
        ).strip()
    except FieldFillError:
        raise
    except Exception as e:
        raise FieldFillError(field_name, e) from e
    if not _plain_values_equivalent(text, actual):
        raise FieldFillError(
            field_name,
            f'写入后页面值不一致：目标={text}，页面={actual}'
        )
    print(f'     已填入{field_name}: {text}（页面显示: {actual}）')
    pause()


def _set_checkbox_field(root, field_name, checked):
    """设置普通复选框或 NCC checkbox_switch，并核验业务状态。"""
    try:
        item, _ = _find_field_item(
            root, field_name, require_input=False
        )
    except Exception as e:
        raise FieldFillError(field_name, e) from e

    def read_state():
        return item.evaluate(
            """el => {
                const result = (checked, evidence) => ({
                    checked: !!checked,
                    evidence: String(evidence || '')
                });
                const parseBoolean = raw => {
                    const value = String(raw ?? '').trim().toLowerCase();
                    if (['true', '1', 'yes', 'y', '是'].includes(value)) {
                        return true;
                    }
                    if (['false', '0', 'no', 'n', '否'].includes(value)) {
                        return false;
                    }
                    return null;
                };
                const role = el.querySelector(
                    '[role="checkbox"], [role="switch"]'
                );
                if (role) {
                    const aria = parseBoolean(
                        role.getAttribute('aria-checked')
                    );
                    if (aria !== null) {
                        return result(aria, `aria-checked=${aria}`);
                    }
                }
                if (el.querySelector(
                    '.is-checked, .u-checkbox-checked, '
                    + '.u-switch-checked, .checked'
                )) {
                    return result(true, 'checked-class');
                }
                const input = el.querySelector('input[type="checkbox"]');
                if (input) {
                    const itemType = String(
                        input.getAttribute('itemtype') || ''
                    ).toLowerCase();
                    if (itemType.includes('checkbox_switch')) {
                        const businessValue = parseBoolean(
                            input.value || input.getAttribute('value')
                        );
                        if (businessValue !== null) {
                            return result(
                                businessValue,
                                `checkbox_switch.value=${businessValue}`
                            );
                        }
                    }
                    return result(
                        !!input.checked, `input.checked=${!!input.checked}`
                    );
                }
                return result(false, '未发现已选中状态');
            }"""
        )

    try:
        current_info = read_state()
        current = bool(current_info.get("checked"))
        if current != checked:
            checkboxes = item.locator("css=input[type='checkbox']")
            if checkboxes.count() > 0:
                checkbox = checkboxes.first
                pause()
                # NCC checkbox_switch 的原生 checked 不变化，不能使用
                # Playwright check()/uncheck()；普通 click 后按业务值核验。
                checkbox.click(
                    force=True, timeout=timing_ms('element_action_timeout')
                )
            else:
                click_targets = item.locator(
                    "css=[role='checkbox'], .u-checkbox, "
                    ".u-checkbox-label, label"
                )
                clicked = False
                for index in range(click_targets.count()):
                    target = click_targets.nth(index)
                    if target.is_visible():
                        pause()
                        target.click(
                            force=True,
                            timeout=timing_ms('element_action_timeout')
                        )
                        clicked = True
                        break
                if not clicked:
                    raise Exception('未找到可点击的复选框')

        deadline = task_now() + timing_seconds('short_element_timeout')
        last_evidence = current_info.get("evidence", '')
        while task_now() < deadline:
            pause()
            final_info = read_state()
            last_evidence = final_info.get("evidence", '')
            if bool(final_info.get("checked")) == checked:
                state_text = '已选中' if checked else '未选中'
                print(
                    f'     {field_name}{state_text}'
                    f'（页面状态: {last_evidence}）'
                )
                task_sleep(timing_seconds('checkbox_commit_wait'))
                pause()
                return
            task_sleep(timing_seconds('field_poll_interval'))
        target_state = '选中' if checked else '取消选中'
        raise Exception(
            f'复选框未能{target_state}，页面状态: {last_evidence or "未知"}'
        )
    except FieldFillError:
        raise
    except Exception as e:
        raise FieldFillError(field_name, e) from e


def fill_internal_collection_fields(frame, drawer, row, previous_row=None):
    """在原位置处理底稿 X 列的内部代收付开关。"""
    flag = excel_text(row["internal_collection_flag"])
    if flag not in ('是', '否'):
        raise FieldFillError(
            '是否涉及内部代收付（含清理）',
            f'底稿“代收代付情况”只能填写“是”或“否”，当前值为“{flag}”'
        )

    previous_flag = (
        excel_text(previous_row["internal_collection_flag"])
        if previous_row is not None else ''
    )
    _set_checkbox_field(
        drawer, '是否涉及内部代收付（含清理）', flag == '是'
    )
    if flag == '否':
        if previous_flag == '是':
            task_sleep(timing_seconds('internal_collection_reset_wait'))
            inherited_fields = (
                '内部代收付（含清理）类型',
                '被代收方（被收回方）',
                '被代收方（被收回方）项目档案',
                '代收（收回）业务实际业务收费类型',
            )
            for field_name in inherited_fields:
                try:
                    _, state = _find_visible_field_input_state(
                        drawer, field_name,
                        timeout=timing_ms('tiny_element_timeout')
                    )
                except Exception:
                    continue
                if state["value"].strip() != '':
                    raise FieldFillError(
                        field_name,
                        f'当前行“代收代付情况”为“否”，但页面仍保留上一条继承值“{state["value"]}”'
                    )
        print('     底稿“代收代付情况”为“否”，跳过内部代收联动字段')
        return


def fill_internal_collection_reference_fields_before_tax_rate(
    frame, drawer, row, previous_row=None
):
    """内部代收付时，在税率前按既定顺序填写四项联动参照字段。"""
    if excel_text(row["internal_collection_flag"]) != '是':
        return
    previous_flag = (
        excel_text(previous_row["internal_collection_flag"])
        if previous_row is not None else ''
    )
    internal_collection_type = excel_text(row["internal_collection_type"])
    _fill_reference_field(
        frame, drawer, '内部代收付（含清理）类型', internal_collection_type,
        '“代收代付类型”',
        previous_value=(
            previous_row["internal_collection_type"]
            if previous_row is not None and previous_flag == '是'
            else None
        )
    )
    _fill_reference_field(
        frame, drawer, '被代收方（被收回方）', row["represented_party"],
        '“被代付账套/被代收账套名称”',
        previous_value=(
            previous_row["represented_party"]
            if previous_row is not None and previous_flag == '是'
            else None
        )
    )
    _fill_reference_field(
        frame, drawer, '被代收方（被收回方）项目档案',
        row["represented_project"],
        '“被代收方（被收回方）项目档案”',
        previous_value=(
            previous_row["represented_project"]
            if previous_row is not None and previous_flag == '是'
            else None
        )
    )
    _fill_reference_field(
        frame, drawer, '代收（收回）业务实际业务收费类型',
        row["actual_charge_type"],
        '“收费类型名称”',
        previous_value=(
            previous_row["actual_charge_type"]
            if previous_row is not None and previous_flag == '是'
            else None
        )
    )


def _click_dropdown_text(
    frame, text, field_name, timeout=None
):
    """在当前可见的下拉菜单中点击指定文本。"""
    if timeout is None:
        timeout = timing_ms('reference_wait')
    selectors = [
        "css=[role='option']",
        "css=.u-select-dropdown-menu-item",
        "css=.rc-select-dropdown-menu-item",
        "css=.ant-select-dropdown-menu-item",
        "css=ul.u-select-dropdown-menu li",
        "css=div.u-select-dropdown li",
        "css=div.u-dropdown li",
    ]
    deadline = task_now() + timeout / 1000
    while task_now() < deadline:
        pause()
        for selector in selectors:
            options = frame.locator(selector)
            try:
                count = options.count()
            except Exception:
                continue
            for index in range(count - 1, -1, -1):
                option = options.nth(index)
                try:
                    if not option.is_visible():
                        continue
                    option_text = (
                        option.get_attribute('title')
                        or option.inner_text(
                            timeout=timing_ms('tiny_element_timeout')
                        )
                        or ''
                    ).strip()
                    if option_text == text:
                        pause()
                        option.click(
                            timeout=timing_ms('element_action_timeout')
                        )
                        print(f'     {field_name}已选择: {text}')
                        return
                except Exception:
                    continue
        task_sleep(timing_seconds('field_poll_interval'))
    seconds = max(1, int(timeout / 1000))
    raise FieldFillError(
        field_name,
        f'等待{seconds}秒后下拉框中未找到选项“{text}”'
    )


def _select_dropdown_field(frame, root, field_name, target):
    """固定下拉也按“打开—匹配—核对”完整重试三次。"""
    last_error = None
    for attempt in range(1, RETRY_MAX + 1):
        click_target = None
        try:
            item, _ = _find_field_item(root, field_name)
            click_targets = [
                item.locator("css=.u-select-selection").first,
                item.locator("css=.u-select").first,
                item.locator("css=input:not(.hidden-input)").first,
                item,
            ]
            for target_locator in click_targets:
                try:
                    if target_locator.count() > 0 and target_locator.is_visible():
                        target_locator.click(
                            timeout=timing_ms('element_action_timeout')
                        )
                        click_target = target_locator
                        break
                except Exception:
                    continue
            if click_target is None:
                raise FieldFillError(field_name, '无法打开下拉框')

            print(
                f'     {field_name}第{attempt}/{RETRY_MAX}次等待下拉选项'
                f'“{target}”......'
            )
            _click_dropdown_text(frame, target, field_name)
            task_sleep(timing_seconds('reference_post_select_wait'))
            selected_item, _ = _find_field_item(root, field_name)
            selected = selected_item.evaluate(
                """(el, target) => {
                    if (String(el.innerText || '').includes(target)) return true;
                    return Array.from(el.querySelectorAll('input'))
                        .some(input => String(input.value || '').includes(target));
                }""",
                target
            )
            if not selected:
                raise FieldFillError(
                    field_name, f'选择失败：目标值={target}'
                )
            pause()
            return
        except TaskCancelledError:
            raise
        except Exception as error:
            last_error = error
            if attempt < RETRY_MAX:
                try:
                    if click_target is not None:
                        click_target.press(
                            'Escape', timeout=timing_ms('tiny_element_timeout')
                        )
                except Exception:
                    pass
                print(
                    f'     {field_name}第{attempt}/{RETRY_MAX}次下拉匹配失败，'
                    f'准备重新打开并选择：{error}'
                )
                task_sleep(timing_seconds('reference_retry_interval'))

    reason = getattr(last_error, 'reason', str(last_error))
    raise FieldFillError(
        field_name,
        f'连续{RETRY_MAX}次仍未找到或选中“{target}”：{reason}'
    )


def _visible_detail_drawer(frame, timeout=None):
    """取得当前可见的“详细信息编辑”侧拉区域。"""
    if timeout is None:
        timeout = timing_ms('long_element_timeout')
    deadline = task_now() + timeout / 1000
    while task_now() < deadline:
        pause()
        bodies = frame.locator(
            "css=div.card-table-modal-form-wrap-body.cf, "
            "div.card-table-modal-form-wrap-body"
        )
        try:
            for index in range(bodies.count() - 1, -1, -1):
                body = bodies.nth(index)
                if body.is_visible():
                    return body
        except Exception:
            pass
        task_sleep(timing_seconds('field_poll_interval'))
    raise Exception('点击“展开”后未出现“详细信息编辑”侧拉框')


def click_latest_expand(frame):
    """点击明细表最新一行的“展开”，并等待侧拉框出现。"""
    try:
        body = frame.locator(SELECTORS["body_table"]).first
        if body.count() > 0:
            body.evaluate('element => element.scrollTop = element.scrollHeight')
    except Exception:
        pass

    candidates = [
        frame.locator(SELECTORS["row_edit_btn"]),
        frame.locator(
            "xpath=//a[contains(@class,'row-btn-open_edit') "
            "or normalize-space(.)='展开']"
        ),
    ]
    for group in candidates:
        try:
            count = group.count()
        except Exception:
            continue
        for index in range(count - 1, -1, -1):
            button = group.nth(index)
            try:
                if button.is_visible():
                    pause()
                    button.click(
                        timeout=timing_ms('long_element_timeout')
                    )
                    drawer = _visible_detail_drawer(frame)
                    print('     已点击最新明细行“展开”')
                    pause()
                    return drawer
            except Exception:
                continue
    raise Exception('未找到最新明细行的“展开”按钮')


def _expand_button_counts(frame):
    """返回明细表“展开”按钮的总数和可见数。"""
    buttons = frame.locator(
        "xpath=//a[contains(@class,'row-btn-open_edit') "
        "or normalize-space(.)='展开']"
    )
    total = buttons.count()
    visible = 0
    for index in range(total):
        try:
            if buttons.nth(index).is_visible():
                visible += 1
        except Exception:
            continue
    return total, visible


def _detail_row_snapshot(frame):
    """取得明细表状态；兼容超过10行后NCC只保留可见行的虚拟滚动表格。"""
    try:
        body = frame.locator(SELECTORS["body_table"]).first
        if body.count() > 0:
            body.evaluate('element => element.scrollTop = element.scrollHeight')
    except Exception:
        pass

    total_buttons, visible_buttons = _expand_button_counts(frame)
    snapshot = {
        "button_total": total_buttons,
        "button_visible": visible_buttons,
        "max_sequence": None,
        "signature": "",
    }
    roots = frame.locator("css=div.selectable-bodys")
    try:
        if roots.count() == 0:
            return snapshot
        table_state = roots.first.evaluate(
            """
            root => {
                const visible = element => {
                    const style = window.getComputedStyle(element);
                    const rect = element.getBoundingClientRect();
                    return style.display !== 'none'
                        && style.visibility !== 'hidden'
                        && rect.width > 0 && rect.height > 0;
                };
                const clean = value => String(value || '')
                    .replace(/\\s+/g, ' ').trim();
                const rows = Array.from(
                    root.querySelectorAll('tr.u-table-row')
                ).filter(visible);
                const sequenceRows = Array.from(root.querySelectorAll(
                    'div.u-table-fixed-left tr.u-table-row'
                )).filter(visible);
                const sequences = [];
                for (const row of sequenceRows) {
                    const text = clean(row.innerText);
                    const matched = text.match(/^(\\d+)(?:\\s|$)/);
                    if (matched) sequences.push(Number(matched[1]));
                }
                const signature = rows.map(row => {
                    const values = Array.from(row.querySelectorAll('input'))
                        .map(input => clean(input.value)).join('¦');
                    return [
                        row.getAttribute('data-row-key') || '',
                        row.getAttribute('aria-rowindex') || '',
                        clean(row.innerText),
                        values,
                    ].join('~');
                }).join('||');
                return {
                    max_sequence: sequences.length
                        ? Math.max(...sequences) : null,
                    signature,
                };
            }
            """
        )
        if isinstance(table_state, dict):
            snapshot.update(table_state)
    except Exception:
        pass
    return snapshot


def _detail_row_added(before, after):
    """判断增行是否完成，不能只依赖虚拟表格中固定不变的按钮数量。"""
    before_sequence = before.get("max_sequence")
    after_sequence = after.get("max_sequence")
    if (
        isinstance(before_sequence, (int, float))
        and isinstance(after_sequence, (int, float))
        and after_sequence > before_sequence
    ):
        return True
    if after.get("button_total", 0) > before.get("button_total", 0):
        return True
    if after.get("button_visible", 0) > before.get("button_visible", 0):
        return True
    return bool(
        after.get("signature")
        and after.get("signature") != before.get("signature")
    )


def click_add_detail_row(frame):
    """点击明细工具栏“增行”，确认新明细行已经生成。"""
    before = _detail_row_snapshot(frame)
    buttons = frame.locator(
        "xpath=//button[normalize-space(.)='增行']"
    )
    clicked = False
    for index in range(buttons.count() - 1, -1, -1):
        button = buttons.nth(index)
        try:
            if button.is_visible() and button.is_enabled():
                pause()
                button.click(timeout=timing_ms('element_action_timeout'))
                clicked = True
                break
        except Exception:
            continue
    if not clicked:
        raise Exception('未找到可点击的明细“增行”按钮')

    if not wait_end(frame):
        raise FieldFillError(
            '明细增行',
            '点击“增行”后页面加载超过'
            f'{format_timing_seconds("page_load_timeout")}秒，已停止继续处理'
        )
    deadline = task_now() + timing_seconds('detail_row_detect_wait')
    stable_signature = None
    stable_since = None
    unobserved_since = None
    while task_now() < deadline:
        pause()
        after = _detail_row_snapshot(frame)
        if _detail_row_added(before, after):
            before_sequence = before.get("max_sequence")
            after_sequence = after.get("max_sequence")
            count_increased = (
                after.get("button_total", 0) > before.get("button_total", 0)
                or after.get("button_visible", 0)
                > before.get("button_visible", 0)
            )
            if (
                isinstance(before_sequence, (int, float))
                and isinstance(after_sequence, (int, float))
                and after_sequence > before_sequence
            ):
                print(
                    f'     已点击“增行”，最新明细序号由 '
                    f'{int(before_sequence)} 增加至 {int(after_sequence)}'
                )
                pause()
                return
            if count_increased:
                print(
                    f'     已点击“增行”，可见展开按钮由 '
                    f'{before.get("button_visible", 0)} 增加至 '
                    f'{after.get("button_visible", 0)}'
                )
                pause()
                return

            current_signature = after.get("signature")
            now = task_now()
            if current_signature != stable_signature:
                stable_signature = current_signature
                stable_since = now
            elif (
                stable_since is not None
                and now - stable_since
                >= timing_seconds('detail_signature_stable_wait')
            ):
                print(
                    '     已点击“增行”，明细虚拟滚动窗口已更新，'
                    '确认新行生成'
                )
                pause()
                return
        else:
            stable_signature = None
            stable_since = None
            if unobserved_since is None:
                unobserved_since = task_now()
            elif (
                task_now() - unobserved_since
                >= timing_seconds('detail_virtual_fallback_wait')
            ):
                # 超过10行时，NCC会复用同一批DOM节点；此时按钮数量、
                # 行文本和序号都可能完全不变，但“增行”请求已经结束。
                # 继续向下展开底部最新行，不能再因前端虚拟滚动而中断。
                print(
                    '     已点击“增行”，虚拟表格未暴露新增序号；'
                    '页面已稳定，继续填写最新明细行'
                )
                task_sleep(timing_seconds('reference_post_select_wait'))
                return
        task_sleep(timing_seconds('field_poll_interval'))
    raise Exception(
        '点击“增行”后'
        f'{format_timing_seconds("detail_row_detect_wait")}秒内未检测到新明细行'
    )


def close_detail_drawer(frame, drawer):
    """点击侧拉框右上角箭头并确认侧拉框关闭。"""
    selectors = [
        "css=span.icon.iconfont.icon-you.icon-celashouqi",
        "css=.icon-celashouqi",
        "xpath=//*[contains(@class,'icon-celashouqi')]",
    ]
    for selector in selectors:
        buttons = frame.locator(selector)
        try:
            count = buttons.count()
        except Exception:
            continue
        for index in range(count - 1, -1, -1):
            button = buttons.nth(index)
            try:
                if button.is_visible():
                    pause()
                    button.click(
                        timeout=timing_ms('element_action_timeout')
                    )
                    drawer.wait_for(
                        state='hidden',
                        timeout=timing_ms('element_action_timeout')
                    )
                    print('     已点击明细右上角箭头')
                    pause()
                    return
            except Exception:
                continue
    raise Exception('未找到或未能点击明细右上角箭头')


def settle_tax_rate_before_close(drawer):
    """税率选择后触发失焦并等待NCC完成内部值绑定，不清空或重填。"""
    try:
        tax_rate_input = _find_field_input(drawer, '税率')
        try:
            tax_rate_input.press(
                'Tab', timeout=timing_ms('short_element_timeout')
            )
        except Exception:
            tax_rate_input.evaluate('(input) => input.blur()')
        print('     税率已触发失焦提交')
    except Exception as error:
        print(f'     税率无法主动触发失焦，将继续等待NCC提交：{error}')
    print(
        '     等待税率绑定稳定（'
        f'{format_timing_seconds("tax_rate_bind_wait")}秒）......'
    )
    task_sleep(timing_seconds('tax_rate_bind_wait'))


def payer_bank_account_value(row):
    """表头“付款银行账号”选填；为空返回空字符串。"""
    return excel_text(row.get("payer_bank_account", ""))


def fill_payer_bank_account(frame, card_root, row):
    """表头“付款银行账号”选填；为空时跳过，不清空系统带出值。"""
    payer_account = payer_bank_account_value(row)
    if payer_account == '':
        print('     底稿“付款银行账号”为空，跳过表头付款银行账号填写')
        return
    _fill_reference_field(
        frame, card_root, '付款银行账号', payer_account,
        '“付款银行账号”'
    )


def header_remark_value(row):
    """表头“备注”选填；为空返回空字符串。"""
    return excel_text(
        normalize_receipt_value(row.get("remark", ""), "remark")
    )


def fill_header_remark(card_root, row):
    """表头“备注”选填；为空时跳过，不清空系统带出值。"""
    remark_text = header_remark_value(row)
    if remark_text == '':
        print('     底稿“备注”为空，跳过表头备注填写')
        return
    _fill_plain_field(
        card_root, '备注', remark_text, '“备注”'
    )


def fill_payer_bank_account_before_save(frame, row):
    """在所有数据（含全部明细）填妥后、点击保存/保存提交前最后一个填入表头“付款银行账号”。"""
    if not is_fill_step_enabled("payer_bank_account"):
        print('     已停用“付款银行账号”填写，保存前跳过表头付款银行账号')
        return
    row = normalize_receipt_row(row)
    payer_account = payer_bank_account_value(row)
    if payer_account != '':
        print('     全部明细已填妥，保存前最后填写表头“付款银行账号”')
    fill_payer_bank_account(frame, frame.locator("body"), row)


def _fill_header_step(frame, card_root, row, transaction_object, step_key):
    """执行单个表头填写步骤（仅表头内部可互换顺序，不含固定最前项和保存前付款银行账号）。"""
    if step_key == "settlement_method":
        _fill_reference_field(
            frame, card_root, '结算方式', row["settlement_method"],
            '“结算方式名称”'
        )
    elif step_key == "transaction_object":
        _select_dropdown_field(
            frame, card_root, '交易对象类型', transaction_object
        )
    elif step_key == "receive_bank_account":
        _fill_reference_field(
            frame, card_root, '收款银行账户', row["bank_account"],
            '“收款银行账号”'
        )
    elif step_key == "remark":
        fill_header_remark(card_root, row)
    else:
        raise FieldFillError('表头填写顺序', f'不支持的表头填写步骤：{step_key}')


def fill_header_fields(frame, row, transaction_object):
    """按当前表头顺序填写每张单据只需执行一次的卡片头字段（付款银行账号在全部明细后、保存前填入）。"""
    row = normalize_receipt_row(row)
    card_root = frame.locator("body")
    settings = active_fill_order_settings()
    disabled_set = set(settings.get("disabled", ()))
    disabled_header = [
        HEADER_FILL_STEP_LABELS[key]
        for key in settings.get("header_order", HEADER_FILL_DEFAULT_ORDER)
        if key in disabled_set
    ]
    if disabled_header:
        print('     本次已停用表头字段：' + '、'.join(disabled_header))
    for step_key in enabled_header_fill_order():
        _fill_header_step(frame, card_root, row, transaction_object, step_key)


def department_reference_value(row):
    """部门编码优先；无编码时使用名称，两者为空不操作 NCC 字段。"""
    code = excel_text(row.get('department_code', ''))
    if code:
        return code, '“部门编码”'
    return excel_text(row.get('department_name', '')), '“部门”'


def fill_department_after_income_item(frame, drawer, row, previous_row=None):
    value, source = department_reference_value(row)
    if not value:
        print('     部门编码和部门均为空，跳过明细部门填写')
        return
    previous = department_reference_value(previous_row)[0] if previous_row is not None else None
    _fill_reference_field(frame, drawer, '部门', value, source, previous_value=previous)


def _fill_detail_step(
    frame, drawer, row, transaction_object, step_key, previous_row=None
):
    """执行单个明细填写步骤（仅明细内部可互换顺序）；若重开侧栏则返回最新 drawer。"""
    if step_key == "summary":
        _fill_plain_field(
            drawer, '摘要', row["summary"], '“摘要”'
        )
    elif step_key == "internal_collection":
        fill_internal_collection_fields(
            frame, drawer, row, previous_row=previous_row
        )
    elif step_key == "customer":
        if transaction_object == '客户':
            customer_value, customer_source, uses_customer_code = (
                customer_reference_value(row)
            )
            previous_customer_value = None
            if previous_row is not None:
                previous_customer_value = customer_reference_value(
                    previous_row
                )[0]
            _fill_reference_field(
                frame, drawer, '客户', customer_value, customer_source,
                previous_value=previous_customer_value,
                retry_ambiguous=False,
                allow_multiple_matches=uses_customer_code,
                customer_name_only=not uses_customer_code
            )
    elif step_key == "receipt_method":
        _fill_reference_field(
            frame, drawer, '收款方式', row["receipt_method"],
            '“收款方式名称”',
            previous_value=(
                previous_row["receipt_method"]
                if previous_row is not None else None
            )
        )
    elif step_key == "expense_period":
        expense_period = format_expense_period(row["expense_period"])
        previous_expense_period = (
            format_expense_period(previous_row["expense_period"])
            if previous_row is not None else None
        )
        _fill_reference_field(
            frame, drawer, '费用所属期', expense_period,
            '“费用所属期间”',
            previous_value=previous_expense_period
        )
    elif step_key == "tax_method":
        _fill_reference_field(
            frame, drawer, '计税方式', row["tax_method"],
            '“计税方式”',
            previous_value=(
                previous_row["tax_method"]
                if previous_row is not None else None
            )
        )
    elif step_key == "accounting_subject":
        drawer = fill_accounting_subject(
            frame, drawer, row["accounting_subject"],
            previous_value=(
                previous_row["accounting_subject"]
                if previous_row is not None else None
            )
        )
    elif step_key == "project_archive":
        if excel_text(row["internal_collection_flag"]) == '是':
            # 内部代收付单据不填写 NCC“项目档案”；同时清除新增行继承值，
            # 避免把上一条非代收代付明细的项目档案带入当前行。
            _clear_optional_reference_field(
                drawer, '项目档案',
                '“项目档案名称”（代收代付为“是”，按规则无需填写）',
                preserve_linked_readonly=True
            )
            print('     底稿“代收代付情况”为“是”，已跳过NCC“项目档案”填写')
        else:
            project_archive_name = excel_text(row["project_archive_name"])
            project_archive_query = project_archive_query_value(project_archive_name)
            if project_archive_query != '':
                if project_archive_query != project_archive_name:
                    print(
                        '     项目档案按“编码/名称”解析：'
                        f'{project_archive_name} -> {project_archive_query}'
                    )
                _fill_reference_field(
                    frame, drawer, '项目档案', project_archive_query,
                    '“项目档案名称”',
                    previous_value=(
                        project_archive_query_value(
                            previous_row["project_archive_name"]
                        )
                        if previous_row is not None else None
                    )
                )
            else:
                _clear_optional_reference_field(
                    drawer, '项目档案', '“项目档案名称”'
                )
    elif step_key == "receipt_original_amount":
        _fill_plain_field(
            drawer, '收款原币金额', row["receipt_original_amount"],
            '“原币收入”'
        )
    elif step_key == "internal_collection_references":
        fill_internal_collection_reference_fields_before_tax_rate(
            frame, drawer, row, previous_row=previous_row
        )
    elif step_key == "new_income_expense_item":
        _fill_reference_field(
            frame, drawer, '新收支项目', row["new_income_expense_item"],
            '“NCC收支项目名称”',
            previous_value=(
                previous_row["new_income_expense_item"]
                if previous_row is not None else None
            )
        )
    elif step_key == "department":
        fill_department_after_income_item(frame, drawer, row, previous_row=previous_row)
    elif step_key == "split_income_type":
        fill_split_income_type_after_income_item(
            frame, drawer, row, previous_row=previous_row
        )
    elif step_key == "business_category":
        validate_community_business_category(
            drawer, row["business_category"]
        )
        business_category = excel_text(row["business_category"])
        if business_category != '':
            _fill_reference_field_if_editable(
                frame, drawer, '经营类别', business_category,
                '“经营类别名称”',
                previous_value=(
                    previous_row["business_category"]
                    if previous_row is not None else None
                )
            )
        else:
            _clear_optional_reference_field(
                drawer, '经营类别', '“经营类别名称”',
                preserve_linked_readonly=True
            )
    elif step_key == "auxiliary_accounting":
        # 默认将辅助核算紧邻税率之前填写。新收支项目、经营类别会触发 NCC 联动，
        # 因此默认不能在这些字段之前先填辅助核算，避免后续联动刷新该值。
        drawer = fill_auxiliary_accounting(
            frame, drawer, row["auxiliary_accounting"]
        )
    elif step_key == "tax_rate":
        tax_rate = format_tax_rate(row["tax_rate"])
        previous_tax_rate = (
            format_tax_rate(previous_row["tax_rate"])
            if previous_row is not None else None
        )
        _fill_reference_field(
            frame, drawer, '税率', tax_rate,
            '“税率”',
            previous_value=previous_tax_rate
        )
        settle_tax_rate_before_close(drawer)
    else:
        raise FieldFillError('明细填写顺序', f'不支持的明细填写步骤：{step_key}')
    return drawer


def fill_detail_row(
    frame, row, transaction_object, previous_row=None
):
    """展开最新明细行，按当前明细顺序填写该行字段并点击右上角箭头退出。"""
    row = normalize_receipt_row(row)
    previous_row = (
        normalize_receipt_row(previous_row)
        if previous_row is not None else None
    )
    drawer = click_latest_expand(frame)
    settings = active_fill_order_settings()
    disabled_set = set(settings.get("disabled", ()))
    disabled_detail = [
        DETAIL_FILL_STEP_LABELS[key]
        for key in settings.get("detail_order", DETAIL_FILL_DEFAULT_ORDER)
        if key in disabled_set
    ]
    if disabled_detail:
        print('     本次已停用明细字段：' + '、'.join(disabled_detail))
    for step_key in enabled_detail_fill_order():
        drawer = _fill_detail_step(
            frame, drawer, row, transaction_object, step_key,
            previous_row=previous_row
        )
    close_detail_drawer(frame, drawer)


def fill_followup_fields(frame, row, transaction_object):
    """兼容首行流程：填写卡片头、第一条明细，并在保存前最后填写付款银行账号。"""
    fill_header_fields(frame, row, transaction_object)
    fill_detail_row(frame, row, transaction_object)
    fill_payer_bank_account_before_save(frame, row)


def _save_result_workbook(data, file_path, respect_task_state=True):
    """
    保存底稿回写结果。原文件被 Excel 占用时按时间设置等待；
    仍不可写则改存带时间戳的“_程序回写”文件。
    一旦降级，本次运行始终写备用文件，避免用启动时的旧快照覆盖原底稿的新修改。
    """
    output_path = getattr(data, '_result_output_path', file_path)

    try:
        data.save(output_path)
        return output_path
    except PermissionError:
        pass

    if os.path.abspath(output_path) == os.path.abspath(file_path):
        unlock_wait = timing_seconds('workbook_unlock_wait')
        retry_interval = timing_seconds('workbook_retry_interval')
        print(
            '     底稿文件正被占用，请关闭 Excel；程序将等待'
            f'{unlock_wait:g}秒后停止重试回写......'
        )
        if respect_task_state:
            clock = task_now
            wait_once = lambda: task_sleep(retry_interval)
        else:
            clock = time.monotonic
            wait_once = lambda: time.sleep(retry_interval)
        deadline = clock() + unlock_wait
        while clock() < deadline:
            # 运行统计属于任务收尾，即使用户此时按了停止也必须尽量落盘。
            wait_once()
            try:
                data.save(file_path)
                print('     Excel 已释放，底稿回写成功')
                return file_path
            except PermissionError:
                continue

    base, ext = os.path.splitext(file_path)
    ext = ext or '.xlsx'
    timestamp = time.strftime('%Y%m%d_%H%M%S')
    fallback_path = f'{base}_程序回写_{timestamp}{ext}'
    suffix = 1
    while os.path.exists(fallback_path):
        fallback_path = f'{base}_程序回写_{timestamp}_{suffix}{ext}'
        suffix += 1
    data.save(fallback_path)
    data._result_output_path = fallback_path
    print(
        '     原底稿仍被占用，回写结果已另存为：'
        f'{fallback_path}'
    )
    return fallback_path


def _runtime_log_next_number(runtime_sheet):
    """从已有A列取得下一次运行序号，兼容数字和“第N次”文本。"""
    largest = 0
    for row_no in range(2, runtime_sheet.max_row + 1):
        value = runtime_sheet.cell(row_no, column=1).value
        if isinstance(value, bool) or value is None:
            continue
        if isinstance(value, (int, float)):
            number = int(value)
        else:
            matched = re.search(r'\d+', str(value))
            number = int(matched.group()) if matched else 0
        largest = max(largest, number)
    return largest + 1


def _runtime_log_next_row(runtime_sheet):
    """忽略仅有格式的空行，保证日志从第2行连续追加。"""
    last_data_row = 1
    for row_no in range(2, runtime_sheet.max_row + 1):
        if any(
            runtime_sheet.cell(row_no, column=column).value not in (None, '')
            for column in range(1, len(RUN_TIME_HEADERS) + 1)
        ):
            last_data_row = row_no
    return last_data_row + 1


def _format_elapsed_time(seconds):
    total_seconds = max(0, int(round(seconds)))
    hours, remainder = divmod(total_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    return f'{hours:02d}:{minutes:02d}:{seconds:02d}'


def _finalize_runtime_log(file_path, run_context):
    """在“运行时间”sheet追加本次统计；同一次运行最多写入一次。"""
    if (
        run_context is None
        or run_context.get('runtime_log_recorded')
        or run_context.get('runtime_log_in_progress')
    ):
        return
    run_context['runtime_log_in_progress'] = True
    run_context.setdefault('finished_at', time.monotonic())
    if run_context.get('source_validated') is False:
        run_context['runtime_log_in_progress'] = False
        return
    if os.path.splitext(file_path)[1].lower() not in ('.xlsx', '.xlsm'):
        run_context['runtime_log_in_progress'] = False
        return

    data = run_context.get('workbook')
    owns_workbook = data is None
    try:
        if data is None:
            data = load_receipt_workbook(file_path)

        row_outcomes = dict(run_context.get('row_outcomes', {}))
        if 'selected_row_numbers' in run_context:
            selected_row_numbers = list(dict.fromkeys(
                int(row_no)
                for row_no in run_context.get('selected_row_numbers', [])
            ))
        else:
            selected_row_numbers = list(row_outcomes)
        success_count = sum(
            1 for row_no in selected_row_numbers
            if row_outcomes.get(row_no) == STATUS_SUCCESS
        )
        abnormal_count = sum(
            1 for row_no in selected_row_numbers
            if row_outcomes.get(row_no) == STATUS_ABNORMAL
        )
        run_data_count = len(selected_row_numbers)
        unfinished_count = max(
            0, run_data_count - success_count - abnormal_count
        )

        if RUN_TIME_SHEET_NAME in data.sheetnames:
            runtime_sheet = data[RUN_TIME_SHEET_NAME]
        else:
            runtime_sheet = data.create_sheet(RUN_TIME_SHEET_NAME)

        header_fill = PatternFill('solid', fgColor='4472C4')
        for column, header in enumerate(RUN_TIME_HEADERS, start=1):
            cell = runtime_sheet.cell(row=1, column=column, value=header)
            cell.font = Font(bold=True, color='FFFFFF')
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center', vertical='center')
        runtime_sheet.freeze_panes = 'A2'
        for column, width in enumerate((12, 12, 14, 32), start=1):
            letter = get_column_letter(column)
            runtime_sheet.column_dimensions[letter].width = width

        pending_position = run_context.get('runtime_log_pending_position')
        if pending_position:
            next_row, run_number = pending_position
        else:
            next_row = _runtime_log_next_row(runtime_sheet)
            run_number = _runtime_log_next_number(runtime_sheet)
            run_context['runtime_log_pending_position'] = (
                next_row, run_number
            )
        elapsed_seconds = max(
            0.0, run_context['finished_at'] - run_context['started_at']
        )
        result_text = f'成功{success_count}条，异常{abnormal_count}条'
        if unfinished_count:
            result_text += f'，未完成{unfinished_count}条'

        runtime_sheet.cell(next_row, 1, run_number)
        runtime_sheet.cell(next_row, 2, run_data_count)
        duration_cell = runtime_sheet.cell(
            next_row, 3, elapsed_seconds / 86400
        )
        duration_cell.number_format = '[h]:mm:ss'
        runtime_sheet.cell(next_row, 4, result_text)
        for column in range(1, len(RUN_TIME_HEADERS) + 1):
            runtime_sheet.cell(next_row, column).alignment = Alignment(
                horizontal='center', vertical='center'
            )

        output_path = _save_result_workbook(
            data, file_path, respect_task_state=False
        )
        run_context['runtime_log_recorded'] = True
        run_context.pop('runtime_log_pending_position', None)
        print(
            f'运行统计已写入“{RUN_TIME_SHEET_NAME}”第{next_row}行：'
            f'第{run_number}次，运行数据{run_data_count}条，'
            f'运行时间{_format_elapsed_time(elapsed_seconds)}，{result_text}'
        )
        if os.path.abspath(output_path) != os.path.abspath(file_path):
            print(f'     运行统计保存在程序回写文件：{output_path}')
    except Exception as error:
        print(f'运行统计写入失败：{error}')
    finally:
        run_context['runtime_log_in_progress'] = False
        if (
            data is not None
            and (owns_workbook or run_context.get('runtime_log_recorded'))
        ):
            try:
                data.close()
            except Exception:
                pass
            if run_context.get('workbook') is data:
                run_context['workbook'] = None


def _record_runtime_outcomes(rows, status):
    """仅记录本次实际成功回写的行；后写状态覆盖处理中标记。"""
    context_local = globals().get('run_context_local')
    if context_local is None:
        return
    run_context = getattr(context_local, 'value', None)
    if run_context is None or status not in (STATUS_SUCCESS, STATUS_ABNORMAL):
        return
    outcomes = run_context.setdefault('row_outcomes', {})
    for row in rows:
        row_no = row['_row'] if isinstance(row, dict) else int(row)
        outcomes[int(row_no)] = status


def write_abnormal(sheet, data, file_path, row_no, reason):
    """按实际底稿布局回写异常及原因。"""
    config = receipt_sheet_layout(sheet, EXCEL_CONFIG)
    sheet.cell(row_no, column=config["status_col"], value=STATUS_ABNORMAL)
    sheet.cell(row_no, column=config["reason_col"], value=reason)
    output_path = _save_result_workbook(data, file_path)
    _record_runtime_outcomes([row_no], STATUS_ABNORMAL)
    print(f'     已在底稿第{row_no}行回写: 异常 / {reason}')
    return output_path


def write_group_result(
    sheet, data, file_path, rows, status, reason='', bill_number=''
):
    """把同一张单据对应的所有底稿行一次性回写。"""
    config = receipt_sheet_layout(sheet, EXCEL_CONFIG)
    row_numbers = []
    for row in rows:
        row_no = row['_row']
        row_numbers.append(row_no)
        sheet.cell(
            row_no, column=config["status_col"], value=status
        )
        sheet.cell(
            row_no, column=config["reason_col"], value=reason
        )
        if bill_number:
            sheet.cell(
                row_no,
                column=config["bill_number_col"],
                value=bill_number
            )
    output_path = _save_result_workbook(data, file_path)
    _record_runtime_outcomes(rows, status)
    rows_text = '、'.join(str(row_no) for row_no in row_numbers)
    if reason:
        print(f'     已回写底稿第{rows_text}行: {status} / {reason}')
    else:
        print(f'     已回写底稿第{rows_text}行: {status}')
    if bill_number:
        print(
            f'     已将单据编号 {bill_number} 回写到底稿第'
            f'{rows_text}行{get_column_letter(config["bill_number_col"])}列'
        )
    return output_path


def format_error_reason(error, step_name=''):
    """生成适合写入“异常原因”列的明确错误原因。"""
    if isinstance(error, FieldFillError):
        return str(error)
    message = str(error).strip() or error.__class__.__name__
    if step_name:
        return f'{step_name}：{message}'
    return message


def write_group_error(
    sheet, data, file_path, rows, failed_row, reason
):
    """
    回写一张单据的异常：出错行写具体原因，并表其余行写统一原因。
    """
    if not rows:
        return
    config = receipt_sheet_layout(sheet, EXCEL_CONFIG)
    failed_row_no = (
        failed_row['_row'] if failed_row is not None else rows[0]['_row']
    )
    detailed_rows = []
    related_rows = []
    for row in rows:
        row_no = row['_row']
        row_reason = reason
        if len(rows) > 1 and row_no != failed_row_no:
            row_reason = '并表中数据存在错误'
            related_rows.append(row_no)
        else:
            detailed_rows.append(row_no)
        sheet.cell(
            row_no, column=config["status_col"], value=STATUS_ABNORMAL
        )
        sheet.cell(
            row_no, column=config["reason_col"], value=row_reason
        )

    output_path = _save_result_workbook(data, file_path)
    _record_runtime_outcomes(rows, STATUS_ABNORMAL)
    detailed_text = '、'.join(str(row_no) for row_no in detailed_rows)
    print(f'     已回写底稿第{detailed_text}行: 异常 / {reason}')
    if related_rows:
        related_text = '、'.join(str(row_no) for row_no in related_rows)
        print(
            f'     已回写并表关联底稿第{related_text}行: '
            '异常 / 并表中数据存在错误'
        )
    return output_path


def cancel_card(frame):
    """异常后点击“取消”及“确定”，确认回到列表页后才允许处理下一张单据。"""
    if _list_page_ready(frame):
        return True

    try:
        drawers = frame.locator(
            "css=div.card-table-modal-form-wrap-body"
        )
        for index in range(drawers.count() - 1, -1, -1):
            drawer = drawers.nth(index)
            if drawer.is_visible():
                close_detail_drawer(frame, drawer)
                break
    except Exception:
        pass

    for attempt in range(1, RETRY_MAX + 1):
        pause()
        if _list_page_ready(frame):
            print('     已检测到列表“新增”按钮')
            return True
        cancel_clicked = False
        buttons = frame.locator(
            "xpath=//button[contains(normalize-space(.),'取消')]"
        )
        for index in range(buttons.count() - 1, -1, -1):
            button = buttons.nth(index)
            try:
                if button.is_visible() and button.is_enabled():
                    pause()
                    button.click(
                        timeout=timing_ms('element_action_timeout')
                    )
                    cancel_clicked = True
                    print('     已点击“取消”，等待确认窗口')
                    break
            except Exception:
                continue
        if not cancel_clicked:
            continue

        confirm_clicked = False
        deadline = task_now() + timing_seconds('cancel_confirm_wait')
        while task_now() < deadline and not confirm_clicked:
            pause()
            confirms = frame.locator(
                "xpath=//button[contains(normalize-space(.),'确定') "
                "or contains(normalize-space(.),'确认')]"
            )
            try:
                for index in range(confirms.count() - 1, -1, -1):
                    confirm = confirms.nth(index)
                    if confirm.is_visible() and confirm.is_enabled():
                        pause()
                        confirm.click(
                            timeout=timing_ms('element_action_timeout')
                        )
                        confirm_clicked = True
                        print('     已点击“确定”，正在返回单据列表')
                        break
            except Exception:
                pass
            if not confirm_clicked:
                task_sleep(timing_seconds('field_poll_interval'))

        list_timeout = (
            timing_ms('list_return_wait')
            if confirm_clicked else timing_ms('tiny_element_timeout')
        )
        if _wait_for_list_page(frame, timeout=list_timeout):
            print('     异常单据已取消，已返回列表')
            return True
        print(f'     第{attempt}/{RETRY_MAX}次取消后仍未返回列表，准备重试')

    print('     异常单据取消失败，未能安全返回单据列表')
    return False


def close_ncc_save_error(frame, timeout=None):
    """保存/提交被 NCC 明确拒绝后，优先关闭错误提示，再退出当前卡片。"""
    if timeout is None:
        timeout = timing_ms('save_error_close_wait')
    deadline = task_now() + timeout / 1000
    while task_now() < deadline:
        pause()
        close_buttons = frame.locator(
            "xpath=//button[normalize-space(.)='关闭' or @title='关闭']"
        )
        try:
            for index in range(close_buttons.count() - 1, -1, -1):
                button = close_buttons.nth(index)
                if button.is_visible() and button.is_enabled():
                    button.click(
                        timeout=timing_ms('element_action_timeout')
                    )
                    print('     已点击保存报错提示的“关闭”')
                    task_sleep(timing_seconds('short_stable_wait'))
                    return True
        except Exception:
            pass
        task_sleep(timing_seconds('field_poll_interval'))
    print('     未找到可点击的保存报错“关闭”按钮，继续尝试取消当前单据')
    return False


def keep_browser_open(browser, context):
    """流程结束后保持浏览器运行，直到用户手动关闭全部浏览器页面。"""
    print('本阶段填写完成，浏览器将保持打开，请检查页面并在需要时手动关闭。')
    while True:
        if stop_event.is_set():
            print('收到停止指令，正在关闭程序打开的浏览器......')
            try:
                context.close()
            except Exception:
                pass
            try:
                browser.close()
            except Exception:
                pass
            break
        try:
            if not browser.is_connected():
                break
            if not any(not page.is_closed() for page in context.pages):
                break
        except Exception:
            break
        time.sleep(timing_seconds('browser_close_poll_interval'))
    print('已检测到浏览器关闭，程序本次任务结束。')


def running_task(
    file_path, save_mode, merge_basis='company_account_date', run_context=None,
    single_company_only=False
):
    """按 GUI 选项分组制单，最后执行保存或提交。"""
    global active_browser, active_context, active_cdp_sessions
    global last_cdp_activity_refresh_at
    active_cdp_sessions = []
    last_cdp_activity_refresh_at = 0.0
    if run_context is None:
        run_context = {
            'started_at': time.monotonic(),
            'row_outcomes': {},
            'source_validated': False,
        }
    else:
        run_context.setdefault('started_at', time.monotonic())
        run_context.setdefault('row_outcomes', {})
        run_context.setdefault('source_validated', False)
    run_context_local.value = run_context
    submit_after_fill = save_mode == 'submit'
    action_text = '保存提交' if submit_after_fill else '保存'
    selected_merge_basis = merge_basis_label(merge_basis)
    if os.path.splitext(file_path)[1].lower() not in ('.xlsx', '.xlsm'):
        raise FieldFillError(
            '底稿文件',
            '请使用新版 .xlsm 底稿；兼容已有 .xlsx 数据，但旧文件没有无损粘贴保护'
        )
    print('正在获取EXCEL底稿数据信息......')
    try:
        rows = get_data(file_path, EXCEL_CONFIG["sheet_name"])
        run_context['source_validated'] = True
        run_context['selected_row_numbers'] = [row['_row'] for row in rows]
        print(f'获取EXCEL底稿数据成功!!!!!! 共 {len(rows)} 行数据')
    except Exception as e:
        print(f'获取EXCEL底稿数据失败: {e}')
        print('请检查所选文件是否为“收款结算单底稿导入模板.xlsx”')
        return

    if not rows:
        print(
            '填写状态筛选后，没有“填写状态”为空或“异常”的待处理数据，'
            '本次不打开浏览器。'
        )
        return

    try:
        data = load_receipt_workbook(file_path)
        sheet = data[EXCEL_CONFIG["sheet_name"]]
        run_context['workbook'] = data
    except KeyError:
        print(f'所选文件中没有工作表【{EXCEL_CONFIG["sheet_name"]}】')
        print('请在 EXCEL_CONFIG["sheet_name"] 中修改为正确的工作表名后重试。')
        return
    except Exception as e:
        print(f'打开底稿文件失败: {e}')
        return

    # AS 是分组控制字段，必须在 build_bill_groups 前全量预检。
    grouping_invalid_rows = []
    for row in rows:
        grouping_error = validate_grouping_fields(row)
        if grouping_error:
            grouping_invalid_rows.append((row, grouping_error))
            write_group_error(
                sheet, data, file_path, [row], row, grouping_error
            )
    if grouping_invalid_rows:
        print(
            '发现并表控制字段错误。为避免错字导致残缺单据，'
            '本次不会处理其他数据；请修正后重试。'
        )
        return

    print(f'本次并表依据：{selected_merge_basis}')
    identified_groups = build_bill_groups(rows, merge_basis=merge_basis)
    bill_groups = []
    for group in identified_groups:
        invalid_row = None
        validation_error = ''
        for row in group:
            validation_error = validate_source_row(row)
            if validation_error:
                invalid_row = row
                break
        if invalid_row is not None:
            write_group_error(
                sheet, data, file_path, group, invalid_row,
                validation_error
            )
            continue
        invalid_row, validation_error = validate_group_header(group)
        if invalid_row is not None:
            write_group_error(
                sheet, data, file_path, group, invalid_row,
                validation_error
            )
            continue
        bill_groups.append(group)

    if not bill_groups:
        print('底稿数据校验完成，没有可继续处理的数据。')
        return

    pending_row_count = sum(len(group) for group in bill_groups)
    print(
        f'并表识别完成：{pending_row_count} 行有效数据，'
        f'将生成 {len(bill_groups)} 张单据；'
        f'依据“{selected_merge_basis}”；完成后执行“{action_text}”。'
    )
    for index, group in enumerate(bill_groups, start=1):
        if len(group) > 1:
            row_numbers = '、'.join(str(row['_row']) for row in group)
            print(f'     并表组{index}: 底稿第{row_numbers}行，共{len(group)}条明细')

    # Chromium 及其渲染器由本进程启动，Windows 会继承此优先级。
    raise_automation_process_priority()
    with sync_playwright() as p:
        print('正在打开浏览器......')
        chromium = p.chromium
        executable_path = resource_path(NCC_CONFIG["chrome_path"])
        if not os.path.isfile(executable_path):
            raise FieldFillError(
                '内置浏览器',
                f'未找到Chrome：{executable_path}。请确认“ms-playwright（勿动）”文件夹完整。'
            )
        browser = chromium.launch(
            executable_path=executable_path,
            headless=False,
            # Windows 将最小化窗口视为“被遮挡”，Chromium 默认会降低渲染器和
            # 定时器优先级，导致 NCC 的下拉、弹窗不再及时刷新。关闭这些后台
            # 节流后，用户可最小化浏览器，自动填写仍会继续执行。
            args=[
                '--start-fullscreen',
                '--start-maximized',
                '--disable-backgrounding-occluded-windows',
                '--disable-renderer-backgrounding',
                '--disable-background-timer-throttling',
                '--disable-features='
                'CalculateNativeWinOcclusion,IntensiveWakeUpThrottling',
            ]
        )
        active_browser = browser
        context = browser.new_context(
            ignore_https_errors=True,
            bypass_csp=True,
            no_viewport=True
        )
        # 让未单独指定 timeout 的 Playwright 操作也服从本批次时间设置。
        # 这里读取的是任务启动时快照，运行中即使配置文件被外部修改也不会变化。
        context.set_default_timeout(timing_ms('long_element_timeout'))
        context.set_default_navigation_timeout(timing_ms('page_load_timeout'))
        active_context = context
        main_page = context.new_page()

        main_page.goto(
            NCC_CONFIG["login_url"],
            timeout=timing_ms('page_load_timeout')
        )
        login(main_page)
        new_page = operate_page(main_page, NCC_CONFIG["node_name"], context)
        keep_page_active_when_minimized(context, new_page)
        frame = get_work_frame(new_page)
        print(f'新页面: {NCC_CONFIG["node_name"]}已确认打开，正在根据明细制单......')

        print(
            f'开始处理 {len(bill_groups)} 张单据'
            '（制单过程中请关闭EXCEL导入模版文件，避免影响数据读写）......'
        )

        next_card_already_open = False
        for bill_index, group in enumerate(bill_groups, start=1):
            card_already_open = next_card_already_open
            next_card_already_open = False
            first_row = group[0]
            active_row = first_row
            active_step = '新增单据'
            row_numbers = '、'.join(str(row['_row']) for row in group)
            first_transaction_object = transaction_object_of(first_row)
            remote_committed = False
            processing_marker_written = False
            print(
                f'正在处理第{bill_index}/{len(bill_groups)}张单据'
                f'（底稿第{row_numbers}行）：'
                f'公司编码={first_row["company_code"]}，'
                f'公司名称={first_row["company_name"]}，'
                f'单据日期={first_row["bill_date"]}，'
                f'明细数={len(group)}'
            )
            try:
                if single_company_only:
                    print(
                        '     已勾选“是否仅有单家公司权限”，'
                        '跳过结算财务组织，等待并填写单据日期'
                    )
                    active_step = '单据日期'
                    fill_bill_date_with_card_recovery(
                        frame, first_row["bill_date"],
                        card_already_open=card_already_open
                    )
                else:
                    active_step = '结算财务组织'
                    fill_settle_org_with_card_recovery(
                        frame, first_row["company_code"],
                        card_already_open=card_already_open
                    )
                    active_step = '单据日期'
                    fill_bill_date(frame, first_row["bill_date"])
                active_step = '卡片头字段'
                fill_header_fields(
                    frame, first_row, first_transaction_object
                )
                active_step = '明细填写'
                fill_detail_row(
                    frame, first_row, first_transaction_object
                )

                previous_detail_row = first_row
                for detail_index, detail_row in enumerate(
                    group[1:], start=2
                ):
                    active_row = detail_row
                    detail_transaction_object = transaction_object_of(detail_row)
                    print(
                        f'     开始填写并表明细第{detail_index}/{len(group)}条'
                        f'（底稿第{detail_row["_row"]}行）'
                    )
                    active_step = '明细增行'
                    click_add_detail_row(frame)
                    active_step = '明细填写'
                    fill_detail_row(
                        frame,
                        detail_row,
                        detail_transaction_object,
                        previous_row=previous_detail_row
                    )
                    previous_detail_row = detail_row

                active_row = first_row
                active_step = '付款银行账号'
                fill_payer_bank_account_before_save(frame, first_row)

                active_row = first_row
                active_step = f'{action_text}前防重复标记'
                processing_status = STATUS_ABNORMAL
                marker_reason = (
                    f'程序已进入“{action_text}”阶段；若程序中断，'
                    '请先在NCC核对结果，禁止直接重跑'
                )
                marker_path = write_group_result(
                    sheet, data, file_path, group,
                    processing_status, reason=marker_reason
                )
                if os.path.abspath(marker_path) != os.path.abspath(file_path):
                    raise ResultWorkbookRedirectedError(
                        '底稿回写',
                        f'原底稿仍被占用，防重复标记只能写入“{marker_path}”；'
                        '本单尚未点击NCC保存。请关闭Excel并改用该回写文件继续。'
                    )
                processing_marker_written = True

                active_step = action_text
                saved_bill_number = save(frame, submit_after_fill)
                remote_committed = True
                result_status = STATUS_SUCCESS
                result_reason = ''
                if not submit_after_fill and not saved_bill_number:
                    active_step = '读取保存后的单据编号'
                    try:
                        saved_bill_number = read_saved_bill_number(frame)
                    except Exception as number_error:
                        result_status = STATUS_SUCCESS
                        result_reason = (
                            'NCC已确认保存成功，但未读取到单据编号；'
                            f'请人工核对，勿重复制单。原因：{number_error}'
                        )
                active_step = '结果回写'
                write_group_result(
                    sheet, data, file_path, group, result_status,
                    reason=result_reason,
                    bill_number=saved_bill_number
                )
                print(
                    f'第{bill_index}张单据已完成“{action_text}”，'
                    f'共{len(group)}条明细。'
                )

            except TaskCancelledError:
                if remote_committed:
                    stop_reason = (
                        f'NCC已确认完成“{action_text}”，但用户在结果回写完成前停止；'
                        '请人工核对，勿重复制单'
                    )
                    try:
                        write_group_result(
                            sheet, data, file_path, group,
                            STATUS_SUCCESS, reason=stop_reason
                        )
                    except Exception as write_error:
                        print(f'     停止后的成功状态回写失败：{write_error}')
                elif processing_marker_written:
                    stop_reason = (
                        f'用户在点击NCC“{action_text}”按钮前停止任务；'
                        '本单尚未发出保存请求，可重新运行'
                    )
                    try:
                        write_group_error(
                            sheet, data, file_path, group, active_row,
                            stop_reason
                        )
                    except Exception as write_error:
                        print(f'     停止后的可重跑状态回写失败：{write_error}')
                raise

            except NccSaveRejectedError as e:
                reason_prefix = f'{e.field_name}结果需人工核对：'
                error_reason = e.reason
                if not error_reason.startswith(reason_prefix):
                    error_reason = reason_prefix + error_reason
                print(f'     {error_reason}')
                close_ncc_save_error(frame)
                returned_to_list = cancel_card(frame)
                try:
                    write_group_result(
                        sheet, data, file_path, group,
                        STATUS_ABNORMAL, reason=error_reason
                    )
                except Exception as write_error:
                    print(
                        '     异常状态回写失败；此前的异常防重复标记仍保留：'
                        f'{write_error}'
                    )
                if not returned_to_list:
                    print(
                        '     保存报错后未能安全返回列表，后续单据不再继续执行。'
                    )
                    break
                if bill_index < len(bill_groups):
                    try:
                        click_add_bill(frame)
                        next_card_already_open = True
                        print('     保存报错单据已跳过并回写，已点击“新增”继续下一张单据')
                    except Exception as add_error:
                        print(
                            '     保存报错单据已回写，但未能点击“新增”，'
                            f'后续单据已停止：{add_error}'
                        )
                        break
                continue

            except NccCommitUncertainError as e:
                reason_prefix = f'{e.field_name}结果需人工核对：'
                error_reason = e.reason
                if not error_reason.startswith(reason_prefix):
                    error_reason = reason_prefix + error_reason
                print(f'     {error_reason}')
                try:
                    write_group_result(
                        sheet, data, file_path, group,
                        STATUS_ABNORMAL, reason=error_reason
                    )
                except Exception as write_error:
                    print(
                        '     异常状态回写失败；此前的异常防重复标记仍保留：'
                        f'{write_error}'
                    )
                break

            except ResultWorkbookRedirectedError as e:
                error_reason = format_error_reason(e, active_step)
                print(f'     {error_reason}')
                pause()
                returned_to_list = cancel_card(frame)
                try:
                    write_group_error(
                        sheet, data, file_path, group, active_row,
                        error_reason
                    )
                except Exception as write_error:
                    print(f'     备用底稿异常回写失败：{write_error}')
                print('     为避免原底稿没有防重复记录，本批次停止。')
                if not returned_to_list:
                    print('     当前卡片未能自动取消，请在浏览器中人工检查。')
                break

            except Exception as e:
                error_reason = format_error_reason(e, active_step)
                if remote_committed:
                    review_reason = (
                        f'NCC已确认完成“{action_text}”，但后处理失败：{error_reason}；'
                        '请人工核对，勿重复制单'
                    )
                    print(f'     {review_reason}')
                    try:
                        write_group_result(
                            sheet, data, file_path, group,
                            STATUS_SUCCESS, reason=review_reason
                        )
                    except Exception as write_error:
                        print(
                            '     成功状态回写失败；此前的异常防重复标记仍保留：'
                            f'{write_error}'
                        )
                    break
                print(f'     处理异常: {error_reason}')
                pause()
                returned_to_list = cancel_card(frame)
                try:
                    write_group_error(
                        sheet, data, file_path, group, active_row,
                        error_reason
                    )
                except Exception as write_error:
                    print(
                        '     异常结果回写失败，但程序将继续清理页面：'
                        f'{write_error}'
                    )
                if not returned_to_list:
                    print(
                        '     无法确认已返回列表，为避免在异常页面误填，'
                        '后续单据不再继续执行。'
                    )
                    remaining_groups = bill_groups[bill_index:]
                    for remaining_group in remaining_groups:
                        try:
                            write_group_error(
                                sheet, data, file_path, remaining_group,
                                remaining_group[0],
                                '异常单据取消后未能返回列表，未执行后续制单'
                            )
                        except Exception as write_error:
                            print(
                                '     后续异常回写失败：'
                                f'{write_error}'
                            )
                    break
                if bill_index < len(bill_groups):
                    print('     将自动点击“新增”，继续填写下一张单据')
                continue

            if bill_index < len(bill_groups):
                try:
                    back(frame)
                except Exception as e:
                    print(f'     无法继续下一张单据: {e}')
                    remaining_groups = bill_groups[bill_index:]
                    for remaining_group in remaining_groups:
                        write_group_error(
                            sheet, data, file_path, remaining_group,
                            remaining_group[0],
                            f'返回列表：前一张单据完成后无法返回列表: {e}'
                        )
                    break

        _finalize_runtime_log(file_path, run_context)
        keep_browser_open(browser, context)


# ============================ GUI 部分 ============================
print('====================欢迎使用收款结算单自动化制单小程序====================')
print('程序正在加载中，请稍等（若出现卡顿无响应请重新打开程序）......')


def _prepare_packaged_tk_runtime():
    """修复部分 Windows Python 环境中 Tcl 未初始化可执行路径的问题。"""
    if not getattr(sys, 'frozen', False) or os.name != 'nt':
        return
    try:
        import ctypes

        runtime_root = sys._MEIPASS
        tcl_dll_path = os.path.join(runtime_root, 'tcl86t.dll')
        tcl_dll = ctypes.WinDLL(tcl_dll_path)
        tcl_dll.Tcl_FindExecutable.argtypes = (ctypes.c_char_p,)
        # _MEIPASS 位于临时目录，路径不受 EXE 所在目录中文字符影响。
        tcl_dll.Tcl_FindExecutable(
            tcl_dll_path.encode('mbcs', errors='replace')
        )
        os.environ['TCL_LIBRARY'] = os.path.join(runtime_root, '_tcl_data')
        os.environ['TK_LIBRARY'] = os.path.join(runtime_root, '_tk_data')
    except Exception:
        pass


_prepare_packaged_tk_runtime()
root = tk.Tk()
style = ttk.Style(root)
root.title('收款结算单工具')
root.resizable(False, False)
style.configure('TLabelframe.Label', font=('微软雅黑', 10, 'bold'))
style.configure('TRadiobutton', font=('微软雅黑', 9))
style.configure('Merge.TRadiobutton', font=('微软雅黑', 8))

frame1 = ttk.Frame(root, padding=(12, 4, 12, 0))
frame1.pack(fill='x')

photo = None
image_path = resource_path('./ms-playwright（勿动）/image（勿动）.png')
try:
    image = Image.open(image_path)
    image = image.resize((150, 115))
    photo = ImageTk.PhotoImage(image)
    label = tk.Label(frame1, image=photo)
except (OSError, ValueError) as image_error:
    print(f'背景图加载失败，改用文字标题：{image_error}')
    label = tk.Label(
        frame1, text='收款结算单工具',
        font=('微软雅黑', 18, 'bold'), width=20, height=5
    )
label.pack(pady=(0, 4))

save_mode_frame = ttk.LabelFrame(
    frame1, text='填写完成后操作', padding=(10, 8)
)
save_mode_frame.pack(pady=(3, 7))

save_mode_var = tk.IntVar(value=0)
save_mode_text = tk.StringVar(value='当前：保存')


def update_save_mode(value):
    mode = '提交' if round(float(value)) == 1 else '保存'
    save_mode_text.set(f'当前：{mode}')


ttk.Label(save_mode_frame, text='保存').grid(
    row=0, column=0, padx=(0, 6)
)
save_mode_slider = tk.Scale(
    save_mode_frame,
    from_=0,
    to=1,
    resolution=1,
    orient=tk.HORIZONTAL,
    showvalue=False,
    variable=save_mode_var,
    command=update_save_mode,
    length=105,
    sliderlength=24,
    width=16,
    highlightthickness=0,
    borderwidth=0,
)
save_mode_slider.grid(row=0, column=1)
ttk.Label(save_mode_frame, text='提交').grid(
    row=0, column=2, padx=(6, 0)
)
ttk.Label(save_mode_frame, textvariable=save_mode_text).grid(
    row=1, column=0, columnspan=3, pady=(4, 1)
)

# ---------- 并表依据选择 ----------
merge_basis_frame = ttk.LabelFrame(
    frame1, text='并表依据选择', padding=(10, 5)
)
merge_basis_frame.pack(fill='x', pady=(0, 2))

merge_basis_var = tk.StringVar(value='company_account_date')

ttk.Radiobutton(
    merge_basis_frame,
    text='同一公司、同一账号、同一日期',
    value='company_account_date',
    variable=merge_basis_var,
    style='Merge.TRadiobutton',
).pack(anchor='w', pady=(0, 5))

ttk.Radiobutton(
    merge_basis_frame,
    text='同一公司、同一账号、同一日期、同一客商',
    value='company_account_date_customer',
    variable=merge_basis_var,
    style='Merge.TRadiobutton',
).pack(anchor='w')

# ---------- 单家公司权限 ----------
single_company_only_var = tk.BooleanVar(value=False)
ttk.Checkbutton(
    frame1,
    text='是否仅有单家公司权限',
    variable=single_company_only_var,
).pack(anchor='w', pady=(5, 1))

# ---------- 填写顺序 / 停用设置（表头与表头、明细与明细各自独立调序） ----------
fill_order_settings_window = None
fill_order_settings_button = None


def _close_fill_order_settings_window():
    """关闭填写顺序设置弹窗并清理模态状态。"""
    global fill_order_settings_window
    window = fill_order_settings_window
    fill_order_settings_window = None
    if window is None:
        return
    try:
        window.grab_release()
    except tk.TclError:
        pass
    try:
        window.destroy()
    except tk.TclError:
        pass


def set_fill_order_settings_running(running):
    """任务运行期间禁止进入或修改填写顺序设置。"""
    if fill_order_settings_button is not None:
        fill_order_settings_button.configure(
            state='disabled' if running else 'normal'
        )
    if running and fill_order_settings_window is not None:
        _close_fill_order_settings_window()


def open_fill_order_settings():
    """打开可持久化的填写顺序与停用设置窗口（表头/明细严格独立）。"""
    global current_fill_order_settings, fill_order_settings_window
    if current_thread is not None and current_thread.is_alive():
        messagebox.showwarning(
            '任务运行中',
            '任务运行期间不能修改填写顺序或停用设置。\n'
            '请等待当前任务结束后再调整。',
            parent=root,
        )
        return
    if fill_order_settings_window is not None:
        try:
            fill_order_settings_window.deiconify()
            fill_order_settings_window.lift()
            fill_order_settings_window.focus_force()
            return
        except tk.TclError:
            fill_order_settings_window = None

    dialog = tk.Toplevel(root)
    fill_order_settings_window = dialog
    dialog.title('填写顺序与停用设置（表头/明细独立调序）')
    dialog.resizable(True, True)
    dialog.transient(root)
    dialog.protocol('WM_DELETE_WINDOW', _close_fill_order_settings_window)

    notice_frame = ttk.Frame(dialog, padding=(12, 10, 12, 6))
    notice_frame.pack(fill='x')
    ttk.Label(
        notice_frame,
        text=(
            '提示：仅允许表头与表头、明细与明细各自内部互换顺序，'
            '不允许表头与明细之间互换顺序。'
        ),
        foreground='#B22222',
        justify='left',
        font=('微软雅黑', 9, 'bold'),
    ).pack(anchor='w')
    ttk.Label(
        notice_frame,
        text=(
            '1. 表头固定最前：“结算财务组织”“单据日期”承担建单与单公司权限切换，不参与顺序互换。\n'
            '2. 保存前最后填入：“付款银行账号”固定在所有数据（含全部明细）填妥后、点击保存前最后一个填入，不参与表头/明细调序。\n'
            '3. 仅标记为[可停用]的选填步骤允许停用；保存时会自动检查联动先后关系，并从下一次运行开始生效。'
        ),
        foreground='#444444',
        justify='left',
        font=('微软雅黑', 9),
    ).pack(anchor='w', pady=(4, 0))

    snapshot = get_fill_order_settings_snapshot()
    local_header_order = list(snapshot['header_order'])
    local_detail_order = list(snapshot['detail_order'])
    local_disabled = set(snapshot['disabled'])
    payer_disabled_var = tk.BooleanVar(
        value='payer_bank_account' in local_disabled
    )

    body_frame = ttk.Frame(dialog, padding=(12, 2, 12, 6))
    body_frame.pack(fill='both', expand=True)
    body_frame.columnconfigure(0, weight=1)
    body_frame.columnconfigure(1, weight=1)
    body_frame.rowconfigure(0, weight=1)

    # ---- 左侧：表头独立调序 + 保存前最后填入（付款银行账号） ----
    left_column = ttk.Frame(body_frame)
    left_column.grid(row=0, column=0, sticky='nsew', padx=(0, 8))
    left_column.rowconfigure(0, weight=1)
    left_column.columnconfigure(0, weight=1)

    header_group = ttk.LabelFrame(
        left_column, text='表头字段（仅表头内部互换）', padding=(8, 6)
    )
    header_group.grid(row=0, column=0, sticky='nsew', pady=(0, 8))
    header_group.columnconfigure(0, weight=1)
    header_group.rowconfigure(0, weight=1)

    header_listbox = tk.Listbox(
        header_group, width=32, height=6,
        activestyle='none', exportselection=False, font=('微软雅黑', 9)
    )
    header_listbox.grid(row=0, column=0, sticky='nsew')

    def selected_header_index():
        selection = header_listbox.curselection()
        return selection[0] if selection else None

    def refresh_header_list(select_index=None):
        previous_index = (
            selected_header_index() if select_index is None else select_index
        )
        header_listbox.delete(0, tk.END)
        for pos, step_key in enumerate(local_header_order, start=1):
            disabled_text = '已停用' if step_key in local_disabled else '已启用'
            allowed_text = (
                '可停用' if step_key in HEADER_FILL_OPTIONAL_STEPS else '必要/联动'
            )
            header_listbox.insert(
                tk.END,
                f'{pos}. [{disabled_text}] [{allowed_text}] '
                f'{HEADER_FILL_STEP_LABELS[step_key]}'
            )
        if previous_index is not None and local_header_order:
            index = min(max(previous_index, 0), len(local_header_order) - 1)
            header_listbox.selection_set(index)
            header_listbox.activate(index)
            header_listbox.see(index)

    def move_header_selected(direction):
        index = selected_header_index()
        if index is None:
            messagebox.showinfo('提示', '请先选择一个表头字段。', parent=dialog)
            return
        target_index = index + direction
        if target_index < 0 or target_index >= len(local_header_order):
            return
        local_header_order[index], local_header_order[target_index] = (
            local_header_order[target_index], local_header_order[index]
        )
        refresh_header_list(target_index)

    def set_header_selected_disabled(disabled):
        index = selected_header_index()
        if index is None:
            messagebox.showinfo('提示', '请先选择一个表头字段。', parent=dialog)
            return
        step_key = local_header_order[index]
        if step_key not in HEADER_FILL_OPTIONAL_STEPS:
            messagebox.showwarning(
                '不能停用',
                f'表头“{HEADER_FILL_STEP_LABELS[step_key]}”为必填或联动步骤，不能停用。',
                parent=dialog,
            )
            return
        if disabled:
            local_disabled.add(step_key)
        else:
            local_disabled.discard(step_key)
        refresh_header_list(index)

    header_actions = ttk.Frame(header_group)
    header_actions.grid(row=1, column=0, sticky='ew', pady=(6, 0))
    ttk.Button(
        header_actions, text='上移', width=6,
        command=lambda: move_header_selected(-1)
    ).pack(side='left')
    ttk.Button(
        header_actions, text='下移', width=6,
        command=lambda: move_header_selected(1)
    ).pack(side='left', padx=(4, 0))
    ttk.Button(
        header_actions, text='停用', width=6,
        command=lambda: set_header_selected_disabled(True)
    ).pack(side='left', padx=(12, 0))
    ttk.Button(
        header_actions, text='启用', width=6,
        command=lambda: set_header_selected_disabled(False)
    ).pack(side='left', padx=(4, 0))

    post_detail_group = ttk.LabelFrame(
        left_column, text='保存前最后填入（全部明细填妥后）', padding=(8, 6)
    )
    post_detail_group.grid(row=1, column=0, sticky='ew')
    ttk.Label(
        post_detail_group,
        text='固定执行位置：全部明细填写完毕 → 付款银行账号 → 点击保存/保存提交',
        foreground='#444444',
        wraplength=275,
        justify='left',
        font=('微软雅黑', 8),
    ).pack(anchor='w')
    payer_status_var = tk.StringVar()

    def sync_payer_disabled():
        if payer_disabled_var.get():
            local_disabled.add('payer_bank_account')
            payer_status_var.set('[已停用] [可停用] 付款银行账号（保存前跳过）')
        else:
            local_disabled.discard('payer_bank_account')
            payer_status_var.set('[已启用] [可停用] 付款银行账号（保存前最后填入）')

    ttk.Label(
        post_detail_group,
        textvariable=payer_status_var,
        font=('微软雅黑', 9),
    ).pack(anchor='w', pady=(4, 2))
    ttk.Checkbutton(
        post_detail_group,
        text='停用“付款银行账号”（保存前跳过填写）',
        variable=payer_disabled_var,
        command=sync_payer_disabled,
    ).pack(anchor='w', pady=(2, 0))

    # ---- 右侧：明细独立调序 ----
    detail_group = ttk.LabelFrame(
        body_frame, text='明细字段（仅明细内部互换）', padding=(8, 6)
    )
    detail_group.grid(row=0, column=1, sticky='nsew')
    detail_group.columnconfigure(0, weight=1)
    detail_group.rowconfigure(0, weight=1)

    detail_list_wrap = ttk.Frame(detail_group)
    detail_list_wrap.grid(row=0, column=0, sticky='nsew')
    detail_list_wrap.columnconfigure(0, weight=1)
    detail_list_wrap.rowconfigure(0, weight=1)

    detail_listbox = tk.Listbox(
        detail_list_wrap, width=38, height=14,
        activestyle='none', exportselection=False, font=('微软雅黑', 9)
    )
    detail_scrollbar = ttk.Scrollbar(
        detail_list_wrap, orient='vertical', command=detail_listbox.yview
    )
    detail_listbox.configure(yscrollcommand=detail_scrollbar.set)
    detail_listbox.grid(row=0, column=0, sticky='nsew')
    detail_scrollbar.grid(row=0, column=1, sticky='ns', padx=(4, 0))

    def selected_detail_index():
        selection = detail_listbox.curselection()
        return selection[0] if selection else None

    def refresh_detail_list(select_index=None):
        previous_index = (
            selected_detail_index() if select_index is None else select_index
        )
        detail_listbox.delete(0, tk.END)
        for pos, step_key in enumerate(local_detail_order, start=1):
            disabled_text = '已停用' if step_key in local_disabled else '已启用'
            allowed_text = (
                '可停用' if step_key in DETAIL_FILL_OPTIONAL_STEPS else '必要/联动'
            )
            detail_listbox.insert(
                tk.END,
                f'{pos:02d}. [{disabled_text}] [{allowed_text}] '
                f'{DETAIL_FILL_STEP_LABELS[step_key]}'
            )
        if previous_index is not None and local_detail_order:
            index = min(max(previous_index, 0), len(local_detail_order) - 1)
            detail_listbox.selection_set(index)
            detail_listbox.activate(index)
            detail_listbox.see(index)

    def move_detail_selected(direction):
        index = selected_detail_index()
        if index is None:
            messagebox.showinfo('提示', '请先选择一个明细字段。', parent=dialog)
            return
        target_index = index + direction
        if target_index < 0 or target_index >= len(local_detail_order):
            return
        local_detail_order[index], local_detail_order[target_index] = (
            local_detail_order[target_index], local_detail_order[index]
        )
        refresh_detail_list(target_index)

    def set_detail_selected_disabled(disabled):
        index = selected_detail_index()
        if index is None:
            messagebox.showinfo('提示', '请先选择一个明细字段。', parent=dialog)
            return
        step_key = local_detail_order[index]
        if step_key not in DETAIL_FILL_OPTIONAL_STEPS:
            messagebox.showwarning(
                '不能停用',
                f'明细“{DETAIL_FILL_STEP_LABELS[step_key]}”为必填或联动步骤，不能停用。',
                parent=dialog,
            )
            return
        if disabled:
            local_disabled.add(step_key)
        else:
            local_disabled.discard(step_key)
        refresh_detail_list(index)

    detail_actions = ttk.Frame(detail_group)
    detail_actions.grid(row=1, column=0, sticky='ew', pady=(6, 0))
    ttk.Button(
        detail_actions, text='上移', width=6,
        command=lambda: move_detail_selected(-1)
    ).pack(side='left')
    ttk.Button(
        detail_actions, text='下移', width=6,
        command=lambda: move_detail_selected(1)
    ).pack(side='left', padx=(4, 0))
    ttk.Button(
        detail_actions, text='停用', width=6,
        command=lambda: set_detail_selected_disabled(True)
    ).pack(side='left', padx=(12, 0))
    ttk.Button(
        detail_actions, text='启用', width=6,
        command=lambda: set_detail_selected_disabled(False)
    ).pack(side='left', padx=(4, 0))

    # ---- 底部操作栏 ----
    bottom_frame = ttk.Frame(dialog, padding=(12, 4, 12, 12))
    bottom_frame.pack(fill='x')

    def restore_default():
        local_header_order[:] = list(HEADER_FILL_DEFAULT_ORDER)
        local_detail_order[:] = list(DETAIL_FILL_DEFAULT_ORDER)
        local_disabled.clear()
        payer_disabled_var.set(False)
        sync_payer_disabled()
        refresh_header_list(0)
        refresh_detail_list(0)

    def save_and_close():
        global current_fill_order_settings
        sync_payer_disabled()
        try:
            current_fill_order_settings = save_fill_order_settings({
                'header_order': list(local_header_order),
                'detail_order': list(local_detail_order),
                'disabled': sorted(local_disabled),
            })
        except Exception as error:
            messagebox.showerror(
                '无法保存设置',
                f'设置未保存：{error}',
                parent=dialog,
            )
            return
        disabled_set = set(current_fill_order_settings['disabled'])
        header_labels = '、'.join(
            HEADER_FILL_STEP_LABELS[key]
            for key in current_fill_order_settings['header_order']
            if key not in disabled_set
        )
        detail_labels = '、'.join(
            DETAIL_FILL_STEP_LABELS[key]
            for key in current_fill_order_settings['detail_order']
            if key not in disabled_set
        )
        payer_state = (
            '已停用' if 'payer_bank_account' in disabled_set
            else '保存前最后填入'
        )
        print(f'已保存表头填写顺序：{header_labels}')
        print(f'已保存明细填写顺序：{detail_labels}')
        print(f'付款银行账号状态：{payer_state}')
        if current_fill_order_settings['disabled']:
            print(
                '已停用字段：'
                + '、'.join(
                    ALL_FILL_STEP_LABELS[key]
                    for key in current_fill_order_settings['disabled']
                )
            )
        messagebox.showinfo(
            '填写顺序设置',
            f'已保存到：\n{fill_order_settings_path()}\n\n下一次运行将使用新顺序。',
            parent=dialog,
        )
        _close_fill_order_settings_window()

    ttk.Button(
        bottom_frame, text='恢复默认', command=restore_default
    ).pack(side='left')
    ttk.Button(
        bottom_frame, text='取消',
        command=_close_fill_order_settings_window, width=10
    ).pack(side='right')
    ttk.Button(
        bottom_frame, text='保存设置', command=save_and_close, width=12
    ).pack(side='right', padx=(0, 8))

    sync_payer_disabled()
    refresh_header_list(0)
    refresh_detail_list(0)
    dialog.grab_set()
    dialog.focus_force()


fill_order_settings_frame = ttk.Frame(frame1)
fill_order_settings_frame.pack(fill='x', pady=(4, 1))
ttk.Label(
    fill_order_settings_frame, text='填写顺序设置：', font=('微软雅黑', 9)
).pack(side='left')
fill_order_settings_button = ttk.Button(
    fill_order_settings_frame,
    text='设置顺序/停用',
    command=open_fill_order_settings,
    width=14,
)
fill_order_settings_button.pack(side='left', padx=(8, 0))

# ---------- 填报时间设置 ----------
time_settings_window = None
time_settings_button = None


def _close_time_settings_window():
    """关闭时间设置弹窗并清理模态状态。"""
    global time_settings_window
    window = time_settings_window
    time_settings_window = None
    if window is None:
        return
    try:
        window.grab_release()
    except tk.TclError:
        pass
    try:
        window.destroy()
    except tk.TclError:
        pass


def set_time_settings_running(running):
    """任务运行期间禁止进入或修改时间设置。"""
    if time_settings_button is not None:
        time_settings_button.configure(
            state='disabled' if running else 'normal'
        )
    if running and time_settings_window is not None:
        _close_time_settings_window()


def open_time_settings():
    """显示全部填报等待参数；保存后供下一次任务启动时快照使用。"""
    global current_time_settings, time_settings_window
    if current_thread is not None and current_thread.is_alive():
        messagebox.showwarning(
            '任务正在运行',
            '任务运行期间不能调整时间，请在本批次结束后再设置。',
            parent=root,
        )
        return
    if time_settings_window is not None:
        try:
            time_settings_window.deiconify()
            time_settings_window.lift()
            time_settings_window.focus_force()
            return
        except tk.TclError:
            time_settings_window = None

    window = tk.Toplevel(root)
    time_settings_window = window
    window.title('填报时间设置')
    window.geometry('720x650')
    window.minsize(620, 480)
    window.transient(root)
    window.protocol('WM_DELETE_WINDOW', _close_time_settings_window)

    ttk.Label(
        window,
        text=(
            '所有数值单位均为秒。保存后从下一次运行开始生效；'
            '任务运行期间设置会被锁定。'
        ),
        foreground='#555555',
        wraplength=660,
        justify='left',
    ).pack(fill='x', padx=14, pady=(12, 7))

    content_frame = ttk.Frame(window)
    content_frame.pack(fill='both', expand=True, padx=12)
    settings_canvas = tk.Canvas(
        content_frame, highlightthickness=0, borderwidth=0
    )
    settings_scrollbar = ttk.Scrollbar(
        content_frame, orient='vertical', command=settings_canvas.yview
    )
    settings_canvas.configure(yscrollcommand=settings_scrollbar.set)
    settings_canvas.pack(side='left', fill='both', expand=True)
    settings_scrollbar.pack(side='right', fill='y')

    settings_body = ttk.Frame(settings_canvas)
    body_window = settings_canvas.create_window(
        (0, 0), window=settings_body, anchor='nw'
    )

    def refresh_scroll_region(_event=None):
        settings_canvas.configure(scrollregion=settings_canvas.bbox('all'))

    def fit_body_width(event):
        settings_canvas.itemconfigure(body_window, width=event.width)

    settings_body.bind('<Configure>', refresh_scroll_region)
    settings_canvas.bind('<Configure>', fit_body_width)

    def scroll_settings(event):
        delta = getattr(event, 'delta', 0)
        if delta:
            settings_canvas.yview_scroll(-1 if delta > 0 else 1, 'units')
        elif getattr(event, 'num', None) == 4:
            settings_canvas.yview_scroll(-1, 'units')
        elif getattr(event, 'num', None) == 5:
            settings_canvas.yview_scroll(1, 'units')
        return 'break'

    # 绑定在 Toplevel 上，鼠标位于输入框或分组标题上时也可滚动。
    window.bind('<MouseWheel>', scroll_settings)
    window.bind('<Button-4>', scroll_settings)
    window.bind('<Button-5>', scroll_settings)

    setting_vars = {}
    current_values = get_time_settings_snapshot()
    groups = []
    for _key, _label, group_name, _default, _minimum, _maximum in TIME_SETTING_SPECS:
        if group_name not in groups:
            groups.append(group_name)

    for group_name in groups:
        group_frame = ttk.LabelFrame(
            settings_body, text=group_name, padding=(10, 7)
        )
        group_frame.pack(fill='x', padx=2, pady=(2, 8))
        group_frame.columnconfigure(1, weight=1)
        row = 0
        for key, label_text, spec_group, default, minimum, maximum in TIME_SETTING_SPECS:
            if spec_group != group_name:
                continue
            ttk.Label(group_frame, text=label_text).grid(
                row=row, column=0, sticky='w', padx=(0, 10), pady=3
            )
            value_var = tk.StringVar(
                value=f'{current_values.get(key, default):g}'
            )
            setting_vars[key] = value_var
            ttk.Entry(
                group_frame, textvariable=value_var, width=13,
                justify='right'
            ).grid(row=row, column=1, sticky='e', pady=3)
            ttk.Label(
                group_frame,
                text=f'秒（{minimum:g}～{maximum:g}）',
                foreground='#777777',
            ).grid(row=row, column=2, sticky='w', padx=(7, 0), pady=3)
            row += 1

    button_frame = ttk.Frame(window, padding=(12, 9, 12, 11))
    button_frame.pack(fill='x')

    def restore_defaults():
        defaults = default_time_settings()
        for key, value_var in setting_vars.items():
            value_var.set(f'{defaults[key]:g}')

    def save_settings_from_dialog():
        global current_time_settings
        candidate = {
            key: value_var.get() for key, value_var in setting_vars.items()
        }
        try:
            saved_values = save_time_settings(candidate)
        except Exception as error:
            messagebox.showerror(
                '时间设置保存失败', str(error), parent=window
            )
            return
        current_time_settings = dict(saved_values)
        messagebox.showinfo(
            '时间设置',
            f'已保存到：\n{timing_config_path()}\n\n下一次运行将使用新设置。',
            parent=window,
        )
        _close_time_settings_window()

    ttk.Button(
        button_frame, text='恢复默认值', command=restore_defaults
    ).pack(side='left')
    ttk.Button(
        button_frame, text='取消', command=_close_time_settings_window,
        width=10
    ).pack(side='right')
    ttk.Button(
        button_frame, text='保存设置', command=save_settings_from_dialog,
        width=12
    ).pack(side='right', padx=(0, 8))

    window.grab_set()
    window.after(0, lambda: settings_canvas.yview_moveto(0))
    window.focus_force()


time_settings_frame = ttk.Frame(frame1)
if ENABLE_TIME_SETTINGS:
    time_settings_frame.pack(fill='x', pady=(4, 1))
ttk.Label(
    time_settings_frame, text='填报等待时间（秒）', font=('微软雅黑', 9)
).pack(side='left')
time_settings_button = ttk.Button(
    time_settings_frame,
    text='时间设置...',
    command=open_time_settings,
    width=14,
)
time_settings_button.pack(side='right')

# ---------- 自有底稿数值转换 ----------
selected_file_path = tk.StringVar(value='')


def open_value_conversion_settings():
    """按对照表方案转换成当前收款结算单标准底稿。"""
    if current_thread is not None and current_thread.is_alive():
        messagebox.showwarning('任务运行中', '请先完成当前填报任务，再转换底稿。')
        return
    template_path = resource_path('收款结算单底稿导入模板.xlsm')
    if not os.path.isfile(template_path):
        messagebox.showerror(
            '缺少正式模板',
            '未找到“收款结算单底稿导入模板.xlsm”。\n'
            '请确认该模板与程序位于同一目录。'
        )
        return
    start_conversion_wizard(
        root,
        selected_file_path,
        template_path,
        target_sheet_name=EXCEL_CONFIG['sheet_name'],
        target_header_row=EXCEL_CONFIG['header_row'],
    )


conversion_frame = ttk.Frame(frame1)
conversion_frame.pack(fill='x', pady=(4, 1))
ttk.Label(
    conversion_frame, text='不想填写底稿？', font=('微软雅黑', 9)
).pack(side='left')
ttk.Button(
    conversion_frame,
    text='设置数值转换',
    command=open_value_conversion_settings,
    width=14,
).pack(side='left', padx=(8, 0))

# ---------- 文件选择区 ----------
file_frame = ttk.LabelFrame(root, text='文件选择', padding=(10, 5))
file_frame.pack(fill='x', padx=12, pady=6)

path_entry = ttk.Entry(file_frame, textvariable=selected_file_path, state='readonly', font=('微软雅黑', 9))
path_entry.grid(row=0, column=0, sticky='ew', padx=(0, 5))


def choose_file():
    file_path = filedialog.askopenfilename(
        title='请选择要处理的文件',
        filetypes=[('宏版收款底稿', '*.xlsm'), ('旧版底稿（无粘贴保护）', '*.xlsx')])
    if file_path:
        selected_file_path.set(file_path)


choose_btn = ttk.Button(file_frame, text='浏览...', command=choose_file, width=8)
choose_btn.grid(row=0, column=1)


def clear_path():
    selected_file_path.set('')


clear_btn = ttk.Button(file_frame, text='清除', command=clear_path, width=6)
clear_btn.grid(row=0, column=2, padx=(5, 0))

file_frame.columnconfigure(0, weight=1)


# ---------- 运行控制 ----------
def _task_entry(
    file_path, save_mode, merge_basis, single_company_only,
    timing_settings_snapshot, fill_order_settings_snapshot=None
):
    """工作线程总边界：把启动、登录等漏网异常安全传回 GUI。"""
    global active_browser, active_context
    if fill_order_settings_snapshot is None:
        fill_order_settings_snapshot = get_fill_order_settings_snapshot()
    run_context = {
        'started_at': time.monotonic(),
        'row_outcomes': {},
        'source_validated': False,
        # 本批次只读快照：GUI/配置文件后续变化不会改变正在运行的节奏与填写顺序。
        'timing_settings': dict(timing_settings_snapshot),
        'fill_order_settings': {
            'header_order': list(fill_order_settings_snapshot['header_order']),
            'detail_order': list(fill_order_settings_snapshot['detail_order']),
            'disabled': list(fill_order_settings_snapshot['disabled']),
        },
    }
    run_context_local.value = run_context
    try:
        running_task(
            file_path, save_mode, merge_basis,
            run_context=run_context,
            single_company_only=single_company_only
        )
    except TaskCancelledError:
        print('任务已按用户要求停止。')
    except Exception as error:
        traceback.print_exc()
        ui_message_queue.put(('error', f'任务失败：{error}'))
    finally:
        # 正常批次已在浏览器保留前写入；异常/停止/提前返回由这里兜底。
        _finalize_runtime_log(file_path, run_context)
        remaining_workbook = run_context.get('workbook')
        if remaining_workbook is not None:
            try:
                remaining_workbook.close()
            except Exception:
                pass
            run_context['workbook'] = None
        run_context_local.value = None
        # 只能在创建 Playwright 的工作线程内关闭，避免跨线程调用同步 API。
        try:
            if active_context is not None:
                active_context.close()
        except Exception:
            pass
        try:
            if active_browser is not None and active_browser.is_connected():
                active_browser.close()
        except Exception:
            pass
        active_context = None
        active_browser = None
        if stop_event.is_set():
            ui_message_queue.put(('info', '任务已停止，程序打开的浏览器已关闭。'))
        ui_message_queue.put(('run_state', 'idle'))


def start():
    global current_thread
    if getattr(root, '_receipt_conversion_busy', False):
        messagebox.showwarning('正在转换', '请等待底稿转换完成后再开始填报。')
        return
    if current_thread is not None and current_thread.is_alive():
        print('当前任务仍在运行或浏览器仍保持打开，请先完成当前任务并关闭浏览器。')
        return
    file_path = selected_file_path.get()
    if not file_path:
        print('请先选择文件！')
        return
    save_mode = 'submit' if save_mode_var.get() == 1 else 'save'
    merge_basis = merge_basis_var.get()
    single_company_only = bool(single_company_only_var.get())
    timing_settings_snapshot = get_time_settings_snapshot()
    fill_order_settings_snapshot = get_fill_order_settings_snapshot()
    disabled_snapshot = set(fill_order_settings_snapshot['disabled'])
    print(
        '本次填写完成后的操作：'
        f'{"保存提交" if save_mode == "submit" else "保存"}'
    )
    print(f'本次并表依据选择：{merge_basis_label(merge_basis)}')
    print(
        '是否仅有单家公司权限：'
        f'{"是（跳过结算财务组织）" if single_company_only else "否"}'
    )
    print(
        '本次时间设置已锁定：'
        f'登录最长{timing_settings_snapshot["login_wait"]:g}秒，'
        f'普通下拉最长{timing_settings_snapshot["reference_wait"]:g}秒，'
        f'辅助核算弹窗最长{timing_settings_snapshot["auxiliary_modal_wait"]:g}秒'
    )
    print(
        '本次表头填写顺序：'
        + '、'.join(
            HEADER_FILL_STEP_LABELS[key]
            for key in fill_order_settings_snapshot['header_order']
            if key not in disabled_snapshot
        )
    )
    print(
        '本次明细填写顺序：'
        + '、'.join(
            DETAIL_FILL_STEP_LABELS[key]
            for key in fill_order_settings_snapshot['detail_order']
            if key not in disabled_snapshot
        )
    )
    print(
        '保存前付款银行账号：'
        + (
            '已停用（跳过）'
            if 'payer_bank_account' in disabled_snapshot
            else '全部明细填妥后、点击保存前最后填入'
        )
    )
    if disabled_snapshot:
        print(
            '本次已停用字段：'
            + '、'.join(
                ALL_FILL_STEP_LABELS[key]
                for key in fill_order_settings_snapshot['disabled']
            )
        )
    reset_pause_clock()
    set_time_settings_running(True)
    set_fill_order_settings_running(True)
    current_thread = threading.Thread(
        target=_task_entry,
        args=(
            file_path, save_mode, merge_basis, single_company_only,
            timing_settings_snapshot, fill_order_settings_snapshot
        ),
        daemon=False
    )
    try:
        current_thread.start()
    except Exception:
        set_time_settings_running(False)
        set_fill_order_settings_running(False)
        raise


def pause_program():
    set_task_paused(True)
    print('任务暂停')


def start_program():
    set_task_paused(False)
    print('任务恢复')


def on_start_button_click():
    start_program()


def on_pause_button_click():
    pause_program()


def stop_program():
    if current_thread is None or not current_thread.is_alive():
        print('当前没有正在运行的任务。')
        return
    if not messagebox.askyesno(
        '停止任务',
        '确定停止当前任务吗？\n'
        '若NCC保存/提交按钮已经点击，程序会把该单据标记为“异常”或“填报成功”，'
        '请先在NCC人工确认后再重跑。'
    ):
        return
    stop_event.set()
    set_task_paused(False)
    print('已请求停止任务，程序会在当前浏览器操作返回后安全退出......')


control_frame = ttk.Frame(root, padding=(12, 2, 12, 8))
control_frame.pack(fill='x')
for column in range(4):
    control_frame.columnconfigure(column, weight=1)

run_button = ttk.Button(
    control_frame, text='开始运行', command=start, width=8
)
run_button.grid(row=0, column=0, padx=(0, 5), sticky='ew')

resume_button = ttk.Button(
    control_frame, text='恢复', command=on_start_button_click, width=8
)
resume_button.grid(row=0, column=1, padx=5, sticky='ew')

pause_button = ttk.Button(
    control_frame, text='暂停', command=on_pause_button_click, width=8
)
pause_button.grid(row=0, column=2, padx=5, sticky='ew')

stop_button = ttk.Button(
    control_frame, text='停止任务', command=stop_program, width=8
)
stop_button.grid(row=0, column=3, padx=(5, 0), sticky='ew')


def poll_ui_messages():
    try:
        while True:
            message_type, message = ui_message_queue.get_nowait()
            if message_type == 'error':
                messagebox.showerror('收款结算单工具', message)
            elif message_type == 'info':
                messagebox.showinfo('收款结算单工具', message)
            elif message_type == 'run_state' and message == 'idle':
                set_time_settings_running(False)
                set_fill_order_settings_running(False)
    except queue.Empty:
        pass
    root.after(timing_ms('gui_poll_interval'), poll_ui_messages)


def _close_when_worker_stops():
    if current_thread is not None and current_thread.is_alive():
        root.after(timing_ms('gui_poll_interval'), _close_when_worker_stops)
        return
    root.destroy()


def on_close():
    global closing_requested
    if current_thread is not None and current_thread.is_alive():
        if closing_requested:
            return
        if not messagebox.askyesno(
            '停止并退出',
            '任务仍在运行。确定停止任务并退出吗？\n'
            '程序会等待当前浏览器操作返回、关闭浏览器后再退出；'
            '若已点击保存/提交，请随后在NCC人工核对结果。'
        ):
            return
        closing_requested = True
        stop_event.set()
        set_task_paused(False)
        print('已请求停止并退出，正在等待工作线程安全结束......')
        _close_when_worker_stops()
        return
    root.destroy()


root.protocol('WM_DELETE_WINDOW', on_close)
root.after(timing_ms('gui_poll_interval'), poll_ui_messages)
# 宽度保持红框尺寸，高度严格按当前控件所需空间计算，避免底部空白。
root.update_idletasks()
root.geometry(f'430x{root.winfo_reqheight()}')
root.mainloop()
