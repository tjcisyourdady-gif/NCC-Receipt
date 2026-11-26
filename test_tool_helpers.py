import ast
import datetime
import json
import math
import os
import pathlib
import re
import sys
import tempfile
import unittest

from openpyxl import load_workbook
from receipt_data_rules import normalize_receipt_value, receipt_sheet_layout
from openpyxl.utils import get_column_letter


SOURCE_PATH = pathlib.Path(__file__).with_name("收款结算单工具.py")
TEMPLATE_PATH = SOURCE_PATH.with_name("收款结算单底稿导入模板.xlsm")
SOURCE = SOURCE_PATH.read_text(encoding="utf-8")
TREE = ast.parse(SOURCE)


def load_config():
    for node in TREE.body:
        if isinstance(node, ast.Assign):
            if any(
                isinstance(target, ast.Name) and target.id == "EXCEL_CONFIG"
                for target in node.targets
            ):
                return ast.literal_eval(node.value)
    raise AssertionError("EXCEL_CONFIG not found")


def load_literal_assignment(name):
    for node in TREE.body:
        if isinstance(node, ast.Assign) and any(
            isinstance(target, ast.Name) and target.id == name
            for target in node.targets
        ):
            return ast.literal_eval(node.value)
    raise AssertionError(f"{name} not found")


def load_helpers():
    timing_specs = load_literal_assignment("TIME_SETTING_SPECS")
    timing_defaults = {
        key: float(default)
        for key, _label, _group, default, _minimum, _maximum in timing_specs
    }
    wanted = {
        "FieldFillError",
        "ReferenceMatchAmbiguousError",
        "SettleOrgInputUnavailableError",
        "BillDateInputUnavailableError",
        "TaskCancelledError",
        "_clean_input_text",
        "excel_text",
        "_format_calendar_date",
        "format_bill_date",
        "format_settlement_date",
        "format_expense_period",
        "format_tax_rate",
        "accounting_subject_code",
        "auxiliary_accounting_code",
        "settlement_org_query_value",
        "project_archive_query_value",
        "fill_settle_org_with_card_recovery",
        "fill_bill_date_with_card_recovery",
        "_normalize_linked_yes_no",
        "validate_community_business_category",
        "transaction_object_of",
        "customer_reference_value",
        "_normalize_reference_text",
        "_customer_name_from_reference_option",
        "_customer_name_option_matches",
        "customer_grouping_key",
        "merge_basis_label",
        "build_bill_groups",
        "fill_auxiliary_accounting",
        "_parse_bill_number",
        "_find_bill_number_once",
        "_detail_row_added",
        "_record_runtime_outcomes",
        "write_group_result",
    }
    nodes = [
        node
        for node in TREE.body
        if isinstance(node, (ast.ClassDef, ast.FunctionDef))
        and node.name in wanted
    ]
    namespace = {
        "receipt_sheet_layout": receipt_sheet_layout,
        "get_column_letter": get_column_letter,
        "re": re,
        "normalize_receipt_value": normalize_receipt_value,
        "EXCEL_CONFIG": load_config(),
        "SETTLE_ORG_CARD_RELOAD_MAX": load_literal_assignment(
            "SETTLE_ORG_CARD_RELOAD_MAX"
        ),
        "BILL_DATE_CARD_RELOAD_MAX": load_literal_assignment(
            "BILL_DATE_CARD_RELOAD_MAX"
        ),
        "timing_seconds": lambda key: timing_defaults[key],
        "timing_ms": lambda key: max(
            1, int(round(timing_defaults[key] * 1000))
        ),
        "format_timing_seconds": lambda key: f"{timing_defaults[key]:g}",
    }
    exec(compile(ast.Module(nodes, type_ignores=[]), str(SOURCE_PATH), "exec"), namespace)
    return namespace


def load_time_helpers():
    """只执行纯时间配置代码，避免导入主程序时启动 Tk 主循环。"""
    assignment_names = {
        "TIME_SETTING_SPECS",
        "TIME_SETTINGS_VERSION",
        "TIME_SETTINGS_FILENAME",
    }
    function_names = {
        "default_time_settings",
        "validate_time_settings",
        "load_time_settings",
        "save_time_settings",
    }
    nodes = []
    for node in TREE.body:
        if isinstance(node, ast.Assign) and any(
            isinstance(target, ast.Name) and target.id in assignment_names
            for target in node.targets
        ):
            nodes.append(node)
        elif isinstance(node, ast.FunctionDef) and node.name in function_names:
            nodes.append(node)
    namespace = {
        "json": json,
        "math": math,
        "os": os,
        "sys": sys,
        "tempfile": tempfile,
    }
    exec(
        compile(ast.Module(nodes, type_ignores=[]), str(SOURCE_PATH), "exec"),
        namespace,
    )
    return namespace


class ToolHelperTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.config = load_config()
        cls.helpers = load_helpers()
        cls.time_helpers = load_time_helpers()

    def test_latest_column_mapping(self):
        fields = self.config["fields"]
        expected = {
            "company_code": 1,
            "expense_period": 6,
            "bill_number": 8,
            "accounting_subject": 11,
            "auxiliary_accounting": 12,
            "represented_party": 13,
            "represented_project": 14,
            "business_category": 15,
            "internal_collection_flag": 16,
            "internal_collection_type": 17,
            "receipt_method": 18,
            "settlement_method": 19,
            "actual_charge_type": 20,
            "new_income_expense_item": 21,
            "project_archive_name": 22,
            "receipt_original_amount": 23,
            "tax_method": 24,
            "tax_rate": 25,
            "split_income_type": 26,
            "merge_flag": 27,
        }
        for field, column in expected.items():
            self.assertEqual(fields[field], column)
        self.assertEqual(self.config["status_col"], 28)
        self.assertEqual(self.config["reason_col"], 29)

    def test_timeouts_are_explicit_and_pause_aware(self):
        specs = load_literal_assignment("TIME_SETTING_SPECS")
        defaults = {key: default for key, _label, _group, default, _min, _max in specs}
        self.assertEqual(defaults["login_wait"], 120.0)
        self.assertEqual(defaults["reference_wait"], 30.0)
        self.assertEqual(defaults["settle_org_option_wait"], 60.0)
        self.assertEqual(defaults["settle_org_input_wait"], 15.0)
        self.assertIn("SETTLE_ORG_CARD_RELOAD_MAX = 3", SOURCE)
        self.assertEqual(defaults["bill_date_input_wait"], 15.0)
        self.assertIn("BILL_DATE_CARD_RELOAD_MAX = 3", SOURCE)
        self.assertEqual(defaults["action_result_wait"], 60.0)
        self.assertEqual(defaults["tax_rate_bind_wait"], 2.5)
        self.assertIn(
            "deadline = task_now() + timing_seconds('login_wait')", SOURCE
        )
        self.assertNotIn("ACTION_WAIT_MS", SOURCE)
        self.assertNotIn("ACTION_WAIT_SECONDS", SOURCE)
        self.assertIn("context.set_default_timeout(timing_ms('long_element_timeout'))", SOURCE)
        self.assertIn(
            "context.set_default_navigation_timeout(timing_ms('page_load_timeout'))",
            SOURCE,
        )
        self.assertIn(
            "save_choose.click(timeout=timing_ms('long_element_timeout'))",
            SOURCE,
        )
        self.assertIn(
            "deadline = task_now() + timing_seconds('cancel_confirm_wait')",
            SOURCE,
        )
        self.assertIn("div.u-modal tr.u-table-row", SOURCE)
        self.assertIn("div.u-dropdown [role='option']", SOURCE)

    def test_time_settings_validate_persist_and_recover(self):
        helpers = self.time_helpers
        defaults = helpers["default_time_settings"]()
        self.assertEqual(len(defaults), len(load_literal_assignment("TIME_SETTING_SPECS")))
        self.assertEqual(defaults["auxiliary_modal_open_wait"], 3.0)

        changed = dict(defaults)
        changed["reference_wait"] = "42.5"
        validated = helpers["validate_time_settings"](changed)
        self.assertEqual(validated["reference_wait"], 42.5)
        with self.assertRaises(ValueError):
            helpers["validate_time_settings"]({"reference_wait": -1})
        with self.assertRaises(ValueError):
            helpers["validate_time_settings"]({"reference_wait": "nan"})

        with tempfile.TemporaryDirectory() as directory:
            config_path = pathlib.Path(directory) / "time-settings.json"
            helpers["save_time_settings"](changed, str(config_path))
            payload = json.loads(config_path.read_text(encoding="utf-8"))
            self.assertEqual(payload["version"], 1)
            self.assertEqual(payload["values"]["reference_wait"], 42.5)
            loaded = helpers["load_time_settings"](str(config_path))
            self.assertEqual(loaded, validated)

            config_path.write_text("{broken", encoding="utf-8")
            recovered = helpers["load_time_settings"](str(config_path))
            self.assertEqual(recovered, defaults)

    def test_time_settings_gui_is_locked_and_snapshotted_per_run(self):
        functions = {
            node.name: ast.get_source_segment(SOURCE, node)
            for node in TREE.body
            if isinstance(node, ast.FunctionDef)
        }
        self.assertIn("text='时间设置...'", SOURCE)
        self.assertIn("window.geometry('720x650')", functions["open_time_settings"])
        self.assertIn("TIME_SETTING_SPECS", functions["open_time_settings"])
        self.assertIn("save_time_settings(candidate)", functions["open_time_settings"])
        self.assertIn(
            "current_thread is not None and current_thread.is_alive()",
            functions["open_time_settings"],
        )
        self.assertIn("set_time_settings_running(True)", functions["start"])
        self.assertIn("get_time_settings_snapshot()", functions["start"])
        self.assertIn(
            "'timing_settings': dict(timing_settings_snapshot)",
            functions["_task_entry"],
        )
        self.assertIn(
            "ui_message_queue.put(('run_state', 'idle'))",
            functions["_task_entry"],
        )
        self.assertIn(
            "set_time_settings_running(False)", functions["poll_ui_messages"]
        )

    def test_input_text_cleaning_removes_edge_hidden_whitespace_only(self):
        clean = self.helpers["_clean_input_text"]
        excel_text = self.helpers["excel_text"]
        dirty = '\u3000\u00a0\u200b  客商  名称 \u2060\ufeff\t'
        self.assertEqual(clean(dirty), '客商  名称')
        self.assertEqual(excel_text(dirty), '客商  名称')
        self.assertEqual(clean('项目 A 01'), '项目 A 01')
        self.assertEqual(excel_text(2203.0), '2203')

    def test_virtual_detail_table_detects_row_15_without_button_count_growth(self):
        row_added = self.helpers["_detail_row_added"]
        before = {
            "button_total": 10,
            "button_visible": 10,
            "max_sequence": 14,
            "signature": "rows-5-to-14",
        }
        after = {
            "button_total": 10,
            "button_visible": 10,
            "max_sequence": 15,
            "signature": "rows-6-to-15",
        }
        self.assertTrue(row_added(before, after))
        self.assertFalse(row_added(before, dict(before)))
        self.assertIn('虚拟表格未暴露新增序号', SOURCE)
        self.assertIn('继续填写最新明细行', SOURCE)

    def test_updated_template_headers_match_mapping(self):
        workbook = load_workbook(TEMPLATE_PATH, read_only=True, data_only=True)
        sheet = workbook["收款结算明细"]
        layout = receipt_sheet_layout(sheet, self.config)
        headers = [cell.value for cell in sheet[2][:layout["reason_col"]]]
        workbook.close()
        expected_headers = load_literal_assignment("EXPECTED_FIELD_HEADERS")
        for field, column in layout["fields"].items():
            if field not in expected_headers:
                continue
            actual = re.sub(r'[\s*＊]+$', '', str(headers[column - 1] or '')).strip()
            allowed = expected_headers[field]
            if not isinstance(allowed, (tuple, list)):
                allowed = (allowed,)
            normalized_allowed = {
                re.sub(r'[\s*＊]+$', '', item).strip() for item in allowed
            }
            self.assertIn(actual, normalized_allowed, field)
        self.assertEqual(headers[layout["fields"]["accounting_subject"] - 1], "会计科目*")
        self.assertEqual(headers[layout["fields"]["auxiliary_accounting"] - 1], "辅助核算")
        self.assertEqual(headers[layout["fields"]["represented_project"] - 1], "被代收方（被收回方）项目档案*")
        self.assertEqual(headers[layout["fields"]["internal_collection_type"] - 1], "代收代付类型")
        self.assertNotIn("对账识别码*", headers)
        self.assertEqual(headers[layout["fields"]["tax_method"] - 1], "计税方式*")
        self.assertEqual(headers[layout["fields"]["tax_rate"] - 1], "税率（非增值税收益等）*")
        self.assertEqual(
            headers[layout["status_col"] - 1:layout["reason_col"]],
            ["填写状态", "异常原因"],
        )

    def test_expense_period_is_month_only(self):
        fn = self.helpers["format_expense_period"]
        self.assertEqual(fn(datetime.date(2026, 7, 4)), "2026-07")
        self.assertEqual(fn("2026-07"), "2026-07")
        self.assertEqual(fn("2026/07"), "2026-07")
        self.assertEqual(fn("2026-7-04"), "2026-07")
        self.assertEqual(fn("20260704"), "2026-07")
        self.assertEqual(fn(46204), "2026-07")

    def test_bill_and_settlement_date_accept_dash_or_slash(self):
        bill = self.helpers["format_bill_date"]
        settlement = self.helpers["format_settlement_date"]
        self.assertEqual(bill("2026-07-01"), "2026-07-01")
        self.assertEqual(bill("2026/07/01"), "2026-07-01")
        self.assertEqual(settlement("2026-07-01"), "2026-07-01")
        self.assertEqual(settlement("2026/07/01"), "2026-07-01")

    def test_tax_rate_and_accounting_code(self):
        self.assertEqual(self.helpers["format_tax_rate"](0.06), "6%")
        self.assertEqual(self.helpers["format_tax_rate"]("13%"), "13%")
        self.assertEqual(
            self.helpers["accounting_subject_code"]("2241002\\AAAAA\\BBBB"),
            "2241002",
        )
        self.assertEqual(
            self.helpers["auxiliary_accounting_code"]("04/场地收入"),
            "04",
        )

    def test_settlement_org_uses_code_from_code_name_format(self):
        lookup = self.helpers["settlement_org_query_value"]
        self.assertEqual(lookup('H1978/示例商业管理有限公司'), 'H1978')
        self.assertEqual(lookup('H1978／示例商业管理有限公司'), 'H1978')
        self.assertEqual(lookup('示例商业管理有限公司'), '示例商业管理有限公司')
        with self.assertRaises(self.helpers["FieldFillError"]):
            lookup('/示例商业管理有限公司')

    def test_settlement_org_card_recovery_state_machine(self):
        recovery = self.helpers["fill_settle_org_with_card_recovery"]
        missing = self.helpers["SettleOrgInputUnavailableError"]
        field_error = self.helpers["FieldFillError"]

        calls = {'fill': 0, 'add': 0, 'cancel': 0}

        def add(_frame):
            calls['add'] += 1

        def cancel(_frame):
            calls['cancel'] += 1
            return True

        def fill_after_one_reload(_frame, _value):
            calls['fill'] += 1
            if calls['fill'] == 1:
                raise missing('结算财务组织', '15秒内未出现')
            return True

        self.helpers['click_add_bill'] = add
        self.helpers['cancel_card'] = cancel
        self.helpers['fill_settle_org'] = fill_after_one_reload
        self.assertTrue(recovery(object(), 'H1978'))
        self.assertEqual(calls, {'fill': 2, 'add': 2, 'cancel': 1})

        calls.update(fill=0, add=0, cancel=0)

        def always_missing(_frame, _value):
            calls['fill'] += 1
            raise missing('结算财务组织', '15秒内未出现')

        self.helpers['fill_settle_org'] = always_missing
        with self.assertRaises(field_error):
            recovery(object(), 'H1978')
        self.assertEqual(calls, {'fill': 4, 'add': 4, 'cancel': 3})

        calls.update(fill=0, add=0, cancel=0)

        def dropdown_failure(_frame, _value):
            calls['fill'] += 1
            raise field_error('结算财务组织', '下拉无数据')

        self.helpers['fill_settle_org'] = dropdown_failure
        with self.assertRaises(field_error):
            recovery(object(), 'H1978')
        self.assertEqual(calls, {'fill': 1, 'add': 1, 'cancel': 0})

        calls.update(fill=0, add=0, cancel=0)
        self.helpers['fill_settle_org'] = always_missing

        def cancel_failure(_frame):
            calls['cancel'] += 1
            return False

        self.helpers['cancel_card'] = cancel_failure
        with self.assertRaises(field_error):
            recovery(object(), 'H1978')
        self.assertEqual(calls, {'fill': 1, 'add': 1, 'cancel': 1})

    def test_single_company_bill_date_card_recovery_state_machine(self):
        recovery = self.helpers["fill_bill_date_with_card_recovery"]
        missing = self.helpers["BillDateInputUnavailableError"]
        field_error = self.helpers["FieldFillError"]
        calls = {'fill': 0, 'add': 0, 'cancel': 0}

        def add(_frame):
            calls['add'] += 1

        def cancel(_frame):
            calls['cancel'] += 1
            return True

        def fill_after_one_reload(_frame, _value):
            calls['fill'] += 1
            if calls['fill'] == 1:
                raise missing('单据日期', '15秒内未出现')
            return True

        self.helpers['click_add_bill'] = add
        self.helpers['cancel_card'] = cancel
        self.helpers['fill_bill_date'] = fill_after_one_reload
        self.assertTrue(recovery(object(), '2026-08-19'))
        self.assertEqual(calls, {'fill': 2, 'add': 2, 'cancel': 1})

        calls.update(fill=0, add=0, cancel=0)

        def always_missing(_frame, _value):
            calls['fill'] += 1
            raise missing('单据日期', '15秒内未出现')

        self.helpers['fill_bill_date'] = always_missing
        with self.assertRaises(field_error):
            recovery(object(), '2026-08-19')
        self.assertEqual(calls, {'fill': 4, 'add': 4, 'cancel': 3})

        calls.update(fill=0, add=0, cancel=0)

        def real_date_error(_frame, _value):
            calls['fill'] += 1
            raise field_error('单据日期', '日期面板中没有目标日期')

        self.helpers['fill_bill_date'] = real_date_error
        with self.assertRaises(field_error):
            recovery(object(), '2026-08-19')
        self.assertEqual(calls, {'fill': 1, 'add': 1, 'cancel': 0})

    def test_customer_lookup_prefers_code_and_allows_name_only(self):
        lookup = self.helpers["customer_reference_value"]
        object_type = self.helpers["transaction_object_of"]
        self.assertEqual(
            lookup({"customer_code": "C001", "customer_name": "甲公司"}),
            ("C001", '“客商编码”', True),
        )
        self.assertEqual(
            lookup({"customer_code": "", "customer_name": "甲公司"}),
            ("甲公司", '“客商名称”', False),
        )
        self.assertEqual(
            object_type({"customer_code": "C001", "customer_name": ""}),
            "客户",
        )
        self.assertEqual(
            object_type({"customer_code": "", "customer_name": "甲公司"}),
            "客户",
        )
        self.assertEqual(
            object_type({"customer_code": "", "customer_name": ""}),
            "人员",
        )

        option_name = self.helpers["_customer_name_from_reference_option"]
        name_matches = self.helpers["_customer_name_option_matches"]
        self.assertEqual(option_name("C001/甲公司"), "甲公司")
        self.assertEqual(option_name("C002／甲公司"), "甲公司")
        self.assertTrue(name_matches("甲公司", "C001/甲公司"))
        self.assertTrue(name_matches("甲公司", "C002/甲公司"))
        self.assertFalse(name_matches("甲公司", "C003/乙公司"))
        self.assertFalse(name_matches("C001", "C001/甲公司"))

    def test_merge_basis_can_optionally_include_customer(self):
        build_groups = self.helpers["build_bill_groups"]
        common = {
            "merge_flag": "是",
            "company_name": "甲公司",
            "bank_account": "6222",
            "merge_key_i": "2026-08-10",
            "customer_name": "客商甲",
        }
        rows = [
            dict(common, _row=3, customer_code="C001"),
            dict(common, _row=4, customer_code="C002", customer_name="客商乙"),
        ]
        basic = build_groups(rows, "company_account_date")
        by_customer = build_groups(
            rows, "company_account_date_customer"
        )
        self.assertEqual([len(group) for group in basic], [2])
        self.assertEqual([len(group) for group in by_customer], [1, 1])

        same_customer = [
            dict(common, _row=3, customer_code="C001"),
            dict(common, _row=4, customer_code="C001", customer_name="别名"),
        ]
        self.assertEqual(
            [len(group) for group in build_groups(
                same_customer, "company_account_date_customer"
            )],
            [2],
        )

        standalone = dict(rows[1], _row=5, merge_flag="否")
        self.assertEqual(
            [len(group) for group in build_groups(
                rows + [standalone], "company_account_date_customer"
            )],
            [1, 1, 1],
        )

    def test_merge_basis_gui_is_wired_to_worker(self):
        self.assertIn("text='并表依据选择'", SOURCE)
        self.assertIn("root.geometry(f'430x{root.winfo_reqheight()}')", SOURCE)
        self.assertNotIn("root.geometry('430x630')", SOURCE)
        self.assertIn("root.resizable(False, False)", SOURCE)
        self.assertIn(
            "merge_basis_var = tk.StringVar(value='company_account_date')",
            SOURCE,
        )
        self.assertIn("company_account_date_customer", SOURCE)
        functions = {
            node.name: ast.get_source_segment(SOURCE, node)
            for node in TREE.body
            if isinstance(node, ast.FunctionDef)
        }
        self.assertIn("merge_basis_var.get()", functions["start"])
        self.assertIn("text='是否仅有单家公司权限'", SOURCE)
        self.assertIn(
            "single_company_only_var = tk.BooleanVar(value=False)", SOURCE
        )
        self.assertIn("single_company_only_var.get()", functions["start"])
        self.assertIn(
            "single_company_only=single_company_only",
            functions["_task_entry"],
        )
        self.assertIn(
            "build_bill_groups(rows, merge_basis=merge_basis)",
            functions["running_task"],
        )

    def test_browser_uses_full_screen_real_viewport(self):
        running = next(
            node
            for node in TREE.body
            if isinstance(node, ast.FunctionDef)
            and node.name == "running_task"
        )
        text = ast.get_source_segment(SOURCE, running)
        self.assertIn("'--start-fullscreen'", text)
        self.assertIn("'--start-maximized'", text)
        self.assertIn("'--disable-backgrounding-occluded-windows'", text)
        self.assertIn("'--disable-renderer-backgrounding'", text)
        self.assertIn("'--disable-background-timer-throttling'", text)
        self.assertIn("IntensiveWakeUpThrottling", text)
        self.assertIn("no_viewport=True", text)
        self.assertIn("keep_page_active_when_minimized(context, new_page)", text)
        self.assertIn("raise_automation_process_priority()", text)
        self.assertIn("Emulation.setFocusEmulationEnabled", SOURCE)
        self.assertIn("Page.setWebLifecycleState", SOURCE)
        self.assertIn("_refresh_page_activity_when_minimized()", SOURCE)
        self.assertIn("SetPriorityClass", SOURCE)

    def test_gui_has_no_capture_button_or_red_warning(self):
        removed_tokens = (
            '最小化时辅助核算偶发填写失败，暂未解决，请避免最小化填报游览器',
            "fg='#d93025'",
            "text='抓取界面'",
            'capture_button',
            'request_auxiliary_structure_capture',
            '_capture_auxiliary_structure',
            '_auxiliary_structure_capture_enabled',
            '_manual_auxiliary_capture_request_path',
            '_capture_requested_auxiliary_structure',
            'NCC_CAPTURE_AUXILIARY',
            '辅助核算界面结构抓取.request',
            'minimize_warning_label',
        )
        for token in removed_tokens:
            self.assertNotIn(token, SOURCE)
        self.assertIn('for column in range(4):', SOURCE)

    def test_reference_dropdown_retries_full_fill_and_match(self):
        functions = {
            node.name: ast.get_source_segment(SOURCE, node)
            for node in TREE.body
            if isinstance(node, ast.FunctionDef)
        }
        generic = functions["_fill_reference_field"]
        retry_loop = generic.index("for attempt in range(1, RETRY_MAX + 1)")
        match_call = generic.index("_click_matching_reference_option(")
        self.assertLess(retry_loop, match_call)
        self.assertIn("_clear_reference_before_fill", generic)
        self.assertIn("except ReferenceMatchAmbiguousError", generic)
        self.assertIn("if not retry_ambiguous", generic)

        organization = functions["fill_settle_org"]
        self.assertIn("timing_ms('settle_org_option_wait')", organization)
        self.assertIn("timing_ms('settle_org_input_wait')", organization)
        self.assertIn("settlement_org_query_value(company_code)", organization)
        self.assertRegex(organization, r"org_input\.fill\(\s*org_query")
        self.assertIn("company_code", organization)
        self.assertNotIn("company_name", organization)
        self.assertIn("_find_settle_org_input", organization)

        add_bill = functions["click_add_bill"]
        self.assertNotIn(
            "frame.locator(SELECTORS[\"settle_org_input\"])", add_bill
        )
        self.assertIn("开始等待新增卡片字段加载", add_bill)
        self.assertNotIn("card_buttons", add_bill)

        organization_locator = functions["_find_settle_org_input"]
        self.assertIn(
            "_find_field_input(body, '结算财务组织'", organization_locator
        )
        self.assertIn("备用规则命中", organization_locator)

        organization_recovery = functions["fill_settle_org_with_card_recovery"]
        self.assertIn(
            "range(SETTLE_ORG_CARD_RELOAD_MAX + 1)", organization_recovery
        )
        self.assertIn("cancel_card(frame)", organization_recovery)
        self.assertIn("click_add_bill(frame)", organization_recovery)
        self.assertIn("if not cancel_card(frame)", organization_recovery)
        running_text = functions["running_task"]
        self.assertIn("fill_settle_org_with_card_recovery(", running_text)
        self.assertIn("if single_company_only:", running_text)
        self.assertIn("fill_bill_date_with_card_recovery(", running_text)
        self.assertIn("跳过结算财务组织，等待并填写单据日期", running_text)
        self.assertLess(
            running_text.index("if single_company_only:"),
            running_text.index("active_step = '单据日期'"),
        )

        detail = functions["fill_detail_row"]
        detail_step = functions["_fill_detail_step"]
        self.assertIn("enabled_detail_fill_order()", detail)
        self.assertIn("_fill_detail_step(", detail)
        self.assertIn("customer_reference_value(row)", detail_step)
        self.assertIn("retry_ambiguous=False", detail_step)
        self.assertIn("allow_multiple_matches=uses_customer_code", detail_step)
        self.assertIn("customer_name_only=not uses_customer_code", detail_step)

        matcher = functions["_click_matching_reference_option"]
        self.assertIn("if customer_name_only", matcher)
        self.assertIn("_customer_name_option_matches", matcher)
        self.assertIn("多个同名客商", matcher)

    def test_blank_auxiliary_accounting_is_optional(self):
        self.helpers["fill_auxiliary_accounting"](
            object(), object(), '', previous_value='04/场地收入'
        )

    def test_community_income_requires_business_category(self):
        normalize = self.helpers["_normalize_linked_yes_no"]
        self.assertEqual(
            normalize('是否社群收入类收支\n是', '是否社群收入类收支'),
            '是',
        )
        self.assertEqual(normalize('否'), '否')
        self.assertEqual(normalize('YES'), '是')
        self.assertEqual(normalize('on'), '')
        self.assertEqual(normalize('checked'), '')
        self.assertEqual(
            normalize('是否社群收入类收支（是/否）', '是否社群收入类收支'),
            '',
        )

        validate = self.helpers["validate_community_business_category"]
        self.helpers["_read_community_income_flag"] = lambda _root: '是'
        with self.assertRaises(self.helpers["FieldFillError"]) as raised:
            validate(object(), '')
        self.assertIn('底稿“经营类别名称”不能为空', str(raised.exception))

        self.helpers["_read_community_income_flag"] = lambda _root: '否'
        self.assertEqual(validate(object(), ''), '否')

        reader = next(
            node
            for node in TREE.body
            if isinstance(node, ast.FunctionDef)
            and node.name == "_read_community_income_flag"
        )
        reader_text = ast.get_source_segment(SOURCE, reader)
        self.assertIn("role === 'radio'", reader_text)
        self.assertNotIn(
            "if (checked === 'true') values.push('是')", reader_text
        )
        self.assertIn(
            "timing_seconds('reference_option_stable_wait')", reader_text
        )

    def test_ncc_checkbox_switch_uses_business_value_and_click(self):
        setter = next(
            node
            for node in TREE.body
            if isinstance(node, ast.FunctionDef)
            and node.name == "_set_checkbox_field"
        )
        setter_text = ast.get_source_segment(SOURCE, setter)
        self.assertIn("checkbox_switch", setter_text)
        self.assertIn("input.value", setter_text)
        self.assertRegex(setter_text, r"checkbox\.click\(\s*force=True")
        self.assertNotIn("checkbox.check(", setter_text)
        self.assertNotIn("checkbox.uncheck(", setter_text)

    def test_saved_bill_number_formats(self):
        fn = self.helpers["_parse_bill_number"]
        number = "F4202608042156390"
        self.assertEqual(fn(f"单据编号： {number}"), number)
        self.assertEqual(fn(f"收款结算 ： {number}"), number)
        self.assertEqual(fn(number), number)

        finder = next(
            node
            for node in TREE.body
            if isinstance(node, ast.FunctionDef)
            and node.name == "_find_bill_number_once"
        )
        finder_text = ast.get_source_segment(SOURCE, finder)
        self.assertIn("contains(normalize-space(.),'单据编号')", finder_text)
        self.assertIn("frame.locator('body')", finder_text)
        self.assertRegex(
            finder_text,
            r"body\.inner_text\(\s*timeout=timing_ms\('tiny_element_timeout'\)",
        )

        class EmptyLocator:
            def count(self):
                return 0

        class BodyLocator:
            def inner_text(self, timeout=None):
                return f"保存成功\n单据编号： {number}\n提交"

            def text_content(self, timeout=None):
                return self.inner_text(timeout)

        class FakeFrame:
            def locator(self, selector):
                return BodyLocator() if selector == 'body' else EmptyLocator()

        self.helpers["SELECTORS"] = {"bill_info": "span.old-bill-title"}
        self.assertEqual(
            self.helpers["_find_bill_number_once"](FakeFrame()), number
        )

    def test_project_archive_uses_code_when_code_and_name_are_supplied(self):
        query = self.helpers["project_archive_query_value"]
        self.assertEqual(query("P000039713/商贸城仓储物流园"), "P000039713")
        self.assertEqual(query(r"P000039713\商贸城仓储物流园"), "P000039713")
        self.assertEqual(query("商贸城仓储物流园"), "商贸城仓储物流园")
        with self.assertRaises(self.helpers["FieldFillError"]):
            query("/商贸城仓储物流园")
        with self.assertRaises(self.helpers["FieldFillError"]):
            query("P000039713/")

    def test_internal_collection_project_uses_existing_s_column_and_order(self):
        function = next(
            node
            for node in TREE.body
            if isinstance(node, ast.FunctionDef)
            and node.name == "fill_internal_collection_reference_fields_before_tax_rate"
        )
        text = ast.get_source_segment(SOURCE, function)
        type_call = text.index("'内部代收付（含清理）类型', internal_collection_type")
        party_call = text.index(
            "frame, drawer, '被代收方（被收回方）', row[\"represented_party\"]"
        )
        project_call = text.index(
            "frame, drawer, '被代收方（被收回方）项目档案'"
        )
        charge_call = text.index(
            "frame, drawer, '代收（收回）业务实际业务收费类型'"
        )
        self.assertLess(type_call, project_call)
        self.assertLess(type_call, party_call)
        self.assertLess(party_call, project_call)
        self.assertLess(project_call, charge_call)
        self.assertIn('row["internal_collection_type"]', text)
        self.assertIn('“代收代付类型”', text)
        self.assertIn('row["represented_project"]', text)
        self.assertIn('“被代收方（被收回方）项目档案”', text)

    def test_internal_collection_skips_ncc_project_archive(self):
        function = next(
            node
            for node in TREE.body
            if isinstance(node, ast.FunctionDef) and node.name == "_fill_detail_step"
        )
        text = ast.get_source_segment(SOURCE, function)
        self.assertIn('row["internal_collection_flag"]) == \'是\'', text)
        self.assertIn('已跳过NCC“项目档案”填写', text)
        self.assertIn('preserve_linked_readonly=True', text)

    def test_saved_bill_number_is_written_to_column_h(self):
        cells = {}

        class FakeSheet:
            def cell(self, row, column, value=None):
                cells[(row, column)] = value
                return type('Cell', (), {'value': value})()

        self.helpers["_save_result_workbook"] = lambda *args: None
        self.helpers["write_group_result"](
            FakeSheet(), object(), "unused.xlsx",
            [{"_row": 3}, {"_row": 4}], "填报成功",
            bill_number="F4202608042156390",
        )
        self.assertEqual(cells[(3, 8)], "F4202608042156390")
        self.assertEqual(cells[(4, 8)], "F4202608042156390")

    def test_auxiliary_accounting_and_tax_rate_detail_order(self):
        detail_order = ast.literal_eval(
            next(
                node.value
                for node in TREE.body
                if isinstance(node, ast.Assign)
                and any(
                    isinstance(target, ast.Name)
                    and target.id == "DETAIL_FILL_DEFAULT_ORDER"
                    for target in node.targets
                )
            )
        )
        order_positions = [
            detail_order.index("expense_period"),
            detail_order.index("tax_method"),
            detail_order.index("accounting_subject"),
            detail_order.index("receipt_original_amount"),
            detail_order.index("internal_collection_references"),
            detail_order.index("new_income_expense_item"),
            detail_order.index("department"),
            detail_order.index("split_income_type"),
            detail_order.index("business_category"),
            detail_order.index("auxiliary_accounting"),
            detail_order.index("tax_rate"),
        ]
        self.assertEqual(order_positions, sorted(order_positions))

        step_function = next(
            node
            for node in TREE.body
            if isinstance(node, ast.FunctionDef) and node.name == "_fill_detail_step"
        )
        row_function = next(
            node
            for node in TREE.body
            if isinstance(node, ast.FunctionDef) and node.name == "fill_detail_row"
        )
        text = ast.get_source_segment(SOURCE, step_function)
        row_text = ast.get_source_segment(SOURCE, row_function)
        self.assertIn("enabled_detail_fill_order()", row_text)
        self.assertIn("close_detail_drawer(frame, drawer)", row_text)
        self.assertLess(
            row_text.index("_fill_detail_step("),
            row_text.index("close_detail_drawer(frame, drawer)"),
        )
        self.assertLess(
            text.index("'税率'"),
            text.index("settle_tax_rate_before_close(drawer)"),
        )
        self.assertLess(
            text.index("validate_community_business_category("),
            text.index("'经营类别'"),
        )
        self.assertRegex(
            text,
            r"_fill_reference_field\(\s*frame, drawer, '费用所属期'",
        )
        self.assertNotRegex(
            text,
            r"_fill_plain_field\(\s*drawer, '费用所属期'",
        )
        self.assertIn("data-codex-aux-before", SOURCE)
        self.assertIn("辅助核算新出现输入框命中", SOURCE)
        self.assertIn("弹窗打开后未检测到新出现的核算内容输入框", SOURCE)
        self.assertIn("会计科目字段暂未渲染，重新展开当前明细后重试", SOURCE)
        self.assertRegex(
            text, r"drawer\s*=\s*fill_auxiliary_accounting\("
        )
        self.assertRegex(
            text,
            r"fill_internal_collection_reference_fields_before_tax_rate\(\s*frame, drawer, row",
        )
        split_function = next(
            node
            for node in TREE.body
            if isinstance(node, ast.FunctionDef)
            and node.name == "fill_split_income_type_after_income_item"
        )
        split_text = ast.get_source_segment(SOURCE, split_function)
        self.assertIn('if state["disabled"]:', split_text)
        self.assertIn('表明系统已自动取值', split_text)
        self.assertIn('_fill_reference_field_if_editable(', split_text)

    def test_tax_rate_settles_before_detail_drawer_closes(self):
        function = next(
            node
            for node in TREE.body
            if isinstance(node, ast.FunctionDef)
            and node.name == "settle_tax_rate_before_close"
        )
        text = ast.get_source_segment(SOURCE, function)
        self.assertRegex(text, r"tax_rate_input\.press\(\s*'Tab'")
        self.assertIn("task_sleep(timing_seconds('tax_rate_bind_wait'))", text)
        self.assertNotIn("clear_before_fill", text)

    def test_auxiliary_confirmation_does_not_require_detail_readback(self):
        functions = {
            node.name: ast.get_source_segment(SOURCE, node)
            for node in TREE.body
            if isinstance(node, ast.FunctionDef)
        }
        fill_text = functions["fill_auxiliary_accounting"]
        input_text = functions["_fill_auxiliary_code_input"]
        once_text = functions["_fill_auxiliary_accounting_once"]
        self.assertEqual(once_text.count("_fill_auxiliary_code_input("), 1)
        self.assertIn("_click_auxiliary_reference_option(", once_text)
        self.assertIn("保留已输入编号并重新打开下拉框", once_text)
        self.assertIn("_fill_auxiliary_accounting_once(", fill_text)
        self.assertNotIn("_auxiliary_value_is_bound(", fill_text)
        self.assertNotIn("未完成绑定", fill_text)
        self.assertNotIn("明细字段未回显", SOURCE)
        self.assertIn("if attempt == 1:", input_text)
        self.assertEqual(input_text.count("_clear_reference_before_fill("), 1)
        self.assertNotIn("_settle_auxiliary_reference_selection(", once_text)
        self.assertIn("直接执行窗口确认", once_text)
        click_text = functions["_click_auxiliary_reference_option"]
        self.assertIn(
            "ul.refer-cascading-item li.refer-li[title]", click_text
        )
        self.assertIn("_dispatch_auxiliary_dom_click(option)", click_text)
        self.assertNotIn("option.click(", click_text)
        self.assertNotIn("_mark_auxiliary_reference_popup(option)", click_text)
        self.assertIn(
            "task_sleep(timing_seconds('auxiliary_option_ready_wait'))", click_text
        )
        confirm_text = functions["_confirm_auxiliary_modal"]
        self.assertEqual(
            confirm_text.count(
                "_dispatch_auxiliary_dom_click(confirm_button)"
            ),
            1,
        )
        dom_click_text = functions["_dispatch_auxiliary_dom_click"]
        self.assertIn("element.click()", dom_click_text)
        self.assertIn("new MouseEvent('mousedown'", dom_click_text)
        self.assertIn(
            "task_sleep(timing_seconds('auxiliary_selection_commit_wait'))",
            confirm_text,
        )
        self.assertIn(
            "task_sleep(timing_seconds('auxiliary_confirm_commit_wait'))",
            confirm_text,
        )
        self.assertIn(
            "task_sleep(timing_seconds('auxiliary_modal_open_wait'))", once_text
        )
        time_defaults = self.time_helpers["default_time_settings"]()
        self.assertEqual(time_defaults["auxiliary_modal_open_wait"], 3.0)
        self.assertEqual(time_defaults["auxiliary_option_ready_wait"], 2.0)
        self.assertEqual(time_defaults["auxiliary_selection_commit_wait"], 3.0)
        self.assertEqual(time_defaults["auxiliary_confirm_commit_wait"], 3.0)
        self.assertNotIn("visible_ncc_system_errors", confirm_text)
        self.assertNotIn("_visible_non_auxiliary_dialog_messages", SOURCE)
        self.assertNotIn("点击“确定”后NCC报错", SOURCE)
        self.assertIn("辅助核算窗口已确认", confirm_text)
        self.assertEqual(
            len(re.findall(r"visible_ncc_system_errors\(", SOURCE)), 2
        )

    def test_save_error_recovery_is_wired(self):
        self.assertIn("visible_ncc_system_errors", SOURCE)
        self.assertIn("结果需人工核对：", SOURCE)
        self.assertNotIn("NCC返回提示，保存结果需人工核对", SOURCE)
        self.assertIsNone(re.search(r"\bSTATUS_REVIEW\b", SOURCE))
        self.assertIn("processing_status = STATUS_ABNORMAL", SOURCE)
        self.assertIn("error_reason = e.reason", SOURCE)
        self.assertIn("NccCommitUncertainError", SOURCE)
        self.assertIn("NccSaveRejectedError", SOURCE)
        self.assertIn("已点击“取消”，等待确认窗口", SOURCE)
        self.assertIn("已点击“确定”，正在返回单据列表", SOURCE)
        self.assertIn("returned_to_list = cancel_card(frame)", SOURCE)
        functions = {
            node.name: ast.get_source_segment(SOURCE, node)
            for node in TREE.body
            if isinstance(node, ast.FunctionDef)
        }
        save_text = functions['save']
        self.assertNotIn("if not _wait_for_old_notifications_to_clear", save_text)
        self.assertIn("before_notifications=before_notifications", save_text)
        self.assertLess(
            save_text.index("save_choose.click("),
            save_text.index("_wait_for_action_result("),
        )
        action_wait_text = functions['_wait_for_action_result']
        self.assertEqual(action_wait_text.count('visible_ncc_messages(frame)'), 1)
        self.assertIn('messages=new_notifications', action_wait_text)
        self.assertIn('raise NccSaveRejectedError(', action_wait_text)
        running_text = functions['running_task']
        rejected = running_text.index('except NccSaveRejectedError as e:')
        close = running_text.index('close_ncc_save_error(frame)', rejected)
        cancel = running_text.index('returned_to_list = cancel_card(frame)', rejected)
        add = running_text.index('click_add_bill(frame)', rejected)
        self.assertLess(close, cancel)
        self.assertLess(cancel, add)


if __name__ == "__main__":
    unittest.main()
