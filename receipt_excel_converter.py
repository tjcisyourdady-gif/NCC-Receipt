# -*- coding: utf-8 -*-
"""按Excel对照表方案转换收款底稿，同一来源列允许复用。"""

from copy import copy
import json
import math
import os
import queue
import re
import shutil
import tempfile
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import PatternFill
from receipt_data_rules import load_receipt_workbook, read_receipt_macro, normalize_receipt_value
from openpyxl.utils import get_column_letter, column_index_from_string


OUTPUT_ONLY_HEADERS = {'单据号', '单据编号', '填写状态', '异常原因'}
OPTIONAL_MAPPING_HEADERS = {
    '部门编码', '部门', '部门名称',
    '付款银行账号', '付款银行账户', '付款账号',
    '备注',
}
TARGET_HEADER_ALIAS_GROUPS = (
    ('收款银行账号', '银行账号'),
    ('付款银行账号', '付款银行账户', '付款账号'),
    ('部门', '部门名称'),
    ('单据号', '单据编号'),
    ('费用所属期间', '费用所属期'),
    ('税率（非增值税收益等）', '税率'),
)
DEFAULT_OUTPUT_FILENAME = '收款结算单底稿导入模板.xlsm'
MAPPING_TABLE_FILENAME = '对照表方案.xlsx'
LATEST_MAPPING_FILENAME = 'receipt_mapping_scheme.json'
IDENTIFIER_HEADERS = {
    '公司编码', '部门编码', '银行账号', '收款银行账号', '收款银行账户',
    '付款银行账号', '付款银行账户', '付款账号',
    '客商编码', '会计科目', '辅助核算',
    '被代付账套/被代收账套名称', '被代收方（被收回方）项目档案',
    '被代付账套/被代付账套名称',
    '项目档案名称',
}
SCIENTIFIC_TEXT_PATTERN = re.compile(
    r'^[+-]?(?:\d+(?:\.\d*)?|\.\d+)[eE][+-]?\d+$'
)
SCIENTIFIC_FORMAT_PATTERN = re.compile(
    r'[0#?][0#?.,]*[eE][+-]?[0#?]+'
)
ERROR_CELL_FILL = PatternFill(fill_type='solid', fgColor='FFFF00')
EXCEL_SAFE_INTEGER_LIMIT = 10 ** 15
EDGE_WHITESPACE = (
    ' \t\r\n\v\f\u00a0\u1680\u180e\u2000\u2001\u2002\u2003'
    '\u2004\u2005\u2006\u2007\u2008\u2009\u200a\u200b\u200c'
    '\u200d\u2028\u2029\u202f\u205f\u2060\u3000\ufeff'
)


class ConversionValidationError(ValueError):
    """映射底稿已生成并标黄，但账号/编码类数据未通过最终验证。"""

    def __init__(self, output_path, issues):
        self.output_path = os.path.abspath(output_path)
        self.issues = list(issues)
        preview = []
        for issue in self.issues[:12]:
            source_note = (
                f'，来源 {issue["source_ref"]}'
                if issue.get('source_ref') else ''
            )
            preview.append(
                f'{issue["target_ref"]}（{issue["header"]}{source_note}）：'
                f'{issue["reason"]}'
            )
        if len(self.issues) > 12:
            preview.append(f'另有{len(self.issues) - 12}处')
        message = (
            f'映射完成后的数据验证未通过，共{len(self.issues)}个单元格。\n'
            '相关错误单元格已在生成底稿中用黄色标记。\n'
            '请把源Excel中的银行账号、编码等标识字段改为文本格式，'
            '重新录入完整值后再次转换。\n\n'
            + '\n'.join(preview)
            + f'\n\n已标记底稿：\n{self.output_path}'
        )
        super().__init__(message)


def clean_edge_text(value):
    """去除普通、全角与不可见的首尾空白。"""
    if value is None:
        return ''
    return str(value).strip(EDGE_WHITESPACE)


def normalize_header(value):
    """比较标题时忽略尾部必填星号和空白。"""
    return re.sub(r'[\s*＊]+$', '', clean_edge_text(value)).strip()


def value_is_blank(value):
    return value is None or (isinstance(value, str) and clean_edge_text(value) == '')


def is_scientific_text(value):
    """仅识别纯数字科学计数法，不把 BGY.E01 等普通编码误判。"""
    if not isinstance(value, str):
        return False
    return SCIENTIFIC_TEXT_PATTERN.fullmatch(clean_edge_text(value)) is not None


def uses_scientific_number_format(number_format):
    """识别 0.00E+00 等 Excel 科学计数格式，忽略引号内的字面量。"""
    text = clean_edge_text(number_format)
    if not text:
        return False
    text = re.sub(r'"[^"]*"', '', text)
    text = re.sub(r'\\.', '', text)
    return SCIENTIFIC_FORMAT_PATTERN.search(text) is not None


def identifier_source_issue(value, number_format='General'):
    """返回账号/编码类源单元格的科学计数或 Excel 精度风险。"""
    if value_is_blank(value):
        return ''
    if is_scientific_text(value):
        return '映射来源是科学计数法文本，无法保证账号或编码完整'
    if not isinstance(value, (int, float)) or isinstance(value, bool):
        return ''
    if isinstance(value, float) and not math.isfinite(value):
        return '映射来源不是有限数值，无法作为账号或编码使用'
    if uses_scientific_number_format(number_format):
        return '映射来源使用了Excel科学计数法格式'
    if 'e' in str(value).lower():
        return '映射来源已被Excel读取为科学计数法数值'
    if float(value).is_integer() and abs(value) >= EXCEL_SAFE_INTEGER_LIMIT:
        return '映射来源是超过Excel 15位安全精度的数值，可能已发生精度丢失'
    return ''


def _mapped_values_equal(expected, actual):
    if value_is_blank(expected) and value_is_blank(actual):
        return True
    return expected == actual


def validate_converted_workbook(
    output_path, target_sheet_name, target_headers, converted_rows,
    source_issues, target_start_row=3,
):
    """在数值转换最后一步复验落盘结果，并把错误目标单元格标为黄色。"""
    target_by_column = {item['column']: item for item in target_headers}
    issues = [dict(issue) for issue in source_issues]
    issue_keys = {
        (issue['target_row'], issue['target_column'], issue['reason'])
        for issue in issues
    }

    def add_issue(target_row, target_column, reason, source_ref=''):
        key = (target_row, target_column, reason)
        if key in issue_keys:
            return
        issue_keys.add(key)
        info = target_by_column[target_column]
        issues.append({
            'target_row': target_row,
            'target_column': target_column,
            'target_ref': (
                f'{target_sheet_name}!'
                f'{get_column_letter(target_column)}{target_row}'
            ),
            'source_ref': source_ref,
            'header': info['header'],
            'reason': reason,
        })

    workbook = load_receipt_workbook(output_path, data_only=False)
    try:
        if target_sheet_name not in workbook.sheetnames:
            raise ValueError(f'生成底稿缺少工作表“{target_sheet_name}”')
        sheet = workbook[target_sheet_name]
        for offset, converted in enumerate(converted_rows):
            target_row = target_start_row + offset
            for target_column, expected in converted.items():
                info = target_by_column[target_column]
                cell = sheet.cell(target_row, target_column)
                if not _mapped_values_equal(expected, cell.value):
                    add_issue(
                        target_row,
                        target_column,
                        '映射值写入后与来源值不一致',
                    )
                if (
                    info['normalized'] not in IDENTIFIER_HEADERS
                    or value_is_blank(cell.value)
                ):
                    continue
                if not isinstance(cell.value, str):
                    add_issue(
                        target_row,
                        target_column,
                        '账号或编码在生成底稿中不是文本类型',
                    )
                elif is_scientific_text(cell.value):
                    add_issue(
                        target_row,
                        target_column,
                        '生成底稿中仍是科学计数法，无法用于NCC填报',
                    )
                if cell.number_format != '@':
                    add_issue(
                        target_row,
                        target_column,
                        '账号或编码在生成底稿中未设置为文本格式',
                    )
        merged_issues = {}
        for issue in issues:
            key = (issue['target_row'], issue['target_column'])
            existing = merged_issues.get(key)
            if existing is None:
                merged_issues[key] = dict(issue)
                continue
            if issue['reason'] not in existing['reason']:
                existing['reason'] += '；' + issue['reason']
            if not existing.get('source_ref') and issue.get('source_ref'):
                existing['source_ref'] = issue['source_ref']
        issues = list(merged_issues.values())
        if issues:
            for issue in issues:
                sheet.cell(
                    issue['target_row'], issue['target_column']
                ).fill = copy(ERROR_CELL_FILL)
            workbook.save(output_path)
    finally:
        workbook.close()
    return issues


def list_source_sheets(file_path):
    workbook = load_workbook(file_path, read_only=True, data_only=False)
    try:
        return list(workbook.sheetnames), workbook.active.title
    finally:
        workbook.close()


def read_nonblank_headers(file_path, sheet_name, header_row):
    """按用户选定标题行读取所有非空列，保留真实列号。"""
    if int(header_row) < 1:
        raise ValueError('标题行号必须大于等于1')
    workbook = load_workbook(file_path, read_only=True, data_only=False)
    try:
        if sheet_name not in workbook.sheetnames:
            raise ValueError(f'Excel中不存在工作表“{sheet_name}”')
        sheet = workbook[sheet_name]
        if header_row > sheet.max_row:
            raise ValueError(f'标题行号{header_row}超过工作表最大行号{sheet.max_row}')
        values = next(sheet.iter_rows(
            min_row=header_row, max_row=header_row,
            min_col=1, max_col=sheet.max_column, values_only=True,
        ))
        headers = []
        for column, value in enumerate(values, start=1):
            text = clean_edge_text(value)
            if text:
                headers.append({
                    'column': column,
                    'header': text,
                    'display': f'{text} 〔{get_column_letter(column)}列〕',
                })
        if not headers:
            raise ValueError(f'工作表“{sheet_name}”第{header_row}行没有读取到非空标题')
        return headers
    finally:
        workbook.close()


def read_target_headers(template_path, sheet_name, header_row=2):
    """读取收款模板字段，标记单据号、状态和异常原因为程序回写列。"""
    workbook = load_workbook(template_path, read_only=True, data_only=False)
    try:
        if sheet_name not in workbook.sheetnames:
            raise ValueError(f'正式模板缺少工作表“{sheet_name}”')
        sheet = workbook[sheet_name]
        values = next(sheet.iter_rows(
            min_row=header_row, max_row=header_row,
            min_col=1, max_col=sheet.max_column, values_only=True,
        ))
        headers = []
        for column, value in enumerate(values, start=1):
            text = clean_edge_text(value)
            if text:
                normalized = normalize_header(text)
                headers.append({
                    'column': column,
                    'header': text,
                    'normalized': normalized,
                    'output_only': normalized in OUTPUT_ONLY_HEADERS,
                })
        if not headers:
            raise ValueError('正式模板没有读取到目标字段')
        return headers
    finally:
        workbook.close()


def assign_unique_mapping(mapping, target_column, source_column):
    """每个目标只有一个来源；允许同一来源映射到多个目标。保留旧API名称。"""
    updated = dict(mapping)
    updated[target_column] = source_column
    return updated


def validate_complete_mapping(mapping, target_headers):
    required = [item for item in target_headers
                if not item['output_only'] and item['normalized'] not in OPTIONAL_MAPPING_HEADERS]
    missing = [item for item in required if item['column'] not in mapping]
    if missing:
        preview = '、'.join(item['header'] for item in missing[:12])
        if len(missing) > 12:
            preview += f'，另有{len(missing) - 12}项'
        raise ValueError(
            '收款结算单底稿的全部输入字段都必须映射；'
            f'当前还缺少{len(missing)}项：{preview}'
        )
    return required


def list_mapping_sheets(table_path):
    """读取实际方案名称；示例页和隐藏的内部页不作为可执行方案。"""
    workbook = load_workbook(table_path, read_only=True, data_only=False)
    try:
        names = [s.title for s in workbook if s.sheet_state == 'visible'
                 and '在此填写无效' not in s.title]
        if not names:
            raise ValueError('对照表中没有可选择的方案sheet（示例页不参与转换）')
        return names
    finally:
        workbook.close()


def _positive_table_integer(value, location):
    text = clean_edge_text(value)
    if isinstance(value, bool) or not re.fullmatch(r'[0-9]+', text):
        raise ValueError(f'{location}必须填写从1开始的整数')
    number = int(text)
    if number < 1 or number > 1048576:
        raise ValueError(f'{location}必须填写1～1048576之间的整数')
    return number


def read_mapping_table(table_path, scheme_name, target_headers):
    """解析C/D/E/F布局；只读取配置，不修改对照表或源数据。"""
    if scheme_name not in list_mapping_sheets(table_path):
        raise ValueError(f'“{scheme_name}”不是可执行的转换方案')
    workbook = load_workbook(table_path, read_only=True, data_only=False)
    try:
        sheet = workbook[scheme_name]
        rows = sheet.iter_rows(min_row=2, min_col=3, max_col=6)
        headings = [clean_edge_text(c.value) for c in next(rows)]
        expected = ['对照表头在第几个sheet', '对照表头在第几行',
                    '对照表头名称', '底稿使用表头名称']
        if headings != expected:
            raise ValueError(f'{scheme_name}!C2:F2表头不正确，请使用随工具提供的新版对照表')
        by_name = {item['normalized']: item for item in target_headers}
        for group in TARGET_HEADER_ALIAS_GROUPS:
            matched_target = next(
                (by_name[alias] for alias in group if alias in by_name), None
            )
            if matched_target is not None:
                for alias in group:
                    by_name.setdefault(alias, matched_target)
        seen = set()
        sources = {}
        sheet_index = header_row = None
        for row_number, cells in enumerate(rows, start=3):
            values = [c.value for c in cells]
            if any(c.data_type == 'f' for c in cells):
                raise ValueError(f'{scheme_name}第{row_number}行配置不允许使用公式')
            if all(value_is_blank(v) for v in values):
                continue
            if row_number == 3:
                sheet_index = _positive_table_integer(values[0], f'{scheme_name}!C3 来源sheet序号')
                header_row = _positive_table_integer(values[1], f'{scheme_name}!D3 表头行号')
            else:
                for column, value, common in [('C', values[0], sheet_index),
                                               ('D', values[1], header_row)]:
                    if not value_is_blank(value):
                        current = _positive_table_integer(value, f'{scheme_name}!{column}{row_number}')
                        if current != common:
                            raise ValueError(f'{scheme_name}!{column}{row_number}与{column}3不一致；'
                                             '每个方案只使用同一个来源sheet和表头行，不支持跨表拼接')
            if not isinstance(values[3], str) or normalize_header(values[3]) not in by_name:
                raise ValueError(f'{scheme_name}!F{row_number}不是正式底稿字段：{values[3]}')
            name = normalize_header(values[3])
            target = by_name[name]
            canonical_name = target['normalized']
            if canonical_name in seen:
                raise ValueError(f'{scheme_name}!F{row_number}底稿字段重复：{name}')
            seen.add(canonical_name)
            if target['output_only']:
                if not value_is_blank(values[2]):
                    raise ValueError(f'{scheme_name}!E{row_number}“{name}”由程序回写，请留空')
                continue
            if canonical_name in OPTIONAL_MAPPING_HEADERS and value_is_blank(values[2]):
                continue
            if not isinstance(values[2], str) or not clean_edge_text(values[2]):
                raise ValueError(f'{scheme_name}!E{row_number}“{name}”必须填写文本类型的来源表头名称')
            sources[target['column']] = clean_edge_text(values[2])
        if sheet_index is None or header_row is None:
            raise ValueError(f'{scheme_name}!C3、D3必须填写来源sheet序号和表头行号')
        missing = [item['header'] for item in target_headers
                   if item['normalized'] not in seen
                   and item['normalized'] not in OPTIONAL_MAPPING_HEADERS]
        if missing:
            raise ValueError('对照表F列缺少底稿字段：' + '、'.join(missing))
        validate_complete_mapping(sources, target_headers)
        return {'sheet_index': sheet_index, 'header_row': header_row,
                'sources': sources, 'scheme_name': scheme_name}
    finally:
        workbook.close()


def parse_source_column_reference(value):
    """解析名称*物理列号；数字从1开始，也支持A～XFD。末尾必填*仍是名称。"""
    text = clean_edge_text(value)
    candidate = text
    if ((candidate.startswith('（') and candidate.endswith('）'))
            or (candidate.startswith('(') and candidate.endswith(')'))):
        inner = candidate[1:-1]
        if re.search(r'[*＊][^*＊]+$', inner):
            candidate = inner
    match = re.fullmatch(r'(.+)[*＊]([^*＊]+)', candidate)
    if match is None:
        return text, None
    name, suffix = clean_edge_text(match.group(1)), clean_edge_text(match.group(2))
    if not name:
        raise ValueError(f'来源配置“{text}”缺少表头名称')
    if re.fullmatch(r'[0-9]+', suffix):
        column = int(suffix)
    elif re.fullmatch(r'[A-Za-z]{1,3}', suffix):
        column = column_index_from_string(suffix.upper())
    else:
        raise ValueError(f'来源配置“{text}”列号无效，请填写“名称*3”或“名称*C”')
    if not 1 <= column <= 16384:
        raise ValueError(f'来源配置“{text}”列号超出Excel范围（1～16384，A～XFD）')
    return name, column


def resolve_table_mapping(source_path, config, target_headers):
    """根据名称重新查找物理列，不依赖上次源文件的列顺序。"""
    names, _ = list_source_sheets(source_path)
    index = config['sheet_index']
    if index > len(names):
        raise ValueError(f'方案指定来源第{index}个sheet，但原始Excel只有{len(names)}个sheet')
    sheet_name = names[index - 1]
    source_headers = read_nonblank_headers(source_path, sheet_name, config['header_row'])
    columns = {}
    by_column = {}
    for item in source_headers:
        columns.setdefault(item['header'], []).append(item['column'])
        by_column[item['column']] = item['header']
    references = {target: parse_source_column_reference(value)
                  for target, value in config['sources'].items()}
    # 即使尚未被映射的字段重名，也需用户处理，不能静默跳过源表异常。
    specified_names = {name for name, column in references.values() if column is not None}
    unresolved = {name: positions for name, positions in columns.items()
                  if len(positions) > 1 and name not in specified_names}
    if unresolved:
        details = '；'.join(
            f'“{name}”（' + '、'.join(f'{get_column_letter(c)}列/{c}' for c in positions) + '）'
            for name, positions in list(unresolved.items())[:12]
        )
        raise ValueError(f'原始Excel“{sheet_name}”第{config["header_row"]}行表头重复，无法确定取哪一列：'
                         f'{details}。请在对照表E列用“名称*列号”明确指定；不用的重名表头请在源表改名')
    mapping = {}
    for target_column, (source_name, explicit_column) in references.items():
        if explicit_column is not None:
            actual = by_column.get(explicit_column)
            if actual != source_name:
                raise ValueError(f'原始Excel“{sheet_name}”第{config["header_row"]}行'
                                 f'{get_column_letter(explicit_column)}列/{explicit_column}表头不符：'
                                 f'配置“{source_name}”，实际“{actual or "空白或不存在"}”；本次不转换')
            mapping = assign_unique_mapping(mapping, target_column, explicit_column)
            continue
        matches = columns.get(source_name, [])
        if not matches:
            raise ValueError(f'原始Excel“{sheet_name}”第{config["header_row"]}行找不到表头“{source_name}”')
        if len(matches) > 1:
            letters = '、'.join(get_column_letter(c) for c in matches)
            raise ValueError(f'原始Excel表头“{source_name}”重复（{letters}列），无法确定取哪一列；'
                             '每项引用都必须用“名称*列号”指定，不能有的指定、有的只写名称')
        mapping = assign_unique_mapping(mapping, target_column, matches[0])
    validate_complete_mapping(mapping, target_headers)
    return sheet_name, mapping


def ask_mapping_sheet(parent, table_path):
    names = list_mapping_sheets(table_path)
    result = {'value': None}
    dialog = tk.Toplevel(parent)
    dialog.title('选择底稿转换方案')
    dialog.transient(parent)
    dialog.resizable(False, False)
    frame = ttk.Frame(dialog, padding=18)
    frame.pack(fill='both', expand=True)
    ttk.Label(frame, text='请选择对照表中的转换方案：').pack(anchor='w')
    choice = tk.StringVar(value=names[0])
    box = ttk.Combobox(frame, textvariable=choice, values=names,
                       state='readonly', width=42)
    box.pack(fill='x', pady=10)
    ttk.Label(frame, text=f'对照表：{table_path}\n'
              'C3：来源sheet序号；D3：表头行；E列：来源表头；F列：底稿字段。\n'
              '重名表头写“名称*列号”（如名称*3），同一来源列可用于多个底稿字段。',
              wraplength=460).pack(anchor='w')

    def confirm():
        result['value'] = choice.get()
        dialog.destroy()

    buttons = ttk.Frame(frame)
    buttons.pack(anchor='e', pady=(14, 0))
    ttk.Button(buttons, text='下一步', command=confirm).pack(side='left')
    ttk.Button(buttons, text='取消', command=dialog.destroy).pack(side='left', padx=(8, 0))
    dialog.grab_set()
    box.focus_set()
    parent.wait_window(dialog)
    return result['value']


def default_output_path(template_path):
    return os.path.join(
        os.path.dirname(os.path.abspath(template_path)), DEFAULT_OUTPUT_FILENAME
    )


def default_mapping_scheme_path(template_path):
    tool_directory = os.path.dirname(os.path.abspath(template_path))
    internal_directory = os.path.join(tool_directory, 'ms-playwright（勿动）')
    if not os.path.isdir(internal_directory):
        internal_directory = tool_directory
    return os.path.join(internal_directory, LATEST_MAPPING_FILENAME)


def _source_signature(source_headers):
    return [
        {'column': int(item['column']), 'header': clean_edge_text(item['header'])}
        for item in source_headers
    ]


def _target_signature(target_headers):
    return [
        {
            'column': int(item['column']),
            'header': normalize_header(item['header']),
            'output_only': bool(item['output_only']),
        }
        for item in target_headers
    ]


def save_latest_mapping_scheme(scheme_path, source_headers, target_headers, mapping):
    """原子覆盖，只保留一份最新的字段方案。"""
    validate_complete_mapping(mapping, target_headers)
    payload = {
        'version': 1,
        'source_fields': _source_signature(source_headers),
        'target_fields': _target_signature(target_headers),
        'mapping': {str(int(target)): int(source) for target, source in mapping.items()},
    }
    scheme_path = os.path.abspath(scheme_path)
    directory = os.path.dirname(scheme_path)
    os.makedirs(directory, exist_ok=True)
    handle, temporary_path = tempfile.mkstemp(
        prefix='.receipt_mapping_', suffix='.json', dir=directory
    )
    try:
        with os.fdopen(handle, 'w', encoding='utf-8') as stream:
            json.dump(payload, stream, ensure_ascii=False, indent=2)
        os.replace(temporary_path, scheme_path)
    except Exception:
        try:
            os.close(handle)
        except OSError:
            pass
        try:
            os.remove(temporary_path)
        except OSError:
            pass
        raise


def load_latest_mapping_scheme(scheme_path, source_headers, target_headers):
    """仅当标题文本、源列号及模板字段都完全一致时复用映射。"""
    if not scheme_path or not os.path.isfile(scheme_path):
        return None
    try:
        with open(scheme_path, 'r', encoding='utf-8') as stream:
            payload = json.load(stream)
        if payload.get('version') != 1:
            return None
        if payload.get('source_fields') != _source_signature(source_headers):
            return None
        if payload.get('target_fields') != _target_signature(target_headers):
            return None
        mapping = {
            int(target): int(source)
            for target, source in payload.get('mapping', {}).items()
        }
        available_sources = {int(item['column']) for item in source_headers}
        if not set(mapping.values()).issubset(available_sources):
            return None
        validate_complete_mapping(mapping, target_headers)
        return mapping
    except (OSError, ValueError, TypeError, KeyError, json.JSONDecodeError):
        return None


def _copy_template_row_style(sheet, source_row, target_row, max_column):
    if source_row == target_row:
        return
    source_height = sheet.row_dimensions[source_row].height
    if source_height is not None:
        sheet.row_dimensions[target_row].height = source_height
    for column in range(1, max_column + 1):
        source = sheet.cell(source_row, column)
        target = sheet.cell(target_row, column)
        if source.has_style:
            target._style = copy(source._style)
        target.font = copy(source.font)
        target.fill = copy(source.fill)
        target.border = copy(source.border)
        target.alignment = copy(source.alignment)
        target.protection = copy(source.protection)
        target.number_format = source.number_format


def _identifier_text(value):
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return clean_edge_text(value)


def convert_custom_workbook(
    source_path, source_sheet_name, header_row, mapping, template_path, output_path,
    target_sheet_name='收款结算明细', target_header_row=2, target_start_row=3,
):
    """读取用户 Excel，清空模板旧数据，按映射写入标准收款底稿。"""
    template_macro = read_receipt_macro(template_path)
    if os.path.splitext(output_path)[1].lower() != '.xlsm':
        raise ValueError('映射结果必须保存为 .xlsm，避免丢失无损粘贴宏')
    if os.path.abspath(source_path) == os.path.abspath(output_path):
        raise ValueError('转换结果不能覆盖用户原始Excel，请选择新的源文件后重试')
    target_headers = read_target_headers(template_path, target_sheet_name, target_header_row)
    validate_complete_mapping(mapping, target_headers)
    target_by_column = {item['column']: item for item in target_headers}
    for target_column in mapping:
        info = target_by_column.get(target_column)
        if info is None:
            raise ValueError(f'目标第{target_column}列不属于正式收款模板')
        if info['output_only']:
            raise ValueError(f'“{info["header"]}”由程序回写，不能设置源字段映射')

    formula_book = load_workbook(source_path, read_only=True, data_only=False)
    value_book = load_workbook(source_path, read_only=True, data_only=True)
    try:
        if source_sheet_name not in formula_book.sheetnames:
            raise ValueError(f'源Excel不存在工作表“{source_sheet_name}”')
        formula_sheet = formula_book[source_sheet_name]
        value_sheet = value_book[source_sheet_name]
        source_columns = sorted(set(mapping.values()))
        first_column, last_column = source_columns[0], source_columns[-1]
        converted_rows, cache_errors = [], []
        numeric_warnings, source_issues = [], []
        formula_rows = formula_sheet.iter_rows(
            min_row=header_row + 1, max_row=formula_sheet.max_row,
            min_col=first_column, max_col=last_column,
        )
        value_rows = value_sheet.iter_rows(
            min_row=header_row + 1, max_row=value_sheet.max_row,
            min_col=first_column, max_col=last_column,
        )
        for row_number, formula_cells, value_cells in zip(
            range(header_row + 1, formula_sheet.max_row + 1),
            formula_rows, value_rows,
        ):
            converted, has_data, row_source_issues = {}, False, []
            for target_column, source_column in mapping.items():
                offset = source_column - first_column
                formula_cell = formula_cells[offset]
                value = value_cells[offset].value
                if formula_cell.data_type == 'f' and value is None:
                    cache_errors.append(
                        f'{source_sheet_name}!{get_column_letter(source_column)}{row_number}'
                    )
                    continue
                if isinstance(value, str):
                    value = clean_edge_text(value)
                info = target_by_column[target_column]
                if info['normalized'] in IDENTIFIER_HEADERS and not value_is_blank(value):
                    source_ref = (
                        f'{source_sheet_name}!'
                        f'{get_column_letter(source_column)}{row_number}'
                    )
                    source_issue = identifier_source_issue(
                        value, formula_cell.number_format
                    )
                    if source_issue:
                        row_source_issues.append({
                            'target_column': target_column,
                            'source_ref': source_ref,
                            'header': info['header'],
                            'reason': source_issue,
                        })
                    if isinstance(value, (int, float)) and not isinstance(value, bool):
                        numeric_warnings.append(f'{info["header"]}（源第{row_number}行）')
                    value = _identifier_text(value)
                converted[target_column] = normalize_receipt_value(value, info['normalized'])
                has_data = has_data or not value_is_blank(value)
            if has_data:
                target_row = target_start_row + len(converted_rows)
                for issue in row_source_issues:
                    issue['target_row'] = target_row
                    issue['target_ref'] = (
                        f'{target_sheet_name}!'
                        f'{get_column_letter(issue["target_column"])}{target_row}'
                    )
                    source_issues.append(issue)
                converted_rows.append(converted)
        if cache_errors:
            preview = '、'.join(cache_errors[:12])
            if len(cache_errors) > 12:
                preview += f'，另有{len(cache_errors) - 12}处'
            raise ValueError(
                '源Excel映射区域含没有缓存结果的公式：'
                f'{preview}。请先用Excel打开并保存，或粘贴为值后重试'
            )
        if not converted_rows:
            raise ValueError('标题行以下没有读取到任何映射字段数据')
    finally:
        formula_book.close()
        value_book.close()

    output_path = os.path.abspath(output_path)
    template_path = os.path.abspath(template_path)
    output_directory = os.path.dirname(output_path)
    os.makedirs(output_directory, exist_ok=True)
    replaces_template = template_path == output_path
    temporary_output = None
    if replaces_template:
        handle, temporary_output = tempfile.mkstemp(
            prefix='.收款结算单映射_', suffix='.xlsm', dir=output_directory
        )
        os.close(handle)
        workbook_source, save_path = template_path, temporary_output
    else:
        shutil.copy2(template_path, output_path)
        workbook_source = save_path = output_path
    workbook = load_receipt_workbook(workbook_source)
    try:
        sheet = workbook[target_sheet_name]
        max_column = max(target_by_column)
        for row in sheet.iter_rows(
            min_row=target_start_row, max_row=max(sheet.max_row, target_start_row),
            min_col=1, max_col=max_column,
        ):
            for cell in row:
                if not isinstance(cell, MergedCell):
                    cell.value = None
        for offset, converted in enumerate(converted_rows):
            target_row = target_start_row + offset
            _copy_template_row_style(sheet, target_start_row, target_row, max_column)
            for target_column, value in converted.items():
                cell = sheet.cell(target_row, target_column, value=value)
                if target_by_column[target_column]['normalized'] in IDENTIFIER_HEADERS:
                    cell.number_format = '@'
            for column in range(1, max_column + 1):
                cell = sheet.cell(target_row, column)
                aligned = copy(cell.alignment)
                aligned.horizontal, aligned.vertical = 'center', 'center'
                cell.alignment = aligned
        workbook.save(save_path)
    except Exception:
        try:
            workbook.close()
        except Exception:
            pass
        if temporary_output:
            try:
                os.remove(temporary_output)
            except OSError:
                pass
        raise
    finally:
        try:
            workbook.close()
        except Exception:
            pass
    try:
        validation_issues = validate_converted_workbook(
            save_path,
            target_sheet_name,
            target_headers,
            converted_rows,
            source_issues,
            target_start_row=target_start_row,
        )
        if read_receipt_macro(save_path) != template_macro:
            raise ValueError('映射生成后宏内容不一致，停止发布转换结果')
    except Exception:
        if temporary_output:
            try:
                os.remove(temporary_output)
            except OSError:
                pass
        raise
    if replaces_template:
        try:
            os.replace(temporary_output, output_path)
        finally:
            # 成功时临时文件已被替换；失败（例如Excel占用）时也不留下额外底稿。
            if os.path.exists(temporary_output):
                os.remove(temporary_output)
    if validation_issues:
        raise ConversionValidationError(output_path, validation_issues)
    return {
        'row_count': len(converted_rows),
        'mapping_count': len(mapping),
        'warnings': list(dict.fromkeys(numeric_warnings)),
        'output_path': output_path,
    }


def ask_source_settings(parent, source_path):
    """选择源工作表；单 Sheet 情况只询问标题行号。"""
    sheet_names, active_name = list_source_sheets(source_path)
    if len(sheet_names) == 1:
        row = simpledialog.askinteger(
            '输入标题行号', f'将读取工作表“{sheet_names[0]}”。\n请输入标题所在行号：',
            parent=parent, minvalue=1, maxvalue=1048576,
        )
        return (sheet_names[0], row) if row else None
    result = {'value': None}
    dialog = tk.Toplevel(parent)
    dialog.title('选择工作表并输入标题行')
    dialog.transient(parent)
    dialog.resizable(False, False)
    frame = ttk.Frame(dialog, padding=16)
    frame.pack(fill='both', expand=True)
    ttk.Label(frame, text='源工作表：').grid(row=0, column=0, sticky='w')
    sheet_var = tk.StringVar(value=active_name)
    ttk.Combobox(
        frame, textvariable=sheet_var, values=sheet_names, state='readonly', width=34
    ).grid(row=0, column=1, padx=(8, 0), pady=(0, 10))
    ttk.Label(frame, text='标题行号：').grid(row=1, column=0, sticky='w')
    row_var = tk.StringVar(value='1')
    row_entry = ttk.Entry(frame, textvariable=row_var, width=12)
    row_entry.grid(row=1, column=1, sticky='w', padx=(8, 0))
    def confirm():
        try:
            row = int(row_var.get().strip())
            if row < 1:
                raise ValueError
        except ValueError:
            messagebox.showerror('行号错误', '标题行号必须是大于0的整数', parent=dialog)
            return
        result['value'] = (sheet_var.get(), row)
        dialog.destroy()
    buttons = ttk.Frame(frame)
    buttons.grid(row=2, column=0, columnspan=2, pady=(16, 0), sticky='e')
    ttk.Button(buttons, text='确定', command=confirm, width=9).pack(side='left')
    ttk.Button(buttons, text='取消', command=dialog.destroy, width=9).pack(side='left', padx=(8, 0))
    dialog.protocol('WM_DELETE_WINDOW', dialog.destroy)
    dialog.grab_set()
    row_entry.focus_set()
    parent.wait_window(dialog)
    return result['value']


class MappingDialog(tk.Toplevel):
    """上方源字段、下方目标字段的字段拖拽映射窗口。"""

    def __init__(self, parent, source_headers, target_headers, on_generate):
        super().__init__(parent)
        self.title('设置数值转换')
        self.geometry('1120x680')
        self.minsize(900, 560)
        self.transient(parent)
        self.source_headers = source_headers
        self.target_headers = target_headers
        self.required_targets = [
            item for item in target_headers
            if not item['output_only'] and item['normalized'] not in OPTIONAL_MAPPING_HEADERS
        ]
        self.on_generate = on_generate
        self.mapping = {}
        self.source_by_column = {item['column']: item for item in source_headers}
        self.source_labels, self.target_slots = {}, {}
        self.selected_source = self.drag_source = None
        self.drag_window = None
        self.generating = False
        self._build_ui()
        self.protocol('WM_DELETE_WINDOW', self._request_close)
        self.grab_set()

    def _build_ui(self):
        outer = ttk.Frame(self, padding=12)
        outer.pack(fill='both', expand=True)
        ttk.Label(
            outer,
            text='把上方源字段拖到下方目标字段的空格中。收款底稿全部输入字段必须映射；同一源字段可重复用于多个目标，双击目标格可清除。',
            font=('微软雅黑', 10, 'bold'),
        ).pack(anchor='w', pady=(0, 8))
        source_box = ttk.LabelFrame(outer, text='自己的Excel标题字段', padding=8)
        source_box.pack(fill='x')
        self.source_canvas = tk.Canvas(source_box, height=165, highlightthickness=0, background='#f8fafc')
        source_scroll = ttk.Scrollbar(source_box, orient='vertical', command=self.source_canvas.yview)
        self.source_canvas.configure(yscrollcommand=source_scroll.set)
        self.source_canvas.pack(side='left', fill='both', expand=True)
        source_scroll.pack(side='right', fill='y')
        source_inner = ttk.Frame(self.source_canvas)
        window = self.source_canvas.create_window((0, 0), window=source_inner, anchor='nw')
        source_inner.bind('<Configure>', lambda _event: self.source_canvas.configure(scrollregion=self.source_canvas.bbox('all')))
        self.source_canvas.bind('<Configure>', lambda event: self.source_canvas.itemconfigure(window, width=event.width))
        for widget in (self.source_canvas, source_inner):
            widget.bind('<MouseWheel>', self._scroll_source_fields)
        for index, item in enumerate(self.source_headers):
            label = tk.Label(
                source_inner, text=item['display'], bg='#dbeafe', fg='#1e3a8a',
                relief='groove', bd=1, padx=8, pady=6, cursor='hand2', font=('微软雅黑', 9),
            )
            label.grid(row=index // 4, column=index % 4, sticky='ew', padx=4, pady=4)
            source_inner.columnconfigure(index % 4, weight=1)
            column = item['column']
            label.bind('<Button-1>', lambda event, col=column: self._select_source(col))
            label.bind('<ButtonPress-1>', lambda event, col=column: self._start_drag(event, col), add='+')
            label.bind('<B1-Motion>', self._drag_motion)
            label.bind('<ButtonRelease-1>', self._drop_source)
            label.bind('<MouseWheel>', self._scroll_source_fields)
            self.source_labels[column] = label
        ttk.Label(outer, text='收款结算单标准字段（目标格在标题上方）', font=('微软雅黑', 10, 'bold')).pack(anchor='w', pady=(12, 5))
        target_box = ttk.Frame(outer)
        target_box.pack(fill='both', expand=True)
        target_canvas = tk.Canvas(target_box, height=270, highlightthickness=0, background='#ffffff')
        horizontal = ttk.Scrollbar(target_box, orient='horizontal', command=target_canvas.xview)
        target_canvas.configure(xscrollcommand=horizontal.set)
        target_canvas.pack(fill='both', expand=True)
        horizontal.pack(fill='x')
        target_inner = ttk.Frame(target_canvas, padding=(4, 12))
        target_canvas.create_window((0, 0), window=target_inner, anchor='nw')
        target_inner.bind('<Configure>', lambda _event: target_canvas.configure(scrollregion=target_canvas.bbox('all')))
        for index, item in enumerate(self.target_headers):
            card = tk.Frame(target_inner, bg='#f8fafc', bd=1, relief='solid', width=178, height=180)
            card.grid(row=0, column=index, padx=5, sticky='ns')
            card.grid_propagate(False)
            slot_text, slot_bg, slot_fg = ('程序回写\n无需映射', '#e5e7eb', '#6b7280') if item['output_only'] else ('拖入源字段', '#ffffff', '#64748b')
            slot = tk.Label(card, text=slot_text, bg=slot_bg, fg=slot_fg, relief='sunken', bd=1, wraplength=150, width=20, height=4, font=('微软雅黑', 8))
            slot.pack(fill='x', padx=8, pady=(10, 8))
            slot._target_column = item['column']
            title = tk.Label(card, text=item['header'], bg='#f8fafc', fg='#111827', wraplength=155, justify='center', font=('微软雅黑', 9, 'bold'))
            title.pack(fill='both', expand=True, padx=6, pady=(2, 8))
            if not item['output_only']:
                slot.bind('<Button-1>', lambda _event, col=item['column']: self._assign_selected(col))
                slot.bind('<Double-Button-1>', lambda _event, col=item['column']: self._clear_target(col))
            self.target_slots[item['column']] = (slot, item)
        footer = ttk.Frame(outer)
        footer.pack(fill='x', pady=(10, 0))
        self.status_var = tk.StringVar(value=f'需要完成 {len(self.required_targets)} 个字段映射；也可以先单击源字段，再单击目标格。')
        ttk.Label(footer, textvariable=self.status_var).pack(side='left')
        self.clear_button = ttk.Button(footer, text='全部清除', command=self._clear_all, width=10)
        self.clear_button.pack(side='right', padx=(8, 0))
        self.generate_button = ttk.Button(footer, text='生成转换底稿', command=self._generate, width=14)
        self.generate_button.pack(side='right', padx=(8, 0))
        self.cancel_button = ttk.Button(footer, text='取消', command=self._request_close, width=9)
        self.cancel_button.pack(side='right')

    def _scroll_source_fields(self, event):
        if event.delta:
            self.source_canvas.yview_scroll((-3 if event.delta > 0 else 3), 'units')
        return 'break'

    def _request_close(self):
        if not self.generating:
            self.destroy()

    def _set_generating(self, active):
        self.generating = active
        state = 'disabled' if active else 'normal'
        self.generate_button.configure(state=state)
        self.clear_button.configure(state=state)
        self.cancel_button.configure(state=state)
        if active:
            self.status_var.set('正在后台生成收款结算单底稿，请稍候……')
        else:
            self._refresh_all()

    def _generation_finished(self, success):
        if self.winfo_exists():
            self._set_generating(False)
            if success:
                self.destroy()

    def _source_used(self, source_column):
        return source_column in self.mapping.values()

    def _select_source(self, source_column):
        self.selected_source = source_column
        self._refresh_source_labels()
        self.status_var.set(f'已选择：{self.source_by_column[source_column]["display"]}；请点击目标格。')

    def _start_drag(self, event, source_column):
        self.drag_source = source_column
        self.drag_window = tk.Toplevel(self)
        self.drag_window.overrideredirect(True)
        tk.Label(self.drag_window, text=self.source_by_column[source_column]['display'], bg='#bfdbfe', fg='#1e3a8a', bd=1, relief='solid', padx=8, pady=5, font=('微软雅黑', 9)).pack()
        self.drag_window.geometry(f'+{event.x_root + 10}+{event.y_root + 10}')

    def _drag_motion(self, event):
        if self.drag_window:
            self.drag_window.geometry(f'+{event.x_root + 10}+{event.y_root + 10}')

    def _target_from_widget(self, widget):
        current = widget
        while current is not None:
            target = getattr(current, '_target_column', None)
            if target is not None:
                return target
            parent_name = current.winfo_parent()
            if not parent_name:
                break
            try:
                current = current.nametowidget(parent_name)
            except Exception:
                break
        return None

    def _drop_source(self, event):
        source_column = self.drag_source
        if self.drag_window:
            self.drag_window.destroy()
        self.drag_window = self.drag_source = None
        if source_column is None:
            return
        target = self._target_from_widget(self.winfo_containing(event.x_root, event.y_root))
        if target is not None:
            self._assign(target, source_column)

    def _assign_selected(self, target_column):
        if self.selected_source is not None:
            self._assign(target_column, self.selected_source)

    def _assign(self, target_column, source_column):
        slot_info = self.target_slots.get(target_column)
        if slot_info is None or slot_info[1]['output_only']:
            return
        try:
            self.mapping = assign_unique_mapping(self.mapping, target_column, source_column)
        except ValueError as error:
            messagebox.showwarning('字段映射', str(error), parent=self)
            return
        self.selected_source = None
        self._refresh_all()

    def _clear_target(self, target_column):
        self.mapping.pop(target_column, None)
        self._refresh_all()

    def _clear_all(self):
        self.mapping.clear()
        self.selected_source = None
        self._refresh_all()

    def _refresh_source_labels(self):
        for column, label in self.source_labels.items():
            if column == self.selected_source:
                label.configure(bg='#fef3c7', fg='#92400e', cursor='hand2')
            else:
                label.configure(bg='#dbeafe', fg='#1e3a8a', cursor='hand2')

    def _refresh_all(self):
        self._refresh_source_labels()
        for target, (slot, item) in self.target_slots.items():
            if item['output_only']:
                continue
            source = self.mapping.get(target)
            if source is None:
                slot.configure(text='拖入源字段', bg='#ffffff', fg='#64748b')
            else:
                slot.configure(text=self.source_by_column[source]['display'], bg='#dcfce7', fg='#166534')
        self.status_var.set(f'当前已完成 {len(self.mapping)}/{len(self.required_targets)} 个字段字段映射。')

    def _generate(self):
        if self.generating:
            return
        missing = [item for item in self.required_targets if item['column'] not in self.mapping]
        if missing:
            preview = '、'.join(item['header'] for item in missing[:10])
            if len(missing) > 10:
                preview += f'，另有{len(missing) - 10}项'
            messagebox.showwarning('映射尚未完成', f'全部收款底稿输入字段都必须映射。\n\n当前还缺少{len(missing)}项：{preview}', parent=self)
            return
        self._set_generating(True)
        try:
            self.on_generate(dict(self.mapping), dialog_parent=self, completion=self._generation_finished)
        except Exception as error:
            self._set_generating(False)
            messagebox.showerror('转换失败', str(error), parent=self)


def start_conversion_wizard(
    parent,
    selected_file_variable,
    template_path,
    target_sheet_name='收款结算明细',
    target_header_row=2,
    scheme_path=None,
):
    """选原始Excel→选择对照表方案→确认→后台生成宏底稿。"""
    if getattr(parent, '_receipt_conversion_busy', False):
        messagebox.showwarning('正在转换', '请等待当前转换完成', parent=parent)
        return
    try:
        read_receipt_macro(template_path)
    except Exception as error:
        messagebox.showerror('正式模板不可用', str(error), parent=parent)
        return
    source_path = filedialog.askopenfilename(
        parent=parent,
        title='请选择待转换的原始Excel（不是对照表）',
        filetypes=[
            ('Excel 工作簿', '*.xlsx *.xlsm'),
            ('Excel xlsx', '*.xlsx'),
            ('Excel xlsm', '*.xlsm'),
        ],
    )
    if not source_path:
        return
    table_path = scheme_path or os.path.join(
        os.path.dirname(os.path.abspath(template_path)), MAPPING_TABLE_FILENAME
    )
    if not os.path.isfile(table_path):
        table_path = filedialog.askopenfilename(
            parent=parent, title=f'未找到{MAPPING_TABLE_FILENAME}，请选择对照表文件',
            filetypes=[('Excel对照表', '*.xlsx')],
        )
        if not table_path:
            return
    try:
        protected = {os.path.normcase(os.path.realpath(p)) for p in (source_path, table_path)}
        if len(protected) < 2:
            raise ValueError('请选择原始业务数据文件，不能把对照表当作待转换数据')
        if os.path.normcase(os.path.realpath(default_output_path(template_path))) in protected:
            raise ValueError('转换结果不能覆盖原始数据或对照表，请将原始数据另存为不同文件名')
        scheme_name = ask_mapping_sheet(parent, table_path)
        if not scheme_name:
            return
        target_headers = read_target_headers(
            template_path, target_sheet_name, target_header_row
        )
        config = read_mapping_table(table_path, scheme_name, target_headers)
        source_sheet_name, mapping = resolve_table_mapping(source_path, config, target_headers)
        header_row = config['header_row']
    except Exception as error:
        messagebox.showerror('转换方案检查失败', str(error), parent=parent)
        return
    if not messagebox.askokcancel(
        '确认转换',
        f'方案：{scheme_name}\n原始Excel：{source_path}\n'
        f'来源：第{config["sheet_index"]}个sheet“{source_sheet_name}”，表头第{header_row}行\n'
        f'已匹配{len(mapping)}个字段。\n\n'
        f'将生成并覆盖：{default_output_path(template_path)}\n'
        '原始Excel及对照表不修改。请先关闭底稿；现有底稿将被覆盖，不再自动备份。',
        parent=parent,
    ):
        return

    def generate(
        mapping, reused=False, dialog_parent=None, completion=None
    ):
        parent._receipt_conversion_busy = True
        output_path = default_output_path(template_path)
        message_parent = dialog_parent or parent
        progress_dialog = None
        progress_bar = None
        if dialog_parent is None:
            progress_dialog = tk.Toplevel(parent)
            progress_dialog.title('正在生成底稿')
            progress_dialog.transient(parent)
            progress_dialog.resizable(False, False)
            progress_dialog.protocol('WM_DELETE_WINDOW', lambda: None)
            progress_frame = ttk.Frame(progress_dialog, padding=18)
            progress_frame.pack(fill='both', expand=True)
            ttk.Label(
                progress_frame,
                text='正在后台生成收款结算单录入底稿，请稍候……',
            ).pack(pady=(0, 12))
            progress_bar = ttk.Progressbar(
                progress_frame, mode='indeterminate', length=310
            )
            progress_bar.pack(fill='x')
            progress_bar.start(12)
            progress_dialog.grab_set()

        result_queue = queue.Queue(maxsize=1)

        def worker():
            try:
                result = convert_custom_workbook(
                    source_path=source_path,
                    source_sheet_name=source_sheet_name,
                    header_row=header_row,
                    mapping=mapping,
                    template_path=template_path,
                    output_path=output_path,
                    target_sheet_name=target_sheet_name,
                    target_header_row=target_header_row,
                )
                result_queue.put((True, result, None))
            except Exception as error:
                result_queue.put((False, None, error))

        def finish(success, result, error):
            parent._receipt_conversion_busy = False
            if progress_bar is not None:
                progress_bar.stop()
            if progress_dialog is not None and progress_dialog.winfo_exists():
                progress_dialog.destroy()
            if not success:
                messagebox.showerror(
                    ('转换底稿数据验证未通过' if isinstance(error, ConversionValidationError)
                     else '转换失败'), str(error), parent=message_parent
                )
                if completion is not None:
                    completion(False)
                return

            selected_file_variable.set(output_path)
            message = (
                f'转换完成：共写入{result["row_count"]}行数据，'
                f'映射{result["mapping_count"]}个字段。\n\n'
                f'转换底稿：\n{output_path}\n\n'
                '该文件已自动放入主界面的文件选择框。\n'
                f'使用对照表方案：{scheme_name}'
            )
            if reused:
                message = '已自动引用上次保存的字段匹配方案。\n\n' + message
            if result['warnings']:
                preview = '、'.join(result['warnings'][:8])
                if len(result['warnings']) > 8:
                    preview += f'，另有{len(result["warnings"]) - 8}处'
                message += (
                    '\n\n注意：以下标识字段在源表中是普通数值格式，已安全转成文本，'
                    f'建议以后在源表中按文本保存以保留前导零：{preview}'
                )
            messagebox.showinfo('转换完成', message, parent=message_parent)
            if completion is not None:
                completion(True)

        def poll_result():
            try:
                payload = result_queue.get_nowait()
            except queue.Empty:
                parent.after(100, poll_result)
                return
            finish(*payload)

        try:
            threading.Thread(
                target=worker,
                name='receipt-workbook-converter',
                daemon=True,
            ).start()
        except Exception as error:
            finish(False, None, error)
            return False
        parent.after(100, poll_result)
        return True

    generate(mapping)
