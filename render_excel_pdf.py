import pathlib

import openpyxl

import re


def find_pattern(text):
    pattern = r'<\$.*?\$>'
    matches = re.findall(pattern, text)
    return matches


def replace_patterns(text, replacement, match):
    text = text.replace(match, replacement)
    return text


def load_excel_as_dict(file_path, sheet_name="Input"):
    file_path = pathlib.Path(file_path)
    if not file_path.is_file():
        return None, f"Check File Path Again! {file_path}"
    try:
        workbook = openpyxl.load_workbook(file_path)
        try:
            sheet = workbook[sheet_name]
        except:
            return None, f"The sheet doesn't exist in the workbook. It should be strictly '{sheet_name}'"
        header_row = [(cell.value).strip() for cell in sheet[1]]
        data_dict = {header: [] for header in header_row}

        for row in sheet.iter_rows(min_row=2, values_only=True):
            for header, cell_value in zip(header_row, row):
                data_dict[header].append(cell_value)

        return data_dict, None
    except FileNotFoundError:
        return None, f"File not found: {file_path}"
