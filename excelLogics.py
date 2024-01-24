import json

import openpyxl
from Models.Signal import Signal
import os

from _publicConst import Const


# Читаем из Экселя и создаем список обьектов Signal
def read_and_save_signals(excel_settings_array, excel_file_path):
    signals_array = []
    workbook = openpyxl.load_workbook(excel_file_path)
    for setting in excel_settings_array:
        sheet = workbook[setting.ExcelSheet]
        row_count = sheet.max_row
        for i in range(setting.StartCount, row_count + setting.StartCount):
            if sheet.cell(row=i, column=1).value is not None:
                user_tree = (setting.TreePath + str(
                    sheet.cell(row=i, column=sheet[setting.Description + "10"].col_idx).value)).replace('/',
                                                                                                        '|').replace(
                    '\\', '/')
                opc_tag = setting.Prefix + "." + setting.Name + "." + str(
                    sheet.cell(row=i, column=sheet[setting.Tag + "10"].col_idx).value) + "." + setting.Postfix
                e_unit = str(sheet.cell(row=i, column=sheet[setting.Unit + "10"].col_idx).value).replace(
                    '\\', '/')
                description = str(sheet.cell(row=i, column=sheet[setting.Description + "10"].col_idx).value)
                _signal = Signal(user_tree, opc_tag, e_unit, description)
                signals_array.append(_signal)

    return signals_array


# Когда будет понимание про Идентификатор узла переписать
# def process_excel_file():
#     input_file_path = Const.FILE_PATH + Const.EXCEL_NAME_OPC  # 'Input/OpcMap12'
#     output_file_path = Const.OUTPUT_FOLDER + Const.EXCEL_NAME_OPC # 'Output/OpcMap.xlsx'
#
#     wb = openpyxl.load_workbook(filename=input_file_path)
#     sheet = wb.active
#
#     last_row = sheet.max_row
#
#     for row_number in range(2, last_row + 1):
#         sheet.cell(row=row_number, column=4, value=Const.govno_BINDING)
#         sheet.cell(row=row_number, column=5, value=Const.govno_ADDRESS)
#         sheet.cell(row=row_number, column=6, value=Const.govno_TYPE)
#
#     if not os.path.exists('Output'):
#         os.makedirs('Output')
#
#     wb.save(output_file_path)
#
#     output_file_absolute_path = os.path.abspath(output_file_path)
#
#     if os.path.exists(output_file_path):
#         print(f"Файл успешно создан по пути: {output_file_absolute_path}")
#     else:
#         print("Не удалось создать файл")
# def process_excel_file():
#     config = load_config()
#     input_file_path = Const.FILE_PATH + Const.EXCEL_NAME_OPC  # 'Input/OpcMap12'
#     output_file_path = Const.OUTPUT_FOLDER + Const.EXCEL_NAME_OPC # 'Output/OpcMap.xlsx'
#     wb = openpyxl.load_workbook(filename=input_file_path)
#     sheet = wb.active
#     last_row = sheet.max_row
#
#     for row_number in range(2, last_row + 1):
#         a_value = sheet.cell(row=row_number, column=1).value
#         print(f"a_value|         {a_value}")
#         for pattern, replacement in config['shitty_wizard'].items():
#             print("____________________________________________")
#             print(f"pattern|         {pattern}")
#             print(f"replacement|         {replacement}")
#             print("____________________________________________")
#             if a_value.startswith(pattern.split("_TAG.")[0]) and a_value.endswith(pattern.split("_TAG.")[1]):
#                 tag = a_value.split(".")[a_value.split(".").index(pattern.split("_TAG.")[0].split(".")[-1]) + 1]
#                 replacement = replacement.replace("_TAG", tag)
#                 sheet.cell(row=row_number, column=7, value=replacement)
#
#     if not os.path.exists('Output'):
#         os.makedirs('Output')
#     wb.save(output_file_path)
#     output_file_absolute_path = os.path.abspath(output_file_path)
#     if os.path.exists(output_file_path):
#         print(f"Файл успешно создан по пути: {output_file_absolute_path}")
#     else:
#         print("Не удалось создать файл")


def load_config():
    with open('Input/config.json', 'r') as config_file:  # Убедитесь, что путь до файла config.js верный
        config_content = config_file.read()
        config_dict = json.loads(config_content)
    return config_dict['shitty_wizard']


def process_excel_file():
    config = load_config()
    input_file_path = Const.FILE_PATH + Const.EXCEL_NAME_OPC  # 'Input/OpcMap12'
    output_file_path = Const.OUTPUT_FOLDER + Const.EXCEL_NAME_OPC  # 'Output/OpcMap.xlsx'

    wb = openpyxl.load_workbook(filename=input_file_path)
    sheet = wb.active
    last_row = sheet.max_row

    for row_number in range(2, last_row + 1):
        a_value = sheet.cell(row=row_number, column=1).value
        sheet.cell(row=row_number, column=4, value=Const.govno_BINDING)
        sheet.cell(row=row_number, column=5, value=Const.govno_ADDRESS)
        sheet.cell(row=row_number, column=6, value=Const.govno_TYPE)
        for pattern, replacement in config.items():
            parts = pattern.split("_TAG")
            if len(parts) == 2 and a_value.startswith(parts[0]) and a_value.endswith(parts[1]):
                tag = a_value.replace(parts[0], "").replace(parts[1], "")
                new_value = replacement.replace("_TAG", tag)

                sheet.cell(row=row_number, column=7, value=new_value)
                break  # Если найдено совпадения шаблона, прерываем цикл

    if not os.path.exists('Output'):
        os.makedirs('Output')
    wb.save(output_file_path)

    output_file_absolute_path = os.path.abspath(output_file_path)
    if os.path.exists(output_file_path):
        print(f"Файл успешно создан по пути: {output_file_absolute_path}")
    else:
        print("Не удалось создать файл")