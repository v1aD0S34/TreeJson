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
def process_excel_file():
    input_file_path = Const.FILE_PATH + Const.EXCEL_NAME_OPC  # 'Input/OpcMap12'
    output_file_path = Const.OUTPUT_FOLDER + Const.EXCEL_NAME_OPC # 'Output/OpcMap.xlsx'

    wb = openpyxl.load_workbook(filename=input_file_path)
    sheet = wb.active

    last_row = sheet.max_row

    for row_number in range(2, last_row + 1):
        sheet.cell(row=row_number, column=4, value=Const.govno_BINDING)
        sheet.cell(row=row_number, column=5, value=Const.govno_ADDRESS)
        sheet.cell(row=row_number, column=6, value=Const.govno_TYPE)

    if not os.path.exists('Output'):
        os.makedirs('Output')

    wb.save(output_file_path)

    output_file_absolute_path = os.path.abspath(output_file_path)

    if os.path.exists(output_file_path):
        print(f"Файл успешно создан по пути: {output_file_absolute_path}")
    else:
        print("Не удалось создать файл")
