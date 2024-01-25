import openpyxl
from Models.Signal import Signal
import os
from _publicConst import Const
from jsonLogics import get_data_from_config


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
# Заполнение OПС-карты (екселя)
def process_excel_file(excel_name_opc, key_name):
    config = get_data_from_config(key_name)
    govno_BINDING = get_data_from_config("binding_OpcMap")
    govno_ADDRESS = get_data_from_config("addressSpace_OpcMap")
    govno_TYPE = get_data_from_config("typeId_OpcMap")
    input_file_path = Const.FILE_PATH + excel_name_opc
    output_file_path = Const.OUTPUT_FOLDER + excel_name_opc
    wb = openpyxl.load_workbook(filename=input_file_path)
    sheet = wb.active
    last_row = sheet.max_row

    for row_number in range(2, last_row + 1):
        a_value = sheet.cell(row=row_number, column=1).value
        sheet.cell(row=row_number, column=4, value=govno_BINDING)
        sheet.cell(row=row_number, column=5, value=govno_ADDRESS)
        sheet.cell(row=row_number, column=6, value=govno_TYPE)
        for pattern, replacement in config.items():
            parts = pattern.split(Const.PATTERN_TAG_OPC)
            if len(parts) == 2 and a_value.startswith(parts[0]) and a_value.endswith(parts[1]):
                tag = a_value.replace(parts[0], "").replace(parts[1], "")
                new_value = replacement.replace(Const.PATTERN_TAG_OPC, tag)
                sheet.cell(row=row_number, column=7, value=new_value)
                break

    if not os.path.exists(Const.OUTPUT_FOLDER):
        os.makedirs(Const.OUTPUT_FOLDER)
    wb.save(output_file_path)

    output_file_absolute_path = os.path.abspath(output_file_path)
    if os.path.exists(output_file_path):
        print(f"Файл успешно создан по пути: {output_file_absolute_path}")
    else:
        print(f"Не удалось создать файл {output_file_absolute_path}")
