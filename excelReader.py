import openpyxl
from Models.Signal import Signal

def read_and_save_signals(excel_settings_array, excel_file_path):
    signals_array = []

    workbook = openpyxl.load_workbook(excel_file_path)

    for setting in excel_settings_array:
        # ПРОХОД ЕЩЕ ПО ОДНОМУ МАССИВУ ЭКСЕЛЯ (10 это костыль )
        sheet = workbook[setting['ExcelSheet']]
        row_count = sheet.max_row  # Предположим, что это количество строк на вашем листе
        print(row_count)
        for i in range(setting['StartCount'], row_count + setting['StartCount']):
            if sheet.cell(row=i, column=1).value is not None:  # Проверяем, что ячейка не пустая
                print(i)
                # Ваш код обработки каждой строки здесь
                print(type(setting['Description'] +"1"))
                user_tree = setting['TreePath'] + str(sheet.cell(row=i, column=sheet[setting['Description']+"10"].col_idx).value)
                opc_tag = setting['Prefix'] + "." + str(sheet.cell(row=i, column=sheet[setting['Tag']+"10"].col_idx).value) + "." + setting['Postfix']
                e_unit = str(sheet.cell(row=i, column=sheet[setting['Unit']+"10"].col_idx).value)
                description = str(sheet.cell(row=i, column=sheet[setting['Description']+"10"].col_idx).value)
                signal = Signal(user_tree, opc_tag, e_unit, description)
                signals_array.append(vars(signal))

    return signals_array

# Пример использования функции для чтения и сохранения состояний из Excel-файла
excel_settings_array = [
    {'Name': 'AI', 'ExcelSheet': 'AI1', 'TreePath': 'AI\\', 'Tag': 'R', 'Unit': 'I', 'Description': 'F', 'Prefix': 'GVL', 'Postfix': 'AI_ToHMI.PV', 'StartCount': 4},
    ]

excel_file_path = 'Input/IO_alpha_CONFIGURATOR_v6.xlsm'
signals = read_and_save_signals(excel_settings_array, excel_file_path)
for signal in signals:
    print(signal)