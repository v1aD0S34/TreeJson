# Чтение данных из файла config.json и создание массива объектов ExcelSettings
import json

from Models.DataSettings import ExcelSettings


def ReadSettingsExcel():
    excel_settings_array = []
    with open('config.json') as json_file:
        data = json.load(json_file)
        excel_settings_list = data["excel_settings"]
        for setting in excel_settings_list:
            excel_setting = ExcelSettings(
                setting["Name"],
                setting["ExcelSheet"],
                setting["TreePath"],
                setting["Tag"],
                setting["Unit"],
                setting["Description"],
                data["prefix"],
                setting['Postfix'],
                setting["StartCount"]
            )
            excel_settings_array.append(excel_setting)

    return excel_settings_array




settings = ReadSettingsExcel()
for setting in settings:
    print(vars(setting))

print(len(settings))
