import json
import os

from Models.DataSettings import ExcelSettings
from _publicConst import Const


# Создаем файл Json  в выходной папке для трендов
def create_json_file(signals):
    data = {"UserTree": [{"Signal": vars(signal)} for signal in signals]}
    output_folder = os.path.join(os.getcwd(), Const.OUTPUT_FOLDER)
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    file_path = os.path.join(output_folder, Const.OUTPUT_JSON_NAME + '.json')
    with open(file_path, 'w', encoding='utf-8') as json_file:
        json.dump(data, json_file, ensure_ascii=False, indent=2)
    if os.path.exists(file_path):
        print(f"Файл успешно создан по пути: {file_path}")
    else:
        print("Возникла проблема при создании файла")


# Создать массив с настройками по каждому листу из Config.json для трендов
def read_settings():
    excel_settings_array = []
    with open(Const.FILE_PATH + Const.CONFIG_NAME, encoding='utf-8') as json_file:
        data = json.load(json_file)
        excel_settings_list = data["excel_settings_IO"]
        for setting in excel_settings_list:
            excel_setting = ExcelSettings(
                setting["Name"],
                setting["ExcelSheet"],
                setting["TreePath"],
                setting["Tag"],
                setting["Unit"],
                setting["Description"],
                data["prefix_Alpha"],
                setting["Postfix_Alpha"],
                setting["StartCount"]
            )
            excel_settings_array.append(excel_setting)

    return excel_settings_array


# Вернуть значения из config.json по названию ключа
def get_data_from_config(data):
    try:
        with open(Const.FILE_PATH + Const.CONFIG_NAME, 'r' , encoding='utf-8') as config_file:
            config_content = config_file.read()
            config_dict = json.loads(config_content)
        return config_dict[data]
    except FileNotFoundError:
        print(f"ОШИБКА: Файл {Const.CONFIG_NAME} не найден в директории {Const.FILE_PATH} !")
        return None
