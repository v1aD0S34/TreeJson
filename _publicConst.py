import json


# Памятка на будущее
# 1. Пути к файлам относительные а не абсолютные(от корня проекта)
# 2. Нужно экранировать символы

class Const:
    FILE_PATH = "Input\\"
    OUTPUT_FOLDER = "Output\\"
    CONFIG_NAME = "config.json"
    OUTPUT_JSON_NAME = "Tree_JSON"

    with open(FILE_PATH + CONFIG_NAME, encoding='utf-8') as json_file:
        data = json.load(json_file)
        EXCEL_NAME_IO = data["nameFile_IO"]
        EXCEL_NAME_OPC = data["nameFile_OpcMap"]
        govno_BINDING = data["binding_OpcMap"]
        govno_ADDRESS = data["addressSpace_OpcMap"]
        govno_TYPE = data["typeId_OpcMap"]