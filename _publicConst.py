import json


# Памятка на будущее
# 1. Пути к файлам относительные а не абсолютные( от корня проекта)
# 2. Нужно экранировать символы

class Const:
    FILE_PATH = "Input\\"
    OUTPUT_FOLDER = "Output\\"
    CONFIG_NAME = "config.json"
    with open(FILE_PATH + CONFIG_NAME) as json_file:
        data = json.load(json_file)
        EXCEL_NAME = data["nameFileTable"]