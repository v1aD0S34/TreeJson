import json


class Const:
    FILE_PATH = "Input\\"
    CONFIG_NAME = "config.json"
    with open(FILE_PATH + CONFIG_NAME) as json_file:
        data = json.load(json_file)
        EXCEL_NAME = data["nameFileTable"]
