from _publicConst import Const
from excelLogics import read_and_save_signals, process_excel_file
from jsonLogics import read_settings, create_json_file, get_data_from_config


def generateTrends():
    try:
        excelSettings = read_settings()
        EXCEL_NAME_IO = get_data_from_config("nameFile_IO")
        UserTree = read_and_save_signals(excelSettings, Const.FILE_PATH + EXCEL_NAME_IO)
        create_json_file(UserTree)
    except Exception as e:
        print("ОШИБКА: генерация трендов провалилась:", e)


def generateOpcMap():
    try:
        EXCEL_NAME_OPC = get_data_from_config("nameFile_OpcMap")
        key_name_json = "shitty_wizard"
        process_excel_file(EXCEL_NAME_OPC, key_name_json)
    except Exception as e:
        print("ОШИБКА: генерация OPC провалилась:", e)
