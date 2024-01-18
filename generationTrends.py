from _publicConst import Const
from excelReader import read_and_save_signals
from jsonLogics import read_settings, create_json_file


def generate():
    excelSettings = read_settings()
    UserTree = read_and_save_signals(excelSettings, Const.FILE_PATH + Const.EXCEL_NAME)
    create_json_file(UserTree)
