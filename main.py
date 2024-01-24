import warnings

from excelLogics import process_excel_file
from generationTrends import generate


if __name__ == "__main__":
    warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
    try:
        generate()
    except Exception as e:
        print("ОШИБКА: генерация трендов провалилась:", e)

    process_excel_file()
    # try:
    #     process_excel_file()
    # except Exception as e:
    #     print("ОШИБКА: генерация OPC провалилась:", e)