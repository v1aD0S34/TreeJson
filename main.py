from excelLogics import process_excel_file
from generationTrends import generate


if __name__ == "__main__":
    try:
        generate()
    except Exception as e:
        print("ОШИБКА: генерация трендов провалилась:", e)

    try:
        process_excel_file()
    except Exception as e:
        print("ОШИБКА: генерация OPC провалилась:", e)