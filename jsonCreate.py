import json
import os


# def create_json_file(signals):
#     data = {"UserTree": [{"Signal": vars(signal)} for signal in signals]}
#     with open('tree.json', 'w', encoding='utf-8') as json_file:
#         json.dump(data, json_file, ensure_ascii=False, indent=2)
def create_json_file(signals):
    data = {"UserTree": [{"Signal": vars(signal)} for signal in signals]}
    output_folder = os.path.join(os.getcwd(), 'Output')
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    file_path = os.path.join(output_folder, 'Tree_JSON.json')
    with open(file_path, 'w', encoding='utf-8') as json_file:
        json.dump(data, json_file, ensure_ascii=False, indent=2)
    if os.path.exists(file_path):
        print(f"Файл успешно создан по пути: {file_path}")
    else:
        print("Возникла проблема при создании файла")