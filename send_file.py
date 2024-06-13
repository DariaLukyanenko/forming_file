# -*- coding: utf-8 -*-
import json
import requests

# Путь к JSON файлу
json_filename = 'ik.json'  # Замените на путь к вашему JSON файлу
excel_filename = 'all_users_data.xlsx'  # Имя файла для сохранения Excel

# URL вашего API
api_url = 'http://192.168.10.193:8080/upload'  # Замените на URL вашего API

# Чтение данных из JSON файла
with open(json_filename, 'r', encoding='utf-8') as file:
    json_data = json.load(file)

# Отправка данных на API
response = requests.post(api_url, files={'file': ('file.json', json.dumps(json_data), 'application/json')})

# Проверка успешности запроса
if response.status_code == 200:
    # Сохранение полученного Excel файла
    with open(excel_filename, 'wb') as f:
        f.write(response.content)
    print(f"Excel файл сохранен как {excel_filename}")
else:
    print(f"Ошибка: {response.status_code}")
    print(response.text)
