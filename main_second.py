import os
import pandas as pd
import requests
import json
import time

# Убедитесь, что библиотека openpyxl установлена
try:
    import openpyxl
except ImportError:
    raise ImportError("Необходима библиотека 'openpyxl'. Установите её с помощью 'pip install openpyxl'.")

# Путь к входному файлу Excel с ИИН
input_file = 'input_iin.xlsx'
# Путь к выходному файлу Excel для результатов
output_file = 'output_results.xlsx'

# Проверка существования файла
if not os.path.exists(input_file):
    print(f"Файл не найден: {input_file}")
    exit()

# Чтение входного файла
df = pd.read_excel(input_file, header=None)

# Поиск индекса строки, с которой начинаются две пустые строки подряд
def find_end_index(df):
    empty_row_count = 0
    for index, value in enumerate(df.iloc[:, 0]):
        if pd.isna(value):
            empty_row_count += 1
            if empty_row_count == 2:
                return index - 1
        else:
            empty_row_count = 0
    return len(df)

end_index = find_end_index(df)
# Убираем .0 и приводим к строковому типу
iin_list = df.iloc[:end_index, 0].dropna().astype(str).str.replace(r'\.0$', '', regex=True).tolist()

# URL и заголовки для запроса
url = "https://aisoip.adilet.gov.kz/rest/debtor/findErd?page=0&size=10"
headers = {
    "Accept": "application/json, text/plain, */*",
    "Accept-Encoding": "gzip, deflate, br, zstd",
    "Accept-Language": "ru-RU,ru;q=0.9",
    "Connection": "keep-alive",
    "Content-Length": "801",
    "Content-Type": "application/json",
    "Cookie": "_ym_uid=1717903665276182295; _ym_d=1717903665; _ym_isad=2",
    "Host": "aisoip.adilet.gov.kz",
    "Origin": "https://aisoip.adilet.gov.kz",
    "Referer": "https://aisoip.adilet.gov.kz/debtors",
    "Sec-Ch-Ua": '"Google Chrome";v="125", "Chromium";v="125", "Not.A/Brand";v="24"',
    "Sec-Ch-Ua-Mobile": "?0",
    "Sec-Ch-Ua-Platform": '"macOS"',
    "Sec-Fetch-Dest": "empty",
    "Sec-Fetch-Mode": "cors",
    "Sec-Fetch-Site": "same-origin",
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36"
}

# Функция для отправки запроса и получения результата
def get_debtor_info(iin):
    data = {
        "action": "findErd",
        "bin": "",
        "captcha": "03AFcWeA4HeDXz5pN_HLQms--Gh3wM-uxVunT5ele68mQuk7sbBacTfgjLVoG6vekpyqtTymmAjxTBc4bEzxw7A1jWzuy2BlJaoA732eyqFw70K7Oa0LMqpeZ4n3PuNiB1P037Udss_dpgHHinPZRpH_A1r8leoVtyCzof6m7FC2iwYKNQWh7kC_ItE9o6lSF-DyYLQhwBK64xhehBMqGImNxu_NKApyfMOEvRJXXZF0Ey4QfeEYaol3mKS0qsloH5AMuP0eVXCM1xg2PtAtiYbozx_bVEfWe79I3kPB_NpzDabpAqAVyw7GKTtM4MGNfvQojZ3GaVev1Px9uErI-Xzi_9o8V5-Wam1EUITd_011nDJpe9ezSFef0Ig1pAeRPmCKv-D5fI_QYujkCFgwlbpKxWd71q578vAC1hZHj_6DVfXjNKmY5wqE0u2Q_7X8oCr2kIXqCpt1g0DSfUnYC7YWbbpznVoM3tNZ5czoOehPQ00ZubwayqqJWXjDOHRw0AVVydwxCVn9uGUuiOgL4-XXBa_xYeL_JeH7xbBjOgou_-qsYJ0ikggniN0zYBILTMt-Ue7YaxpBSjUtRYktoMvCUOf-H-cVcFGwsEpMAm1Sdz8bfSYzrRS9vMBbsQRScHjUrWgvaC7vvbwe7XkzbND4JjTt88Jr5fvT7ki6SxNoL__sp-c5KrM1I",
        "docNum": "",
        "fullName": "",
        "iin": iin,
        "searchType": 0
    }

    response = requests.post(url, headers=headers, json=data)

    if response.status_code == 200:
        return response.json()
    else:
        return {"error": f"Запрос не удался с кодом состояния: {response.status_code}"}

# Список для хранения результатов
results = []

# Обработка каждого ИИН с паузой в 5 секунд
for iin in iin_list:
    result = get_debtor_info(iin)
    if "content" in result and result["content"]:
        content = result["content"][0]
        debtor_info = {
            "iin": iin,
            "debtorFullName": content.get("debtorFullName", ""),
            "banStartDate": content.get("banStartDate", ""),
            "full_response": json.dumps(result, ensure_ascii=False)
        }
    else:
        debtor_info = {
            "iin": iin,
            "debtorFullName": "Долгов и ограничений нет",
            "banStartDate": "",
            "full_response": json.dumps(result, ensure_ascii=False)
        }
    results.append(debtor_info)
    print(f"Обработан ИИН: {iin}, ожидание 5 секунд перед следующим запросом...")
    time.sleep(5)  # Пауза в 5 секунд

# Создание DataFrame с результатами
results_df = pd.DataFrame(results)

# Запись результатов в выходной файл
results_df.to_excel(output_file, index=False)

print(f"Результаты успешно сохранены в файл {output_file}")
