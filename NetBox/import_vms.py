import pandas as pd
import requests
import json
import os

# --- Конфигурация ---
NETBOX_URL = "https://netbox.axioma-ipc.ru"
# API токен Netbox
NETBOX_TOKEN = "878fce3ad774ca50e0023f5d780c880a01488070"
# Путь к Excel файлу
EXCEL_FILE_PATH = "kln_address.xlsx"
# Название листа с данными
SHEET_NAME = "Tech"

# --- Заголовки запросов API ---
HEADERS = {
    "Authorization": f"Token {NETBOX_TOKEN}",
    "Content-Type": "application/json",
    "Accept": "application/json",
}

# --- Функция для создания виртуальной машины в Netbox ---
def create_virtual_machine(vm_data):
    """
    Отправляет запрос на создание виртуальной машины в Netbox.
    :param vm_data: Словарь с данными виртуальной машины.
    :return: True, если успешно, False в противном случае.
    """
    url = f"{NETBOX_URL}/api/virtualization/virtual-machines/"

# debug:

    print(f"--- Отправка запроса ---")
    print(f"URL: {url}")
    print(f"Метод: POST") # проверкаPOST
    print(f"Заголовки: {HEADERS}")
    print(f"Тело запроса (JSON): {json.dumps(vm_data, indent=2)}") # Печатаем красиво форматированный JSON
    print(f"--- Начало выполнения запроса ---")
    
    try:
        response = requests.post(url, headers=HEADERS, data=json.dumps(vm_data))
        response.raise_for_status()  # Вызовет исключение для неудовлетворительных ответов (4xx или 5xx)

        print(f"Успешно создана VM: {vm_data.get('name')}")
        return True
    
    except requests.exceptions.HTTPError as errh:
        print(f"Ошибка HTTP: {errh}")
        try:
            error_details = response.json()
            print("Детали ошибки:")
            print(json.dumps(error_details, indent=2, ensure_ascii=False))
        except json.JSONDecodeError:
            print("Не удалось декодировать ответ как JSON.")
            print(f"Тело ответа: {response.text}")
    except requests.exceptions.ConnectionError as errc:
        print(f"Ошибка подключения: {errc}")
    except requests.exceptions.Timeout as errt:
        print(f"Ошибка таймаута: {errt}")
    except requests.exceptions.RequestException as err:
        print(f"Неожиданная ошибка запроса: {err}")
    except requests.exceptions.RequestException as e:
        print(f"Ошибка при создании VM '{vm_data.get('name')}': {e}")
        if response:
            try:
                error_details = response.json()
                print(f"Детали ошибки API: {error_details}")
            except json.JSONDecodeError:
                print(f"Не удалось декодировать ответ API: {response.text}")
        return False

# --- Основная логика скрипта ---
def import_vms_from_excel(excel_path, sheet_name):
    """
    Читает данные из Excel и импортирует виртуальные машины в Netbox.
    """
    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name)
        print(f"Успешно прочитан файл: {excel_path}, лист: {sheet_name}")
    except FileNotFoundError:
        print(f"Ошибка: Файл '{excel_path}' не найден.")
        return
    except ValueError as ve:
        print(f"Ошибка: Лист '{sheet_name}' не найден в файле '{excel_path}'.")
        print(f"Доступные листы: {pd.ExcelFile(excel_path).sheet_names}")
        return
    except Exception as e:
        print(f"Неожиданная ошибка при чтении Excel файла: {e}")
        return

    # Проверка наличия необходимых столбцов
    required_columns = ['name', 'role', 'description', 'serial']
    if not all(col in df.columns for col in required_columns):
        missing_cols = [col for col in required_columns if col not in df.columns]
        print(f"Ошибка: Отсутствуют обязательные столбцы в Excel файле: {', '.join(missing_cols)}")
        print(f"Найденные столбцы: {df.columns.tolist()}")
        return

    for index, row in df.iterrows():
        vm_name = row['name']
        vm_role_name = row['role']
        vm_description = row['description']
        vm_serial = row['serial']
        vm_status = row['status']
        vm_ip_primary = row['ip_primary']
        vm_vcpus = row['vcpus']
        vm_memory = row['memory']
        vm_disk = row['disk']

        # Проверка на пустые значения в обязательных полях
        if pd.isna(vm_name) or pd.isna(vm_role_name) or pd.isna(vm_status):
            print(f"Пропуск строки {index + 2}: Отсутствует 'name' или 'role' или 'status'.")
            continue

        # Формирование данных для API
        # Важно: 'role' в API Netbox ожидает ID роли, но мы можем указать ее имя
        # Netbox найдет ID по имени, если роль существует.
        vm_payload = {
            "name": str(vm_name),
            #"site": {"id": 2},
            "cluster": {"id": 4},
            "role": {"name": str(vm_role_name)},
            "description": str(vm_description) if not pd.isna(vm_description) else "",
            "serial": str(vm_serial) if not pd.isna(vm_serial) else "",
            #"platform": {"id": 1},
            # "primary_ip4": {
            #     "address": str(vm_ip_primary)  if not pd.isna(vm_ip_primary) else "",
            #     "description": ""
            # },
            "vcpus": vm_vcpus if not pd.isna(vm_vcpus) else "",
            "memory": vm_memory if not pd.isna(vm_memory) else "",
            "disk": vm_disk if not pd.isna(vm_disk) else "",            
            #"status": {"label": str(vm_status) if not pd.isna(vm_status) else ""},
            
            # "tenant": {"name": "YourTenantName"},
            # "config_template": {"name": "YourConfigTemplateName"},
            # ...
        }

        print(f"Попытка импорта VM: {vm_name}...")
        create_virtual_machine(vm_payload)

# --- Запуск импорта ---
if __name__ == "__main__":
    # Проверка наличия файла и токена
    if NETBOX_URL == "http://netbox-instance.com":
        print("Ошибка: Пожалуйста, обновите NETBOX_URL в скрипте.")
    elif NETBOX_TOKEN == "api_token":
        print("Ошибка: Пожалуйста, обновите NETBOX_TOKEN в скрипте.")
    elif not os.path.exists(EXCEL_FILE_PATH):
        print(f"Ошибка: Excel файл '{EXCEL_FILE_PATH}' не найден.")
    else:
        import_vms_from_excel(EXCEL_FILE_PATH, SHEET_NAME)
        print("\nИмпорт завершен.")