import pandas as pd
import requests
import os

# --- Конфигурация. Config ---
NETBOX_URL = "https://netbox.axioma-ipc.ru"
NETBOX_TOKEN = "878fce3ad774ca50e0023f5d780c880a01488070"  # API токен Netbox (admin - API Token and create Permission)
EXCEL_FILE_PATH = "kln_address.xlsx"  # Excel with VM attributes
SHEET_NAME = "Prod"  # Название листа в Excel файле

# --- Настройки API Netbox ---
NETBOX_API_URL = f"{NETBOX_URL}/api"
HEADERS = {
    "Authorization": f"Token {NETBOX_TOKEN}",
    "Content-Type": "application/json",
    "Accept": "application/json",
}

# --- Функции для работы с Netbox API ---

def netbox_api_request(method, url, **kwargs):
    """Выполняет запрос к API Netbox."""
    try:
        response = requests.request(method, url, headers=HEADERS, **kwargs)
        response.raise_for_status()  # Вызов исключения для кодов ошибок (4xx или 5xx)
        return response.json()
    except requests.exceptions.RequestException as e:
        print(f"Ошибка при запросе к Netbox API: {e}")
        if hasattr(e, 'response') and e.response is not None:
            try:
                error_details = e.response.json()
                print(f"Детали ошибки: {error_details}")
            except:
                print(f"Тело ответа: {e.response.text}")
        return None

def get_virtual_machine_by_name(name):
    """Получение VM по имени."""
    url = f"{NETBOX_API_URL}/virtualization/virtual-machines/"
    params = {"name": name}
    try:
        response = netbox_api_request("GET", url, params=params)
        if response and response.get("count", 0) > 0:
            return response["results"][0]
        return None
    except Exception as e:
        print(f"Ошибка при получении VM '{name}': {e}")
        return None

def update_virtual_machine(vm_id, payload):
    """Обновление существующей VM."""
    url = f"{NETBOX_API_URL}/virtualization/virtual-machines/{vm_id}/"
    print(f"Обновление VM (ID: {vm_id})...")
    return netbox_api_request("PATCH", url, json=payload)

def create_virtual_machine(payload):
    """Создание новой VM."""
    url = f"{NETBOX_API_URL}/virtualization/virtual-machines/"
    print(f"Создание VM '{payload.get('name')}'...")
    return netbox_api_request("POST", url, json=payload)

def get_ip_by_address(address):
    """Получение IP-адреса по его значению."""
    url = f"{NETBOX_API_URL}/ipam/ip-addresses/"
    params = {"address": address}
    try:
        response = netbox_api_request("GET", url, params=params)
        if response and response.get("count", 0) > 0:
            return response["results"][0]
        return None
    except Exception as e:
        print(f"Ошибка при получении IP-адреса '{address}': {e}")
        return None

def create_ip_address(payload):
    """Создание нового IP-адреса."""
    url = f"{NETBOX_API_URL}/ipam/ip-addresses/"
    print(f"Создание IP-адреса '{payload.get('address')}'...")
    return netbox_api_request("POST", url, json=payload)

def get_subnet_by_network_and_prefix(network, prefix_length):
    """Получает подсеть по сети и длине префикса."""
    url = f"{NETBOX_API_URL}/ipam/subnets/"
    params = {"contains": network, "prefix_length": prefix_length}
    try:
        response = netbox_api_request("GET", url, params=params)
        if response and response.get("count", 0) > 0:
            return response["results"][0]
        return None
    except Exception as e:
        print(f"Ошибка при получении подсети '{network}/{prefix_length}': {e}")
        return None

def create_subnet(payload):
    """Создание новой подсети."""
    url = f"{NETBOX_API_URL}/ipam/subnets/"
    print(f"Создание подсети '{payload.get('prefix')}'...")
    return netbox_api_request("POST", url, json=payload)

def get_vrf_by_name(name):
    """Получает VRF по имени."""
    url = f"{NETBOX_API_URL}/ipam/vrfs/"
    params = {"name": name}
    try:
        response = netbox_api_request("GET", url, params=params)
        if response and response.get("count", 0) > 0:
            return response["results"][0]
        return None
    except Exception as e:
        print(f"Ошибка при получении VRF '{name}': {e}")
        return None

def get_vm_interfaces(vm_id):
    """Получение интерфейсов VM."""
    url = f"{NETBOX_API_URL}/virtualization/interfaces/"
    params = {"virtual_machine_id": vm_id}
    try:
        response = netbox_api_request("GET", url, params=params)
        if response:
            return response["results"]
        return []
    except Exception as e:
        print(f"Ошибка при получении интерфейсов VM ID {vm_id}: {e}")
        return []

def create_vm_interface(vm_id, interface_name="eth0"):
    """Создание интерфейса для VM."""
    url = f"{NETBOX_API_URL}/virtualization/interfaces/"
    payload = {
        "virtual_machine": {"id": vm_id},
        "name": interface_name,
        "type": {"value": "virtual"},
    }
    return netbox_api_request("POST", url, json=payload)

def assign_ip_to_interface(interface_id, ip_address_id):
    """Назначение IP адреса интерфейсу."""
    url = f"{NETBOX_API_URL}/ipam/ip-addresses/{ip_address_id}/"
    payload = {
        "assigned_object_type": "virtualization.vminterface",
        "assigned_object_id": interface_id,
    }
    return netbox_api_request("PATCH", url, json=payload)

def get_device_role_by_name(name):
    """Получение роли устройства по имени."""
    url = f"{NETBOX_API_URL}/dcim/device-roles/"
    params = {"name": name}
    try:
        response = netbox_api_request("GET", url, params=params)
        if response and response.get("count", 0) > 0:
            return response["results"][0]
        return None
    except Exception as e:
        print(f"Ошибка при получении роли '{name}': {e}")
        return None

def create_device_role(name, slug=None, color="9e9e9e"):
    """Создание роли устройства."""
    if not slug:
        slug = name.lower().replace(" ", "-")
    
    url = f"{NETBOX_API_URL}/dcim/device-roles/"
    payload = {
        "name": name,
        "slug": slug,
        "color": color,
        "description": f"Автоматически созданная роль {name}",
    }
    print(f"Создание роли устройства '{name}'...")
    return netbox_api_request("POST", url, json=payload)

def get_cluster_by_name(name):
    """Получение кластера по имени."""
    url = f"{NETBOX_API_URL}/virtualization/clusters/"
    params = {"name": name}
    try:
        response = netbox_api_request("GET", url, params=params)
        if response and response.get("count", 0) > 0:
            return response["results"][0]
        return None
    except Exception as e:
        print(f"Ошибка при получении кластера '{name}': {e}")
        return None

def get_cluster_type_by_name(name="VMware"):
    """Получение типа кластера по имени."""
    url = f"{NETBOX_API_URL}/virtualization/cluster-types/"
    params = {"name": name}
    try:
        response = netbox_api_request("GET", url, params=params)
        if response and response.get("count", 0) > 0:
            return response["results"][0]
        return None
    except Exception as e:
        print(f"Ошибка при получении типа кластера '{name}': {e}")
        return None

def create_cluster(name, site_id=2):
    """Создание кластера."""
    # Сначала получаем или создаем тип кластера
    cluster_type = get_cluster_type_by_name("VMware")
    if not cluster_type:
        # Если нет типа кластера, используем ID = 1 (обычно VMware)
        cluster_type = {"id": 1}
    
    url = f"{NETBOX_API_URL}/virtualization/clusters/"
    payload = {
        "name": name,
        "type": {"id": cluster_type["id"]},
        "site": {"id": site_id},
        "description": f"Автоматически созданный кластер {name}",
    }
    print(f"Создание кластера '{name}'...")
    return netbox_api_request("POST", url, json=payload)

def create_ip_address(address, description=""):
    """Создание IP адреса."""
    # Определяем подсеть для IP адреса
    if '/' in address:
        ip_with_mask = address
    else:
        # Если маска не указана, добавляем /24 по умолчанию
        ip_with_mask = f"{address}/24"
    
    url = f"{NETBOX_API_URL}/ipam/ip-addresses/"
    payload = {
        "address": ip_with_mask,
        "description": description,
    }
    print(f"Создание IP адреса '{ip_with_mask}'...")
    return netbox_api_request("POST", url, json=payload)

# --- Основная функция импорта ---

def import_vms_from_excel(excel_path, sheet_name):
    """Импорт или обновление VM из Excel файла."""
    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name)
    except FileNotFoundError:
        print(f"Ошибка: Excel файл '{excel_path}' не найден.")
        return
    except Exception as e:
        print(f"Ошибка при чтении Excel файла: {e}")
        return

    # Статистика
    total_records = len(df)
    processed_count = 0
    skipped_count = 0
    skipped_records = []

    for index, row in df.iterrows():
        vm_name = row["name"]
        vm_role_name = row["role"]
        vm_description = row.get("description")
        vm_serial = row.get("serial")
        vm_platform_id = row.get("platform_id", 1)  # По умолчанию platform_id = 1
        vm_site_id = row.get("site_id", 2)        # По умолчанию site_id = 2
        vm_cluster_name = row.get("cluster")        # Получаем имя кластера из файла
        vm_vcpus = row.get("vcpus")
        vm_memory = row.get("memory")
        vm_disk = row.get("disk")
        vm_primary_ip4 = row.get("ip_primary")
        vm_primary_ip4_description = row.get("ip_primary_description", "")
        vm_status = row.get("status")
        vm_tenant_name = row.get("tenant_name")
        vm_vrf_name = row.get("vrf_name") # Новое поле для VRF

        # Проверка обязательных полей
        if pd.isna(vm_name) or pd.isna(vm_role_name):
            skipped_count += 1
            skipped_records.append({
                'row': index + 2,
                'name': vm_name,
                'role': vm_role_name,
                'reason': 'Отсутствуют обязательные поля name или role'
            })
            continue

        existing_vm = get_virtual_machine_by_name(vm_name)

        # Проверяем и создаем роль, если она не существует
        role_obj = get_device_role_by_name(vm_role_name)
        if not role_obj:
            print(f"Роль '{vm_role_name}' не найдена, создаем...")
            role_obj = create_device_role(vm_role_name)
            if not role_obj:
                print(f"Ошибка создания роли '{vm_role_name}', пропускаем VM '{vm_name}'")
                continue

        # Проверяем и создаем кластер, если он не существует
        cluster_obj = get_cluster_by_name(vm_cluster_name) if not pd.isna(vm_cluster_name) else None
        if not cluster_obj and not pd.isna(vm_cluster_name):
            print(f"Кластер '{vm_cluster_name}' не найден, создаем...")
            cluster_obj = create_cluster(vm_cluster_name, vm_site_id)
            if not cluster_obj:
                print(f"Ошибка создания кластера '{vm_cluster_name}', пропускаем VM '{vm_name}'")
                continue
        
        # Формирование базового payload
        payload = {
            "name": str(vm_name),
            "role": {"id": role_obj["id"]},  # Используем ID роли для избежания дублирования
            "description": str(vm_description) if not pd.isna(vm_description) else "",
            "serial": str(vm_serial) if not pd.isna(vm_serial) else "",
            "vcpus": int(vm_vcpus) if pd.notna(vm_vcpus) else None,
            "memory": int(vm_memory) if pd.notna(vm_memory) else None,
            "disk": int(vm_disk) if pd.notna(vm_disk) else None,
        }
        
        # Добавляем кластер, если он найден или создан
        if cluster_obj:
            payload["cluster"] = {"id": cluster_obj["id"]}
        elif not pd.isna(vm_cluster_name):
            payload["cluster"] = {"name": str(vm_cluster_name)}

        # Пропускаем status, так как он может быть неправильным
        # if vm_status:
        #     payload["status"] = str(vm_status)

        if vm_tenant_name:
            payload["tenant"] = {"name": str(vm_tenant_name)}
        
        if vm_vrf_name: # Добавляем VRF, если указан
            vrf = get_vrf_by_name(vm_vrf_name)
            if vrf:
                payload["vrf"] = {"id": vrf["id"]}
            else:
                print(f"Строка {index + 2}: VRF '{vm_vrf_name}' не найден в Netbox. VM не будет привязана к VRF.")

        # Обработка primary_ip4
        if not pd.isna(vm_primary_ip4) and vm_primary_ip4:
            ip_address_str = str(vm_primary_ip4)
            # Оставляем маску подсети как есть, если она указана в файле
            existing_ip = get_ip_by_address(ip_address_str)

            if existing_ip:
                ip_id = existing_ip["id"]
                
                # Проверяем, назначен ли уже IP адрес
                if existing_ip.get("assigned_object_id") and existing_ip.get("assigned_object_type") == "virtualization.vminterface":
                    # IP уже назначен, проверяем, является ли он primary для VM
                    if existing_vm and existing_vm.get("primary_ip4") and existing_vm["primary_ip4"]["id"] == ip_id:
                        print(f"IP '{ip_address_str}' уже является primary_ip4 для VM '{vm_name}'")
                    else:
                        # Устанавливаем как primary_ip4
                        payload["primary_ip4"] = {"id": ip_id}
                else:
                    # IP не назначен, нужно создать интерфейс и назначить IP
                    if existing_vm:
                        vm_id = existing_vm["id"]
                        
                        # Получаем интерфейсы VM
                        interfaces = get_vm_interfaces(vm_id)
                        
                        # Если нет интерфейсов, создаем
                        if not interfaces:
                            interface = create_vm_interface(vm_id, "eth0")
                            if interface:
                                interface_id = interface["id"]
                            else:
                                print(f"Ошибка создания интерфейса для VM '{vm_name}'")
                                continue
                        else:
                            interface_id = interfaces[0]["id"]
                        
                        # Назначаем IP адрес интерфейсу
                        if assign_ip_to_interface(interface_id, ip_id):
                            payload["primary_ip4"] = {"id": ip_id}
                        else:
                            print(f"Ошибка назначения IP адреса интерфейсу для VM '{vm_name}'")
            else:
                # IP адрес не найден, создаем его
                print(f"IP адрес '{ip_address_str}' не найден в Netbox, создаем...")
                new_ip = create_ip_address(ip_address_str, vm_primary_ip4_description)
                if new_ip:
                    ip_id = new_ip["id"]
                    
                    # Если VM уже существует, создаем интерфейс и назначаем IP
                    if existing_vm:
                        vm_id = existing_vm["id"]
                        
                        # Получаем интерфейсы VM
                        interfaces = get_vm_interfaces(vm_id)
                        
                        # Если нет интерфейсов, создаем
                        if not interfaces:
                            interface = create_vm_interface(vm_id, "eth0")
                            if interface:
                                interface_id = interface["id"]
                            else:
                                print(f"Ошибка создания интерфейса для VM '{vm_name}'")
                                continue
                        else:
                            interface_id = interfaces[0]["id"]
                        
                        # Назначаем IP адрес интерфейсу
                        if assign_ip_to_interface(interface_id, ip_id):
                            payload["primary_ip4"] = {"id": ip_id}
                        else:
                            print(f"Ошибка назначения IP адреса интерфейсу для VM '{vm_name}'")
                    else:
                        # VM еще не создана, просто добавляем IP в payload
                        payload["primary_ip4"] = {"id": ip_id}
                else:
                    print(f"Ошибка создания IP адреса '{ip_address_str}'")

        # Импорт или обновление VM
        processed_count += 1
        if existing_vm:
            vm_id = existing_vm["id"]
            # Сравнение полей для определения, нужно ли обновление
            update_needed = False
            for key, value in payload.items():
                # Специальная обработка для вложенных словарей (site, cluster, role, platform, tenant, vrf)
                if isinstance(value, dict) and key in existing_vm and isinstance(existing_vm[key], dict):
                    if value.get("id") != existing_vm[key].get("id") and value.get("name") != existing_vm[key].get("name"):
                        update_needed = True
                        break
                elif key == "primary_ip4":
                    # Сравниваем ID IP-адреса
                    if isinstance(value, dict) and isinstance(existing_vm.get(key), dict):
                        if value.get("id") != existing_vm[key].get("id"):
                            update_needed = True
                            break
                    elif isinstance(value, dict) and existing_vm.get(key) is None:
                        update_needed = True
                        break
                    elif value is None and existing_vm.get(key) is not None:
                         update_needed = True
                         break
                elif value != existing_vm.get(key):
                    # Проверяем, что значение не является пустым, если в Netbox оно тоже пустое
                    if not (pd.isna(value) and existing_vm.get(key) is None or existing_vm.get(key) == ""):
                        update_needed = True
                        break
            
            if update_needed:
                update_virtual_machine(vm_id, payload)
            else:
                print(f"VM '{vm_name}' уже актуальна. Изменения не требуются.")
        else:
            create_virtual_machine(payload)

    # Вывод итоговой статистики
    print(f"\n{'='*60}")
    print(f"ИТОГОВАЯ СТАТИСТИКА ОБРАБОТКИ ЛИСТА '{sheet_name}':")
    print(f"{'='*60}")
    print(f"Всего записей в файле: {total_records}")
    print(f"Успешно обработано: {processed_count}")
    print(f"Пропущено: {skipped_count}")
    
    if skipped_records:
        print(f"\nПРОПУЩЕННЫЕ ЗАПИСИ:")
        print(f"{'Строка':<6} {'Имя VM':<25} {'Роль':<10} {'Причина'}")
        print("-" * 70)
        for record in skipped_records:
            print(f"{record['row']:<6} {str(record['name'])[:24]:<25} {str(record['role'])[:9]:<10} {record['reason']}")
    
    print(f"{'='*60}")

# --- Update run ---
if __name__ == "__main__":
    # Проверка доступности URL, наличия файла и токена. Check: URL, File and Token exist
    if NETBOX_URL == "http://netbox-instance.com":
        print("Ошибка: Пожалуйста, обновите NETBOX_URL в скрипте.")
    elif NETBOX_TOKEN == "api_token":
        print("Ошибка: Пожалуйста, обновите NETBOX_TOKEN в скрипте.")
    elif not os.path.exists(EXCEL_FILE_PATH):
        print(f"Ошибка: Excel файл '{EXCEL_FILE_PATH}' не найден.")
    else:
        import_vms_from_excel(EXCEL_FILE_PATH, SHEET_NAME)
        print("\nИмпорт завершен.")