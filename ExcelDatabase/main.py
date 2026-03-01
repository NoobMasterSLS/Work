import os
import shutil
import ctypes
from ctypes import wintypes
from config import DATA_FOLDER, TEMPLATE_FILE
from helpers import list_companies
from file_dialog import select_excel_file, save_excel_file_dialog
from excel_io import read_excel_and_save_to_json, load_from_json_and_create_excel, read_json_company

def get_desktop_path():
    """
    возвращает путь к рабочему столу текущего пользователя в Windows
    если не удаётся (например, не Windows), возвращает стандартный путь
    """
    try:
        buf = ctypes.create_unicode_buffer(260)
        # CSIDL_DESKTOP = 0x0000
        if ctypes.windll.shell32.SHGetFolderPathW(None, 0x0000, None, 0, buf) == 0:
            return buf.value
    except:
        pass
    #запасной вариант
    fallback = os.path.join(os.path.expanduser("~"), "Desktop")
    if os.path.exists(fallback):
        return fallback
    return os.path.expanduser("~")

def display_company_info(company_name):
    try:
        data = read_json_company(company_name)
    except Exception as e:
        print(f"❌ Ошибка при чтении данных: {e}")
        return

    print("\n" + "="*60)
    print(f"ИНФОРМАЦИЯ О КОМПАНИИ: {company_name}")
    print("="*60)
    print("\n--- Общая информация ---")
    for key, value in data['company_info'].items():
        print(f"{key}: {value}")

    employees = data['employees']
    print(f"\n--- Сотрудники ({len(employees)}) ---")
    if not employees:
        print("Нет данных о сотрудниках.")
    else:
        headers = ['№', 'Фамилия', 'Имя', 'Отчество', 'Дата рожд.', 'Дата приёма', 'Телефон', 'Должность', 'Оклад']
        col_widths = [5, 15, 10, 15, 12, 12, 12, 15, 10]
        header_line = ' '.join([f"{h:<{w}}" for h, w in zip(headers, col_widths)])
        print(header_line)
        print('-' * len(header_line))

        for i, emp in enumerate(employees, 1):
            row = [
                str(i),
                emp.get('last_name', '')[:15],
                emp.get('first_name', '')[:10],
                emp.get('middle_name', '')[:15],
                emp.get('birth_date', '')[:12],
                emp.get('hire_date', '')[:12],
                emp.get('phone', '')[:12],
                emp.get('position', '')[:15],
                emp.get('salary', '')[:10]
            ]
            row_line = ' '.join([f"{str(cell):<{w}}" for cell, w in zip(row, col_widths)])
            print(row_line)
    print("="*60)

def console_menu():
    # получаем точный путь к рабочему столу
    desktop_path = get_desktop_path()

    while True:
        print("\n" + "="*50)
        print("ГЛАВНОЕ МЕНЮ")
        print("1. Выбрать Excel-файл и сохранить в базу данных")
        print("2. Загрузить данные компании из базы в Excel")
        print("3. Просмотреть информацию о компании из базы")
        print("4. Создать новый шаблон компании")
        print("5. Выйти")
        choice = input("Выберите пункт (1-5): ").strip()

        if choice == '1':
            method = input("Выберите способ: 1 - через проводник, 2 - ручной ввод пути (по умолчанию 1): ").strip()
            if method == '2':
                file_path = input("Введите путь к существующему Excel-файлу (например, C:\\Users\\...\\файл.xlsx): ").strip()
                if not os.path.exists(file_path):
                    print("❌ Файл не найден.")
                    continue
            else:
                file_path = select_excel_file()
                if not file_path:
                    print("❌ Выбор файла отменён или проводник недоступен.")
                    continue
            try:
                company = read_excel_and_save_to_json(file_path)
                print(f"Компания '{company}' успешно сохранена/обновлена.")
            except Exception as e:
                print(f"❌ Ошибка при обработке файла: {e}")

        elif choice == '2':
            companies = list_companies()
            if not companies:
                print("В базе нет сохранённых компаний.")
                continue
            print("Доступные компании:")
            for idx, comp in enumerate(companies, 1):
                print(f"{idx}. {comp}")
            comp_choice = input("Введите номер или название компании: ").strip()

            selected_company = None
            if comp_choice.isdigit():
                idx = int(comp_choice)
                if 1 <= idx <= len(companies):
                    selected_company = companies[idx-1]
                else:
                    print("❌ Неверный номер.")
                    continue
            else:
                found = [c for c in companies if c.lower() == comp_choice.lower()]
                if found:
                    selected_company = found[0]
                else:
                    print("❌ Компания с таким названием не найдена.")
                    continue

            save_method = input("Выберите способ: 1 - через проводник, 2 - ручной ввод пути (по умолч. 1): ").strip()
            if save_method == '2':
                out_path = input("Введите путь для сохранения Excel-файла (например, C:\\Users\\...\\файл.xlsx): ").strip()
                if not out_path:
                    print("❌ Путь не указан.")
                    continue
            else:
                default_filename = f"{selected_company}.xlsx"
                out_path = save_excel_file_dialog(default_filename, desktop_path)
                if not out_path:
                    print("❌ Сохранение отменено или через проводник недоступен.")
                    continue

            try:
                load_from_json_and_create_excel(selected_company, out_path)
            except Exception as e:
                print(f"❌ Ошибка при создании Excel: {e}")

        elif choice == '3':
            companies = list_companies()
            if not companies:
                print("В базе нет сохранённых компаний.")
                continue
            print("Доступные компании:")
            for idx, comp in enumerate(companies, 1):
                print(f"{idx}. {comp}")
            comp_choice = input("Введите номер или название компании: ").strip()

            selected_company = None
            if comp_choice.isdigit():
                idx = int(comp_choice)
                if 1 <= idx <= len(companies):
                    selected_company = companies[idx-1]
                else:
                    print("❌ Неверный номер.")
                    continue
            else:
                found = [c for c in companies if c.lower() == comp_choice.lower()]
                if found:
                    selected_company = found[0]
                else:
                    print("❌ Компания с таким названием не найдена.")
                    continue

            display_company_info(selected_company)

        elif choice == '4':
            if not os.path.exists(TEMPLATE_FILE):
                print(f"❌ Файл шаблона '{TEMPLATE_FILE}' не найден в текущей папке.")
                continue

            method = input("Выберите способ: 1 - через проводник, 2 - ручной ввод пути (по умолчанию 1): ").strip()
            if method == '2':
                user_input = input("Введите путь для сохранения нового файла (например, C:\\Users\\...\\Новая_компания.xlsx): ").strip()
                if not user_input:
                    print("❌ Путь не указан.")
                    continue

                # определяем, является ли введённый путь папкой
                if os.path.isdir(user_input):
                    target_dir = user_input
                    base_name = "Новая_компания"
                else:
                    # это путь к файлу (или несуществующий путь)
                    target_dir = os.path.dirname(user_input)
                    base_name_with_ext = os.path.basename(user_input)
                    if not target_dir:  # если ввели только имя файла без пути
                        target_dir = os.getcwd()
                    # если имя файла не содержит расширение .xlsx, добавляем его
                    if not base_name_with_ext.lower().endswith('.xlsx'):
                        base_name_with_ext += '.xlsx'
                    # отделяем имя от расширения для генерации уникальности
                    base_name = os.path.splitext(base_name_with_ext)[0]
                    extension = '.xlsx'

                # проверяем, существует ли целевая директория
                if not os.path.exists(target_dir):
                    print(f"❌ Указанная папка не существует: {target_dir}")
                    continue

                # генерируем уникальное имя файла в целевой папке
                extension = '.xlsx'
                counter = 0
                while True:
                    if counter == 0:
                        filename = f"{base_name}{extension}"
                    else:
                        filename = f"{base_name} {counter}{extension}"
                    full_path = os.path.join(target_dir, filename)
                    if not os.path.exists(full_path):
                        break
                    counter += 1

                out_path = full_path
                print(f"Будет создан файл: {out_path}")
            else:
                # режим проводника: генерируем уникальное имя на рабочем столе и открываем проводник
                base_name = "Новая_компания"
                extension = ".xlsx"
                counter = 0
                while True:
                    if counter == 0:
                        filename = f"{base_name}{extension}"
                    else:
                        filename = f"{base_name} {counter}{extension}"
                    full_path = os.path.join(desktop_path, filename)
                    if not os.path.exists(full_path):
                        break
                    counter += 1
                out_path = save_excel_file_dialog(filename, desktop_path)
                if not out_path:
                    print("❌ Создание отменено или через проводник недоступен.")
                    continue

            try:
                shutil.copy2(TEMPLATE_FILE, out_path)
                print(f"✅ Новый шаблон создан: {out_path}")
            except Exception as e:
                print(f"❌ Ошибка при создании файла: {e}")

        elif choice == '5':
            print("Выход.")
            break
        else:
            print("❌ Неверный ввод. Пожалуйста, выберите 1, 2, 3, 4 или 5.")

if __name__ == "__main__":
    console_menu()