import os
import sys
import shutil

# Настройка Django
import django
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'config.settings')
django.setup()

from constants import TEMPLATE_FILE
from helpers import list_companies_from_db
from file_dialog import select_excel_file, save_excel_file_dialog
from excel_io import read_excel_and_save_to_db, load_from_db_and_create_excel, display_company_info_from_db
from companies.models import Company

def get_desktop_path():
    """возвращает путь к рабочему столу текущего пользователя в Windows"""
    try:
        import ctypes
        from ctypes import wintypes
        buf = ctypes.create_unicode_buffer(260)
        if ctypes.windll.shell32.SHGetFolderPathW(None, 0x0000, None, 0, buf) == 0:
            return buf.value
    except:
        pass
    fallback = os.path.join(os.path.expanduser("~"), "Desktop")
    if os.path.exists(fallback):
        return fallback
    return os.path.expanduser("~")

def delete_company(company_name):
    """удаляет компанию и всех её сотрудников с подтверждением"""
    try:
        company = Company.objects.get(name=company_name)
    except Company.DoesNotExist:
        print(f"❌ Компания '{company_name}' не найдена.")
        return

    print(f"\nВы действительно хотите удалить компанию '{company_name}' и всех её сотрудников?")
    confirm = input("Введите 'ДА' для подтверждения: ").strip()
    if confirm == 'ДА':
        company.delete()
        print(f"✅ Компания '{company_name}' удалена.")
    else:
        print("❌ Удаление отменено.")

def console_menu():
    desktop_path = get_desktop_path()

    while True:
        print("\n" + "="*50)
        print("ГЛАВНОЕ МЕНЮ")
        print("1. Выбрать Excel-файл и сохранить в базу данных")
        print("2. Загрузить данные компании из базы в Excel")
        print("3. Просмотреть информацию о компании из базы")
        print("4. Создать новый шаблон компании")
        print("5. Удалить компанию")
        print("6. Выйти")
        choice = input("Выберите пункт (1-6): ").strip()

        if choice == '1':
            method = input("Выберите способ: 1 - через проводник, 2 - ручной ввод пути (по умолчанию 1): ").strip()
            if method == '2':
                file_path = input("Введите путь к существующему Excel-файлу: ").strip()
                if not os.path.exists(file_path):
                    print("❌ Файл не найден.")
                    continue
            else:
                file_path = select_excel_file()
                if not file_path:
                    print("❌ Выбор файла отменён.")
                    continue
            try:
                company = read_excel_and_save_to_db(file_path)
                if company is None:
                    # ошибка уже выведена внутри функции (например, отсутствие таблиц)
                    continue
                print(f"Компания '{company}' успешно сохранена/обновлена.")
            except Exception as e:
                print(f"❌ Ошибка при обработке файла: {e}")

        elif choice == '2':
            companies = list_companies_from_db()
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
                out_path = input("Введите путь для сохранения Excel-файла: ").strip()
                if not out_path:
                    print("❌ Путь не указан.")
                    continue
            else:
                default_filename = f"{selected_company}.xlsx"
                out_path = save_excel_file_dialog(default_filename, desktop_path)
                if not out_path:
                    print("❌ Сохранение отменено.")
                    continue

            try:
                load_from_db_and_create_excel(selected_company, out_path)
            except Exception as e:
                print(f"❌ Ошибка при создании Excel: {e}")

        elif choice == '3':
            companies = list_companies_from_db()
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

            display_company_info_from_db(selected_company)

        elif choice == '4':
            if not os.path.exists(TEMPLATE_FILE):
                print(f"❌ Файл шаблона '{TEMPLATE_FILE}' не найден в текущей папке.")
                continue

            method = input("Выберите способ: 1 - через проводник, 2 - ручной ввод пути (по умолчанию 1): ").strip()
            if method == '2':
                user_input = input("Введите путь для сохранения нового файла: ").strip()
                if not user_input:
                    print("❌ Путь не указан.")
                    continue
                if os.path.isdir(user_input):
                    target_dir = user_input
                    base_name = "Новая_компания"
                else:
                    target_dir = os.path.dirname(user_input) or os.getcwd()
                    base_name_with_ext = os.path.basename(user_input)
                    if not base_name_with_ext.lower().endswith('.xlsx'):
                        base_name_with_ext += '.xlsx'
                    base_name = os.path.splitext(base_name_with_ext)[0]
                if not os.path.exists(target_dir):
                    print(f"❌ Указанная папка не существует: {target_dir}")
                    continue
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
                    print("❌ Создание отменено.")
                    continue

            try:
                shutil.copy2(TEMPLATE_FILE, out_path)
                print(f"✅ Новый шаблон создан: {out_path}")
            except Exception as e:
                print(f"❌ Ошибка при создании файла: {e}")

        elif choice == '5':
            companies = list_companies_from_db()
            if not companies:
                print("В базе нет сохранённых компаний.")
                continue
            print("Доступные компании:")
            for idx, comp in enumerate(companies, 1):
                print(f"{idx}. {comp}")
            comp_choice = input("Введите номер или название компании для удаления: ").strip()

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

            delete_company(selected_company)

        elif choice == '6':
            print("Выход.")
            break
        else:
            print("❌ Неверный ввод. Пожалуйста, выберите 1-6.")

if __name__ == "__main__":
    console_menu()