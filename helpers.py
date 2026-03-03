import re
from companies.models import Company
from django.db.utils import OperationalError

def sanitize_filename(name):
    """заменяет недопустимые символы в имени файла на подчёркивания"""
    return re.sub(r'[\\/*?:"<>|]', '_', name)

def list_companies_from_db():
    """возвращает список названий компаний из БД, либо пустой список, если таблицы нет"""
    try:
        return list(Company.objects.values_list('name', flat=True))
    except OperationalError as e:
        if 'no such table' in str(e):
            print("\n❌ База данных не инициализирована. Пожалуйста, выполните данную команду в консоль:")
            print("   python manage.py migrate\n")
            return []
        else:
            # если это другая ошибка
            raise