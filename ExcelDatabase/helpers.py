import re
import os
from config import DATA_FOLDER

def sanitize_filename(name):
    """заменяет недопустимые символы в имени файла на подчёркивания"""
    return re.sub(r'[\\/*?:"<>|]', '_', name)

def list_companies():
    """возвращает список названий компаний (без расширения) из папки данных"""
    companies = []
    for file in os.listdir(DATA_FOLDER):
        if file.endswith('.json'):
            companies.append(os.path.splitext(file)[0])
    return companies