import os

# константы
TEMPLATE_FILE = 'template.xlsx'   # шаблон Excel (в папке проекта)
DATA_FOLDER = 'company_data'       # папка для JSON-файлов

# создаём папку при импорте
os.makedirs(DATA_FOLDER, exist_ok=True)