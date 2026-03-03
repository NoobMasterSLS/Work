"""
модуль для работы с проводником для выбора и сохранения файлов
использует Tkinter с созданием скрытого корневого окна и отображением поверх всех окон
"""

import tkinter as tk
from tkinter import filedialog
import time
import os  # добавлен для работы с путями

def tkinter_available():
    """проверяет, может ли Tkinter создать окно"""
    try:
        root = tk.Tk()
        root.withdraw()
        root.update()
        root.destroy()
        return True
    except Exception as e:
        print(f"⚠️ Tkinter недоступен: {e}")
        return False

def select_excel_file():
    """открывает проводник для выбора файла Excel через Tkinter с показом окна"""
    if not tkinter_available():
        print("❌ Tkinter не работает. Используйте ручной ввод пути.")
        return ""

    try:
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)  # поверх всех окон
        root.update()
        time.sleep(0.1)

        file_path = filedialog.askopenfilename(
            title="Выберите Excel-файл",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )

        root.destroy()
        return file_path
    except Exception as e:
        print(f"❌ Ошибка при открытии проводника: {e}")
        return ""

def save_excel_file_dialog(initial_filename, initial_dir=None):
    """открывает проводник для сохранения файла Excel через Tkinter."""
    if not tkinter_available():
        print("❌ Tkinter не работает. Используйте ручной ввод пути.")
        return ""

    try:
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        root.update()
        time.sleep(0.1)

        # если начальная папка не указана, используем текущую
        if initial_dir is None:
            initial_dir = os.getcwd()

        file_path = filedialog.asksaveasfilename(
            title="Сохранить Excel-файл как",
            defaultextension=".xlsx",
            initialfile=initial_filename,
            initialdir=initial_dir,
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )

        root.destroy()
        return file_path
    except Exception as e:
        print(f"❌ Ошибка при открытии проводника для сохранения: {e}")
        return ""