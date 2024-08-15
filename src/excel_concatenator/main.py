from src.excel_concatenator.app import ExcelConcatenatorApp
import tkinter as tk
import os
import sys

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), 'src')))

def main():
    """
    Главная функция для запуска приложения.
    """
    app = ExcelConcatenatorApp()  # Инициализируем приложение
    app.main_screen.mainloop()  # Запускаем основной цикл обработки событий

# Импортирование основного файла для удобного запуска


# Добавление текущей директории в sys.path для запуска приложения из пакета
if __name__ == "__main__":
    main()

