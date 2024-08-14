from src.excel_concatenator.app import ExcelConcatenatorApp
import tkinter as tk

def main():
    """
    Главная функция для запуска приложения.
    """
    root = tk.Tk()
    app = ExcelConcatenatorApp(root)
    root.mainloop()

# Импортирование основного файла для удобного запуска
import os
import sys

# Добавление текущей директории в sys.path для запуска приложения из пакета
if __name__ == "__main__":
    main()

