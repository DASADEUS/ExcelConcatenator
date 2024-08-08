from excel_concatenator.app import ExcelConcatenatorApp
import tkinter as tk

def main():
    """
    Главная функция для запуска приложения.
    """
    root = tk.Tk()
    app = ExcelConcatenatorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()

