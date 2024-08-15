import os
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import threading
import time

from src.excel_concatenator.utils import resource_path
from src.excel_concatenator.files_processing import concatenate_files, save_file


class ExcelConcatenatorApp:
    def __init__(self):
        """
        Инициализация приложения для объединения Excel-файлов.

        Args:
            main_screen (tk.Tk): Основное окно приложения.
        """
        self.main_screen = tk.Tk()  # Создаем главное окно приложения
        self.loading_screen = None  # Окно индикатора загрузки
        self.selected_files = []  # Список выбранных файлов
        self.skip_top_rows = 0
        self.header_rows = 1
        self.skip_bottom_rows = 0
        # self.include_filename_column = tk.BooleanVar(value=True)  # Переменная для состояния чекбокса
        self.setup_main_screen()  # Настраиваем основной экран
        self.show_screen(self.main_screen)  # Отображаем главный экран

    def show_screen_widgets(self, widgets):
        """
        Отображает экран приложения, восстанавливая его виджеты.
        """
        for widget in widgets:
            widget.pack()

    def hide_screen_widgets(self, widgets):
        """
        Скрывает экран приложения, удаляя его виджеты.
        """
        for widget in widgets:
            widget.pack_forget()

    def clear_screen_widgets(self,root):
        """
        Очищает экран от всех виджетов.
        """
        for widget in root.winfo_children():
            widget.destroy()  # Удаляем каждый виджет

    def hide_screen(self,root):
        """Скрывает все виджеты и окно."""
        root.withdraw()

    def show_screen(self,root):
        """Показывает окно."""
        if root == self.main_screen:
            self.clear_screen_widgets(root=self.main_screen)
            self.setup_main_screen()
        else:
            root.deiconify()

    def update_screen(self,root):
        """Обновляет окно."""
        root.update()

    def setup_main_screen(self):
        """
        Настраивает виджеты главного экрана приложения.
        """
        self.clear_screen_widgets(root=self.main_screen)
        self.main_screen.title("Excel Concatenator")
        self.main_screen.geometry("800x600")

        # Создаем главный фрейм
        main_frame = tk.Frame(self.main_screen)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Инструкция 1
        label1 = tk.Label(
            main_frame,
            text="Выберите файлы или папку для объединения нескольких файлов или целой папки по строкам.",
            wraplength=750  # Ограничиваем ширину текста для лучшего отображения
        )
        label1.pack(anchor="center", pady=(0, 10))

        # Картинка
        image_widget = self.display_image(resource_path('assets/files_concatination_scheme.png'), root=main_frame, size=(500, 300))
        image_widget.pack(pady=(0, 10))

        # Инструкция 2
        label2 = tk.Label(
            main_frame,
            text="Поддерживаются файлы формата .xlsx .csv .xls .xlsm"
        )
        label2.pack(anchor="center", pady=(0, 20))

        # Фрейм для кнопок
        buttons_frame = tk.Frame(main_frame)
        buttons_frame.pack(fill=tk.X, pady=10)

        # Внутренний фрейм для выравнивания кнопок
        buttons_inner_frame = tk.Frame(buttons_frame)
        buttons_inner_frame.pack(side=tk.TOP, anchor="center")

        # Кнопка "Выбрать файлы"
        select_files_btn = tk.Button(
            buttons_inner_frame,
            text="Выбрать файлы",
            command=self.file_selection_window
        )
        select_files_btn.pack(side=tk.LEFT, padx=10)

        # Кнопка "Выбрать папку"
        select_folder_btn = tk.Button(
            buttons_inner_frame,
            text="Выбрать папку",
            command=self.folder_selection_window
        )
        select_folder_btn.pack(side=tk.LEFT, padx=10)

        # Кнопка "Завершить"
        exit_btn = tk.Button(
            main_frame,
            text="Завершить",
            command=self.main_screen.quit
        )
        exit_btn.pack(pady=(10, 0))

    def setup_loading_screen(self):
        """
        создаеь окно индикатора загрузки.
        """
        self.loading_screen = tk.Toplevel()
        self.loading_screen.geometry("400x200")
        self.loading_screen.title("Пожалуйста, подождите")
        loading_label = tk.Label(self.loading_screen, text="Обработка файлов, пожалуйста, подождите...")
        loading_label.pack(expand=True)

    def display_image(self, image_path, root, size=(20, 20)):
        """
        Отображает изображение на заданном родительском виджете с заданным размером.

        Args:
            image_path (str): Путь к изображению.
            parent (tk.Widget): Родительский виджет, на котором будет отображаться изображение.
            size (tuple): Размер изображения в формате (ширина, высота).

        Returns:
            tk.Label: Виджет метки с изображением.
        """
        try:
            image = Image.open(image_path)
            resized_image = image.resize(size)
            photo = ImageTk.PhotoImage(resized_image)

            # Удаляем старый логотип, если он существует
            if hasattr(self, 'image_label') and self.image_label.winfo_exists():
                self.image_label.config(image=photo)
                self.image_label.image = photo  # Обновляем изображение
            else:
                self.image_label = tk.Label(root, image=photo)
                self.image_label.image = photo  # Предотвращаем удаление объекта
                self.image_label.pack(pady=20)

            return self.image_label

        except Exception as e:
            messagebox.showerror("Ошибка загрузки изображения", f"Не удалось загрузить изображение: {e}")

    def file_selection_window(self):
        """
        Открывает окно для выбора нескольких файлов и обрабатывает выбор.
        """
        downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
        try:
            files = filedialog.askopenfilenames(
                initialdir=downloads_path,
                title="Выберите файлы Excel или CSV",
                filetypes=[
                    ("Supported filetypes", "*.xlsx;*.xls;*.xlsm;*.csv"),  # Все поддерживаемые форматы
                    ("Excel files", "*.xlsx;*.xls;*.xlsm"),  # Файлы Excel
                    ("CSV files", "*.csv"),  # Файлы CSV
                    ("Все файлы", "*.*")  # Все файлы
                ]
            )
            self.input_processing(files)  # Обработка выбора файлов

        except Exception as e:
            messagebox.showerror("Ошибка выбора файлов", f"Произошла ошибка при выборе файлов: {str(e)}")
            self.show_screen(self.main_screen)  # Возвращаемся на главный экран в случае ошибки

    def folder_selection_window(self):
        """
        Открывает окно для выбора папки с файлами и обрабатывает выбор.
        """
        downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
        try:
            folder = filedialog.askdirectory(
                initialdir=downloads_path,
                title="Выберите папку с файлами"
            )
            self.input_processing(folder)  # Обработка выбора папки

        except Exception as e:
            messagebox.showerror("Ошибка выбора папки", f"Произошла ошибка при выборе папки: {str(e)}")
            self.show_screen(self.main_screen)  # Возвращаемся на главный экран в случае ошибки

    def input_processing(self, input_path):
        """
        Обрабатывает выбранные пользователем файлы или папку с файлами Excel и CSV.

        Args:
            input_path (str or tuple): Путь к выбранной папке или список выбранных файлов.
        """
        try:
            # Проверяем, что input_path не пустой
            if not input_path:
                self.show_screen(self.main_screen)
                return

            # Поддерживаемые расширения файлов
            supported_extensions = ('.xlsx', '.xls', '.xlsm', '.csv', 'xlsb')

            # Если input_path - это список файлов
            if isinstance(input_path, tuple):
                files = input_path
                if not files:
                    raise ValueError("Выбор файлов отменен.")
                # Фильтруем выбранные файлы, оставляя только поддерживаемые форматы
                self.selected_files = [file for file in files if file.lower().endswith(supported_extensions)]
                invalid_files = [file for file in files if not file.lower().endswith(supported_extensions)]
            # Если input_path - это путь к папке
            elif isinstance(input_path, str) and os.path.isdir(input_path):
                folder = input_path
                files = [os.path.join(folder, file) for file in os.listdir(folder)]
                self.selected_files = [file for file in files if file.lower().endswith(supported_extensions)]
                invalid_files = [file for file in files if not file.lower().endswith(supported_extensions)]
            else:
                raise ValueError("Неверный формат входных данных. Ожидался путь к папке или список файлов.")

            # Проверяем на наличие поддерживаемых файлов
            if not self.selected_files:
                raise ValueError("Вы не выбрали ни одного поддерживаемого файла.")
            elif len(self.selected_files) == 1:
                raise ValueError("Необходимо выбрать более одного файла для объединения.")
            elif invalid_files:
                invalid_files_list = "\n".join(os.path.basename(file) for file in invalid_files)
                raise ValueError(f"Некорректные файлы:\n{invalid_files_list}")

            # Переходим к экрану подтверждения выбора
            self.selection_confirmation_frame()

        except ValueError as ve:
            messagebox.showerror("Ошибка выбора файлов", str(ve))
            self.show_screen(self.main_screen)  # Возвращаемся на главный экран в случае ошибки

        except Exception as e:
            # Обрабатываем любые другие непредвиденные ошибки
            messagebox.showerror("Ошибка", f"Произошла ошибка: {str(e)}")
            self.show_screen(self.main_screen)  # Возвращаемся на главный экран в случае ошибки

    def selection_confirmation_frame(self):
        """
        Отображает окно подтверждения выбранных файлов.
        """
        self.clear_screen_widgets(root=self.main_screen)  # Очищаем главный экран перед отображением нового контента

        # Создаем фрейм для основного содержимого
        content_frame = tk.Frame(self.main_screen)
        content_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # Создаем холст и вертикальный скроллбар
        canvas = tk.Canvas(content_frame)
        scrollbar = tk.Scrollbar(content_frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)

        # Функция для обновления области прокрутки при изменении размера scrollable_frame
        def update_scroll_region(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        scrollable_frame.bind("<Configure>", update_scroll_region)

        # Помещаем scrollable_frame на холст
        canvas.create_window((0, 0), window=scrollable_frame, anchor=tk.NW)
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Запрещаем изменение размера фрейма, содержащего прокрутку
        content_frame.grid_rowconfigure(0, weight=1)
        content_frame.grid_columnconfigure(0, weight=1)
        content_frame.grid_columnconfigure(1, weight=0)

        # Добавляем метку о выбранных файлах
        selection_label = tk.Label(scrollable_frame, text="Вы выбрали следующие файлы:")
        selection_label.pack(pady=10)

        # Добавляем список выбранных файлов
        for file in self.selected_files:
            file_label = tk.Label(scrollable_frame, text=os.path.basename(file), anchor="w", justify=tk.LEFT)
            file_label.pack(fill=tk.BOTH, padx=10)

        # Создаем фрейм для кнопок и чекбокса
        controls_frame = tk.Frame(self.main_screen)
        controls_frame.pack(side=tk.BOTTOM, pady=20, fill=tk.X)

        tk.Label(controls_frame, text="Количество строк заголовков:").pack(anchor="w", padx=10)
        self.header_rows_entry = tk.Entry(controls_frame)
        self.header_rows_entry.pack(padx=10, fill=tk.X)
        self.header_rows_entry.insert(0, str(self.header_rows))  # Устанавливаем текущее значение

        tk.Label(controls_frame, text="Пропустить строк сверху:").pack(anchor="w", padx=10)
        self.skip_top_rows_entry = tk.Entry(controls_frame)
        self.skip_top_rows_entry.pack(padx=10, fill=tk.X)
        self.skip_top_rows_entry.insert(0, str(self.skip_top_rows))  # Устанавливаем текущее значение

        tk.Label(controls_frame, text="Пропустить строк снизу:").pack(anchor="w", padx=10)
        self.skip_bottom_rows_entry = tk.Entry(controls_frame)
        self.skip_bottom_rows_entry.pack(padx=10, fill=tk.X)
        self.skip_bottom_rows_entry.insert(0, str(self.skip_bottom_rows))  # Устанавливаем текущее значение

        # Чекбокс для включения столбца с именами файлов
        self.include_filename_column = tk.BooleanVar(value=True)
        self.include_filename_checkbox_entry = tk.Checkbutton(
            controls_frame,
            text="Добавить столбец с именами файлов",
            variable=self.include_filename_column
        )
        self.include_filename_checkbox_entry.pack(side=tk.TOP, pady=10)

        # Создаем фрейм для кнопок
        buttons_frame = tk.Frame(controls_frame)
        buttons_frame.pack(side=tk.BOTTOM, pady=10, fill=tk.X)

        # Кнопка "Назад"
        back_button = tk.Button(
            buttons_frame,
            text="Назад",
            command=lambda: self.show_screen(self.main_screen)
        )
        back_button.pack(side=tk.LEFT, padx=20)

        # Кнопка "Объединить и сохранить"
        save_button = tk.Button(
            buttons_frame,
            text="Объединить и сохранить",
            command=self.concatenate_save
        )
        save_button.pack(side=tk.RIGHT, padx=20)

    def savepath_selection_window(self):
        """
        Открывает диалоговое окно для выбора места сохранения объединенного файла.
        """
        downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
        try:
            save_path = filedialog.asksaveasfilename(
            title="Сохранить результат",
            initialdir=downloads_path,
            defaultextension=".xlsx",
            filetypes=[
                ("Excel files (xlsx)", "*.xlsx"),   # Возможность выбора формата .xlsx
                ("CSV files (csv)", "*.csv"),      # Возможность выбора формата .csv
                ("Все файлы", "*.*")         # Возможность выбора всех форматов
            ]
            )
            if not save_path:  # Проверяем, если пользователь нажал "Отмена"
                return None  # Возвращаем None, чтобы показать, что сохранение было отменено

            return save_path
        except Exception as e:
            messagebox.showerror(title="Ошибка сохранения файла", message=f"Произошла ошибка при сохранении файла: {str(e)}")
            return None

    def concatenate_save(self):
        """
        Выполняет сохранение объединенного файла в фоновом режиме.
        """

        save_path = self.savepath_selection_window()
        if save_path is not None:
            self.setup_loading_screen()
            self.show_screen(self.loading_screen)
            self.update_screen(self.loading_screen)

            try:
                # Объединяем файлы один раз

                concatenation_result = concatenate_files(files=self.selected_files,
                                                         add_filename_column=self.include_filename_column.get(),
                                                         skip_top_rows=int(self.skip_top_rows_entry.get()),
                                                         header_rows=int(self.header_rows_entry.get()),
                                                         skip_bottom_rows=int(self.skip_bottom_rows_entry.get()))

                while True:
                    try:
                        # Пытаемся сохранить файл в выбранный формат
                        save_file(data=concatenation_result,save_path=save_path,csv_delimiter=';')
                        self.clear_screen_widgets(self.loading_screen)
                        self.show_screen(self.main_screen)  # Возвращаемся на главный экран
                        messagebox.showinfo("Успешное сохранение", f"Файл успешно сохранен:\n{save_path}")
                        break  # Выходим из цикла, если сохранение прошло успешно

                    except ValueError as e:
                        # Если ошибка связана с неподдерживаемым форматом, повторяем выбор места сохранения
                        self.hide_screen(self.loading_screen)
                        messagebox.showerror(message=f"{str(e)}.")
                        save_path = self.savepath_selection_window()  # Повторно открываем окно выбора пути сохранения файла

                        if save_path is None:
                            break
                        self.show_screen(self.loading_screen)



            except Exception as e:
                # Обработка других ошибок
                messagebox.showerror("Ошибка объединения файлов", f"Произошла ошибка при объединении файлов: {str(e)}")
