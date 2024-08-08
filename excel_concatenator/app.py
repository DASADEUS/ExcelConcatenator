import os
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import pandas as pd
import threading

class ExcelConcatenatorApp:
    def __init__(self, root):
        """
        Инициализация класса приложения.

        Args:
            root (tk.Tk): Главный корневой элемент tkinter.
        """
        self.root = root
        self.root.title("Excel Concatenator")  # Устанавливаем заголовок окна
        self.files = []  # Список для хранения выбранных файлов
        self.loading_window = None  # Окно индикатора загрузки
        self.add_filename_column = tk.BooleanVar()  # Переменная для чекбокса "Добавить столбец с названием файла"
        self.init_main_screen()  # Инициализируем главный экран

    def init_main_screen(self):
        """
        Инициализирует главный экран с выбором файлов или папки и отображает логотип.
        """
        self.root.geometry("800x600")  # Устанавливаем размер окна 800x600 пикселей
        self.clear_screen()  # Очищаем экран от предыдущих виджетов

        # Текстовая метка с инструкцией
        self.label_m1 = tk.Label(self.root, text="Объединение возможно по принципу конкатинации по строкам")
        self.label_m1.pack(pady=10)

        # Загружаем и отображаем изображение логотипа
        self.load_image('assets/files_concatination_scheme.png')

        # Текстовая метка с инструкцией
        self.label_m2 = tk.Label(self.root, text="Выберите файлы формата xlsx или папку с ними для объединения")
        self.label_m2.pack(pady=10)

        # Фрейм для размещения кнопок "Выбрать файлы" и "Выбрать папку" на одном уровне
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=5)

        # Кнопка для выбора файлов
        self.btn_select_files = tk.Button(button_frame, text="Выбрать файлы", command=self.select_files)
        self.btn_select_files.grid(row=0, column=0, padx=5)

        # Кнопка для выбора папки
        self.btn_select_folder = tk.Button(button_frame, text="Выбрать папку", command=self.select_folder)
        self.btn_select_folder.grid(row=0, column=1, padx=5)

        # Кнопка для завершения работы приложения
        self.btn_exit = tk.Button(self.root, text="Завершить", command=self.root.quit)
        self.btn_exit.pack(pady=20)

    def load_image(self, image_path):
        """
        Загружает и отображает изображение на главном экране, изменяя его размер до 400x200 пикселей.

        Args:
            image_path (str): Путь к изображению.
        """
        try:
            image = Image.open(image_path)  # Открываем изображение
            image = image.resize((400, 200))  # Изменяем размер изображения на 400x200 пикселей
            photo = ImageTk.PhotoImage(image)  # Преобразуем изображение для tkinter

            # Удаляем старое изображение, если оно есть
            if hasattr(self, 'image_label'):
                self.image_label.destroy()

            # Создаем метку с изображением
            self.image_label = tk.Label(self.root, image=photo)
            self.image_label.image = photo  # Необходимо для предотвращения сборщика мусора
            self.image_label.pack(pady=10)
        except Exception as e:
            messagebox.showerror("Ошибка загрузки изображения", f"Не удалось загрузить изображение: {e}")

    def select_files(self):
        """
        Открывает диалоговое окно для выбора файлов Excel.
        """
        downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")  # Путь к папке Загрузки
        try:
            files = filedialog.askopenfilenames(
                initialdir=downloads_path,
                title="Выберите файлы Excel",
                filetypes=(("Excel files", "*.xlsx"), ("Все файлы", "*.*")),
                multiple=True
            )

            if not files:
                raise ValueError("Выбор файлов отменен.")

            self.files = [file for file in files if file.endswith('.xlsx')]  # Фильтруем файлы по расширению .xlsx
            non_excel_files = [file for file in files if not file.endswith('.xlsx')]

            if non_excel_files:
                non_excel_files_list = "\n".join(os.path.basename(file) for file in non_excel_files)
                raise ValueError(f"Неподходящие файлы:\n{non_excel_files_list}")

            if not self.files:
                raise ValueError("Ни один из выбранных файлов не является файлом Excel.")
            elif len(self.files) == 1:
                raise ValueError("Необходимо выбрать более одного файла для объединения.")

            self.confirm_selection()  # Переходим к подтверждению выбора

        except ValueError as ve:
            messagebox.showerror("Ошибка выбора файлов", str(ve))
            self.init_main_screen()  # Возвращаемся на главный экран в случае ошибки

        except FileNotFoundError:
            messagebox.showerror("Ошибка файла", "Не удалось найти указанный файл.")
            self.init_main_screen()  # Возвращаемся на главный экран в случае ошибки

        except PermissionError:
            messagebox.showerror("Ошибка доступа", "Отказано в доступе к файлу или папке.")
            self.init_main_screen()  # Возвращаемся на главный экран в случае ошибки

        except Exception as e:
            messagebox.showerror("Ошибка выбора файлов", f"Произошла ошибка при выборе файлов: {str(e)}")
            self.init_main_screen()  # Возвращаемся на главный экран в случае ошибки

    def get_excel_files_from_folder(self, folder):
        """
        Получает файлы Excel из указанной папки.

        Args:
            folder (str): Путь к папке.

        Returns:
            tuple: Список файлов Excel и список неподходящих файлов.
        """
        files = [os.path.join(folder, file) for file in os.listdir(folder)]
        excel_files = [file for file in files if file.endswith('.xlsx')]
        non_excel_files = [file for file in files if not file.endswith('.xlsx')]
        return excel_files, non_excel_files

    def select_folder(self):
        """
        Открывает диалоговое окно для выбора папки.
        """
        downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")  # Путь к папке Загрузки
        try:
            folder = filedialog.askdirectory(
                initialdir=downloads_path,
                title="Выберите папку с файлами Excel"
            )

            if not folder:
                raise ValueError("Выбор папки отменен.")

            self.files, non_excel_files = self.get_excel_files_from_folder(folder)

            if not self.files:
                raise ValueError("В выбранной папке нет файлов Excel.")
            elif len(self.files) == 1:
                raise ValueError("В выбранной папке только один файл. Необходимо выбрать папку с более чем одним файлом.")
            elif non_excel_files:
                non_excel_files_list = "\n".join(os.path.basename(file) for file in non_excel_files)
                raise ValueError(f"Неподходящие файлы:\n{non_excel_files_list}")

            self.confirm_selection()  # Переходим к подтверждению выбора

        except ValueError as ve:
            messagebox.showerror("Ошибка выбора папки", str(ve))
            self.init_main_screen()  # Возвращаемся на главный экран в случае ошибки

        except FileNotFoundError:
            messagebox.showerror("Ошибка файла", "Не удалось найти указанную папку.")
            self.init_main_screen()  # Возвращаемся на главный экран в случае ошибки

        except PermissionError:
            messagebox.showerror("Ошибка доступа", "Отказано в доступе к папке.")
            self.init_main_screen()  # Возвращаемся на главный экран в случае ошибки

        except Exception as e:
            messagebox.showerror("Ошибка выбора папки", f"Произошла ошибка при выборе папки: {str(e)}")
            self.init_main_screen()  # Возвращаемся на главный экран в случае ошибки

    def confirm_selection(self):
        """
        Показывает экран подтверждения выбранных файлов.
        """
        self.clear_screen()  # Очищаем экран

        # Создаем фрейм для основного содержимого
        content_frame = tk.Frame(self.root)
        content_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # Создаем холст и вертикальный скроллбар
        canvas = tk.Canvas(content_frame)
        scrollbar = tk.Scrollbar(content_frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)

        # Функция для обновления скролла
        def on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        scrollable_frame.bind("<Configure>", on_frame_configure)

        canvas.create_window((0, 0), window=scrollable_frame, anchor=tk.NW)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.configure(yscrollcommand=scrollbar.set)

        # Метка с текстом о выбранных файлах
        self.label = tk.Label(scrollable_frame, text="Вы выбрали следующие файлы:")
        self.label.pack(pady=10)

        # Отображаем выбранные файлы
        for file in self.files:
            file_name = os.path.basename(file)  # Получаем только имя файла
            file_label = tk.Label(scrollable_frame, text=file_name, anchor=tk.W)
            file_label.pack(pady=2, padx=10, anchor=tk.W)

        # Фрейм для чекбокса и кнопок
        controls_frame = tk.Frame(self.root)
        controls_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=10)

        # Чекбокс для добавления столбца с названием файла
        self.checkbox = tk.Checkbutton(
            controls_frame,
            text="Добавить столбец с названием файла",
            variable=self.add_filename_column
        )
        self.checkbox.pack(pady=5)

        # Кнопка "Назад" для возврата на главный экран
        self.btn_back = tk.Button(controls_frame, text="Назад", command=self.init_main_screen)
        self.btn_back.pack(side=tk.LEFT, padx=20)

        # Кнопка "Сохранить результат" для сохранения объединенных данных
        self.btn_save = tk.Button(controls_frame, text="Объединить и сохранить", command=self.save_result)
        self.btn_save.pack(side=tk.RIGHT, padx=20)

    def show_loading(self):
        """
        Отображает индикатор выполнения в отдельном окне.
        """
        if self.loading_window is None:
            self.loading_window = tk.Toplevel(self.root)
            self.loading_window.title("Загрузка")
            self.loading_window.geometry("300x100")
            self.loading_window.grab_set()  # Блокируем основной интерфейс
            tk.Label(self.loading_window, text="Пожалуйста, подождите...", font=("Arial", 14)).pack(pady=20)
            self.root.update()  # Обновляем основной интерфейс

    def hide_loading(self):
        """
        Скрывает индикатор выполнения.
        """
        if self.loading_window:
            self.loading_window.destroy()
            self.loading_window = None

    def save_result(self):
        """
        Открывает диалоговое окно для выбора места сохранения файла и сохраняет объединенные данные.
        """
        location = os.path.join(os.path.expanduser("~"), "Downloads")  # Путь к папке Загрузки
        save_location = filedialog.asksaveasfilename(
            title="Сохранить результат",
            initialdir=location,
            defaultextension=".xlsx",
            filetypes=(("Excel files", "*.xlsx"), ("Все файлы", "*.*"))
        )

        if save_location:
            # Создаем поток для выполнения операции сохранения, чтобы не блокировать интерфейс
            threading.Thread(target=self.save_result_in_background, args=(save_location,)).start()

    def save_result_in_background(self, save_location):
        """
        Сохраняет результат в фоновом режиме.

        Args:
            save_location (str): Путь к файлу для сохранения результата.
        """
        self.show_loading()  # Отображаем индикатор загрузки
        self.concatenate_excel(self.files, save_location)  # Объединяем и сохраняем файлы
        self.hide_loading()  # Скрываем индикатор загрузки
        self.init_main_screen()  # Возвращаемся на главный экран

    def concatenate_excel(self, files, save_location):
        """
        Объединяет файлы Excel и сохраняет результат в указанный файл.

        Args:
            files (list): Список путей к файлам Excel.
            save_location (str): Путь к файлу для сохранения результата.
        """
        dfs = []  # Список для хранения данных из файлов Excel

        for file in files:
            try:
                df = pd.read_excel(file, engine='openpyxl')  # Читаем файл Excel
                if self.add_filename_column.get():
                    df['Source File'] = os.path.basename(file)  # Добавляем столбец с названием файла
                dfs.append(df)  # Добавляем DataFrame в список
            except Exception as e:
                messagebox.showerror("Ошибка чтения файла", f"Не удалось прочитать файл {file}: {e}")
                return

        if not dfs:
            messagebox.showwarning("Предупреждение", "Нет файлов для объединения.")
            return

        try:
            concatenated_df = pd.concat(dfs, ignore_index=True, sort=False, axis=0)  # Объединяем данные
            concatenated_df.to_excel(save_location, index=False, engine='openpyxl')  # Сохраняем в Excel
            messagebox.showinfo("Успех", f"Файлы успешно объединены и сохранены в {save_location}")
        except Exception as e:
            messagebox.showerror("Ошибка при объединении", f"Произошла ошибка при объединении файлов: {e}")

    def clear_screen(self):
        """
        Очищает текущий экран.
        """
        for widget in self.root.winfo_children():
            widget.destroy()  # Удаляем все виджеты с экрана
