import pandas as pd
import os

import pandas as pd

def concatenate_files(files, add_filename_column=False, skip_top_rows:int = 0, header_rows:int = 1,skip_bottom_rows:int = 0):
    concatenation_result = pd.DataFrame()

    for file in files:
        # Определяем тип файла по расширению
        file_extension = file.split('.')[-1].lower()

        try:
            if file_extension in ['xlsx', 'xls', 'xlsm', 'xlsb']:
                df = pd.read_excel(file)  # Читаем Excel файлы
            elif file_extension == 'csv':
                df = pd.read_csv(file, encoding='utf-8')  # Читаем CSV файлы с кодировкой UTF-8
            else:
                raise ValueError(f"Unsupported file format: {file_extension}")

            # Если необходимо, добавляем колонку с именем файла
            if add_filename_column:
                df['Файл'] = file.split('/')[-1]

            # Объединяем датафреймы
            concatenation_result = pd.concat([concatenation_result, df], ignore_index=True)

        except Exception as e:
            raise ValueError(f"Ошибка при обработке файла {file}: {e}")

    return concatenation_result

def save_file(data, path):
    """
    Сохраняет объединенные данные в формате .xlsx или .csv.

    Args:
        data (pd.DataFrame): Датафрейм, содержащий объединенные данные.
        path (str): Путь для сохранения файла.
    """
    # Определяем расширение файла
    file_extension = path.split('.')[-1].lower()

    # Проверяем формат файла
    if file_extension not in ['xlsx', 'csv']:
        raise ValueError("Неподдерживаемый формат файла. Пожалуйста, выберите .xlsx или .csv")

    try:
        if file_extension == 'xlsx':
            # Сохраняем в формате Excel (.xlsx)
            data.to_excel(path, index=False)
        elif file_extension == 'csv':
            # Сохраняем в формате CSV (.csv) с кодировкой UTF-8
            data.to_csv(path, index=False, encoding='utf-8')
    except Exception as e:
        # Обрабатываем все другие ошибки
        raise ValueError(f"Произошла ошибка при сохранении файла: {str(e)}")


