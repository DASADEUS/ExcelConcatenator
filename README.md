# ExcelConcatenator
ExcelConcatenator — это простое приложение на Python для объединения нескольких файлов Excel в один. Приложение предоставляет графический интерфейс для выбора файлов или папки и поддерживает опциональное добавление столбца с именем исходного файла.
## Структура

```
ExcelConcatenator/
├── .gitignore
├── README.md
├── requirements.txt
├── setup.py
├── src/
│   └── excel_concatenator/
│       ├── __init__.py
│       ├── app.py
│       ├── files_processing.py
│       ├── utils.py
│       └── main.py
└── assets/
    └── files_concatination_scheme.png
```

## Установка Для Windows

### 1. Клонирование репозитория

```bash
git clone https://github.com/yourusername/ExcelConcatenator.git
cd ExcelConcatenator
```

### 2. Установка виртуального окружения
Рекомендуется использовать виртуальное окружение для изоляции зависимостей.
```bash
python -m venv venv
venv\Scripts\activate
```

Проверка версии Python:
```bash
python --version
```
Проверка PYTHONPATH
```bash
echo PYTHONPATH
или 
echo $env:PYTHONPATH
```
Если необходимо явно установить PYTHONPATH, используйте
```bash
set PYTHONPATH=C:\Users\YourUserName\GitHub\ExcelConcatenator
или
$env:PYTHONPATH="C:\Users\YourUserName\GitHub\ExcelConcatenator"
```

### 3. Установка зависимостей
```bash
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

Проверка установленных зависимостей
```bash
pip list
```

### 4. Запуск приложения с использованием python

```bash
python -m excel_concatenator\main.py
или
python -m excel_concatenator.main
```
### 5. Создание исполняемого файла (exe)

Для создания исполняемого файла вы можете использовать pyinstaller.

### Установка
Если PyInstaller еще не установлен, установите его:
```bash
pip install pyinstaller
```

### Создание исполняемого файла
Выполните следующую команду для создания исполняемого файла:
```bash
pyinstaller --clean --onefile --windowed --name excel_concatenator --add-data "assets/files_concatination_scheme.png;assets" --exclude-module torch src\excel_concatenator\main.py
```
После завершения процесса создания, исполняемый файл будет находиться в папке dist.

Возможно возникновение проблемы с лимитом рекурсии
Можно добавить в начало файла .spec увеличение лимита и запустить его
```bash
import sys
sys.setrecursionlimit(1000)
```
### 6. Использование

 1.Запустите приложение.

 2.Выберите файлы или папку с файлами Excel для объединения.

 3.Выберите, нужно ли добавлять столбец с именем исходного файла (установите/снимите галочку).

 4.Нажмите "Сохранить результат" и выберите место для сохранения объединенного файла.

