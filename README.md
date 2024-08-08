# ExcelConcatenator

ExcelConcatenator - это простое приложение на Python для объединения нескольких файлов Excel в один. Приложение предоставляет графический интерфейс для выбора файлов или папки и поддерживает опциональное добавление столбца с именем исходного файла.

## Структура

```
ExcelConcatenator/
├── .gitignore
├── README.md
├── requirements.txt
├── setup.py
├── main.py
├── excel_concatenator/
│   ├── __init__.py
│   ├── app.py
│   └── utils.py
└── assets/
    └── logo.png
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
#Windows:
python -m venv venv
#macOS/Linux:
venv\Scripts\activate
```

### 3. Установка зависимостей
```bash
pip install --upgrade pip
pip install -r requirements.txt
```

### 4. Установка пакета
```bash
pip install .
или
python setup.py install
```

### 6. Проверка установленных зависимостей
```bash
pip list
```

## Запуск

### 1. С использованием python

```bash
python -m main
```
### 2. С использованием командной строки (если установлен через setup.py)

```bash
excel_concatenator
```
## Создание исполняемого файла (exe)

Для создания исполняемого файла вы можете использовать pyinstaller.

### Установка

```bash
pip install pyinstaller
```

### Создание исполняемого файла

```bash
pyinstaller --onefile -w -n excel_concatenator --add-data 'assets;assets' main.py

```
Исполняемый файл будет находиться в папке dist.


### Использование

#### 1.Запустите приложение.
#### 2.Выберите файлы или папку с файлами Excel для объединения.
#### 3.Выберите, нужно ли добавлять столбец с именем исходного файла (установите/снимите галочку).
#### 4.Нажмите "Сохранить результат" и выберите место для сохранения объединенного файла.
