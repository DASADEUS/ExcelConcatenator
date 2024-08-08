from setuptools import setup, find_packages

# Чтение зависимостей из файла requirements.txt
with open("requirements.txt") as f:
    required_packages = f.read().splitlines()

setup(
    name="ExcelConcatenator",
    version="0.1",
    packages=find_packages(),
    install_requires=required_packages,
    include_package_data=True,  # Чтобы включить не-Python файлы, такие как изображения
    package_data={
        '': ['assets/*'],  # Указывает на включение всех файлов из папки assets
    },
    entry_points={
        "console_scripts": [
            "excel_concatenator=excel_concatenator.app:main",
        ],
    },
)
