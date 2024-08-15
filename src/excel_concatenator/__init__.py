

"""
ExcelConcatenator Package
"""

# Импортирование необходимых классов и функций
from src.excel_concatenator.app import ExcelConcatenatorApp
from src.excel_concatenator.utils import resource_path  # Пример функции из utils.py
from src.excel_concatenator.files_processing import read_file_excel_formats, save_file
__all__ = ['ExcelConcatenatorApp', 'resource_path']

# Определение версии пакета
__version__ = '1.2.0'


