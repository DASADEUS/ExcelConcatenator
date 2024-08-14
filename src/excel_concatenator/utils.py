import sys
import os

def resource_path(relative_path):
    """ Получить абсолютный путь к ресурсу, который будет работать как при сборке в один файл, так и в стандартном режиме. """
    try:
        # PyInstaller создает временную папку _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath("")

    return os.path.join(base_path, relative_path)

