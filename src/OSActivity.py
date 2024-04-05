import os
import shutil


def create_folder(folder_path: str) -> None:
    """
    Создает папку
    :param folder_path: Путь до папки
    """
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)


def remove_folder(folder_path: str) -> None:
    """
    Удаляет папку
    :param folder_path: Путь до папки
    """
    if os.path.exists(folder_path):
        shutil.rmtree(folder_path)
