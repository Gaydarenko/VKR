"""
Проверка файлов
"""
import os.path
import json

from message_for_user import MessageError as Me


class CheckFiles:
    """
    Получение путей и их проверка
    """

    def __init__(self):
        self.paths = None
        self.get_paths()
        self.check_files()

    def get_paths(self) -> None:
        """
        Получение путей к файлам с данными из служебного файла Data.txt.
        При неудаче выводится окно с ошибкой
        :return: None
        """
        try:
            with open('Data.txt', 'r', encoding='utf-8') as file:
                self.paths = json.load(file)
        except json.decoder.JSONDecodeError:
            Me.message_window('Некорретный формат файла Data.txt')
        except FileNotFoundError:
            Me.message_window('Файл Data.txt не найден')
        except Exception as except_:
            print(except_)
            Me.message_window('Что-то пошло не так!!!')

    def check_files(self) -> None:
        """
        Производится проверка наличия файлов по указанным путям.
        :return: None
        """
        for file in self.paths:
            if not os.path.exists(self.paths[file]):
                Me.message_window(f'Файл {file} не найден')
