import os.path
import json
# install openpyxl, pandas
import openpyxl


class CheckFiles:

    def __init__(self):
        self.get_paths()
        self.check_files()

    def get_paths(self) -> None:
        """
        Получение путей к файлам с данными из файла служебного файла Data.txt.
        При неудаче выводится окно с ошибкой
        :return: None
        """
        try:
            with open('Data.txt', 'w', encoding='utf-8') as file:
                self.paths = json.load(file)
        except json.decoder.JSONDecodeError:
            "Реализовать вывод окна с ошибкой: Некорретный формат файла Data.txt"
        except FileNotFoundError:
            "Реализовать вывод окна с ошибкой: Файл Data.txt не найден"
        except Exception:
            "Реализовать вывод окна с ошибкой: Что-то пошло не так!!!"

    def check_files(self) -> None:
        """
        Производится проверка наличия файлов Distributors.xlsx и
        :return: None
        """
        for file in self.paths:
            if not os.path.exists(self.paths[file]):
                f"Реализовать вывод окна с ошибкой: Файл {file}.xlsx не найден"


class Distributors:

    def __init__(self, path):
        self.cell = None
        self.month = None
        self.path = path
        self.get_cell()
        self.is_valid_data_cell()

    def get_cell(self):
        """
        Получение содержимого ячейки А1 из файла.
        :return: None
        """
        workbook = openpyxl.load_workbook(self.path)
        table = workbook.active
        self.cell = table.cell(row=1, column=1)

    def is_valid_data_cell(self):
        """
        Проверка первой ячейки таблицы. Содержимое ячейки должно быть в формате даты.
        :return: bool
        """
        ...

    def get_month(self):
        """
        Получение месяца в цифровом коде из содержимого ячейки.
        :return: None
        """
        self.month = self.cell.value.month


