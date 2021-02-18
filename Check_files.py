import os.path
import json
# install openpyxl, pandas
import openpyxl


class Check_files:

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

    # def read_from_xlsx(self):

