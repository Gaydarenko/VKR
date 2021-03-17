import unittest
import shutil
from openpyxl import load_workbook, Workbook
import random
import datetime as dt
import json
import os
import time
import sys

from check_files import CheckFiles as Cf
from distributors import Distributors as Distr


BASIC_DATA = {
    "Distributors": "data\\Дистрибьюторы.xlsx",
    "Contractors": "data\\Контрагенты.xlsx",
    "Reports": "data\\Отчеты.xlsx",
    "Database": "data\\pearl.xlsx",
    "KAM": "data\\Торговые_представители.xlsx",
    "basic_tables_archive": "data\\archive\\basic_tables\\",
    "sales_representatives_archive": "data\\archive\\sales_representatives\\",
    "sales_representatives": "to_send\\"
}

DISTRIBUTORS = [[dt.datetime.now(), ],
                ["Азазель",	"azazelik@mail.ru"],
                ["Гайдаренко Е.Г.", "gaydarenko@mail.ru"],
                ["Жена", "wife@love.me"],
                ["Кот и кошка", "cat@112.55"],
                ["Мама", "mother@family.com"],
                ["Папа", "father@google.com"]]

CONTRACTORS = [["Контрагент", "Дистрибьютор", "ДистрибьютерРегион", "Регион", "Федеральный округ", "ИНН", "Class", "ID"],
               ["АНРО", "Анима", "Санкт-Петербург г", "РФ", "Северо-Запад", "6802420243", "D", "2"],
               ["ВетТрейд", "Мурманская станция", "Мурманская обл", "РФ", "Северо-Запад", "6802420244", "С", "1"]]

REPORTS = [["Report", "Year", "Month", "Date", "Контрагент", "Дистрибьютор", "ДистрибьютерРегион", "Регион", "Федеральный округ", "ИНН", "Vet Category DHP", "Vet Category Bravecto", "Pet Category Value", "Адрес", "Сеть", "Clients_Type", "СББЖ", "KAM", "IFP", "UIN", "GPF", "Type", "Sum-количествоОборот", "Sum-Дозы", "Sum-СтоимостьБезНДСОборот1"], ]

SALES_REPRESENTATIVES = [["KAM", "email"],
                         ["Ivanov", "azazelik@mail.ru"],
                         ["Pugach", "gaydarenko@mail.ru"],
                         ["Kucc", "gaydarenko@mail.ru"]]

BASIC_TABLE = [["Year", "Month", "Date", "Контрагент", "Дистрибьютор", "ДистрибьютерРегион", "Регион", "Федеральный округ", "ИНН", "Vet Category DHP", "Vet Category Bravecto", "Pet Category Value", "Адрес", "Сеть", "Clients_Type", "СББЖ", "KAM", "IFP", "UIN", "GPF", "Type", "Sum-количествоОборот", "Sum-Дозы", "Sum-СтоимостьБезНДСОборот1"],
               ["2016", "12", "2016-12", "АНРО", "Анима", "Санкт-Петербург г", "РФ", "Северо-Запад", "6802420243", "D", "2", "3", "Санкт-Петербург, Грибалевой 7", "АнимаТрейд", "Vet", "Прочие", "Pugach", "Nobivac Rabies 10x1ds", 	"153698", "Rabies (Alu)", "Bio", "0.5", "5", "375"],
               ["2016", "12", "2016-12", "ВетТрейд", "Мурманская станция", "Мурманская обл", "РФ", "Северо-Запад", "6802420244", "С", "1", "1", "Мурманск, Грибалевой 8", "Прочие", "Pet", "СББЖ", "Pugach", "Vasotop P 0.625mg 3x28tab", "153699", "Vasotop", "Pharma", "1", "3", "375"]]

FILES = {
    "Distributors": DISTRIBUTORS,
    "Contractors": CONTRACTORS,
    "Reports": REPORTS,
    "Database": BASIC_TABLE,
    "KAM": SALES_REPRESENTATIVES,
}


def create_data_txt() -> None:
    """
    Создание файла Data.txt в формате json.
    :return: None
    """
    with open("Data.txt", "w", encoding="utf-8") as file:
        json.dump(BASIC_DATA, file, ensure_ascii=False, indent=0)


def create_data_txt_no_json() -> None:
    """
    Создание файла Data.txt c данными в формате строки.
    :return: None
    """
    with open("Data.txt", "w", encoding="utf-8") as file:
        file.write(str(BASIC_DATA))


def create_data_files() -> None:
    """
    Создание всех необходимых для работв приложения ключевых файлов.
    :return: None
    """
    if not os.path.exists('data'):
        os.mkdir("data")
        for key in FILES:
            wb = Workbook()
            ws = wb.active
            for row in FILES[key]:
                ws.append(row)
            wb.save(BASIC_DATA[key])
        os.makedirs(BASIC_DATA["basic_tables_archive"])
        os.makedirs(BASIC_DATA["sales_representatives_archive"])
        os.makedirs(BASIC_DATA["sales_representatives"])


def rm_data_txt() -> None:
    """
    Удаление файла Data.txt
    :return: None
    """
    if os.path.exists("Data.txt"):
        os.remove("Data.txt")


def rm_all_data() -> None:
    """
    Удаление с содержимым папок data и to_send.
    :return: None
    """
    if os.path.exists("data"):
        shutil.rmtree("data")
    if os.path.exists("to_send"):
        shutil.rmtree("to_send")


def save_data() -> None:
    """
    Перемещение ключевых данных во временную папку для сохранности во время тестов.
    :return: None
    """
    os.mkdir("Tests/temp")
    shutil.move("Data.txt", "Tests/temp/Data.txt")
    shutil.move("data", "Tests/temp/data")
    shutil.move("to_send", "Tests/temp/to_send")


def load_data() -> None:
    """
    Перемещение ключевых данных из временной папки после тестов.
    :return: None
    """
    shutil.move("Tests/temp/Data.txt", "Data.txt")
    shutil.move("Tests/temp/data", "data")
    shutil.move("Tests/temp/to_send", "to_send")
    shutil.rmtree("Tests/temp")


class TestsCheckFiles(unittest.TestCase):
    """
    Тестирование различных модулей приложения.
    """
    # def init(self):
    #     self.start()
    #     self.test_get_paths()
    #     self.test_check_files()
    #     self.test_get_cell_a1()
    #     self.finish()

    def start(self):
        save_data()
        os.chdir("..")

    # def test_get_paths(self) -> None:
    #     """
    #     Тест метода get_path
    #     :return: None
    #     """
    #     # файл не существует и переменная не определена
    #     rm_data_txt()
    #     with self.assertRaises(SystemExit):
    #         check = Cf()
    #     with self.assertRaises(UnboundLocalError):
    #         print(check)
    #
    #     # данные в файле не в формте json
    #     create_data_txt_no_json()
    #
    #     with self.assertRaises(SystemExit):
    #         paths = Cf()
    #     with self.assertRaises(UnboundLocalError):
    #         print(paths)
    #     rm_data_txt()
    #     print("1")

    def test_check_files(self) -> None:
        """
        Тест метода check_files.
         + тест метода get_path при условии корретности данных.
        :return:
        """
        create_data_txt()
        create_data_files()
        self.check = Cf()
        self.assertEqual(self.check.paths, BASIC_DATA)
        print("2")


    def test_get_cell_a1(self) -> None:
        """
        Проверка методов get_cell_a1 и is_valid_data_cell.
        :return: None
        """
        a = Distr(BASIC_DATA)
        self.assertTrue(isinstance(a.date_in_file.value, dt.datetime))
        print("3")

    def finish(self):
        rm_data_txt()
        rm_all_data()
        load_data()


if __name__ == '__main__':
    unittest.main()
