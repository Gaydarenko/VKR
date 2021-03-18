import os
import shutil
import json
from openpyxl import Workbook
import datetime as dt


PATHS = {
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
               ["2016", "12", "2016-12", "АНРО", "Анима", "Санкт-Петербург г", "РФ", "Северо-Запад", "6802420243", "D", "2", "3", "Санкт-Петербург, Грибалевой 7", "АнимаТрейд", "Vet", "Прочие", "Pugach", "Nobivac Rabies 10x1ds", "153698", "Rabies (Alu)", "Bio", "0.5", "5", "375"],
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
        json.dump(PATHS, file, ensure_ascii=False, indent=0)


def create_data_txt_no_json() -> None:
    """
    Создание файла Data.txt c данными в формате строки.
    :return: None
    """
    with open("Data.txt", "w", encoding="utf-8") as file:
        file.write(str(PATHS))


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
            wb.save(PATHS[key])
        os.makedirs(PATHS["basic_tables_archive"])
        os.makedirs(PATHS["sales_representatives_archive"])
        os.makedirs(PATHS["sales_representatives"])


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
    os.makedirs("Tests/temp")
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