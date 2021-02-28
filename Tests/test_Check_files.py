import unittest
from shutil import copyfile, rmtree
from openpyxl import load_workbook, Workbook
import pandas as pd
import random
import datetime as dt
import json
import os
import time

import Check_files as cf


BASIC_DATA = {"Distributors": r"data\Дистрибьюторыю.xlsx",
              "Contractors": r"data\Контрагенты.xlsx",
              "Reports": r"data\Отчеты.xlsx",
              "Database": r"data\pearl.xlsx"}


def create_files_for_test():
    try:
        os.mkdir('data')
    except OSError:
        print('Не удалось создать директорию data')
    for key in BASIC_DATA:
        wb = Workbook()
        wb.save(BASIC_DATA[key])


def del_files_after_test():
    rmtree('data')


class TestsCheckFiles(unittest.TestCase):

    def test_get_paths(self):
        create_files_for_test()
        copyfile('../Data.txt', 'data/Temporary.txt')
        with open('../Data.txt', 'w', encoding='utf-8') as file:
            json.dump(BASIC_DATA, file, ensure_ascii=False)



        copyfile('data/Temporary.txt', '../Data.txt')
        del_files_after_test()


