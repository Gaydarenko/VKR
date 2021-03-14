"""
Работа с файлом дистрибьютеров
"""
import datetime as dt
from openpyxl import load_workbook

from message_for_user import MessageError as Me


class Distributors:
    """
    Работа с файлом дистрибьютеров
    """

    def __init__(self, path: str):
        self.date_in_file = None
        self.month = None
        self.debtors = []
        self.path = path
        self.wb_distributors = load_workbook(self.path)
        self.distributors_table = self.wb_distributors.active

        self.basic_run()

    def basic_run(self) -> None:
        """
        Порядок запуска методов при основном обращении
        :return: None
        """
        self.get_cell_a1()
        self.is_valid_data_cell()
        self.check_month_in_file()
        self.get_debtors()

    def get_cell_a1(self) -> None:
        """
        Получение содержимого ячейки А1 из файла формата xlsx.
        :return: None
        """
        self.date_in_file = self.distributors_table.cell(row=1, column=1)

    def is_valid_data_cell(self) -> None:
        """
        Проверка первой ячейки таблицы. Содержимое ячейки должно быть в формате даты.
        :return: None
        """
        if not isinstance(self.date_in_file.value, dt.datetime):
            Me.message_window('В файле Дистрибьютеры.xlsx в ячейке А1 '
                              'отсутствует дата в нужном формате (ДД.ММ.ГГГГ).')

    def get_debtors(self) -> None:
        """
        Формирование списка "должников".
        Критерием является незакрашенность ячейки с именем дистрибьютера.
        :return: None
        """
        for i in range(2, self.distributors_table.max_row + 1):
            if self.distributors_table.cell(row=i, column=1).fill.fgColor.value in ["00FFFFFF", "00000000", 0, "FFFFFFFF"]:
                self.debtors.append(self.distributors_table.cell(row=i, column=2).value)

    def check_month_in_file(self) -> None:
        """
        Сравнение указанного в файле месяца с текущим.
        Если не совпадает, то закрасить весь файл в белый
        :return: None
        """
        current_month = dt.date.today().month
        month_in_file = self.date_in_file.value.month
        if month_in_file != current_month:

            for cell_row in self.distributors_table["A2": f"A{self.distributors_table.max_row + 1}"]:
                cell_row[0].fill.fgColor.value = '00FFFFFF'

            self.wb_distributors.save(self.path)

    @staticmethod
    def set_month_in_file(path) -> None:
        """
        Запись текущей даты в файл.
        :param path: путь к файлу с таблицей дистрибьютеров
        :return: None
        """
        wb_distributor = load_workbook(path)
        distributor_table = wb_distributor.active
        distributor_table.cell(row=1, column=1).value = dt.datetime.today()
        wb_distributor.save(path)

    @staticmethod
    def form_status_data(path: str) -> dict:
        """
        Формирование словаря с данными о ходе выполнения общей задачи.
        :param path: путь к файлу с таблицей дистрибьютеров
        :return: словарь с данными
        """
        statuses = ["не прислали доклад",
                    "доклад без замечаний",
                    "требует участия человека",
                    "некорретных докладов",
                    "невозможно определить статус",
                    "Всего"]
        progress = {status: 0 for status in statuses}

        wb_distributor = load_workbook(path)
        distributor_table = wb_distributor.active

        for i in range(2, distributor_table.max_row + 1):
            color = distributor_table.cell(row=i, column=1).fill.fgColor.value
            if color in ["00FFFFFF", "00000000", 0, "FFFFFFFF"]:
                status = 0
            elif color in ["FF92D050", ]:
                status = 1
            elif color in ["FFFFC000", ]:
                status = 2
            elif color in ["FFFF0000", ]:
                status = 3
            else:
                status = 4
            progress[statuses[status]] += 1

        progress[statuses[5]] = distributor_table.max_row - 1
        return progress
