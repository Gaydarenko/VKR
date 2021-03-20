"""
Программа выполненная в объеме ВКР
"""
from email_ import Email
from check_files import CheckFiles
from distributors import Distributors
from work_with_files import BasicTable
from message_for_user import ProgressReport


class Service:
    """
    Обеспечивает работу без использования графического интерфейса.
    """
    @staticmethod
    def main():
        checks = CheckFiles()  # получение путей и проверка наличия необходимых файлов
        distributors = Distributors(checks.paths)  # анализ содержимого файла с информацией о дистрибьютерах
        debtors = distributors.debtors  # получение списка интересующих дистрибьютеров
        email = Email(debtors)  # скачивание прикрепленных файлов
        distributor_path = checks.paths['Distributors']  # получение пути к файлу с информацией о дистрибьюторам
        Distributors.set_month_in_file(distributor_path)  # запись в файл текущей даты
        basic_table = BasicTable(checks.paths)  # запись данных в основную таблицу и в отчеты
        colors = basic_table.distributor_color  # получение словаря цветовых статусов для дистрибьютеров
        Distributors.coloring(distributor_path, colors)  # изменение цвета заливки для указанных дистрибьютеров
        status_data = Distributors.form_status_data(
            distributor_path)  # получение словаря с данными о ходе выполнения общей задачи
        BasicTable.form_report_for_sr(checks.paths, status_data)  # формирование докладов для торговых представителей
        Email.sender(checks.paths)  # подготовка черновиков с прикрепленными докладами для торговых представителей
        ProgressReport.progress_window(status_data)  # вывод окна с информацией о прогрессе за текущий месяц


if __name__ == '__main__':
    Service.main()
