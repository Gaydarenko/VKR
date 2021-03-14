"""
Программа выполненая в объеме ВКР
"""
from email_ import Email
from check_files import CheckFiles
from distributors import Distributors
from work_with_files import BasicTable
from message_for_user import ProgressReport

checks = CheckFiles()  # получение путей и проверка наличия необходимых файлов
distributor_path = checks.paths['Distributors']  # получение пути к файлу с информацией о дистрибьюторам
distributors = Distributors(distributor_path)  # анализ содержимого фйла с информацией о дистрибьютерах
debtors = distributors.debtors  # формирование списка интересующих дистрибьютеров
email = Email(debtors)  # скачивание прикрепленных файлов
Distributors.set_month_in_file(distributor_path)  # запись в файл текущей даты
basic_table = BasicTable(checks.paths)
status_data = Distributors.form_status_data(distributor_path)
message = ProgressReport(status_data)
