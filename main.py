"""
Программа выполненая в объеме ВКР
"""
from email_ import Email
from Check_files import CheckFiles, Distributors
from work_with_files import BasicTable


checks = CheckFiles()   # получение путей и проверка наличия необходимых файлов
distributor_path = checks.paths['Distributors']     # получение пути к файлу с информацией о дистрибьюторам
distributors = Distributors(distributor_path)   # анализ содержимого фйла с информацией о дистрибьютерах
debtors = distributors.debtors  # формирование списка интересующих дистрибьютеров
# print(debtors)
email = Email(debtors)  # скачивание прикрепленных файлов
distributors.set_month_in_file()    # запись в файл текущей даты

