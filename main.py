from Check_files import CheckFiles, Distributors
from email import Email


checks = CheckFiles()   # получение путей и проверка наличия необходимых файлов
distributor_path = checks.paths['Distributors']     # получение пути к файлу с информацией о дистрибьюторам
distributors = Distributors(distributor_path)   # анализ содержимого фйла с информацией о дистрибьютерах
debtors = distributors.debtors  # формирование списка интересующих дистрибьютеров
# print(debtors)
email = Email(debtors)  # скачивание прикрепленных файлов
