"""
Работа с электронной почтой посредством Outlook.
"""
# install pywin32
import os
import datetime as dt
from shutil import rmtree
import win32com.client


class Email:

    def __init__(self, debtors_email):
        self.outlook = win32com.client.Dispatch('outlook.application').GetNamespace("MAPI")
        self.inbox = self.outlook.GetDefaultFolder(6)
        self.messages = None
        self.output_dir = 'email_files'     # название папки для скачанных файлов

        received_dt = dt.datetime.today().replace(day=1, hour=0, minute=0, second=0)
        self.received_dt = received_dt.strftime("%m/%d/%Y %H:%M %p")

        self.run(debtors_email)

    def reader(self) -> None:
        """
        Функция сохраняет прикрепленный файл на диск.
        :return: None
        """
        try:
            for message in list(self.messages):
                try:
                    for attachment in message.Attachments:
                        # выявлена проблема в Outlook - прикрепленный файл имеет расширение .xlsx
                        # Это видно в любом браузере, но в Outlook он имеет расширение .xls_
                        new_filename = attachment.FileName[:-1] + 'x'  # временная заплатка
                        path = os.path.join(os.getcwd(), self.output_dir, new_filename)
                        attachment.SaveAsFile(path)
                except Exception as except_:
                    print("Ошибка на этапе сохранения файла - " + str(except_))
        except Exception as except_:
            print("Ошибка на этапе обработки email - " + str(except_))

    def email_filter(self, debtor_email: str) -> None:
        """
        Функция осуществляет получение и фильтрацию сообщений.
        Сюда выведены все фильтры для удобства дальнейшей работы с кодом.
        :param debtor_email: Email дистрибьютера, от которого ожидается сообщение.
        :return: None
        """
        self.messages = self.inbox.Items
        self.messages = self.messages.Restrict(f"[ReceivedTime] >= '{self.received_dt}'")
        self.messages = self.messages.Restrict(f"[SenderEmailAddress] = '{debtor_email}'")

    def run(self, debtors_email) -> None:
        """
        Функция запускает фильтрацию emails и скачивание данных для каждого дистрибьютера,
        от которого ожидается доклад.
        :param debtors_email: Список email-ов адресатов, чьи письма нужно скачать.
        :return: None
        """
        if os.path.exists("email_files"):
            rmtree("email_files")
        os.mkdir('email_files')

        for email in debtors_email:
            self.email_filter(email)
            self.reader()
