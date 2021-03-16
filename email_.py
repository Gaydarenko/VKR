"""
Работа с электронной почтой посредством Outlook.
"""
import os
import datetime as dt
from shutil import rmtree
import win32com.client

from message_for_user import MessageError as Me
from work_with_files import OtherFiles


class Email:
    """
    Работа с электронными письмами.
    """

    def __init__(self, debtors_email):
        self.outlook = win32com.client.Dispatch("outlook.application").GetNamespace("MAPI")
        self.inbox = self.outlook.GetDefaultFolder(6)
        self.messages = None
        self.output_dir = "email_files"  # название папки для скачанных файлов

        received_dt = dt.datetime.today().replace(day=1, hour=0, minute=0, second=0, microsecond=0)
        self.received_dt = received_dt.strftime("%m/%d/%Y %H:%M %p")

        self.run(debtors_email)

    def reader(self, email: str) -> None:
        """
        Функция сохраняет прикрепленный файл на диск.
        :param email: Строка с email, который используется как имя файла.
        :return: None
        """
        try:
            for message in list(self.messages):
                try:
                    for attachment in message.Attachments:
                        path = os.path.join(os.getcwd(), self.output_dir, email)
                        attachment.SaveAsFile(path)
                except Exception as except_:
                    print("Ошибка на этапе сохранения файла - " + str(except_))
                    Me.message_window("Ошибка во время работы с email на этапе сохранения файла")
        except Exception as except_:
            print("Ошибка на этапе обработки писем - " + str(except_))
            Me.message_window("Ошибка во время работы с email на этапе обработки писем")

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
            self.reader(email + ".xlsx")

    @staticmethod
    def sender(paths: dict) -> None:
        """
        Формирование писем в черновиках с прикрепленными файлами.
        :param paths:
        :return:
        """
        path = paths["sales_representatives"]
        path_kam = paths["KAM"]
        files = list(os.walk(path))[0][2]
        for file in files:
            recipient = OtherFiles.sales_representatives(path_kam, file)
            if recipient:
                outlook = win32com.client.Dispatch("outlook.application")
                mail = outlook.CreateItem(0)
                mail.To = recipient
                mail.Subject = "Доклад за прошлый месяц."
                mail.Attachments.Add(os.path.join(os.getcwd(), path, file))
                mail.Save()
