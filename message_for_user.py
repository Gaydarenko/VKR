"""
Всё для вывода сообщений пользователю.
"""
import tkinter as tk
from sys import exit as sys_exit


class MessageError:
    """
    Выод сообщения об ошибке
    """

    @staticmethod
    def message_window(text):
        """
        Вывод окна ошибки с переданным текстом проблемы и прекратит выполнение программы.
        :return:
        """
        window = tk.Tk()
        window.title("Ошибка!!!")
        label = tk.Label(text=f"{text}", height=2, width=40)
        label.pack()
        window.mainloop()
        sys_exit()


class ProgressReport:
    """
    Формирование и вывод сообщения для пользователя по итогам текущего месяца.
    """

    @staticmethod
    def progress_window(data) -> None:
        """
        Вывод окна с переданными данными.
        :return: None
        """
        window = tk.Tk()
        window.title("Обобщенная информация")
        label = tk.Label(text="За текущий месяц:")
        label.pack(padx=10, anchor='w')
        for status in data:
            percent = round(data[status] / data['Всего'] * 100, 2)
            label = tk.Label(text=f"   - {status}: {data[status]} ({percent}%)")
            label.pack(padx=15, anchor="w")
        window.mainloop()
