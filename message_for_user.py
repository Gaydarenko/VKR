"""
Всё для вывода сообщений с ошибками
"""
import tkinter as tk
from sys import exit as sys_exit


class MessageError:

    # def __init__(self, text: str):
    #     self.text = text

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

    def __init__(self, data: dict):
        self.data = data
        self.progress_window()

    def progress_window(self):
        window = tk.Tk()
        window.title("Обобщенная информация")
        label = tk.Label(text="За текущий месяц:")
        label.pack(padx=10, anchor='w')
        for status in self.data:
            percent = round(self.data[status] / self.data['Всего'] * 100, 2)
            label = tk.Label(text=f"   - {status}: {self.data[status]} ({percent}%)")
            label.pack(padx=15, anchor="w")
        window.mainloop()
