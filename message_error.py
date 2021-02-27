"""
Всё для вывода сообщений с ошибками
"""
import tkinter as tk
from sys import exit as sys_exit


class MessageError:

    def __init__(self, text: str):
        self.text = text

    def message_window(self):
        """
        Вывод окна ошибки с переданным текстом проблемы и прекратит выполнение программы.
        :return:
        """
        window = tk.Tk()
        window.title("Ошибка!!!")
        label = tk.Label(text=f"{self.text}", height=2, width=40)
        label.pack()
        window.mainloop()
        sys_exit()
