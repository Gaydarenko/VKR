import os
import openpyxl


class BasicTable:

    def __init__(self, path: str):
        # self.dir_path = 'email_files'
        self.basic_table_path = path    # путь к базовой таблице pearl.xlsx
        # self.files_path = []
        # self.data_for_write_1_row = {}
        self.column_names_dict = dict()

        # self.start()
        self.form_column_names_dict()

    def form_column_names_dict(self) -> None:
        """
        Метод формирует словарь с названиями столбуов таблицы.
        Далее этот копия этого словаря будет использоваться для формирования набора данных
         для записи в главную таблицу.
        :return: None
        """
        self.workbook = openpyxl.load_workbook(self.basic_table_path)
        table = self.workbook.active
        self.column_names_dict = {table.cell(row=1, column=i).value: None for i in range(1, table.max_column + 1)}
        # print(self.column_names_dict)

    def start(self) -> None:
        """
        Стартовый метод, в котором все по порядку пишется.
        Потом раскидывается по другим методам.
        :return: None
        """
        self.files_path = list(os.walk(self.dir_path))[0]
        for file in self.files_path[2]:
            ...


    def get_table_from_file(self):
        # return os.walk(self.dir_path)
        ...

    def is_valid_data(self):
        ...


if __name__ == '__main__':
    a = BasicTable('email_files/шАПКА.xlsx')
    print(a)
    # b = a.get_table_from_file()
    # print(list(b))
    # filelist = list(os.walk('email_files'))[0]
    # print(filelist)
    # for file in filelist[3]:


