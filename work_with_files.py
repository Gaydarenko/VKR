import os
import openpyxl


class BasicTable:

    def __init__(self, path: str):
        self.dir_path = 'email_files'   # путь с скачанным файлам
        self.basic_table_path = path    # путь к базовой таблице pearl.xlsx
        self.workbook = openpyxl.load_workbook(self.basic_table_path)
        self.basic_table = self.workbook.active
        self.column_names_list = list()
        self.downloaded_files = list()
        self.data_for_write_dict = dict()

        self.form_column_names_dict()
        self.get_downloaded_files()
        self.workbook.save(r"data\pearl")   # Расширение файла не указано специально - временная мера,
                                            # чтобы пока не заменять исходный файл

    def form_column_names_dict(self) -> None:
        """
        Формирование словаря с названиями столбуов таблицы.
        Далее копия этого словаря будет использоваться для формирования набора данных.
         для записи в главную таблицу.
        :return: None
        """
        self.column_names_list = [self.basic_table.cell(row=1, column=j).value
                                  for j in range(1, self.basic_table.max_column + 1)]
        self.data_for_write_dict = {key: None for key in self.column_names_list}
        # print(self.column_names_dict)

    def get_downloaded_files(self) -> None:
        """
        Перебираются все скачанные файлы. Для каждого файла запускается метод form_data_for_basic_table.
        :return: None
        """
        self.downloaded_files = list(os.walk(self.dir_path))[0][2]
        # print(self.downloaded_files)
        for file in self.downloaded_files:
            self.form_data_for_basic_table(file)

    def form_data_for_basic_table(self, source_file: str) -> None:
        """
        Формируется словарь с данными, которые необходимо записать в главную таблицу,
         и вызывается метод записывающий эти данные (write_to_basic_table).
        В случае отсутствия каких-нибудь необходимых данных, запись в главную таблицу производится,
         но при этом производится запись в файл отчетов.
        :param source_file: Строка с именем файла.
        :return: None
        """
        wb_source_file = openpyxl.load_workbook(os.path.join(self.dir_path, source_file))
        table = wb_source_file.active
        column_names = [table.cell(row=1, column=j).value for j in range(1, table.max_column + 1)]
        for i in range(2, table.max_row + 1):
            data_for_write = self.data_for_write_dict.copy()
            for column in data_for_write:
                try:
                    data_for_write[column] = table.cell(row=i, column=column_names.index(column)+1).value
                except ValueError:
                    data_for_write[column] = '-----'
                    # add_note_to_reports()
            self.write_to_basic_table(data_for_write)

    def write_to_basic_table(self, data: dict) -> None:
        """
        Производится запись данных в файл основной таблицы в самый конец
        :param data: Словарь данных для записи в основную таблицу. Ключами выступают название столбцов.
        :return: None
        """
        start_row = self.basic_table.max_row
        for j in range(1, len(self.column_names_list) + 1):
            self.basic_table.cell(row=start_row + 1, column=j).value = data[self.column_names_list[j-1]]

#
# if __name__ == '__main__':
#     a = BasicTable('data/pearl.xlsx')
#     print(a)
