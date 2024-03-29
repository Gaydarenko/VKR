import os
import unittest

from check_files import CheckFiles as Cf
from Tests.data_for_tests import PATHS, create_data_txt, create_data_txt_no_json,\
    create_data_files, rm_data_txt, rm_all_data, save_data, load_data


class TestsCheckFiles(unittest.TestCase):
    """
    Тестирование различных модулей приложения.
    """
    # def init(self):
    #     self.start()
    #     self.test_check_files()
    #     self.finish()

    @classmethod
    def setUpClass(cls) -> None:
        """
        Перенос ключевых файлов в безопасное место
         и формирование замены этих файлов стестовыми данными.
        :return: None
        """
        os.chdir("..")
        save_data()
        print("c0")

    # После выполнения конкретно этого метода работа класса заканчивается несмотря даже на init.
    # def test_get_paths(self) -> None:
    #     """
    #     Тест метода get_path
    #     :return: None
    #     """
    #     # файл не существует и переменная не определена
    #     rm_data_txt()
    #     with self.assertRaises(SystemExit):
    #         check = Cf()
    #     with self.assertRaises(UnboundLocalError):
    #         print(check)
    #
    #     # данные в файле не в формте json
    #     create_data_txt_no_json()
    #
    #     with self.assertRaises(SystemExit):
    #         paths = Cf()
    #     with self.assertRaises(UnboundLocalError):
    #         print(paths)
    #     rm_data_txt()
    #     print("c1")

    def test_check_files(self) -> None:
        """
        Тест метода check_files.
         + тест метода get_path при условии корретности данных.
        :return:
        """
        create_data_txt()
        create_data_files()
        self.check = Cf()
        self.assertEqual(self.check.paths, PATHS)
        print("c2")

    @classmethod
    def tearDownClass(cls) -> None:
        """
        Удаление тестовых файлов и перемещение обратно рабочих файлов.
        :return: None
        """
        rm_data_txt()
        rm_all_data()
        load_data()
        print("c99")


if __name__ == '__main__':
    unittest.main()
