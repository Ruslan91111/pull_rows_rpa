import os
import pandas as pd
import numpy as np
import logging


# Logger
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
fh = logging.FileHandler('./log.txt')
fh.setLevel(logging.INFO)
fh.setFormatter(formatter)
logger.addHandler(fh)


class MakerFiles:
    """Создание xlsx-файлов в папке."""
    def __init__(self, rows: int, cols: int):
        self.rows = rows
        self.cols = cols

    def make_files_in_dir(self, number: int, start_number: int, dir_with_files: str):
        """Создать заданное количество файлов xlsx в папке.
        Принимает два числа: количество необходимых файлов и цифра,
        с которой начинается нумерация файлов."""
        logger.info(f"Starting to create xlsx files.")
        start_number = start_number
        for item in range(number):
            df = pd.DataFrame(np.random.randint(
                low=0, high=100, size=(self.rows, self.cols)),
            )  # columns=list('COLS')
            df.to_excel(f'{dir_with_files}/excel file number {start_number}.xlsx')
            start_number += 1
        logger.info(f"Created {number} xlsx files in directory {dir_with_files}. ")


class CommonFile:
    """Работа с общим файлом xlsx."""

    def __init__(self, path_for_example: str):
        self.example_for_common_file = path_for_example

    def create_and_fill_common_file(self, path_to_files: str,
                                    name_for_common_file: str = 'main.xlsx'):
        """Создать общий файл xlsx, куда подтянуть данные из всех файлов xlsx в папке."""
        current_main_df = pd.read_excel(self.example_for_common_file)
        # Перебираем файлы в папке, читаем добавляем в общий датафрейм.
        for filename in os.listdir(path_to_files):
            path_to_filename = path_to_files + '\\' + filename
            if path_to_filename == self.example_for_common_file:
                continue
            df = pd.read_excel(path_to_files + '/' + filename)
            buf_df = current_main_df._append(df, ignore_index=True)
            current_main_df = buf_df
            logger.info(f"Data from {filename} append to common dataframe.")

        # Записываем общий датафрейм в Excel файл.
        current_main_df.to_excel(name_for_common_file, index=False)
        logger.info(f"Data from xlsx-files dumped in common {name_for_common_file}")


def check_the_files(path_to_files: str, path_to_common: str,):
    """Перебрать файлы в папке, пройтись по строкам в файлах и если
    строки уже есть в общем файле, то пропустить, если нет, то добавить."""
    current_main_df = pd.read_excel(path_to_common)
    logger.info(f"Start to check the xlsx files on new rows")
    for filename in os.listdir(path_to_files):
        logger.info(f"Checking file {filename}.")
        df = pd.read_excel(path_to_files + '/' + filename)
        for index, row in df.iterrows():
            # Проверяем есть ли строка в общем файле, если нет, то добавляем.
            if row.to_dict() not in current_main_df.to_dict(orient='records'):
                current_main_df = current_main_df._append(row)
                logger.info(f"Row {row} append.")

    # Записываем общий датафрейм в Excel файл.
    current_main_df.to_excel(path_to_common, index=False)
    logger.info(f"All rows dumped in {path_to_common}.")


def validate_input(input_text):
    """Проверка введенных пользователем значений. """
    if not os.path.isfile(input_text) and not os.path.isdir(input_text):
        raise ValueError("Путь указан неверно.")


def check_the_dir(dir_with_files):
    """Проверить наличие папки, если нет, то создать."""
    try:
        os.mkdir(dir_with_files)
    except FileExistsError:
        print('dir exists')


def create_files():
    """Создать файлы """
    first_input = input('Желаете создать файлы xlsx? \n'
                        'Введите: Да или Нет \n->')
    if first_input.strip(" ' ,.").lower() == 'да':
        path_to_all_files = input("Введите полный путь до директории, в которой "
                                  "будут находиться xlsx файлы:\n-> ")
        number_of_rows = int(input("Введите число - количество строк, которые"
                                   " необходимо создать в xlsx файлах:\n-> "))
        number_of_cols = int(input("Введите число -  количество колонок которые"
                                   " необходимо создать в xlsx файлах:\n-> "))
        number_of_files = int(input("Введите число - количество xlsx файлов,"
                                    " которые необходимо создать:\n-> "))
        start_number_for_file = int(input("Введите число - с которого начнется "
                                          "нумерация xlsx файлов:\n-> "))

        check_the_dir(path_to_all_files)
        maker_file = MakerFiles(number_of_rows, number_of_cols)
        maker_file.make_files_in_dir(number_of_files, start_number_for_file, path_to_all_files)
        return


def create_common_file():
    """Создать общий файл, куда будут подтягиваться строки из других файлов."""
    first_input = input('Желаете создать общий (Консолидированный) файл xlsx? '
                        'Введите: Да или Нет\n->  ')
    if first_input.strip(" ' ,.").lower() == 'да':
        path_to_example_for_main = input("Введите полный путь до xlsx файла: "
                                         "который будет взят за образец, "
                                         "в том числе название файла с расширением:\n-> ")
        validate_input(path_to_example_for_main)
        name_for_common_file = input("Введите название создаваемого "
                                     "общего консолидированного файла\n-> ")
        path_to_dir = input("Введите полный путь до директории с xlsx файлами:\n-> ")
        validate_input(path_to_dir)
        common_file = CommonFile(path_to_example_for_main)
        common_file.create_and_fill_common_file(path_to_dir, name_for_common_file)


def check_rows_in_files():
    path_to_main = input("Введите полный путь до общего xlsx файла, в который\n"
                         "будут сохраняться строки из xlsx файлов менеджеров:\n-> ")
    validate_input(path_to_main)
    path_to_dir = input("Введите полный путь до директории с xlsx файлами: ")
    validate_input(path_to_dir)
    check_the_files(path_to_dir, path_to_main)


if __name__ == '__main__':
    create_files()
    create_common_file()
    check_rows_in_files()
