"""
Desktop app. Choose files xlsx, choose consolidated file xlsx.
Check the files for new rows, if row isn't in consolidated file,
append in this consolidated file.
"""
import logging
import sys
from pathlib import Path

import pandas as pd
from PyQt6.QtWidgets import (QApplication, QFileDialog, QWidget,
                             QListWidget, QPushButton, QLabel, QVBoxLayout)

# Logger
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
file_handler = logging.FileHandler('log_gui_app.txt')
file_handler.setLevel(logging.INFO)
file_handler.setFormatter(formatter)
console_handler = logging.StreamHandler()
logger.addHandler(file_handler)
logger.addHandler(console_handler)


class MainWindow(QWidget):
    """Main class for desktop app."""
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        # Title and size of window.
        self.setWindowTitle('Проверить на наличие новых строк в файлах.')
        self.setGeometry(100, 100, 400, 100)

        # Set up vertical layout for window.
        layout = QVBoxLayout()
        self.setLayout(layout)

        # Атрибуты, чтобы сохранить список файлов и путь к консолид. файл
        self.file_list_paths = []
        self.path_to_consolidated = ""

        # all files - button
        button_files_for_check = QPushButton('Поиск файлов для проверки')
        button_files_for_check.clicked.connect(self.open_file_dialog)

        # Виджет список файлов для проверки - чтобы наглядно отображалось в приложении
        self.file_list = QListWidget(self)

        # Consolidated file - button
        button_consolidated_file = QPushButton('Выбрать консолидированный файл')
        button_consolidated_file.clicked.connect(self.open_consolidated_file)

        # Виджет в виде списка для консолидированного файла -
        # чтобы наглядно отображалось в приложении
        self.consolidated_file = QListWidget(self)

        # check the files - button
        button_check_the_files = QPushButton('Осуществить проверку файлов на наличие новых строк')
        button_check_the_files.clicked.connect(self.check_the_rows)

        # Добавляем виджеты в макет
        layout.addWidget(QLabel('Файлы, выбранные для проверки на наличие новых строк'))
        layout.addWidget(self.file_list)
        layout.addWidget(button_files_for_check)
        layout.addWidget(QLabel('Консолидированный файл'))
        layout.addWidget(self.consolidated_file)
        layout.addWidget(button_consolidated_file)
        layout.addWidget(button_check_the_files)
        self.show()

    def open_file_dialog(self):
        filenames, _ = QFileDialog.getOpenFileNames(self, "Select Files")
        if filenames:
            self.file_list.addItems([str(Path(filename)) for filename in filenames])
            self.file_list_paths.extend([str(Path(filename)) for filename in filenames])

    def open_consolidated_file(self):
        filenames, _ = QFileDialog.getOpenFileNames(self, "Select Files")
        if filenames:
            self.consolidated_file.addItems([str(Path(filename)) for filename in filenames])
            self.path_to_consolidated = filenames[0]

    def check_the_rows(self):
        current_main_df = pd.read_excel(self.path_to_consolidated)
        # Цикл по выбранным файлам
        for file in self.file_list_paths:

            df = pd.read_excel(file)
            logger.info(f"{df}")
            # Цикл по строкам файла
            for index, row in df.iterrows():
                if row.to_dict() not in current_main_df.to_dict(orient='records'):
                    print('current', current_main_df)
                    print(row)
                    current_main_df = current_main_df._append(row)
                    print('*************', current_main_df)

        current_main_df.to_excel(self.path_to_consolidated, index=False)
        print("Все строки сохранены в файле ", self.path_to_consolidated)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    sys.exit(app.exec())
