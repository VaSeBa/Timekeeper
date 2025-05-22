import sys
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QPushButton, QFileDialog, QLabel,
                             QProgressBar, QTextEdit, QMessageBox)
from PyQt6.QtCore import Qt, QThread, pyqtSignal


class ExcelComparator(QMainWindow):
    """Главное окно приложения"""

    def __init__(self):
        super().__init__()
        self.initUI()
        self.file1_path = None  # Путь к первому файлу
        self.file2_path = None  # Путь ко второму файлу

    def initUI(self):
        """Инициализация пользовательского интерфейса"""
        self.setWindowTitle("Сравнение Excel-файлов")
        self.setGeometry(300, 300, 600, 400)

        # Центральный виджет и основной лейаут
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # Блок выбора первого файла
        file1_layout = QHBoxLayout()
        self.btn_file1 = QPushButton("Выбрать файл 1", self)
        self.btn_file1.clicked.connect(lambda: self.select_file(1))
        self.label_file1 = QLabel("Файл 1 не выбран")
        file1_layout.addWidget(self.btn_file1)
        file1_layout.addWidget(self.label_file1)

        # Блок выбора второго файла
        file2_layout = QHBoxLayout()
        self.btn_file2 = QPushButton("Выбрать файл 2", self)
        self.btn_file2.clicked.connect(lambda: self.select_file(2))
        self.label_file2 = QLabel("Файл 2 не выбран")
        file2_layout.addWidget(self.btn_file2)
        file2_layout.addWidget(self.label_file2)

        # Добавление блоков в основной лейаут
        layout.addLayout(file1_layout)
        layout.addLayout(file2_layout)

        # Прогресс-бар
        self.progress = QProgressBar(self)
        self.progress.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.progress)

        # Кнопка запуска сравнения
        self.btn_compare = QPushButton("Сравнить файлы", self)
        self.btn_compare.clicked.connect(self.start_comparison)
        layout.addWidget(self.btn_compare)

        # Текстовое поле для лога
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        layout.addWidget(self.log)

        # Строка состояния
        self.statusBar().showMessage('Готово')

    def select_file(self, file_num):
        """Выбор файла через диалоговое окно"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Выберите файл", "", "Excel Files (*.xlsx *.xls)")

        if file_path:
            if file_num == 1:
                self.file1_path = file_path
                self.label_file1.setText(file_path.split('/')[-1])
            else:
                self.file2_path = file_path
                self.label_file2.setText(file_path.split('/')[-1])

    def start_comparison(self):
        """Запуск процесса сравнения в отдельном потоке"""
        if not self.file1_path or not self.file2_path:
            QMessageBox.warning(self, "Ошибка", "Выберите оба файла!")
            return

        # Создание и настройка рабочего потока
        self.worker = CompareWorker(self.file1_path, self.file2_path)
        # Подключение сигналов потока к слотам GUI
        self.worker.progress_updated.connect(self.update_progress)
        self.worker.message_received.connect(self.log_message)
        self.worker.finished.connect(self.on_completion)
        self.worker.error_occurred.connect(self.show_error)
        self.worker.start()

        # Блокировка кнопки во время выполнения
        self.btn_compare.setEnabled(False)
        self.progress.setValue(0)
        self.log.clear()

    def update_progress(self, value):
        """Обновление прогресс-бара"""
        self.progress.setValue(value)

    def log_message(self, message):
        """Добавление сообщения в лог"""
        self.log.append(message)

    def show_error(self, message):
        """Обработка ошибок"""
        QMessageBox.critical(self, "Ошибка", message)
        self.btn_compare.setEnabled(True)

    def on_completion(self, output_path):
        """Действия по завершении обработки"""
        self.btn_compare.setEnabled(True)
        self.progress.setValue(100)
        QMessageBox.information(self, "Готово",
                                f"Файл отчёта сохранён как:\n{output_path}")


class CompareWorker(QThread):
    """Рабочий поток для обработки файлов"""
    progress_updated = pyqtSignal(int)  # Сигнал обновления прогресса
    message_received = pyqtSignal(str)  # Сигнал для сообщений в лог
    finished = pyqtSignal(str)  # Сигнал завершения работы
    error_occurred = pyqtSignal(str)  # Сигнал об ошибке

    def __init__(self, file1_path, file2_path):
        super().__init__()
        self.file1_path = file1_path
        self.file2_path = file2_path

    def run(self):
        """Основной метод выполнения задачи"""
        try:
            self.message_received.emit("Начало обработки файлов...")

            # Загрузка данных с преобразованием в строки
            df1 = pd.read_excel(self.file1_path, dtype=str)
            df2 = pd.read_excel(self.file2_path, dtype=str)
            self.progress_updated.emit(10)
            self.message_received.emit("Файлы загружены")

            # Проверка наличия всех необходимых столбцов
            required_columns = ['id', 'ФИО', 'должность'] + [str(i) for i in range(1, 32)]
            for col in required_columns:
                if col not in df1.columns:
                    raise ValueError(f"Столбец '{col}' отсутствует в первом файле")
                if col not in df2.columns:
                    raise ValueError(f"Столбец '{col}' отсутствует во втором файле")
            self.progress_updated.emit(20)

            # Подготовка данных для сравнения
            df1['original_index'] = df1.index  # Сохраняем оригинальные индексы
            df2['original_index'] = df2.index
            merged_df = pd.merge(df1, df2, on='id', suffixes=('_file1', '_file2'))

            report_data = []  # Данные для отчета
            highlight_cells = []  # Ячейки для подсветки

            total_rows = len(merged_df)
            for idx, (_, row) in enumerate(merged_df.iterrows()):
                original_row = int(row['original_index_file1']) + 2  # Строка в Excel
                for day in range(1, 32):
                    col_name = str(day)
                    val1 = row[f'{col_name}_file1'] or ''  # Обработка NaN и None
                    val2 = row[f'{col_name}_file2'] or ''

                    # Сравнение значений
                    if str(val1) != str(val2):
                        report_data.append([
                            row['id'],
                            row['ФИО_file1'],
                            day,
                            str(val1),
                            str(val2)
                        ])
                        # Определение позиции ячейки для подсветки
                        col_idx = df1.columns.get_loc(col_name) + 1
                        highlight_cells.append((original_row, col_idx))

                # Обновление прогресса
                progress = 20 + int(70 * (idx + 1) / total_rows)
                self.progress_updated.emit(progress)

            # Сохранение результатов
            self.message_received.emit("Сохранение результатов...")
            wb = load_workbook(self.file1_path)
            ws = wb.active

            # Подсветка несовпадающих ячеек
            yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
            for row, col in highlight_cells:
                ws.cell(row=row, column=col).fill = yellow_fill

            # Создание листа отчета
            if 'Отчет' in wb.sheetnames:
                del wb['Отчет']
            report_sheet = wb.create_sheet('Отчет')
            report_headers = ['ID', 'ФИО', 'День', 'Файл 1', 'Файл 2']
            report_sheet.append(report_headers)

            # Заполнение отчета данными
            for entry in report_data:
                report_sheet.append(entry)

            # Автоматическая настройка ширины столбцов
            for column in report_sheet.columns:
                max_length = 0
                for cell in column:
                    try:
                        max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                adjusted_width = (max_length + 2)
                report_sheet.column_dimensions[cell.column_letter].width = adjusted_width

            # Сохранение файла
            output_path = self.file1_path.replace('.xlsx', '_сравнение.xlsx')
            wb.save(output_path)

            self.finished.emit(output_path)
            self.message_received.emit("Обработка завершена успешно!")

        except Exception as e:
            self.error_occurred.emit(str(e))  # Отправка ошибки в GUI


if __name__ == '__main__':
    # Создание и запуск приложения
    app = QApplication(sys.argv)
    ex = ExcelComparator()
    ex.show()
    sys.exit(app.exec())
