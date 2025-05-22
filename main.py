import sys
import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QPushButton, QFileDialog, QLabel,
                             QProgressBar, QTextEdit, QMessageBox)
from PyQt6.QtCore import Qt, QThread, pyqtSignal


class ExcelComparator(QMainWindow):
    """Главное окно приложения для сравнения Excel-файлов"""

    def __init__(self):
        super().__init__()
        self.file1_path = None
        self.file2_path = None
        self.current_worker = None
        self.init_ui()

    def init_ui(self):
        """Инициализация пользовательского интерфейса"""
        self.setWindowTitle("Сравнение Excel-файлов")
        self.setMinimumSize(800, 600)
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # Панели выбора файлов
        self.init_file_selectors(layout)

        # Элементы управления
        self.init_controls(layout)

        # Лог и прогресс
        self.init_progress_log(layout)
        self.statusBar().showMessage('Готово к работе')

    def init_file_selectors(self, layout):
        """Инициализация компонентов выбора файлов"""
        file1_layout = QHBoxLayout()
        self.btn_file1 = QPushButton("Выбрать базовый файл", self)
        self.btn_file1.clicked.connect(lambda: self.select_file(1))
        self.label_file1 = QLabel("Файл не выбран", self)
        file1_layout.addWidget(self.btn_file1)
        file1_layout.addWidget(self.label_file1)
        file2_layout = QHBoxLayout()
        self.btn_file2 = QPushButton("Выбрать файл для сравнения", self)
        self.btn_file2.clicked.connect(lambda: self.select_file(2))
        self.label_file2 = QLabel("Файл не выбран", self)
        file2_layout.addWidget(self.btn_file2)
        file2_layout.addWidget(self.label_file2)
        layout.addLayout(file1_layout)
        layout.addLayout(file2_layout)

    def init_controls(self, layout):
        """Инициализация элементов управления"""
        control_layout = QHBoxLayout()
        self.btn_compare = QPushButton("Сравнить файлы", self)
        self.btn_compare.clicked.connect(self.start_comparison)
        self.btn_abort = QPushButton("Прервать", self)
        self.btn_abort.clicked.connect(self.abort_processing)
        self.btn_abort.setEnabled(False)
        control_layout.addWidget(self.btn_compare)
        control_layout.addWidget(self.btn_abort)
        layout.addLayout(control_layout)

    def init_progress_log(self, layout):
        """Инициализация прогресс-бара и лога"""
        self.progress = QProgressBar(self)
        self.progress.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.progress)
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        layout.addWidget(self.log)

    def select_file(self, file_num):
        """Выбор файла через диалоговое окно"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите файл",
            "",
            "Excel Files (*.xlsx *.xls *.xlsm);;All Files (*)"
        )

        if file_path:
            try:
                pd.read_excel(file_path, nrows=1)
                if file_num == 1:
                    self.file1_path = file_path
                    self.label_file1.setText(os.path.basename(file_path))
                else:
                    self.file2_path = file_path
                    self.label_file2.setText(os.path.basename(file_path))
            except Exception as e:
                QMessageBox.critical(self, "Ошибка",
                                     f"Невозможно прочитать файл:\n{str(e)}")

    def start_comparison(self):
        """Запуск процесса сравнения"""
        if not all([self.file1_path, self.file2_path]):
            QMessageBox.warning(self, "Ошибка", "Необходимо выбрать оба файла!")
            return

        if self.file1_path == self.file2_path:
            QMessageBox.warning(self, "Ошибка", "Файлы должны быть разными!")
            return

        self.current_worker = CompareWorker(
            self.file1_path,
            self.file2_path
        )

        self.current_worker.progress_updated.connect(self.update_progress)
        self.current_worker.message_received.connect(self.log_message)
        self.current_worker.finished.connect(self.on_completion)
        self.current_worker.error_occurred.connect(self.show_error)
        self.current_worker.start()
        self.toggle_controls(False)

    def abort_processing(self):
        """Прерывание выполнения операции"""
        if self.current_worker and self.current_worker.isRunning():
            self.current_worker.terminate()
            self.log_message("Процесс прерван пользователем")
            self.toggle_controls(True)

    def toggle_controls(self, enable):
        """Переключение состояния элементов управления"""
        self.btn_compare.setEnabled(enable)
        self.btn_abort.setEnabled(not enable)
        self.progress.setValue(0)
        self.log.clear()

    def update_progress(self, value):
        """Обновление прогресс-бара"""
        self.progress.setValue(value)

    def log_message(self, message):
        """Добавление сообщения в лог с временной меткой"""
        from datetime import datetime
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log.append(f"[{timestamp}] {message}")

    def show_error(self, message):
        """Обработка ошибок"""
        QMessageBox.critical(self, "Критическая ошибка", message)
        self.toggle_controls(True)

    def on_completion(self, output_path):
        """Действия по завершении обработки"""
        self.toggle_controls(True)
        if output_path:
            QMessageBox.information(self, "Готово", f"Результаты сравнения сохранены в:\n{output_path}")
            self.progress.setValue(100)


class CompareWorker(QThread):
    """Поток для сравнения Excel-файлов"""
    progress_updated = pyqtSignal(int)
    message_received = pyqtSignal(str)
    finished = pyqtSignal(str)
    error_occurred = pyqtSignal(str)

    def __init__(self, file1_path, file2_path):
        super().__init__()
        self.file1_path = file1_path
        self.file2_path = file2_path
        self._is_running = True

    def run(self):
        """Основная логика сравнения файлов"""
        try:
            self.message_received.emit("Инициализация процесса сравнения...")
            df1 = self.load_data(self.file1_path)
            df2 = self.load_data(self.file2_path)
            self.validate_data(df1, df2)
            merged_df = self.merge_dataframes(df1, df2)
            report_data, highlight_info = self.process_differences(merged_df)
            output_path = self.generate_report(report_data, highlight_info)
            self.finished.emit(output_path)

        except Exception as e:
            self.error_occurred.emit(f"Ошибка обработки: {str(e)}")
        finally:
            self._is_running = False

    def load_data(self, file_path):
        """Загрузка данных из файла"""
        self.message_received.emit(f"Загрузка {os.path.basename(file_path)}...")
        try:
            df = pd.read_excel(
                file_path,
                dtype=str,
                keep_default_na=False,
                engine='openpyxl')
            df.columns = df.columns.astype(str)
            return df
        except Exception as e:
            raise ValueError(f"Ошибка чтения файла {file_path}: {str(e)}")

    def validate_data(self, df1, df2):
        """Проверка структуры данных"""
        required_columns = {'id', 'ФИО', 'должность'} | \
                           {str(i) for i in range(1, 32)}
        for df, num in zip([df1, df2], ['первом', 'втором']):
            missing = required_columns - set(df.columns)
            if missing:
                raise ValueError(
                    f"В {num} файле отсутствуют колонки: {', '.join(missing)}")

    def merge_dataframes(self, df1, df2):
        """Объединение данных из двух файлов"""
        self.message_received.emit("Сопоставление данных...")
        df1['original_index'] = df1.index
        df2['original_index'] = df2.index
        merged = pd.merge(
            df1, df2,
            on='id',
            suffixes=('_base', '_compare'),
            how='outer',
            indicator=True)

        # Обработка отсутствующих записей
        missing_in_base = merged[merged['_merge'] == 'right_only']
        missing_in_compare = merged[merged['_merge'] == 'left_only']
        if not missing_in_base.empty:
            self.message_received.emit(
                f"Найдено {len(missing_in_base)} записей отсутствующих в базовом файле")

        if not missing_in_compare.empty:
            self.message_received.emit(
                f"Найдено {len(missing_in_compare)} записей отсутствующих в файле сравнения")
        return merged[merged['_merge'] == 'both']

    def process_differences(self, merged_df):
        """Обработка расхождений с разделением на категории"""
        self.message_received.emit("Поиск различий...")
        vv_diff = []  # Для различий "ВВ"
        dp_diff = []  # Для различий "ДП"
        other_diff = []  # Все остальные случаи
        highlight_info = {'base': [], 'compare': []}
        total_rows = len(merged_df)
        day_columns = [str(i) for i in range(1, 32)]

        for idx, (_, row) in enumerate(merged_df.iterrows()):
            if not self._is_running:
                break
            base_index = row['original_index_base'] + 2
            compare_index = row['original_index_compare'] + 2
            for day in day_columns:
                base_val = str(row[f'{day}_base']).strip()
                compare_val = str(row[f'{day}_compare']).strip()
                if base_val != compare_val:
                    entry = [
                        row['id'],
                        row['ФИО_base'],
                        day,
                        base_val,
                        compare_val
                    ]

                    # Классификация различий
                    if "ВВ" in (base_val, compare_val):
                        vv_diff.append(entry)
                    elif "ДП" in (base_val, compare_val):
                        dp_diff.append(entry)
                    else:
                        other_diff.append(entry)

                    # Сохранение позиций для подсветки
                    base_col = merged_df.columns.get_loc(f'{day}_base') + 1
                    compare_col = merged_df.columns.get_loc(f'{day}_compare') + 1
                    highlight_info['base'].append((base_index, base_col))
                    highlight_info['compare'].append((compare_index, compare_col))

            # Обновление прогресса
            progress = int(20 + 70 * (idx + 1) / total_rows)
            self.progress_updated.emit(progress)
        return {'vv': vv_diff, 'dp': dp_diff, 'other': other_diff}, highlight_info

    def generate_report(self, report_data, highlight_info):
        """Генерация итогового отчета"""
        self.message_received.emit("Формирование отчетов...")
        try:
            self.highlight_differences(self.file1_path, highlight_info['base'])
            self.highlight_differences(self.file2_path, highlight_info['compare'])
            output_path = self.create_report_file(report_data)
            return output_path
        except Exception as e:
            raise ValueError(f"Ошибка создания отчетов: {str(e)}")

    def highlight_differences(self, file_path, cells):
        """Подсветка различий в исходном файле"""

        if not cells:
            return
        wb = load_workbook(file_path)
        ws = wb.active
        yellow_fill = PatternFill(
            start_color='FFFF00',
            end_color='FFFF00',
            fill_type='solid'
        )

        for row, col in cells:
            ws.cell(row=row, column=col).fill = yellow_fill
        wb.save(file_path)
        wb.close()

    def create_report_file(self, report_data):
        """Создание файла с тремя листами отчетов"""
        output_dir = os.path.dirname(self.file1_path)
        base_name = os.path.splitext(os.path.basename(self.file1_path))[0]
        output_name = f"Сравнение_{base_name}.xlsx"
        output_path = os.path.join(output_dir, output_name)
        wb = load_workbook(self.file1_path)

        # Удаляем старые версии отчетов
        for sheet in ['ВВ', 'ДП', 'Остальные']:
            if sheet in wb.sheetnames:
                del wb[sheet]

        # Создаем новые листы
        sheets_config = {
            'ВВ': {
                'data': report_data['vv'],
                'color': 'FF0000',  # Красный
                'header': 'Различия с отметками ВВ'
            },
            'ДП': {
                'data': report_data['dp'],
                'color': '0000FF',  # Синий
                'header': 'Различия с отметками ДП'
            },

            'Остальные': {
                'data': report_data['other'],
                'color': '008000',  # Зеленый
                'header': 'Прочие различия'
            }
        }

        for sheet_name, config in sheets_config.items():
            ws = wb.create_sheet(sheet_name)

            # Заголовки
            headers = ['ID', 'ФИО', 'День', 'Базовый файл', 'Файл сравнения']
            ws.append(headers)

            # Стиль заголовков
            header_font = Font(bold=True, color=config['color'])
            for cell in ws[1]:
                cell.font = header_font

            # Данные
            for row in config['data']:
                ws.append(row)

            # Автонастройка ширины столбцов
            for col in ws.columns:
                max_length = 0
                column = get_column_letter(col[0].column)
                for cell in col:
                    try:
                        max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                adjusted_width = (max_length + 2) * 1.2
                ws.column_dimensions[column].width = adjusted_width

        wb.save(output_path)
        wb.close()
        return output_path


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ExcelComparator()
    window.show()
    sys.exit(app.exec())
