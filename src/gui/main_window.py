"""Главное окно приложения.

GUI полностью отделен от бизнес-логики: окно работает только через
сервисный слой, репозиторий правил и стандартный logging.
"""

from __future__ import annotations

import logging
import threading
from pathlib import Path

from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QDragEnterEvent, QDropEvent, QFont
from PyQt6.QtWidgets import (
    QComboBox,
    QFileDialog,
    QHeaderView,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from src.core.models import TeacherOverride
from src.core.overrides import TeacherOverridesRepository
from src.core.service import JournalConversionService
from src.gui.logging_handler import QtLogEmitter, QtTextEditLogHandler
from src.utils.constants import APP_GEOMETRY, APP_TITLE, get_overrides_path

LOGGER = logging.getLogger(__name__)


class ConversionNotifier(QWidget):
    """Qt-объект для сигналов завершения фоновой конвертации."""

    conversion_complete = pyqtSignal(bool, str)


class JournalConverterMainWindow(QMainWindow):
    """Главное окно настольного приложения."""

    def __init__(self) -> None:
        super().__init__()
        self.notifier = ConversionNotifier()
        self.notifier.conversion_complete.connect(self.conversion_finished)

        self.log_emitter = QtLogEmitter()
        self.log_emitter.log_message.connect(self.append_log)
        self.log_handler = QtTextEditLogHandler(self.log_emitter)

        self.conversion_thread: threading.Thread | None = None
        self.excel_files: list[str] = []
        self.input_folder: str | None = None

        self.project_root = Path(__file__).resolve().parents[2]
        self.overrides_repository = TeacherOverridesRepository(
            get_overrides_path(self.project_root),
        )
        self.overrides: list[TeacherOverride] = self.overrides_repository.load()
        self.conversion_service = JournalConversionService()

        self._configure_logging()
        self._init_ui()

    def _configure_logging(self) -> None:
        """Настраивает корневой логгер для вывода в GUI."""

        root_logger = logging.getLogger()
        root_logger.setLevel(logging.INFO)
        self.log_handler.setFormatter(logging.Formatter("%(asctime)s | %(levelname)s | %(message)s"))

        if not any(isinstance(handler, QtTextEditLogHandler) for handler in root_logger.handlers):
            root_logger.addHandler(self.log_handler)

    def _init_ui(self) -> None:
        """Создает визуальную структуру окна и вкладок."""

        x_pos, y_pos, width, height = APP_GEOMETRY
        self.setWindowTitle(APP_TITLE)
        self.setGeometry(x_pos, y_pos, width, height)
        self.setStyleSheet(self.get_stylesheet())

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        tabs = QTabWidget()
        tabs.addTab(self.create_upload_tab(), "Загрузка файлов")
        tabs.addTab(self.create_mapping_tab(), "Актуализация учителей")
        tabs.addTab(self.create_report_tab(), "Отчет")
        main_layout.addWidget(tabs)

    def create_upload_tab(self) -> QWidget:
        """Создает вкладку выбора файлов и запуска конвертации."""

        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setSpacing(15)

        self.drag_drop_area = QWidget()
        self.drag_drop_area.setMinimumHeight(200)
        self.drag_drop_area.setStyleSheet(
            "QWidget { border: 2px dashed #208080; border-radius: 6px; background-color: #f0f8f8; }",
        )
        self.drag_drop_area.setAcceptDrops(True)
        self.drag_drop_area.dragEnterEvent = self.drag_enter_event
        self.drag_drop_area.dropEvent = self.drop_event

        drag_label = QLabel(
            'Перетащите Excel файлы сюда\nили кликните "Выбрать папку"',
        )
        drag_label.setFont(QFont("Arial", 13))
        drag_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        drag_label.setStyleSheet("QLabel { border: none; background: transparent; }")

        drag_layout = QVBoxLayout(self.drag_drop_area)
        drag_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        drag_layout.addWidget(drag_label)

        layout.addWidget(self.drag_drop_area, 1)
        layout.addWidget(QLabel("Выбранные файлы:"))

        self.file_list = QListWidget()
        self.file_list.setMaximumHeight(150)
        layout.addWidget(self.file_list)

        button_layout = QHBoxLayout()

        self.select_folder_button = QPushButton("Выбрать папку")
        self.select_folder_button.setMinimumHeight(45)
        self.select_folder_button.clicked.connect(self.select_folder)
        button_layout.addWidget(self.select_folder_button)

        self.clear_button = QPushButton("Очистить")
        self.clear_button.setMinimumHeight(45)
        self.clear_button.clicked.connect(self.clear_files)
        button_layout.addWidget(self.clear_button)

        self.convert_button = QPushButton("Начать работу")
        self.convert_button.setMinimumHeight(45)
        self.convert_button.setStyleSheet(
            "QPushButton { background-color: #208080; color: white; border: none; "
            "border-radius: 6px; font-weight: bold; font-size: 14px; } "
            "QPushButton:hover { background-color: #1a6666; } "
            "QPushButton:pressed { background-color: #155555; }",
        )
        self.convert_button.clicked.connect(self.start_conversion)
        button_layout.addWidget(self.convert_button)

        layout.addLayout(button_layout)
        return widget

    def create_mapping_tab(self) -> QWidget:
        """Создает вкладку управления правилами замены учителей."""

        widget = QWidget()
        layout = QVBoxLayout(widget)

        info_label = QLabel(
            "Укажите основного учителя, если в файле указан замещающий.\n"
            "(Достаточно части предмета, например «англ» или «литератур»)",
        )
        info_label.setStyleSheet("color: #555; font-style: italic;")
        layout.addWidget(info_label)

        form_layout = QHBoxLayout()

        self.combo_num = QComboBox()
        self.combo_num.addItems([str(number) for number in range(1, 12)])

        self.combo_letter = QComboBox()
        self.combo_letter.addItems(["а", "б", "в", "г", "д", "е", "ж", "з", "и", "к", "л", "м", "НДО"])
        self.combo_letter.setEditable(True)

        self.edit_subject = QLineEdit()
        self.edit_subject.setPlaceholderText("Предмет (напр. Литература)")

        self.edit_teacher = QLineEdit()
        self.edit_teacher.setPlaceholderText("ФИО основного учителя")

        add_button = QPushButton("Добавить")
        add_button.clicked.connect(self.add_override)
        add_button.setStyleSheet("background-color: #4CAF50; color: white;")

        form_layout.addWidget(QLabel("Класс:"))
        form_layout.addWidget(self.combo_num)
        form_layout.addWidget(QLabel("-"))
        form_layout.addWidget(self.combo_letter)
        form_layout.addWidget(self.edit_subject)
        form_layout.addWidget(self.edit_teacher)
        form_layout.addWidget(add_button)

        layout.addLayout(form_layout)

        self.table_overrides = QTableWidget(0, 3)
        self.table_overrides.setHorizontalHeaderLabels(
            ["Класс", "Предмет (ключевое слово)", "Основной учитель"],
        )
        self.table_overrides.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Stretch,
        )
        layout.addWidget(self.table_overrides)

        remove_button = QPushButton("Удалить выбранное правило")
        remove_button.clicked.connect(self.remove_selected_override)
        layout.addWidget(remove_button)

        self.update_overrides_table()
        return widget

    def create_report_tab(self) -> QWidget:
        """Создает вкладку с текстовым логом выполнения."""

        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.addWidget(QLabel("Лог выполнения:"))

        self.report_text = QTextEdit()
        self.report_text.setReadOnly(True)
        self.report_text.setFont(QFont("Courier New", 10))
        self.report_text.setStyleSheet(
            "QTextEdit { background-color: #ffffff; color: #333333; border: 1px solid #d0d0d0; }",
        )
        layout.addWidget(self.report_text)
        return widget

    def update_overrides_table(self) -> None:
        """Обновляет таблицу правил teacher mapping."""

        self.table_overrides.setRowCount(0)
        for row_index, override in enumerate(self.overrides):
            self.table_overrides.insertRow(row_index)
            self.table_overrides.setItem(row_index, 0, QTableWidgetItem(override.class_name))
            self.table_overrides.setItem(row_index, 1, QTableWidgetItem(override.subject))
            self.table_overrides.setItem(row_index, 2, QTableWidgetItem(override.teacher))

    def persist_overrides(self) -> None:
        """Сохраняет текущие правила в JSON."""

        self.overrides_repository.save(self.overrides)

    def add_override(self) -> None:
        """Добавляет или обновляет правило замены учителя."""

        class_number = self.combo_num.currentText().strip()
        class_letter = self.combo_letter.currentText().strip()
        subject = self.edit_subject.text().strip()
        teacher = self.edit_teacher.text().strip()

        if not subject or not teacher:
            QMessageBox.warning(self, "Ошибка", "Заполните поля Предмет и ФИО учителя.")
            return

        full_class = f"{class_number}-{class_letter}" if class_letter else class_number

        for override in self.overrides:
            if override.class_name == full_class and override.subject.lower() == subject.lower():
                override.teacher = teacher
                self.persist_overrides()
                self.update_overrides_table()
                self.edit_teacher.clear()
                return

        self.overrides.append(
            TeacherOverride(
                class_name=full_class,
                subject=subject,
                teacher=teacher,
            ),
        )
        self.persist_overrides()
        self.update_overrides_table()
        self.edit_teacher.clear()

    def remove_selected_override(self) -> None:
        """Удаляет выбранное правило из таблицы и JSON."""

        current_row = self.table_overrides.currentRow()
        if current_row >= 0:
            del self.overrides[current_row]
            self.persist_overrides()
            self.update_overrides_table()

    def drag_enter_event(self, event: QDragEnterEvent) -> None:
        """Разрешает drag-and-drop для Excel-файлов."""

        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def drop_event(self, event: QDropEvent) -> None:
        """Обрабатывает drop списка Excel-файлов в окно."""

        files = [
            url.toLocalFile()
            for url in event.mimeData().urls()
            if url.toLocalFile().endswith(".xlsx")
        ]
        if files:
            self.excel_files = files
            self.input_folder = str(Path(files[0]).parent)
            self.update_file_list()
        event.accept()

    def select_folder(self) -> None:
        """Открывает диалог выбора папки с Excel-файлами."""

        folder = QFileDialog.getExistingDirectory(self, "Выберите папку с Excel файлами", "")
        if folder:
            self.input_folder = folder
            excel_files = sorted(Path(folder).glob("*.xlsx"))
            if excel_files:
                self.excel_files = [str(file_path) for file_path in excel_files]
                self.update_file_list()
            else:
                QMessageBox.warning(
                    self,
                    "Ошибка",
                    "В выбранной папке не найдены Excel файлы (.xlsx).",
                )

    def update_file_list(self) -> None:
        """Отображает список выбранных файлов в интерфейсе."""

        self.file_list.clear()
        for file_path in self.excel_files:
            self.file_list.addItem(QListWidgetItem(Path(file_path).name))

    def clear_files(self) -> None:
        """Очищает выбранные файлы и текущую папку."""

        self.excel_files = []
        self.input_folder = None
        self.file_list.clear()

    def start_conversion(self) -> None:
        """Запускает конвертацию в отдельном потоке."""

        if not self.excel_files:
            QMessageBox.warning(self, "Ошибка", "Файлы не выбраны.")
            return

        self.report_text.clear()
        self.conversion_thread = threading.Thread(target=self.run_conversion, daemon=True)
        self.conversion_thread.start()

    def run_conversion(self) -> None:
        """Выполняет конвертацию в фоне и сообщает о результате через сигнал."""

        try:
            result = self.conversion_service.convert_files(
                excel_files=self.excel_files,
                input_folder=self.input_folder,
                overrides=self.overrides,
            )
            self.notifier.conversion_complete.emit(result.success, result.message)
        except Exception as error:  # noqa: BLE001
            LOGGER.exception("Ошибка запуска пакетной конвертации: %s", error)
            self.notifier.conversion_complete.emit(False, str(error))

    def append_log(self, message: str) -> None:
        """Добавляет строку лога в текстовый виджет."""

        self.report_text.append(message)
        scroll_bar = self.report_text.verticalScrollBar()
        scroll_bar.setValue(scroll_bar.maximum())

    def conversion_finished(self, success: bool, message: str) -> None:
        """Показывает пользователю финальный статус операции."""

        if success:
            QMessageBox.information(self, "Успех", message)
        else:
            QMessageBox.critical(self, "Ошибка", message)

    @staticmethod
    def get_stylesheet() -> str:
        """Возвращает базовую тему приложения."""

        return """
            QMainWindow { background-color: #ffffff; }
            QLabel { color: #333; font-family: Arial; }
            QPushButton {
                background-color: #e0e0e0;
                color: #333;
                border: 1px solid #c0c0c0;
                border-radius: 6px;
                padding: 8px;
                font-weight: 500;
                font-size: 12px;
            }
            QPushButton:hover { background-color: #d0d0d0; }
            QPushButton:pressed { background-color: #b0b0b0; }
            QPushButton:disabled { background-color: #f0f0f0; color: #aaa; }
            QListWidget { border: 1px solid #d0d0d0; border-radius: 4px; background-color: #fafafa; }
            QListWidget::item { padding: 5px; border-bottom: 1px solid #efefef; }
            QListWidget::item:selected { background-color: #208080; color: white; }
            QTabWidget::pane { border: 1px solid #d0d0d0; }
            QTabBar::tab {
                background-color: #e8e8e8;
                color: #333;
                padding: 8px 20px;
                border: 1px solid #d0d0d0;
                border-bottom: none;
            }
            QTabBar::tab:selected { background-color: #208080; color: white; }
            QLineEdit, QComboBox { padding: 5px; border: 1px solid #ccc; border-radius: 4px; }
            QTableWidget { border: 1px solid #ccc; border-radius: 4px; }
        """
