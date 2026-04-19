"""Интеграция стандартного logging с интерфейсом PyQt6."""

from __future__ import annotations

import logging

from PyQt6.QtCore import QObject, pyqtSignal


class QtLogEmitter(QObject):
    """Объект-сигнализатор для безопасной доставки логов в GUI."""

    log_message = pyqtSignal(str)


class QtTextEditLogHandler(logging.Handler):
    """Обработчик логов, пересылающий записи в QTextEdit через сигнал Qt."""

    def __init__(self, emitter: QtLogEmitter) -> None:
        super().__init__()
        self.emitter = emitter

    def emit(self, record: logging.LogRecord) -> None:
        """Преобразует запись лога в строку и отправляет в интерфейс."""

        try:
            message = self.format(record)
            self.emitter.log_message.emit(message)
        except Exception:  # noqa: BLE001
            self.handleError(record)
