"""Точка входа в приложение конвертации школьных журналов."""

from __future__ import annotations

import sys

from PyQt6.QtWidgets import QApplication

from src.gui.main_window import JournalConverterMainWindow


def main() -> int:
    """Запускает Qt-приложение."""

    app = QApplication(sys.argv)
    window = JournalConverterMainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
