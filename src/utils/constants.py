"""Константы проекта.

В модуле собраны строковые литералы, значения по умолчанию и настройки,
которые используются сразу в нескольких слоях приложения.
"""

from __future__ import annotations

from pathlib import Path

MONTHS_MAP: dict[str, int] = {
    "сен": 9,
    "окт": 10,
    "ноя": 11,
    "дек": 12,
    "янв": 1,
    "фев": 2,
    "мар": 3,
    "апр": 4,
    "май": 5,
    "июн": 6,
    "июл": 7,
    "авг": 8,
}

GRADES_COL_START: int = 3
OVERRIDES_FILE_NAME: str = "teacher_overrides.json"
EMPTY_TRIMESTER_MARKERS: set[str] = {"нпа", "за", "а/з"}

DEFAULT_LESSON_COLUMNS: dict[str, int] = {"дата": 21, "тема": 22, "дз": 23}
DEFAULT_TEACHER_NAME: str = "Не указан"
DEFAULT_SUBJECT_NAME: str = "Не указан"
DEFAULT_TOPIC_NAME: str = "Не указана"
DEFAULT_HOMEWORK_TEXT: str = "Не задано"

OUTPUT_GRADES_SHEET_NAME: str = "Оценки и посещение"
OUTPUT_LESSONS_SHEET_NAME: str = "Уроки"

APP_TITLE: str = "Конвертер журналов"
APP_GEOMETRY: tuple[int, int, int, int] = (100, 100, 800, 600)


def get_overrides_path(base_dir: Path) -> Path:
    """Возвращает путь к JSON-файлу с переопределениями учителей."""

    return base_dir / OVERRIDES_FILE_NAME
