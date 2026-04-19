"""Вспомогательные функции для работы с файлами, датами и идентификаторами."""

from __future__ import annotations

import hashlib
from datetime import datetime
from pathlib import Path

from src.utils.patterns import CLASS_PATTERNS, FILE_DATE_PATTERN


def extract_class_from_filename(file_name: str) -> str | None:
    """Извлекает обозначение класса из имени Excel-файла."""

    base_name = Path(file_name).stem.strip()
    for pattern in CLASS_PATTERNS:
        match = pattern.search(base_name)
        if match:
            return match.group(1)
    return None


def extract_date_from_filename(file_name: str) -> datetime:
    """Извлекает дату выгрузки журнала из имени файла."""

    match = FILE_DATE_PATTERN.search(file_name)
    if match:
        try:
            return datetime.strptime(match.group(1), "%Y%m%d")
        except ValueError:
            return datetime.now()
    return datetime.now()


def generate_lesson_id(subject: str, class_num: str, key_str: str) -> str:
    """Генерирует стабильный идентификатор урока."""

    key = f"{subject}_{class_num}_{key_str}"
    return hashlib.md5(key.encode()).hexdigest()[:12]


def get_trimester_by_date(date_obj: datetime) -> int | None:
    """Определяет номер триместра по дате без изменения исходной логики."""

    month = date_obj.month
    day = date_obj.day

    if month in [9, 10]:
        return 1
    if month == 11:
        return 1 if day <= 23 else 2
    if month in [12, 1, 2]:
        return 2
    if month in [3, 4, 5, 6, 7, 8]:
        return 3
    return None
