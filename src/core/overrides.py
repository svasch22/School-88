"""Работа с JSON-правилами переопределения учителей."""

from __future__ import annotations

import json
import logging
from pathlib import Path

from src.core.models import TeacherOverride

LOGGER = logging.getLogger(__name__)


class TeacherOverridesRepository:
    """Репозиторий для чтения и записи teacher mapping правил."""

    def __init__(self, file_path: Path) -> None:
        self.file_path = file_path

    def load(self) -> list[TeacherOverride]:
        """Загружает правила из JSON-файла."""

        if not self.file_path.exists():
            return []

        try:
            with self.file_path.open("r", encoding="utf-8") as file:
                payload = json.load(file)
            return [TeacherOverride.from_dict(item) for item in payload]
        except Exception as error:  # noqa: BLE001
            LOGGER.exception("Не удалось загрузить JSON с переопределениями: %s", error)
            return []

    def save(self, overrides: list[TeacherOverride]) -> None:
        """Сохраняет правила в JSON-файл."""

        with self.file_path.open("w", encoding="utf-8") as file:
            json.dump(
                [override.to_dict() for override in overrides],
                file,
                ensure_ascii=False,
                indent=4,
            )
