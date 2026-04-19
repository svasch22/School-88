"""Доменные модели, используемые в ядре приложения."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass(slots=True)
class TeacherOverride:
    """Правило переопределения учителя."""

    class_name: str
    subject: str
    teacher: str

    def to_dict(self) -> dict[str, str]:
        """Сериализует правило в формат JSON-файла."""

        return {
            "class": self.class_name,
            "subject": self.subject,
            "teacher": self.teacher,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "TeacherOverride":
        """Создает правило из словаря."""

        return cls(
            class_name=str(data.get("class", "")).strip(),
            subject=str(data.get("subject", "")).strip(),
            teacher=str(data.get("teacher", "")).strip(),
        )


@dataclass(slots=True)
class ConversionStats:
    """Статистика обработки одного Excel-файла."""

    file_name: str
    total_sheets: int
    processed_sheets: int
    records_count: int

    def to_dict(self) -> dict[str, int | str]:
        """Преобразует статистику в словарь для отчетов."""

        return {
            "файл": self.file_name,
            "листов_всего": self.total_sheets,
            "листов_обработано": self.processed_sheets,
            "записей": self.records_count,
        }


@dataclass(slots=True)
class ConversionResult:
    """Результат пакетной конвертации."""

    success: bool
    message: str
    output_file: str | None = None
    details: dict[str, Any] = field(default_factory=dict)
