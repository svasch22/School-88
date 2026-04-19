"""Предкомпилированные регулярные выражения проекта."""

from __future__ import annotations

import re

CLASS_PATTERNS: tuple[re.Pattern[str], ...] = (
    re.compile(r"\d+\s+(\d+-НДО)", re.IGNORECASE),
    re.compile(r"\d+\s+(\d+-[а-яА-Я])", re.IGNORECASE),
    re.compile(r"\d+\s+(\d+\s+НДО)", re.IGNORECASE),
    re.compile(r"^\d+\s+(\d+)", re.IGNORECASE),
)

FILE_DATE_PATTERN: re.Pattern[str] = re.compile(r"(\d{8})")
DAY_MONTH_PATTERN: re.Pattern[str] = re.compile(r"(?<!\d)[2-5](?!\d)|\bн\b", re.IGNORECASE)
SUBJECT_TRAILING_DIGITS_PATTERN: re.Pattern[str] = re.compile(r"\d{6,8}\s*$")

SUBJECT_FALLBACK_PATTERNS: tuple[re.Pattern[str], ...] = (
    re.compile(r"^([А-Яа-я\s]+?)\s+\d+[-а-яА-Я]*\s+УП\s+"),
    re.compile(r"^([А-Яа-я\s]+?)\s+\d+[-а-яА-Я]"),
    re.compile(r"^([А-Яа-я\s]+?)\s*\("),
)

GROUP_NUMBER_PATTERN: re.Pattern[str] = re.compile(r"(\d+)\s*гр", re.IGNORECASE)
TEACHER_LINE_PATTERN: re.Pattern[str] = re.compile(r"Учитель:\s*(.+?)\s+\d{2}\.\d{2}")
