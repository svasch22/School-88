"""Сервисный слой приложения.

Модуль инкапсулирует пользовательский сценарий пакетной конвертации.
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

from src.core.converters import BatchFolderConverter
from src.core.models import ConversionResult, TeacherOverride


class JournalConversionService:
    """Фасад над ядром конвертации для GUI и других интеграций."""

    def convert_files(
        self,
        excel_files: list[str],
        input_folder: str | None,
        overrides: list[TeacherOverride],
    ) -> ConversionResult:
        """Конвертирует выбранные Excel-файлы и сохраняет итоговый отчет."""

        if not excel_files:
            return ConversionResult(success=False, message="Файлы не выбраны")

        resolved_input_folder = Path(input_folder) if input_folder else Path(excel_files[0]).parent
        output_file = resolved_input_folder.parent / f"журнал_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

        converter = BatchFolderConverter(resolved_input_folder, overrides=overrides)
        converter.excel_files = [Path(file_path) for file_path in excel_files]

        if converter.convert_all_files() and converter.save_results(output_file):
            return ConversionResult(
                success=True,
                message=f"Готово: {output_file}",
                output_file=str(output_file),
                details={"files": converter.file_results},
            )

        return ConversionResult(
            success=False,
            message="Ошибка конвертации",
            details={"files": converter.file_results},
        )
