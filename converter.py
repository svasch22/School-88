"""Обратносуместимый запуск приложения.

Оставлен как тонкая обертка над новой точкой входа `main.py`
"""

from __future__ import annotations

from main import main


if __name__ == "__main__":
    raise SystemExit(main())
