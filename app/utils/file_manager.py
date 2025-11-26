"""Управление файлами"""

import os

from pathlib import Path
from typing import List

from app.config.settings import settings
from app.utils.logger import setup_logger


logger = setup_logger(__name__)


def ensure_temp_dir() -> Path:
    """
    Создает временную директорию если её нет

    Returns:
        Path к временной директории
    """

    temp_dir = settings.TEMP_DIR
    temp_dir.mkdir(parents=True, exist_ok=True)
    return temp_dir


def cleanup_files(file_paths: List[str]) -> None:
    """
    Удаление временных файлов

    Args:
        file_paths: Список путей к файлам для удаления
    """

    for file_path in file_paths:
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
                logger.debug(f"Удален файл: {file_path}")
        except Exception as e:
            logger.warning(f"Не удалось удалить {file_path}: {e}")


def get_file_size_mb(file_path: str) -> float:
    """
    Получить размер файла в МБ

    Args:
        file_path: Путь к файлу

    Returns:
        Размер файла в МБ
    """

    return os.path.getsize(file_path) / (1024 * 1024)


def is_valid_document(filename: str) -> bool:
    """
    Проверка валидности документа

    Args:
        filename: Имя файла

    Returns:
        True если файл валиден
    """

    return any(filename.lower().endswith(ext) for ext in settings.ALLOWED_EXTENSIONS)