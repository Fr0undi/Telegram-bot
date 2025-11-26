"""Настройка логирования"""

import logging
import sys

from app.config.settings import settings


def setup_logger(name: str = __name__) -> logging.Logger:
    """
    Настройка логгера

    Args:
        name: Имя логгера

    Returns:
        Настроенный логгер
    """

    logger = logging.getLogger(name)
    logger.setLevel(getattr(logging, settings.LOG_LEVEL))

    # Формат логов
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )

    # Консольный обработчик
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    # Файловый обработчик
    if settings.LOG_FILE:
        # Создаем директорию для логов если её нет
        settings.LOG_FILE.parent.mkdir(parents=True, exist_ok=True)

        file_handler = logging.FileHandler(
            settings.LOG_FILE,
            encoding='utf-8'
        )
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

    return logger