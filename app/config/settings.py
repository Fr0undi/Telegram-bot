"""Настройки приложения"""

import os

from pathlib import Path
from dotenv import load_dotenv

# Загрузка переменных окружения
load_dotenv()

class Settings:
    """Класс с настройками приложения"""

    # Базовые настройки
    BASE_DIR = Path(__file__).resolve().parent.parent.parent
    TEMP_DIR = BASE_DIR / "temp"

    # Telegram Bot
    BOT_TOKEN: str = os.getenv("BOT_TOKEN", "")

    # Настройки документов
    ALLOWED_EXTENSIONS = ['.docx', '.doc']
    MAX_FILE_SIZE_MB = 50

    # Настройки логирования
    LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO")
    LOG_FILE = BASE_DIR / "logs" / "bot.log"

    # ГОСТ настройки форматирования
    GOST_FONT_NAME = "Times New Roman"
    GOST_FONT_SIZE = 14
    GOST_LINE_SPACING = 1.5
    GOST_LEFT_MARGIN_CM = 3      # Левое поле 3 см
    GOST_RIGHT_MARGIN_CM = 1.5   # Правое поле 1.5 см (исправлено с 1)
    GOST_TOP_MARGIN_CM = 2       # Верхнее поле 2 см
    GOST_BOTTOM_MARGIN_CM = 2    # Нижнее поле 2 см
    GOST_INDENT_CM = 1.25        # Красная строка 1.25 см

    def validate(self):
        """Проверка обязательных настроек"""

        if not self.BOT_TOKEN:
            raise ValueError("BOT_TOKEN не установлен в переменных окружения")
        return True

settings = Settings()