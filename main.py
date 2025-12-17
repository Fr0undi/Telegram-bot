"""Главная точка входа для Telegram бота"""

import sys

from telegram.ext import Application, CommandHandler, MessageHandler, filters

from app.config.settings import settings
from app.handlers.command_handler import start, help_command
from app.handlers.document_handler import handle_document
from app.utils.logger import setup_logger
from app.utils.file_manager import ensure_temp_dir


# Настройка логирования
logger = setup_logger(__name__)


async def error_handler(update, context):
    """Обработчик глобальных ошибок"""

    logger.error(f"Исключение при обработке обновления: {context.error}", exc_info=context.error)

    # Отправляем сообщение об ошибке пользователю
    if update and update.message:
        await update.message.reply_text(
            "❌ Произошла непредвиденная ошибка. Пожалуйста, попробуйте позже."
        )


def main():
    """Запуск бота"""

    try:
        # Валидация конфигурации
        logger.info("Проверка конфигурации...")
        settings.validate()

        # Создание временной директории
        logger.info("Создание временной директории...")
        ensure_temp_dir()

        # Создание приложения
        logger.info("Инициализация приложения...")
        application = (
            Application.builder()
            .token(settings.BOT_TOKEN)
            .read_timeout(300)  # 5 минут на чтение ответа
            .write_timeout(300)  # 5 минут на отправку данных
            .connect_timeout(60)  # 1 минута на подключение
            .pool_timeout(300)  # 5 минут для пула соединений
            .build()
        )

        # Регистрация обработчиков команд
        logger.info("Регистрация обработчиков...")
        application.add_handler(CommandHandler("start", start))
        application.add_handler(CommandHandler("help", help_command))

        # Регистрация обработчика документов
        application.add_handler(MessageHandler(filters.Document.ALL, handle_document))

        # Регистрация обработчика ошибок
        application.add_error_handler(error_handler)

        # Запуск бота
        logger.info("=" * 50)
        logger.info("  Бот успешно запущен!")
        logger.info("=" * 50)

        application.run_polling()

    except ValueError as e:
        logger.error(f"Ошибка конфигурации: {e}")
        sys.exit(1)
    except Exception as e:
        logger.error(f"Критическая ошибка: {e}", exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()