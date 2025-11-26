"""Обработчик документов"""

from telegram import Update
from telegram.ext import ContextTypes

from app.services.formatter import format_document
from app.utils.file_manager import cleanup_files, ensure_temp_dir, is_valid_document
from app.utils.logger import setup_logger


logger = setup_logger(__name__)


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """
    Обработчик документов

    Args:
        update: Объект Update от Telegram
        context: Контекст приложения
    """

    input_path = None
    output_path = None

    try:
        # Получаем информацию о файле
        document = update.message.document

        # Проверяем, что это Word документ
        if not is_valid_document(document.file_name):
            await update.message.reply_text(
                "❌ Пожалуйста, отправьте файл в формате .docx"
            )
            logger.warning(
                f"Пользователь {update.effective_user.id} отправил "
                f"неподдерживаемый файл: {document.file_name}"
            )
            return

        # Отправляем сообщение о начале обработки
        processing_msg = await update.message.reply_text("⏳ Обрабатываю документ...")
        logger.info(
            f"Начало обработки документа {document.file_name} "
            f"от пользователя {update.effective_user.id}"
        )

        # Создаем временную директорию
        temp_dir = ensure_temp_dir()

        # Скачиваем файл
        file = await context.bot.get_file(document.file_id)
        input_path = str(temp_dir / f"input_{document.file_id}.docx")
        await file.download_to_drive(input_path)

        logger.info(f"Файл скачан: {input_path}")

        # Обрабатываем документ
        output_path = str(temp_dir / f"output_{document.file_id}.docx")
        success = format_document(input_path, output_path)

        if success:
            # Отправляем обработанный файл
            with open(output_path, 'rb') as f:
                await update.message.reply_document(
                    document=f,
                    filename=f"formatted_{document.file_name}"
                )
            await processing_msg.edit_text("✅ Документ успешно отформатирован!")
            logger.info(
                f"Документ успешно обработан для пользователя {update.effective_user.id}"
            )
        else:
            await processing_msg.edit_text(
                "❌ Произошла ошибка при обработке документа. "
                "Проверьте, что файл не поврежден."
            )
            logger.error(
                f"Ошибка обработки документа для пользователя {update.effective_user.id}"
            )

    except Exception as e:
        logger.error(f"Ошибка при обработке документа: {e}", exc_info=True)
        await update.message.reply_text(
            "❌ Произошла непредвиденная ошибка при обработке файла. "
            "Пожалуйста, попробуйте позже."
        )

    finally:
        # Удаляем временные файлы
        if input_path or output_path:
            cleanup_files([path for path in [input_path, output_path] if path])