"""Сервис форматирования документов по ГОСТ"""

import os
import re
import traceback

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from app.config.settings import settings
from app.utils.logger import setup_logger


logger = setup_logger(__name__)


def _find_title_page_end(doc: Document) -> int:
    """
    Находит индекс параграфа, где заканчивается титульный лист.
    Титульник заканчивается перед "Содержание" или "Оглавление".

    Returns:
        Индекс первого параграфа после титульника
    """
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip().upper()
        if text in ['СОДЕРЖАНИЕ', 'ОГЛАВЛЕНИЕ']:
            return i
    return 0  # Если не нашли, считаем что титульника нет


def _is_main_heading(text: str) -> bool:
    """
    Проверяет, является ли текст главным заголовком (Heading 1).
    Требует разрыв страницы и шрифт 16pt BOLD.
    """
    text_upper = text.strip().upper()

    # Специальные разделы
    if text_upper in ['СОДЕРЖАНИЕ', 'ОГЛАВЛЕНИЕ', 'ВВЕДЕНИЕ', 'ЗАКЛЮЧЕНИЕ',
                      'СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ', 'БИБЛИОГРАФИЧЕСКИЙ СПИСОК',
                      'АННОТАЦИЯ', 'РЕФЕРАТ', 'ПРИЛОЖЕНИЕ', 'ПРИЛОЖЕНИЯ']:
        return True

    # Главы: "Глава 1", "ГЛАВА 2", "Глава 1." и т.д.
    if re.match(r'^ГЛАВА\s+\d+', text_upper):
        return True

    return False


def _is_subheading(text: str) -> bool:
    """
    Проверяет, является ли текст подзаголовком (Heading 2).
    Формат: 1.1, 1.2, 2.1 и т.д.
    """
    return bool(re.match(r'^\d+\.\d+', text.strip()))


def _is_figure_caption(text: str) -> bool:
    """Проверяет, является ли текст подписью к рисунку"""
    return bool(re.match(r'^Рисунок\s+\d*\s*[-–—]', text.strip(), re.IGNORECASE))


def _is_table_caption(text: str) -> bool:
    """Проверяет, является ли текст подписью к таблице"""
    return bool(re.match(r'^Таблица\s+\d+\s*[-–—]', text.strip(), re.IGNORECASE))


def _add_page_break_before(paragraph) -> None:
    """Добавляет разрыв страницы перед параграфом"""
    p = paragraph._element
    pPr = p.get_or_add_pPr()

    # Проверяем, нет ли уже разрыва
    existing_break = pPr.find(qn('w:pageBreakBefore'))
    if existing_break is None:
        page_break = OxmlElement('w:pageBreakBefore')
        page_break.set(qn('w:val'), 'true')
        pPr.insert(0, page_break)


def format_document(input_file: str, output_file: str = None) -> bool:
    """
    Основная функция для форматирования документа по ГОСТ

    Args:
        input_file: Путь к входному файлу .docx
        output_file: Путь к выходному файлу. Если None, создается автоматически

    Returns:
        True если успешно, False если ошибка
    """

    # Автоматическое создание имени для выходного файла
    if not output_file:
        base, ext = os.path.splitext(input_file)
        output_file = f"{base}_formatted{ext}"

    # Проверка существования файла
    if not os.path.exists(input_file):
        logger.error(f"Файл '{input_file}' не найден!")
        return False

    try:
        # 1. Загрузка документа
        logger.info("Загрузка документа...")
        doc = Document(input_file)
        logger.info("Документ успешно загружен")

        # 2. Определяем конец титульного листа
        title_end_idx = _find_title_page_end(doc)
        logger.info(f"Титульный лист заканчивается на параграфе {title_end_idx}")

        # 3. Применение параметров страницы
        logger.info("Применение параметров страницы...")
        for section in doc.sections:
            section.left_margin = Cm(settings.GOST_LEFT_MARGIN_CM)
            section.right_margin = Cm(settings.GOST_RIGHT_MARGIN_CM)
            section.top_margin = Cm(settings.GOST_TOP_MARGIN_CM)
            section.bottom_margin = Cm(settings.GOST_BOTTOM_MARGIN_CM)

        # 4. Настройка нумерации страниц
        logger.info("Настройка нумерации страниц...")
        _add_page_numbers(doc)

        # 5. Форматирование документа (кроме титульника)
        logger.info("Форматирование документа...")
        _format_document_content(doc, title_end_idx)

        # 6. Финальная проверка
        logger.info("Финальная проверка...")
        _print_statistics(doc)

        # 7. Сохранение результата
        logger.info("Сохранение документа...")
        doc.save(output_file)
        logger.info(f"Документ успешно сохранен: {output_file}")

        return True

    except Exception as e:
        logger.error(f"Ошибка при обработке документа: {str(e)}")
        logger.debug(f"Детали ошибки: {traceback.format_exc()}")
        return False


def _format_document_content(doc: Document, title_end_idx: int) -> None:
    """
    Форматирует содержимое документа, пропуская титульный лист.

    Args:
        doc: Документ
        title_end_idx: Индекс первого параграфа после титульника
    """

    for i, paragraph in enumerate(doc.paragraphs):
        # Пропускаем титульный лист
        if i < title_end_idx:
            continue

        text = paragraph.text.strip()
        if not text:
            continue

        # Определяем тип параграфа и форматируем
        if _is_main_heading(text):
            _format_main_heading(paragraph)
        elif _is_subheading(text):
            _format_subheading(paragraph)
        elif _is_figure_caption(text) or _is_table_caption(text):
            _format_caption(paragraph)
        else:
            _format_regular_paragraph(paragraph)


def _format_main_heading(paragraph) -> None:
    """
    Форматирует главный заголовок (Введение, Глава X, Заключение и т.д.)
    - 16pt BOLD
    - По центру
    - Без красной строки
    - Разрыв страницы перед
    """
    # Разрыв страницы перед заголовком
    _add_page_break_before(paragraph)

    # Выравнивание по центру
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Убираем отступы
    paragraph.paragraph_format.first_line_indent = Cm(0)
    paragraph.paragraph_format.left_indent = Cm(0)
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = settings.GOST_LINE_SPACING

    # Шрифт 16pt BOLD
    for run in paragraph.runs:
        run.font.name = settings.GOST_FONT_NAME
        run.font.size = Pt(16)
        run.font.bold = True
        # Устанавливаем шрифт для кириллицы
        run._element.rPr.rFonts.set(qn('w:eastAsia'), settings.GOST_FONT_NAME)


def _format_subheading(paragraph) -> None:
    """
    Форматирует подзаголовок (1.1, 1.2, 2.1 и т.д.)
    - 14pt BOLD
    - По центру
    - Без красной строки
    - Без разрыва страницы
    """
    # Выравнивание по центру
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Убираем отступы
    paragraph.paragraph_format.first_line_indent = Cm(0)
    paragraph.paragraph_format.left_indent = Cm(0)
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = settings.GOST_LINE_SPACING

    # Шрифт 14pt BOLD
    for run in paragraph.runs:
        run.font.name = settings.GOST_FONT_NAME
        run.font.size = Pt(settings.GOST_FONT_SIZE)
        run.font.bold = True
        run._element.rPr.rFonts.set(qn('w:eastAsia'), settings.GOST_FONT_NAME)


def _format_caption(paragraph) -> None:
    """
    Форматирует подпись к рисунку или таблице
    - 14pt
    - По центру
    - Без красной строки
    """
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    paragraph.paragraph_format.first_line_indent = Cm(0)
    paragraph.paragraph_format.left_indent = Cm(0)
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = settings.GOST_LINE_SPACING

    for run in paragraph.runs:
        run.font.name = settings.GOST_FONT_NAME
        run.font.size = Pt(settings.GOST_FONT_SIZE)
        run.font.bold = False
        run._element.rPr.rFonts.set(qn('w:eastAsia'), settings.GOST_FONT_NAME)


def _format_regular_paragraph(paragraph) -> None:
    """
    Форматирует обычный абзац
    - 14pt
    - По ширине
    - Красная строка 1.25 см
    - Межстрочный интервал 1.5
    """
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    paragraph.paragraph_format.first_line_indent = Cm(settings.GOST_INDENT_CM)
    paragraph.paragraph_format.left_indent = Cm(0)
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = settings.GOST_LINE_SPACING

    for run in paragraph.runs:
        run.font.name = settings.GOST_FONT_NAME
        run.font.size = Pt(settings.GOST_FONT_SIZE)
        run._element.rPr.rFonts.set(qn('w:eastAsia'), settings.GOST_FONT_NAME)


def _add_page_numbers(doc: Document) -> None:
    """
    Добавляет нумерацию страниц.
    Титульник = страница 1, но номер не отображается.
    Номера видны со страницы 2.
    """
    for section in doc.sections:
        # Включаем "Different First Page" чтобы скрыть номер на титульнике
        section.different_first_page_header_footer = True

        # Настраиваем footer для остальных страниц
        footer = section.footer
        footer.is_linked_to_previous = False

        # Очищаем существующий footer
        for para in footer.paragraphs:
            para.clear()

        # Добавляем номер страницы
        paragraph = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Создаем поле PAGE
        run = paragraph.add_run()
        fld_char_begin = OxmlElement('w:fldChar')
        fld_char_begin.set(qn('w:fldCharType'), 'begin')

        instr_text = OxmlElement('w:instrText')
        instr_text.set(qn('xml:space'), 'preserve')
        instr_text.text = 'PAGE'

        fld_char_end = OxmlElement('w:fldChar')
        fld_char_end.set(qn('w:fldCharType'), 'end')

        run._r.append(fld_char_begin)
        run._r.append(instr_text)
        run._r.append(fld_char_end)

        # Форматируем номер страницы
        run.font.name = settings.GOST_FONT_NAME
        run.font.size = Pt(settings.GOST_FONT_SIZE)


def _print_statistics(doc: Document) -> None:
    """Вывод статистики форматирования"""

    center_count = 0
    justify_count = 0
    left_count = 0
    table_count = len(doc.tables)

    for paragraph in doc.paragraphs:
        if paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER:
            center_count += 1
        elif paragraph.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY:
            justify_count += 1
        else:
            left_count += 1

    logger.info("Статистика форматирования:")
    logger.info(f"  - По центру: {center_count} параграфов")
    logger.info(f"  - По ширине: {justify_count} параграфов")
    logger.info(f"  - По левому краю: {left_count} параграфов")
    logger.info(f"  - Таблиц: {table_count}")