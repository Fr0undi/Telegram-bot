"""Сервис форматирования документов по ГОСТ"""

import os
import re
import traceback

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from app.config import settings
from app.utils.logger import setup_logger


logger = setup_logger(__name__)


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

        # 2. Применение параметров страницы
        logger.info("Применение параметров страницы...")
        section = doc.sections[0]
        section.left_margin = Cm(settings.GOST_LEFT_MARGIN_CM)
        section.right_margin = Cm(settings.GOST_RIGHT_MARGIN_CM)
        section.top_margin = Cm(settings.GOST_TOP_MARGIN_CM)
        section.bottom_margin = Cm(settings.GOST_BOTTOM_MARGIN_CM)

        # 3. Настройка нумерации страниц
        logger.info("Настройка нумерации страниц...")
        _add_page_numbers(doc)

        # 4. Удаление интервалов перед и после абзацев
        logger.info("Удаление интервалов перед и после абзацев...")
        for paragraph in doc.paragraphs:
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)

        # 5. Форматирование всего текста
        logger.info("Применение настроек шрифта ко всему тексту...")
        for paragraph in doc.paragraphs:
            for run in paragraph.runs:
                run.font.name = settings.GOST_FONT_NAME
                run.font.size = Pt(settings.GOST_FONT_SIZE)

        # 6. Обработка списков и отступов
        logger.info("Обработка списков и отступов...")
        _format_lists(doc)

        # 7. Форматирование заголовков
        logger.info("Форматирование заголовков...")
        _format_headings(doc)

        # 8. Форматирование основного текста
        logger.info("Форматирование основного текста по ширине...")
        _format_main_text(doc)

        # 9. Форматирование абзацев
        logger.info("Форматирование абзацев...")
        _format_paragraphs(doc)

        # 10. Обработка таблиц и рисунков
        logger.info("Обработка таблиц и рисунков...")
        _format_tables_and_figures(doc)

        # 11. Форматирование библиографического списка
        logger.info("Форматирование библиографического списка...")
        _format_bibliography(doc)

        # 12. Генерация оглавления
        logger.info("Генерация оглавления...")
        _generate_toc(doc)

        # 13. Финальная проверка
        logger.info("Финальная проверка...")
        _print_statistics(doc)

        # 14. Сохранение результата
        logger.info("Сохранение документа...")
        doc.save(output_file)
        logger.info(f"Документ успешно сохранен: {output_file}")

        return True

    except Exception as e:
        logger.error(f"Ошибка при обработке документа: {str(e)}")
        logger.debug(f"Детали ошибки: {traceback.format_exc()}")
        return False


def _add_page_numbers(doc: Document) -> None:
    """Добавление нумерации страниц"""

    for section in doc.sections:
        footer = section.footer
        paragraph = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        run = paragraph.add_run()
        fld_char = OxmlElement('w:fldChar')
        fld_char.set(qn('w:fldCharType'), 'begin')
        instr_text = OxmlElement('w:instrText')
        instr_text.set(qn('xml:space'), 'preserve')
        instr_text.text = 'PAGE'
        fld_char2 = OxmlElement('w:fldChar')
        fld_char2.set(qn('w:fldCharType'), 'end')

        run._r.append(fld_char)
        run._r.append(instr_text)
        run._r.append(fld_char2)


def _is_list_item(text: str) -> bool:
    """Проверка, является ли текст элементом списка"""

    return bool(
        re.match(r'^\d+\.\s', text) or
        re.match(r'^\d+\)\s', text) or
        re.match(r'^[а-я]\.\s', text.lower()) or
        re.match(r'^[а-я]\)\s', text.lower()) or
        re.match(r'^[a-z]\.\s', text.lower()) or
        re.match(r'^[a-z]\)\s', text.lower())
    )


def _format_lists(doc: Document) -> None:
    """Форматирование списков"""

    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()

        if _is_list_item(text):
            # Убираем ВСЕ отступы для списков
            paragraph.paragraph_format.first_line_indent = Cm(0)
            paragraph.paragraph_format.left_indent = Cm(0)
            paragraph.paragraph_format.right_indent = Cm(0)

            # Выравнивание по ширине для списков
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

            # Убираем табуляцию после номера
            if len(paragraph.runs) > 0:
                first_run = paragraph.runs[0]
                first_run_text = first_run.text

                if '\t' in first_run_text:
                    first_run.text = first_run_text.replace('\t', ' ')

                # Убеждаемся, что после цифры стоит только один пробел
                cleaned_text = re.sub(r'^(\d+\.)\s+', r'\1 ', first_run.text)
                cleaned_text = re.sub(r'^(\d+\))\s+', r'\1 ', cleaned_text)
                first_run.text = cleaned_text


def _format_headings(doc: Document) -> None:
    """Форматирование заголовков"""

    for paragraph in doc.paragraphs:
        text = paragraph.text.strip().upper()

        # Заголовки разделов
        if text in ['ВВЕДЕНИЕ', 'ЗАКЛЮЧЕНИЕ', 'СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ',
                    'СОДЕРЖАНИЕ', 'ОГЛАВЛЕНИЕ', 'АННОТАЦИЯ', 'РЕФЕРАТ']:
            for run in paragraph.runs:
                run.font.bold = True
                run.font.size = Pt(16)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.first_line_indent = Cm(0)

        # Заголовки подразделов
        elif re.match(r'^\d+\.\d+', paragraph.text):
            for run in paragraph.runs:
                run.font.bold = True
                run.font.size = Pt(14)
            paragraph.paragraph_format.first_line_indent = Cm(0)


def _format_main_text(doc: Document) -> None:
    """Форматирование основного текста по ширине"""

    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()

        # Проверяем, не является ли элемент особым случаем
        is_special = (
            paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER or
            re.match(r'^Таблица \d+ –', text) or
            re.match(r'^Рисунок \d+ –', text) or
            text.upper() in ['ВВЕДЕНИЕ', 'ЗАКЛЮЧЕНИЕ', 'СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ',
                            'СОДЕРЖАНИЕ', 'ОГЛАВЛЕНИЕ', 'АННОТАЦИЯ', 'РЕФЕРАТ'] or
            re.match(r'^\d+\.\d+', text)
        )

        if not is_special:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY


def _format_paragraphs(doc: Document) -> None:
    """Форматирование абзацев (красная строка и междустрочный интервал)"""

    for paragraph in doc.paragraphs:
        paragraph.paragraph_format.line_spacing = settings.GOST_LINE_SPACING

        text = paragraph.text.strip()

        is_list = _is_list_item(text)
        is_special = (
            text.upper() in ['ВВЕДЕНИЕ', 'ЗАКЛЮЧЕНИЕ', 'СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ',
                           'СОДЕРЖАНИЕ', 'ОГЛАВЛЕНИЕ', 'АННОТАЦИЯ', 'РЕФЕРАТ'] or
            re.match(r'^\d+\.\d+', text) or
            paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER
        )

        # Красная строка только для обычных абзацев
        if not is_list and not is_special and text:
            paragraph.paragraph_format.first_line_indent = Cm(settings.GOST_INDENT_CM)
        else:
            paragraph.paragraph_format.first_line_indent = Cm(0)


def _format_tables_and_figures(doc: Document) -> None:
    """Обработка таблиц и рисунков"""

    # Обработка таблиц
    for i, table in enumerate(doc.tables, 1):
        paragraph = doc.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run(f'Таблица {i} – ')
        run.font.size = Pt(12)
        run.font.name = settings.GOST_FONT_NAME

    # Поиск и подпись рисунков
    figure_count = 0
    figure_paragraphs = []

    for i, paragraph in enumerate(doc.paragraphs):
        if paragraph._element.xpath('.//w:drawing') or paragraph._element.xpath('.//a:blip'):
            figure_count += 1
            logger.debug(f"Найден рисунок {figure_count} в параграфе {i}")
            figure_paragraphs.append(i)

    # Добавляем подписи после рисунков
    for figure_index, para_index in enumerate(sorted(figure_paragraphs, reverse=True), 1):
        try:
            figure_number = len(figure_paragraphs) - figure_index + 1

            if para_index + 1 < len(doc.paragraphs):
                caption_paragraph = doc.paragraphs[para_index + 1].insert_paragraph_before()
            else:
                caption_paragraph = doc.add_paragraph()

            caption_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = caption_paragraph.add_run(f'Рисунок {figure_number}')
            run.font.size = Pt(12)
            run.font.name = settings.GOST_FONT_NAME
            run.font.bold = True
            caption_paragraph.paragraph_format.space_after = Pt(12)

            logger.debug(f"Добавлена подпись 'Рисунок {figure_number}'")
        except Exception as e:
            logger.warning(f"Ошибка при добавлении подписи для рисунка: {e}")


def _format_bibliography(doc: Document) -> None:
    """Форматирование библиографического списка"""

    bibliography_started = False

    for paragraph in doc.paragraphs:
        text = paragraph.text.strip().upper()

        if 'СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ' in text or 'БИБЛИОГРАФИЧЕСКИЙ СПИСОК' in text:
            bibliography_started = True
            continue

        if bibliography_started and paragraph.text.strip():
            text = paragraph.text.strip()
            text = re.sub(r'\s+', ' ', text)

            paragraph.clear()
            run = paragraph.add_run(text)
            run.font.name = settings.GOST_FONT_NAME
            run.font.size = Pt(settings.GOST_FONT_SIZE)

            paragraph.paragraph_format.first_line_indent = Cm(-1.25)
            paragraph.paragraph_format.left_indent = Cm(1.25)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY


def _generate_toc(doc: Document) -> None:
    """Генерация оглавления"""

    toc_paragraph = None
    for paragraph in doc.paragraphs:
        if 'СОДЕРЖАНИЕ' in paragraph.text.upper() or 'ОГЛАВЛЕНИЕ' in paragraph.text.upper():
            toc_paragraph = paragraph
            break

    if not toc_paragraph:
        toc_paragraph = doc.paragraphs[0].insert_paragraph_before('СОДЕРЖАНИЕ')
        toc_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        toc_paragraph.runs[0].font.bold = True
        toc_paragraph.runs[0].font.size = Pt(16)

    doc.add_paragraph()

    headings = []
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if (text.upper() in ['ВВЕДЕНИЕ', 'ЗАКЛЮЧЕНИЕ', 'СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ'] or
            re.match(r'^\d+\.', text)):
            headings.append(text)

    for heading in headings:
        toc_item = doc.add_paragraph()
        toc_item.paragraph_format.left_indent = Cm(0)
        toc_item.paragraph_format.first_line_indent = Cm(0)

        run = toc_item.add_run(f'{heading} ')
        run.font.name = settings.GOST_FONT_NAME
        run.font.size = Pt(settings.GOST_FONT_SIZE)

        run = toc_item.add_run('1')
        run.font.name = settings.GOST_FONT_NAME
        run.font.size = Pt(settings.GOST_FONT_SIZE)


def _print_statistics(doc: Document) -> None:
    """Вывод статистики форматирования"""

    center_count = 0
    justify_count = 0
    left_count = 0
    list_count = 0
    table_count = len(doc.tables)
    figure_count = sum(1 for p in doc.paragraphs if p._element.xpath('.//w:drawing'))

    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()

        if _is_list_item(text):
            list_count += 1
        elif paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER:
            center_count += 1
        elif paragraph.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY:
            justify_count += 1
        else:
            left_count += 1

    logger.info("Статистика форматирования:")
    logger.info(f"  - По центру: {center_count} параграфов")
    logger.info(f"  - По ширине: {justify_count} параграфов")
    logger.info(f"  - По левому краю: {left_count} параграфов")
    logger.info(f"  - Элементов списка: {list_count}")
    logger.info(f"  - Таблиц: {table_count}")
    logger.info(f"  - Рисунков: {figure_count}")