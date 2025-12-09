"""Сервис форматирования документов по ГОСТ"""

import os
import re
import traceback
import copy

from docx import Document
from docx.shared import Pt, Cm, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn, nsmap
from docx.oxml import OxmlElement

from app.config.settings import settings
from app.utils.logger import setup_logger


logger = setup_logger(__name__)

# Константы для заголовков разделов
SECTION_HEADINGS = [
    'ВВЕДЕНИЕ', 'ЗАКЛЮЧЕНИЕ', 'СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ',
    'СПИСОК ЛИТЕРАТУРЫ', 'БИБЛИОГРАФИЧЕСКИЙ СПИСОК',
    'СОДЕРЖАНИЕ', 'ОГЛАВЛЕНИЕ', 'АННОТАЦИЯ', 'РЕФЕРАТ',
    'СПИСОК СОКРАЩЕНИЙ', 'ГЛОССАРИЙ', 'ОПРЕДЕЛЕНИЯ'
]

# Константы для приложений
APPENDIX_PATTERNS = [
    r'^ПРИЛОЖЕНИЕ\s+[А-Я]',
    r'^Приложение\s+[А-Я]',
]


def format_document(input_file: str, output_file: str = None) -> bool:
    """
    Основная функция для форматирования документа по ГОСТ

    Args:
        input_file: Путь к входному файлу .docx
        output_file: Путь к выходному файлу. Если None, создается автоматически

    Returns:
        True если успешно, False если ошибка
    """

    if not output_file:
        base, ext = os.path.splitext(input_file)
        output_file = f"{base}_formatted{ext}"

    if not os.path.exists(input_file):
        logger.error(f"Файл '{input_file}' не найден!")
        return False

    try:
        logger.info("Загрузка документа...")
        doc = Document(input_file)
        logger.info("Документ успешно загружен")

        # 1. Применение параметров страницы
        logger.info("Применение параметров страницы...")
        _apply_page_settings(doc)

        # 2. Настройка стилей документа
        logger.info("Настройка стилей документа...")
        _setup_document_styles(doc)

        # 3. Форматирование всего текста базовым шрифтом
        logger.info("Применение базового шрифта...")
        _apply_base_font(doc)

        # 4. Форматирование заголовков
        logger.info("Форматирование заголовков...")
        _format_headings(doc)

        # 5. Форматирование основного текста
        logger.info("Форматирование основного текста...")
        _format_main_text(doc)

        # 6. Форматирование списков
        logger.info("Форматирование списков...")
        _format_lists(doc)

        # 7. Форматирование таблиц
        logger.info("Форматирование таблиц...")
        _format_tables(doc)

        # 8. Форматирование рисунков
        logger.info("Форматирование рисунков...")
        _format_figures(doc)

        # 9. Форматирование формул
        logger.info("Форматирование формул...")
        _format_formulas(doc)

        # 10. Форматирование библиографии
        logger.info("Форматирование библиографии...")
        _format_bibliography(doc)

        # 11. Форматирование приложений
        logger.info("Форматирование приложений...")
        _format_appendixes(doc)

        # 12. Добавление нумерации страниц
        logger.info("Добавление нумерации страниц...")
        _add_page_numbers(doc)

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


def _apply_page_settings(doc: Document) -> None:
    """Применение параметров страницы по ГОСТ"""

    for section in doc.sections:
        # Поля страницы
        section.left_margin = Cm(settings.GOST_LEFT_MARGIN_CM)
        section.right_margin = Cm(settings.GOST_RIGHT_MARGIN_CM)
        section.top_margin = Cm(settings.GOST_TOP_MARGIN_CM)
        section.bottom_margin = Cm(settings.GOST_BOTTOM_MARGIN_CM)

        # Размер страницы A4
        section.page_width = Cm(21)
        section.page_height = Cm(29.7)


def _setup_document_styles(doc: Document) -> None:
    """Настройка стилей документа"""

    styles = doc.styles

    # Настройка стиля Normal
    try:
        normal_style = styles['Normal']
        normal_style.font.name = settings.GOST_FONT_NAME
        normal_style.font.size = Pt(settings.GOST_FONT_SIZE)
        normal_style.paragraph_format.line_spacing = settings.GOST_LINE_SPACING
        normal_style.paragraph_format.space_before = Pt(0)
        normal_style.paragraph_format.space_after = Pt(0)

        # Установка шрифта для кириллицы
        _set_font_eastasia(normal_style._element, settings.GOST_FONT_NAME)
    except Exception as e:
        logger.warning(f"Не удалось настроить стиль Normal: {e}")


def _set_font_eastasia(element, font_name: str) -> None:
    """Безопасная установка шрифта для восточноазиатских символов"""
    try:
        rPr = element.get_or_add_rPr() if hasattr(element, 'get_or_add_rPr') else None
        if rPr is not None:
            rFonts = rPr.find(qn('w:rFonts'))
            if rFonts is None:
                rFonts = OxmlElement('w:rFonts')
                rPr.insert(0, rFonts)
            rFonts.set(qn('w:eastAsia'), font_name)
    except Exception:
        pass


def _apply_base_font(doc: Document) -> None:
    """Применение базового шрифта ко всему документу"""

    for paragraph in doc.paragraphs:
        # Убираем интервалы перед и после
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)

        # Междустрочный интервал
        paragraph.paragraph_format.line_spacing = settings.GOST_LINE_SPACING

        # Применяем шрифт к каждому run
        for run in paragraph.runs:
            run.font.name = settings.GOST_FONT_NAME
            run.font.size = Pt(settings.GOST_FONT_SIZE)
            # Для кириллицы
            _set_run_font(run, settings.GOST_FONT_NAME)


def _set_run_font(run, font_name: str) -> None:
    """Безопасная установка шрифта для run"""
    try:
        r = run._element
        rPr = r.get_or_add_rPr()
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = OxmlElement('w:rFonts')
            rPr.insert(0, rFonts)
        rFonts.set(qn('w:ascii'), font_name)
        rFonts.set(qn('w:hAnsi'), font_name)
        rFonts.set(qn('w:eastAsia'), font_name)
        rFonts.set(qn('w:cs'), font_name)
    except Exception:
        pass


def _is_section_heading(text: str) -> bool:
    """Проверка, является ли текст заголовком раздела"""

    text_upper = text.strip().upper()
    return text_upper in SECTION_HEADINGS


def _is_numbered_heading(text: str) -> tuple:
    """
    Проверка, является ли текст нумерованным заголовком

    Returns:
        (is_heading, level) - кортеж из bool и уровня заголовка (1, 2, 3...)
    """

    text = text.strip()

    # Заголовок первого уровня: "1 НАЗВАНИЕ" или "1. НАЗВАНИЕ"
    if re.match(r'^\d+\.?\s+[А-ЯA-Z]', text):
        # Проверяем, что это не подзаголовок (нет второй точки)
        if not re.match(r'^\d+\.\d+', text):
            return (True, 1)

    # Заголовок второго уровня: "1.1 Название"
    if re.match(r'^\d+\.\d+\.?\s+', text):
        if not re.match(r'^\d+\.\d+\.\d+', text):
            return (True, 2)

    # Заголовок третьего уровня: "1.1.1 Название"
    if re.match(r'^\d+\.\d+\.\d+\.?\s+', text):
        return (True, 3)

    return (False, 0)


def _is_appendix_heading(text: str) -> bool:
    """Проверка, является ли текст заголовком приложения"""

    for pattern in APPENDIX_PATTERNS:
        if re.match(pattern, text.strip()):
            return True
    return False


def _format_headings(doc: Document) -> None:
    """Форматирование заголовков"""

    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()

        if not text:
            continue

        # Заголовки разделов (ВВЕДЕНИЕ, ЗАКЛЮЧЕНИЕ и т.д.)
        if _is_section_heading(text):
            _format_section_heading(paragraph)
            continue

        # Приложения
        if _is_appendix_heading(text):
            _format_appendix_heading(paragraph)
            continue

        # Нумерованные заголовки
        is_heading, level = _is_numbered_heading(text)
        if is_heading:
            _format_numbered_heading(paragraph, level)


def _format_section_heading(paragraph) -> None:
    """Форматирование заголовка раздела (ВВЕДЕНИЕ, ЗАКЛЮЧЕНИЕ и т.д.)"""

    # Текст заголовка капсом
    original_text = paragraph.text.strip()
    paragraph.clear()
    run = paragraph.add_run(original_text.upper())

    # Стиль шрифта
    run.font.name = settings.GOST_FONT_NAME
    run.font.size = Pt(settings.GOST_FONT_SIZE)
    run.font.bold = True
    _set_run_font(run, settings.GOST_FONT_NAME)

    # Выравнивание по центру
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Без отступа красной строки
    paragraph.paragraph_format.first_line_indent = Cm(0)
    paragraph.paragraph_format.left_indent = Cm(0)

    # Интервал перед заголовком
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)


def _format_numbered_heading(paragraph, level: int) -> None:
    """Форматирование нумерованного заголовка"""

    for run in paragraph.runs:
        run.font.name = settings.GOST_FONT_NAME
        run.font.bold = True
        run.font.size = Pt(settings.GOST_FONT_SIZE)
        _set_run_font(run, settings.GOST_FONT_NAME)

    # Для заголовков первого уровня - по центру
    if level == 1:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.paragraph_format.first_line_indent = Cm(0)
    else:
        # Для подзаголовков - по ширине с красной строкой
        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        paragraph.paragraph_format.first_line_indent = Cm(settings.GOST_INDENT_CM)


def _format_appendix_heading(paragraph) -> None:
    """Форматирование заголовка приложения"""

    for run in paragraph.runs:
        run.font.name = settings.GOST_FONT_NAME
        run.font.size = Pt(settings.GOST_FONT_SIZE)
        run.font.bold = True
        _set_run_font(run, settings.GOST_FONT_NAME)

    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.first_line_indent = Cm(0)


def _is_list_item(text: str) -> bool:
    """Проверка, является ли текст элементом списка"""

    # Нумерованные списки: 1. 1) а. а) a. a)
    patterns = [
        r'^\d+\.\s',           # 1.
        r'^\d+\)\s',           # 1)
        r'^[а-яa-z]\.\s',      # а. или a.
        r'^[а-яa-z]\)\s',      # а) или a)
        r'^[-–—•]\s',          # маркеры списка
    ]

    text_lower = text.strip().lower()

    for pattern in patterns:
        if re.match(pattern, text_lower):
            return True

    return False


def _format_lists(doc: Document) -> None:
    """Форматирование списков"""

    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()

        if _is_list_item(text):
            # Убираем красную строку
            paragraph.paragraph_format.first_line_indent = Cm(0)
            paragraph.paragraph_format.left_indent = Cm(0)

            # Выравнивание по ширине
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

            # Очистка лишних пробелов и табуляций
            if paragraph.runs:
                first_run = paragraph.runs[0]
                # Заменяем табуляцию на пробел
                first_run.text = first_run.text.replace('\t', ' ')
                # Убираем множественные пробелы
                first_run.text = re.sub(r'\s+', ' ', first_run.text)


def _format_main_text(doc: Document) -> None:
    """Форматирование основного текста"""

    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()

        if not text:
            continue

        # Пропускаем специальные элементы
        if _is_section_heading(text):
            continue

        is_heading, _ = _is_numbered_heading(text)
        if is_heading:
            continue

        if _is_appendix_heading(text):
            continue

        if _is_list_item(text):
            continue

        # Проверяем подписи к таблицам и рисункам
        if re.match(r'^Таблица\s+\d+', text) or re.match(r'^Рисунок\s+\d+', text):
            continue

        # Проверяем формулы
        if _is_formula_line(text):
            continue

        # Обычный текст
        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        paragraph.paragraph_format.first_line_indent = Cm(settings.GOST_INDENT_CM)
        paragraph.paragraph_format.left_indent = Cm(0)


def _format_tables(doc: Document) -> None:
    """Форматирование таблиц по ГОСТ"""

    table_counter = 0

    for table in doc.tables:
        table_counter += 1

        # Форматирование содержимого таблицы
        for row in table.rows:
            for cell in row.cells:
                # Вертикальное выравнивание по центру
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

                for paragraph in cell.paragraphs:
                    # Шрифт в таблице может быть меньше (12pt)
                    for run in paragraph.runs:
                        run.font.name = settings.GOST_FONT_NAME
                        run.font.size = Pt(12)
                        _set_run_font(run, settings.GOST_FONT_NAME)

                    # Междустрочный интервал одинарный в таблицах
                    paragraph.paragraph_format.line_spacing = 1.0
                    paragraph.paragraph_format.space_before = Pt(0)
                    paragraph.paragraph_format.space_after = Pt(0)

        # Центрирование таблицы
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Проверяем и форматируем подписи таблиц
    _format_table_captions(doc)


def _format_table_captions(doc: Document) -> None:
    """Форматирование подписей к таблицам"""

    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()

        # Поиск подписи таблицы
        if re.match(r'^Таблица\s+\d+', text):
            # Форматирование подписи
            for run in paragraph.runs:
                run.font.name = settings.GOST_FONT_NAME
                run.font.size = Pt(settings.GOST_FONT_SIZE)
                run.font.bold = False
                _set_run_font(run, settings.GOST_FONT_NAME)

            # Подпись таблицы - слева без отступа
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            paragraph.paragraph_format.first_line_indent = Cm(0)
            paragraph.paragraph_format.left_indent = Cm(0)
            paragraph.paragraph_format.space_after = Pt(6)


def _format_figures(doc: Document) -> None:
    """Форматирование рисунков по ГОСТ"""

    figure_paragraphs = []

    # Находим все параграфы с изображениями
    for i, paragraph in enumerate(doc.paragraphs):
        # Проверяем наличие изображений
        has_image = _has_drawing(paragraph)

        if has_image:
            # Центрируем рисунок
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.first_line_indent = Cm(0)
            figure_paragraphs.append(i)

    # Форматируем существующие подписи рисунков
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()

        if re.match(r'^Рисунок\s+\d+', text):
            for run in paragraph.runs:
                run.font.name = settings.GOST_FONT_NAME
                run.font.size = Pt(settings.GOST_FONT_SIZE)
                run.font.bold = False
                _set_run_font(run, settings.GOST_FONT_NAME)

            # Подпись рисунка - по центру
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.first_line_indent = Cm(0)
            paragraph.paragraph_format.space_before = Pt(6)
            paragraph.paragraph_format.space_after = Pt(12)

    logger.info(f"Найдено рисунков: {len(figure_paragraphs)}")


def _is_formula_line(text: str) -> bool:
    """Проверка, является ли строка формулой"""

    text = text.strip()

    # Формула с номером в скобках справа: "F = ma, (1)"
    if re.search(r'\(\d+\)\s*$', text):
        return True

    # Формула с нумерацией вида (1.1)
    if re.search(r'\(\d+\.\d+\)\s*$', text):
        return True

    # Строка "где" после формулы
    if text.lower().startswith('где ') or text.lower() == 'где':
        return True

    return False


def _format_formulas(doc: Document) -> None:
    """Форматирование формул по ГОСТ"""

    formula_count = 0

    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()

        if _is_formula_line(text):
            formula_count += 1

            # Формула по центру
            if not text.lower().startswith('где'):
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                paragraph.paragraph_format.first_line_indent = Cm(0)
                paragraph.paragraph_format.space_before = Pt(6)
                paragraph.paragraph_format.space_after = Pt(6)
            else:
                # "где" с красной строкой
                paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                paragraph.paragraph_format.first_line_indent = Cm(settings.GOST_INDENT_CM)

            # Форматирование шрифта
            for run in paragraph.runs:
                run.font.name = settings.GOST_FONT_NAME
                run.font.size = Pt(settings.GOST_FONT_SIZE)
                _set_run_font(run, settings.GOST_FONT_NAME)

    # Также обрабатываем OMML формулы (Office Math)
    _format_omml_formulas(doc)

    logger.info(f"Обработано формул: {formula_count}")


def _format_omml_formulas(doc: Document) -> None:
    """Обработка встроенных формул Office Math (OMML)"""

    omml_count = 0

    # Namespace для математических формул
    MATH_NS = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'

    for paragraph in doc.paragraphs:
        # Проверяем наличие OMML формул через итерацию по элементам
        has_omml = False
        for elem in paragraph._element.iter():
            if elem.tag == f'{MATH_NS}oMath' or 'oMath' in elem.tag:
                has_omml = True
                break

        if has_omml:
            omml_count += 1

            # Центрируем параграф с формулой
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.first_line_indent = Cm(0)
            paragraph.paragraph_format.space_before = Pt(6)
            paragraph.paragraph_format.space_after = Pt(6)

    if omml_count > 0:
        logger.info(f"Найдено OMML формул: {omml_count}")


def _format_bibliography(doc: Document) -> None:
    """Форматирование библиографического списка"""

    bibliography_started = False
    source_count = 0

    for paragraph in doc.paragraphs:
        text = paragraph.text.strip().upper()

        # Начало библиографии
        if any(heading in text for heading in
               ['СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ', 'СПИСОК ЛИТЕРАТУРЫ', 'БИБЛИОГРАФИЧЕСКИЙ СПИСОК']):
            bibliography_started = True
            continue

        # Конец библиографии (начало приложений)
        if bibliography_started and _is_appendix_heading(paragraph.text):
            bibliography_started = False
            continue

        if bibliography_started and paragraph.text.strip():
            source_count += 1
            text = paragraph.text.strip()

            # Убираем множественные пробелы
            text = re.sub(r'\s+', ' ', text)

            # Переформатируем абзац
            paragraph.clear()
            run = paragraph.add_run(text)
            run.font.name = settings.GOST_FONT_NAME
            run.font.size = Pt(settings.GOST_FONT_SIZE)
            _set_run_font(run, settings.GOST_FONT_NAME)

            # Выступающий отступ (висячий отступ)
            paragraph.paragraph_format.first_line_indent = Cm(0)
            paragraph.paragraph_format.left_indent = Cm(0)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    logger.info(f"Обработано источников в библиографии: {source_count}")


def _format_appendixes(doc: Document) -> None:
    """Форматирование приложений"""

    appendix_started = False

    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()

        if _is_appendix_heading(text):
            appendix_started = True
            _format_appendix_heading(paragraph)
            continue

        if appendix_started and text:
            # Применяем базовое форматирование к тексту приложений
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            paragraph.paragraph_format.first_line_indent = Cm(settings.GOST_INDENT_CM)


def _add_page_numbers(doc: Document) -> None:
    """Добавление нумерации страниц снизу по центру"""

    for section in doc.sections:
        footer = section.footer
        footer.is_linked_to_previous = False

        # Очищаем footer если там уже что-то есть
        if footer.paragraphs:
            paragraph = footer.paragraphs[0]
            paragraph.clear()
        else:
            paragraph = footer.add_paragraph()

        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Создаем поле с номером страницы
        run = paragraph.add_run()

        # Настройка шрифта для номера страницы
        run.font.name = settings.GOST_FONT_NAME
        run.font.size = Pt(settings.GOST_FONT_SIZE)

        # XML для поля PAGE
        fld_char_begin = OxmlElement('w:fldChar')
        fld_char_begin.set(qn('w:fldCharType'), 'begin')

        instr_text = OxmlElement('w:instrText')
        instr_text.set(qn('xml:space'), 'preserve')
        instr_text.text = ' PAGE '

        fld_char_separate = OxmlElement('w:fldChar')
        fld_char_separate.set(qn('w:fldCharType'), 'separate')

        fld_char_end = OxmlElement('w:fldChar')
        fld_char_end.set(qn('w:fldCharType'), 'end')

        # Добавляем элементы
        run._r.append(fld_char_begin)
        run._r.append(instr_text)
        run._r.append(fld_char_separate)
        run._r.append(fld_char_end)

    logger.info("Нумерация страниц добавлена")


def _print_statistics(doc: Document) -> None:
    """Вывод статистики форматирования"""

    stats = {
        'center': 0,
        'justify': 0,
        'left': 0,
        'right': 0,
        'lists': 0,
        'headings': 0,
        'tables': len(doc.tables),
        'figures': 0,
        'formulas': 0
    }

    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()

        if not text:
            continue

        if _is_list_item(text):
            stats['lists'] += 1
        elif _is_section_heading(text) or _is_numbered_heading(text)[0]:
            stats['headings'] += 1
        elif _is_formula_line(text):
            stats['formulas'] += 1
        elif paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER:
            stats['center'] += 1
        elif paragraph.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY:
            stats['justify'] += 1
        elif paragraph.alignment == WD_ALIGN_PARAGRAPH.RIGHT:
            stats['right'] += 1
        else:
            stats['left'] += 1

        # Подсчет рисунков
        if _has_drawing(paragraph):
            stats['figures'] += 1

    logger.info("=" * 50)
    logger.info("Статистика форматирования:")
    logger.info(f"  Заголовков: {stats['headings']}")
    logger.info(f"  По центру: {stats['center']}")
    logger.info(f"  По ширине: {stats['justify']}")
    logger.info(f"  По левому краю: {stats['left']}")
    logger.info(f"  Элементов списка: {stats['lists']}")
    logger.info(f"  Таблиц: {stats['tables']}")
    logger.info(f"  Рисунков: {stats['figures']}")
    logger.info(f"  Формул: {stats['formulas']}")
    logger.info("=" * 50)


def _has_drawing(paragraph) -> bool:
    """Проверка наличия рисунка в параграфе"""
    for elem in paragraph._element.iter():
        if 'drawing' in elem.tag.lower() or 'pict' in elem.tag.lower():
            return True
    return False
