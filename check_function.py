# check_function.py

import os
import traceback
import win32com.client
from win32com.client import constants
import re


def check_headers(file_path):
    discrepancies = {}

    # Нормализуем путь к файлу
    normalized_path = os.path.normpath(os.path.abspath(file_path))

    # Список требуемых заголовков, исключая "СОДЕРЖАНИЕ"
    required_headings = ['ВВЕДЕНИЕ', 'ЗАКЛЮЧЕНИЕ']

    # Списки для хранения найденных заголовков, их позиций и страниц
    found_headings = []
    heading_positions = {}
    heading_pages = {}

    # Открываем Word через COM
    word_app = win32com.client.gencache.EnsureDispatch('Word.Application')
    word_app.Visible = False  # Не отображать окно Word

    # Открываем документ
    try:
        doc = word_app.Documents.Open(normalized_path)
    except Exception as e:
        print(f"Ошибка при открытии документа: {e}")
        traceback.print_exc()
        word_app.Quit()
        return {}

    try:
        # Проверяем общее количество страниц в документе
        total_pages = doc.ComputeStatistics(2)  # wdStatisticPages = 2

        # Если документ более 10 страниц, проверяем наличие заголовка "СОДЕРЖАНИЕ"
        content_heading_found = False
        if total_pages > 10:
            for para in doc.Content.Paragraphs:
                text = para.Range.Text.strip()
                if text.upper() == "СОДЕРЖАНИЕ":
                    content_heading_found = True
                    break

        # Проходим по всем абзацам основного текста документа
        for para in doc.Content.Paragraphs:
            text = para.Range.Text.strip()
            if not text:
                continue  # Пропускаем пустые абзацы

            # Приводим текст к верхнему регистру для корректного сравнения
            heading_text = text.upper()

            # Пропускаем заголовок "СОДЕРЖАНИЕ" и его форматирование, если он найден в документе
            if content_heading_found and heading_text == "СОДЕРЖАНИЕ":
                continue

            # Проверяем, является ли абзац заголовком (центрирован и входит в список требуемых заголовков)
            alignment = para.Format.Alignment
            if alignment == constants.wdAlignParagraphCenter and heading_text in required_headings:
                # Добавляем найденный заголовок в списки
                found_headings.append(heading_text)
                heading_positions[heading_text] = len(found_headings) - 1  # Позиция заголовка в списке

                # Получаем номер страницы параграфа
                page_number = para.Range.Information(constants.wdActiveEndAdjustedPageNumber)
                heading_pages[heading_text] = page_number
        # Проверяем наличие всех необходимых заголовков
        for heading in required_headings:
            if heading not in found_headings:
                discrepancies.setdefault('Общие', []).append(f"Раздел '{heading}' не найден в документе.")

        # Проверяем правильный порядок заголовков
        for i, heading in enumerate(required_headings):
            if heading in found_headings:
                found_index = heading_positions[heading]
                if found_index != i:
                    correct_prev_heading = required_headings[i - 1] if i > 0 else None
                    correct_next_heading = required_headings[i + 1] if i + 1 < len(required_headings) else None
                    message = f"Раздел '{heading}' находится не на правильном месте."
                    if correct_prev_heading and correct_prev_heading in found_headings:
                        message += f" Он должен быть после '{correct_prev_heading}'."
                    elif correct_next_heading and correct_next_heading in found_headings:
                        message += f" Он должен быть перед '{correct_next_heading}'."
                    page = heading_pages.get(heading, 'Неизвестно')
                    discrepancies.setdefault(page, []).append(message)

    except Exception as e:
        print(f"Ошибка при обработке документа: {e}")
        traceback.print_exc()
    finally:
        # Закрываем документ и приложение Word
        doc.Close(False)
        word_app.Quit()

    return discrepancies


def check_formatting(file_path):
    discrepancies = {}
    normalized_path = os.path.normpath(os.path.abspath(file_path))
    word_app = win32com.client.gencache.EnsureDispatch('Word.Application')
    word_app.Visible = False

    try:
        doc = word_app.Documents.Open(normalized_path)
    except Exception as e:
        print(f"Ошибка при открытии документа: {e}")
        traceback.print_exc()
        word_app.Quit()
        return {}

    try:
        # Проверка размеров полей (для первого раздела)
        section = doc.Sections(1)
        page_setup = section.PageSetup

        # Поля в миллиметрах (1 пункт = 0.3527778 мм)
        left_margin_mm = page_setup.LeftMargin * 0.3527778
        right_margin_mm = page_setup.RightMargin * 0.3527778
        top_margin_mm = page_setup.TopMargin * 0.3527778
        bottom_margin_mm = page_setup.BottomMargin * 0.3527778

        margin_discrepancies = []

        if abs(left_margin_mm - 30.0) > 0.5:
            margin_discrepancies.append(f"Левое поле должно быть 30 мм.")
        if abs(right_margin_mm - 15.0) > 0.5:
            margin_discrepancies.append(f"Правое поле должно быть 15 мм.")
        if abs(top_margin_mm - 20.0) > 0.5:
            margin_discrepancies.append(f"Верхнее поле должно быть 20 мм.")
        if abs(bottom_margin_mm - 20.0) > 0.5:
            margin_discrepancies.append(f"Нижнее поле должно быть 20 мм.")

        if margin_discrepancies:
            discrepancies.setdefault('Общие', []).extend(margin_discrepancies)

        # Эталонное значение абзацного отступа
        standard_indent_cm = 1.25  # Значение отступа по умолчанию
        other_lines_indent_cm = 0  # для остальных строк
        allowed_endings = {',', ';'}
        paragraphs = list(doc.Content.Paragraphs)

        # Найдем страницы, где "СОДЕРЖАНИЕ" находится вверху страницы
        pages_with_soderzhanie = set()
        paragraphs = list(doc.Content.Paragraphs)

        for i, para in enumerate(paragraphs):
            text = para.Range.Text.strip().upper()
            if text == 'СОДЕРЖАНИЕ' or text == 'СПИСОК ИСПОЛЬЗУЕМЫХ ИСТОЧНИКОВ':
                # Проверяем, является ли этот абзац первым на странице
                page_number = para.Range.Information(constants.wdActiveEndAdjustedPageNumber)
                is_first_on_page = para.Range.Information(constants.wdFirstCharacterLineNumber) == 1
                if is_first_on_page:
                    pages_with_soderzhanie.add(page_number)

        # Проходим по всем абзацам основного текста документа
        for para in paragraphs:
            text = para.Range.Text.strip()
            if not text:
                continue  # Пропускаем пустые абзацы

            # Получаем номер страницы параграфа
            page_number = para.Range.Information(constants.wdActiveEndAdjustedPageNumber)

            # Пропускаем проверку на страницах с "СОДЕРЖАНИЕМ" вверху
            if page_number in pages_with_soderzhanie:
                continue  # Пропускаем абзацы на этой странице

            # Пропускаем абзацы внутри таблиц
            if para.Range.Tables.Count > 0:
                continue  # Пропускаем абзацы в таблицах

            # Пропускаем абзацы без буквенно-цифровых символов
            if not any(char.isalnum() for char in text):
                continue

            # Инициализируем список несоответствий для страницы
            discrepancies.setdefault(page_number, [])

            # Проверка списка
            line = text
            snippet = line[:30] + '...' if len(line) > 30 else line

            # Проверяем, является ли абзац частью списка
            if para.Range.ListFormat.ListType != 0:
                # Абзац из автоматического списка Word
                current_ending = None
                item_text = line.strip()

                # Проверка окончания для автоматического списка
                if para.Next() is None or para.Next().Range.ListFormat.ListType == 0:
                    # Если это последний элемент списка
                    if not item_text.endswith('.'):
                        discrepancies[page_number].append(
                            f"Последний элемент списка должен заканчиваться точкой: '{snippet}'"
                        )
                elif item_text[-1] not in allowed_endings:
                    discrepancies[page_number].append(
                        f"Некорректное окончание в элементе списка: '{snippet}'"
                    )

                indent = para.Format.LeftIndent / 28.35  # В см
                if indent != 0:
                    discrepancies[page_number].append(
                        f"Неверный отступ в элементе списка '{snippet}': ожидалось {standard_indent_cm} см"
                    )

            elif line.startswith("- ") or re.match(r'^([а-я])\)', line) or re.match(r'^(\d+)\)', line):
                # Абзац из ручного списка
                bullet = line[:2].strip()  # Маркер списка
                item_text = line[2:].strip()

                # Проверка окончания для ручного списка
                if para.Next() is None or not re.match(r'^(-|[а-я]\)|\d+\))', para.Next().Range.Text.strip()):
                    # Если это последний элемент списка
                    if not item_text.endswith('.'):
                        discrepancies[page_number].append(
                            f"Последний элемент списка должен заканчиваться точкой: '{snippet}'"
                        )
                elif item_text[-1] not in allowed_endings:
                    discrepancies[page_number].append(
                        f"Некорректное окончание в элементе списка: '{snippet}'"
                    )

                indent = para.Format.LeftIndent / 28.35  # В см
                if indent != 0:
                    discrepancies[page_number].append(
                        f"Неверный отступ в элементе списка '{snippet}': ожидалось {standard_indent_cm} см"
                    )

            # Проверка абзацев, начинающихся с "Листинг Х - " или "Таблица Х - "
            lower_text = text.lower()
            if (lower_text.startswith('листинг') or lower_text.startswith('таблица')) and '–' in text:
                # Проверяем, что отступ равен 0 см
                first_line_indent = para.Format.FirstLineIndent  # В пунктах
                first_line_indent_cm = first_line_indent * 0.03527778  # Переводим в сантиметры

                if abs(first_line_indent_cm) > 0.05:
                    snippet = text[:30] + '...' if len(text) > 30 else text
                    discrepancies[page_number].append(
                        f"Отступ должен быть 0 см в абзаце '{snippet}'"
                    )
                continue  # Продолжаем к следующему абзацу

            # Проверка на текст вида "Рисунок X -"
            if lower_text.startswith('рисунок') and '–' in text:
                continue  # Пропускаем подписи к рисункам

            # Проверка абзацного отступа для остальных абзацев
            first_line_indent = para.Format.FirstLineIndent  # В пунктах
            first_line_indent_cm = first_line_indent * 0.03527778  # Переводим в сантиметры

            # Вывод отладочной информации
            snippet = text[:30] + '...' if len(text) > 30 else text

            # Пропускаем абзацный отступ, если текст жирный и выровнен по центру
            is_bold_and_centered = (para.Range.Font.Bold == -1) and (
                        para.Format.Alignment == constants.wdAlignParagraphCenter)
            if is_bold_and_centered:
                continue  # Пропускаем этот абзац

            # Проверка абзацного отступа для первой строки
            first_line_indent = para.Format.FirstLineIndent * 0.03527778  # В сантиметрах

            if abs(first_line_indent - first_line_indent_cm) > 0.01:
                snippet = text[:30] + '...' if len(text) > 30 else text
                discrepancies[page_number].append(
                    f"Отступ первой строки должен быть {first_line_indent_cm} см в абзаце '{snippet}'"
                )

            # Проверка отступа для второй и последующих строк
            first_line_indent_cm = para.Format.FirstLineIndent / 28.35  # Преобразование в сантиметры
            left_indent_cm = para.Format.LeftIndent / 28.35  # Преобразование в сантиметры

            # Подсчёт строк в абзаце
            lines_in_paragraph = para.Range.ComputeStatistics(win32com.client.constants.wdStatisticLines)

            # Проверяем отступ для абзацев с более чем одной строкой
            if lines_in_paragraph > 1 and abs(left_indent_cm - other_lines_indent_cm) > 0.01:
                snippet = text[:30] + '...' if len(text) > 30 else text
                discrepancies[page_number].append(
                    f"Отступ для второй и последующих строк должен быть {other_lines_indent_cm} см в абзаце '{snippet}'"
                )

            # Проверка междустрочного интервала
            line_spacing_rule = para.Format.LineSpacingRule
            line_spacing = para.Format.LineSpacing

            is_correct_spacing = False
            # Проверяем, что междустрочный интервал равен 1,5
            if line_spacing_rule == constants.wdLineSpace1pt5:
                is_correct_spacing = True
            elif line_spacing_rule == constants.wdLineSpaceMultiple and abs(line_spacing - 1.5) < 0.05:
                is_correct_spacing = True

            if not is_correct_spacing:
                discrepancies[page_number].append(
                    f"Междустрочный интервал должен быть 1,5 в абзаце '{snippet}'."
                )

            # Проверка выравнивания абзаца
            alignment = para.Format.Alignment
            if alignment != constants.wdAlignParagraphJustify:
                snippet = text[:30] + '...' if len(text) > 30 else text
                discrepancies[page_number].append(
                    f"Абзац должен быть выровнен по ширине: '{snippet}'"
                )

            # Проверка шрифта
            font_name = para.Range.Font.Name
            if font_name != 'Times New Roman':
                discrepancies[page_number].append(
                    f"Шрифт должен быть 'Times New Roman' в абзаце '{snippet}'."
                )

            for word in para.Range.Words:
                color_value = word.Font.Color
                if color_value not in [0, 1, 9999999]:  # 0 и 1 - черный и автоматический черный в Word
                    red = color_value & 0xFF
                    green = (color_value >> 8) & 0xFF
                    blue = (color_value >> 16) & 0xFF
                    if (red, green, blue) != (0, 0, 0):
                        snippet = text[:30] + '...' if len(text) > 30 else text
                        discrepancies[page_number].append(
                            f"Цвет текста должен быть чёрным в абзаце '{snippet}'."
                        )
                        # color_issue_found = True
                        break  # Прерываем проверку цвета слов, если уже найдена ошибка в абзаце

            # Проверка размера шрифта для каждого слова
            words = para.Range.Words
            font_size_correct = True  # Флаг для отслеживания, все ли слова соответствуют размеру

            for word in words:
                word_size = word.Font.Size
                if word_size < 12:
                    font_size_correct = False  # Если хотя бы одно слово имеет размер меньше 12, отметим ошибку
                    break  # Прерываем цикл, так как ошибка уже найдена

            # Если размер шрифта не соответствует требованиям для хотя бы одного слова
            if not font_size_correct:
                discrepancies[page_number].append(
                    f"Размер шрифта должен быть не менее 12 пт в абзаце '{snippet}'."
                )

        # Удаляем страницы без несоответствий
        pages_to_remove = []
        for page, issues in discrepancies.items():
            if isinstance(page, int) and not issues:
                pages_to_remove.append(page)
        for page in pages_to_remove:
            del discrepancies[page]

        # Если несоответствий нет, добавляем сообщение
        # if not discrepancies:
        #     discrepancies['Общие'] = ["Форматирование соответствует требованиям."]

    except Exception as e:
        print(f"Ошибка при проверке форматирования: {e}")
        traceback.print_exc()
    finally:
        doc.Close(False)
        word_app.Quit()

    return discrepancies


def check_illustrations(file_path):
    discrepancies = {}

    try:
        # Открываем документ Word
        word_app = win32com.client.Dispatch("Word.Application")
        word_app.Visible = False
        doc = word_app.Documents.Open(file_path)

        # Инициализация для проверки номеров рисунков
        illustration_number = 1

        # Массив для возможных форм упоминания "рисунок"
        illustration_name_mas = ["рисунок", "рисунка", "рисункам", "рисунок", "рисунками", "рисунке", "рисунках"]

        flag = True

        for inline_shape in doc.InlineShapes:
            if inline_shape.Type == constants.wdInlineShapePicture:
                inline_range = inline_shape.Range
                page_number = inline_range.Information(win32com.client.constants.wdActiveEndPageNumber)

                # Проверка правильности нумерации рисунков
                title_paragraph = inline_range.Paragraphs(1).Next()  # Следующий абзац после рисунка

                if title_paragraph is not None:
                    text = title_paragraph.Range.Text.strip()
                    snippet = text[:30] + '...' if len(text) > 30 else text

                    # Формируем правильный заголовок
                    expected_label = f"Рисунок {illustration_number} – "

                    if text.startswith(expected_label):
                        title = text[len(expected_label):]

                        # Проверка, что заголовок начинается с заглавной буквы и не оканчивается точкой
                        if not title[0].isupper() or title.endswith('.'):
                            discrepancies.setdefault(page_number, []).append(
                                f"Неправильное оформление наименования рисунка {illustration_number}: '{snippet}'"
                            )
                        if not title_paragraph.Format.Alignment == constants.wdAlignParagraphCenter:
                            discrepancies.setdefault(page_number, []).append(
                                f"Неправильное выравнивание подписи рисунка {illustration_number}: '{snippet}'"
                            )
                    else:
                        flag = False
                        discrepancies.setdefault(page_number, []).append(
                            f"Неправильный заголовок после рисунка {illustration_number}. Правильный заголовок: Рисунок {illustration_number} – ... "
                        )

                # Проверка упоминания рисунка в абзаце перед рисунком
                preceding_paragraph = inline_range.Paragraphs(1).Previous()

                if preceding_paragraph is not None:
                    preceding_text = preceding_paragraph.Range.Text.strip().lower()

                    if any(f"{form} {illustration_number}" in preceding_text for form in illustration_name_mas):
                        mentioned = True
                    else:
                        mentioned = False

                    if not mentioned:
                        discrepancies.setdefault(page_number, []).append(
                            f"Отсутствует упоминание рисунка {illustration_number} в абзаце перед ним."
                        )

                # Увеличение номера рисунка
                illustration_number += 1

        # # Проверка сквозной нумерации рисунков
        # actual_illustration_count = sum(
        #     1 for inline_shape in doc.InlineShapes if inline_shape.Type == constants.wdInlineShapePicture
        # )
        # expected_illustration_count = illustration_number - 1

        if not flag:
            discrepancies.setdefault("Общие", []).append(
                f"Нарушена сквозная нумерация иллюстраций."
            )

        # Закрываем документ
        doc.Close(SaveChanges=False)
        word_app.Quit()

    except Exception as e:
        print(f"Ошибка при проверке иллюстраций: {e}")
        traceback.print_exc()

    return discrepancies


def check_tables(file_path):
    discrepancies = {}

    try:
        # Открываем документ Word
        word_app = win32com.client.Dispatch("Word.Application")
        word_app.Visible = False
        doc = word_app.Documents.Open(file_path)

        # Инициализация для проверки номеров таблиц и листингов
        table_number = 1
        listing_number = 1

        table_name_mas = ["таблица", "таблицы", "таблиц", "таблице", "таблицам", "таблицу", "таблицы", "таблицей", "таблицами", "таблице", "таблицах"]
        listing_name_mas = ["листинг", "листинги", "листинга", "листингов", "листингу", "листингам", "листингом", "листингами", "листинге", "листингах"]

        for table in doc.Tables:
            table_range = table.Range
            table_position = table_range.Start

            # Определяем страницу, на которой находится таблица
            start_page = table_range.Information(3)  # 3 соответствует wdStartPageNumber

            # Определяем, таблица это или листинг (если только один столбец, считаем листингом)
            is_listing = table.Columns.Count == 1
            expected_label = f"Листинг {listing_number} – " if is_listing else f"Таблица {table_number} – "

            # Проверка текста перед таблицей/листингом
            preceding_paragraph = table_range.Paragraphs(1).Previous()
            if preceding_paragraph is not None:
                text = preceding_paragraph.Range.Text.strip()

                # Проверка формата текста перед таблицей/листингом
                if text.startswith(expected_label):
                    title = text[len(expected_label):]

                    # Проверка, что заголовок начинается с заглавной буквы и не оканчивается точкой
                    if not title[0].isupper() or title.endswith('.'):
                        discrepancies.setdefault(start_page, []).append(
                            f"Неправильное оформление наименования {'листинга' if is_listing else 'таблицы'} {listing_number if is_listing else table_number}: '{text}'"
                        )
                else:
                    discrepancies.setdefault(start_page, []).append(
                        f"Неправильный заголовок перед {'листингом' if is_listing else 'таблицей'} {listing_number if is_listing else table_number}. Правильный заголовок: {'Листинг' if is_listing else 'Таблица'} {listing_number if is_listing else table_number} – ... "
                    )

                # Проверка упоминания таблицы/листинга в предыдущих абзацах
                mentioned = False
                previous_paragraph = preceding_paragraph
                while previous_paragraph is not None:
                    previous_text = previous_paragraph.Range.Text.strip().lower()

                    # Проверяем, что абзац не является заголовком таблицы или листинга
                    if previous_text and not previous_text.startswith(
                            f"таблица {table_number} –") and not previous_text.startswith(
                            f"листинг {listing_number} –"):
                        # Проверяем, что ближайший предыдущий абзац содержит упоминание таблицы или листинга
                        if (not is_listing and any(
                                f"{form} {table_number}" in previous_text for form in table_name_mas)) or \
                                (is_listing and any(
                                    f"{form} {listing_number}" in previous_text for form in listing_name_mas)):
                            mentioned = True
                            break
                    previous_paragraph = previous_paragraph.Previous()

                if not mentioned:
                    discrepancies.setdefault(start_page, []).append(
                        f"Отсутствует упоминание {'листинга' if is_listing else 'таблицы'} {listing_number if is_listing else table_number} перед его заголовком."
                    )

            # Увеличение номера таблицы или листинга
            if is_listing:
                listing_number += 1
            else:
                table_number += 1

        # Проверка сквозной нумерации таблиц и листингов
        actual_table_count = len(doc.Tables)
        if actual_table_count != table_number - 1:
            discrepancies.setdefault("Общие", []).append(
                "Нарушена сквозная нумерация таблиц."
            )

        # Закрываем документ
        doc.Close(SaveChanges=False)
        word_app.Quit()

    except Exception as e:
        print(f"Ошибка при проверке таблиц: {e}")
        traceback.print_exc()

    return discrepancies



def check_page_numbering(file_path):
    discrepancies = {"Общие": []}
    try:
        # Открываем документ Word
        word_app = win32com.client.Dispatch("Word.Application")
        word_app.Visible = False
        doc = word_app.Documents.Open(file_path)

        # Получаем общее количество страниц в документе
        page_count = doc.ComputeStatistics(win32com.client.constants.wdStatisticPages)
        incorrect_numbering_found = False

        # Проверяем нумерацию страниц, начиная с первой страницы
        for page_number in range(1, page_count + 1):

            # Инициализируем подмножество для текущей страницы, если оно не существует
            if page_number not in discrepancies:
                discrepancies[page_number] = []

            try:
                # Переходим к нужной секции и проверяем колонтитул
                section_index = doc.Range().GoTo(win32com.client.constants.wdGoToPage,
                                                 win32com.client.constants.wdGoToAbsolute,
                                                 page_number).Sections(1).Index
                footer = doc.Sections(section_index).Footers(win32com.client.constants.wdHeaderFooterPrimary)
                footer_text = footer.Range.Text.strip()

                # Проверка на отсутствие нумерации
                if not footer_text:
                    discrepancies[page_number].append(f"Отсутствует нумерация на странице {page_number}.")
                    incorrect_numbering_found = True
                    continue

                # Проверка выравнивания номера страницы по центру
                if footer.Range.ParagraphFormat.Alignment != win32com.client.constants.wdAlignParagraphCenter:
                    discrepancies[page_number].append(
                        f"Номер на странице {page_number} проставлен не по центру."
                    )
                    incorrect_numbering_found = True

            except Exception as inner_e:
                discrepancies[page_number].append(f"Ошибка при проверке страницы {page_number}: {inner_e}")

        # Если ошибок не найдено, выводим сообщение о корректной нумерации
        if not incorrect_numbering_found:
            discrepancies["Общие"].append(
                "Нумерация корректна. Убедитесь в отсутствии номера страницы на титульном листе.")

        # Закрываем документ
        doc.Close(SaveChanges=False)
        word_app.Quit()

    except Exception as e:
        print(f"Ошибка при проверке нумерации страниц: {e}")
        discrepancies["Общие"].append(f"Ошибка при проверке нумерации страниц: {e}")

    return discrepancies


def check(file_path):
    try:
        # Нормализуем путь к файлу
        normalized_path = os.path.normpath(os.path.abspath(file_path))

        # Открытие документа
        # word_app = win32com.client.Dispatch('Word.Application')
        # word_app.Visible = False
        # doc = word_app.Documents.Open(normalized_path)
        #
        # # Проверка количества страниц
        # page_count = check_page_count(doc)  # Теперь используем нашу функцию для подсчёта
        # print(f"Количество страниц: {page_count}")
        #
        # if page_count > 100:
        #     return {"Общие": ["Объём документа превышает допустимый, выберите другой документ"]}

        discrepancies = {}

        # Проверка заголовков
        header_discrepancies = check_headers(normalized_path)
        for key, value in header_discrepancies.items():
            discrepancies.setdefault(key, []).extend(value)

        # Проверка форматирования
        formatting_discrepancies = check_formatting(normalized_path)
        for key, value in formatting_discrepancies.items():
            if key in discrepancies:
                discrepancies[key].extend(value)
            else:
                discrepancies[key] = value

        # Проверка иллюстраций
        illustration_discrepancies = check_illustrations(normalized_path)
        for key, value in illustration_discrepancies.items():
            if key in discrepancies:
                discrepancies[key].extend(value)
            else:
                discrepancies[key] = value

        # Проверка таблиц
        table_discrepancies = check_tables(normalized_path)
        for key, value in table_discrepancies.items():
            discrepancies.setdefault(key, []).extend(value)

        # Проверка нумерации страниц
        pages_discrepancies = check_page_numbering(normalized_path)
        for key, value in pages_discrepancies.items():
            discrepancies.setdefault(key, []).extend(value)

        if not discrepancies:
            discrepancies['Общие'] = ["Несоответствий не найдено."]

        # Закрытие документа
        # doc.Close(False)
        # word_app.Quit()

        return discrepancies
    except Exception as e:
        print(f"Ошибка при проверке файла: {e}")
        traceback.print_exc()
        return None


# def get_page_count(doc):
#     try:
#         # Используем метод ComputeStatistics для подсчета страниц
#         return doc.ComputeStatistics(win32com.client.constants.wdStatisticPages)
#     except Exception as e:
#         print(f"Ошибка при подсчете страниц: {e}")
#         return 0