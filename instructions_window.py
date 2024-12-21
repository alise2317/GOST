import customtkinter as CTk
import re


def show_instructions_window(parent):
    # Создаем новое окно
    instructions_window = CTk.CTkToplevel(parent)
    instructions_window.title("Руководство пользователя")
    instructions_window.geometry("900x700")
    instructions_window.configure(fg_color="#E6E6FA")
    instructions_window.resizable(True, True)
    instructions_window.minsize(600, 400)

    # Окно появится на переднем плане, но не будет оставаться всегда сверху
    instructions_window.attributes("-topmost", True)
    instructions_window.after(100, lambda: instructions_window.attributes("-topmost", False))

    # Создаем Frame внутри instructions_window
    main_frame = CTk.CTkFrame(instructions_window, fg_color="#E6E6FA")
    main_frame.pack(fill="both", expand=True)

    # Настраиваем строки и столбцы для растягивания в main_frame
    main_frame.grid_rowconfigure(0, weight=0)  # Для заголовка
    main_frame.grid_rowconfigure(1, weight=1)  # Для текстового поля
    main_frame.grid_columnconfigure(0, weight=1)

    # Добавляем заголовок
    header_label = CTk.CTkLabel(master=main_frame, text="Руководство пользователя", font=("Arial", 40, "bold"))
    header_label.grid(row=0, column=0, pady=(10, 0))

    # Создаем текстовое поле с возможностью прокрутки
    instructions_textbox = CTk.CTkTextbox(master=main_frame, font=("Arial", 14), wrap='word', fg_color="#E6E6FA",
                                          border_width=0)
    instructions_textbox.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")

    # Получаем доступ к внутреннему виджету Text
    tk_text_widget = instructions_textbox._textbox

    # Создаем теги для различных стилей
    tk_text_widget.tag_configure('heading', font=("Arial", 24, "bold"))
    tk_text_widget.tag_configure('normal', font=("Arial", 18))
    tk_text_widget.tag_configure('list_item', font=("Arial", 18), lmargin1=30, lmargin2=50)
    tk_text_widget.tag_configure('bold', font=("Arial", 18, "bold", "italic"))
    tk_text_widget.tag_configure('italic', font=("Arial", 18, "italic"))

    # Определяем список абзацев
    text_paragraphs = [
        "Привет, студент! Ты попал в приложение, которое поможет сэкономить твоё ценное время и быстро расправиться с отчётом по курсовой, лабораторной и любой другой работе.",
        "Ниже приведено подробное руководство пользователя, которое ответит на все интересующие тебя вопросы.",
        "Данное приложение предназначено для проверки документов Microsoft Word в формате `.docx` на соответствие требованиям стандарта ГОСТ 7.32-2017.",
        "Основные функции приложения:",
        "- Выбор файла для проверки",
        "- Проверка документа на соответствие ГОСТ 7.32-2017",
        "- Просмотр результатов проверки",
        "- Доступ к оригинальному тексту ГОСТ 7.32-2017",
        "- Получение справочной информации",
        "- Получение инструкции по использованию",
        "",
        "1. Выбор файла для проверки",
        "Перед выбором файла для проверки необходимо сохранить несохранённые файлы, открытые в приложении Word и закрыть приложение",
        "Чтобы проверить файл, необходимо нажать на иконку рядом с полем ввода в центре экрана. В открывшемся окне перейдите к расположению вашего файла. Выберите файл и нажмите \"Открыть\".",
        "**Важно: Приложение поддерживает только файлы формата .docx. Файлы других форматов не могут быть проверены (всё равно потом ошибки в файле docx исправлять...).**",
        "Название выбранного файла отобразится в поле ввода. Если название файла длиннее 30 символов, используйте горизонтальную полосу прокрутки под полем ввода; она поможет просмотреть полное название файла и удостовериться, что выбран нужный документ.",
        "",
        "2. Проверка документа",
        "После выбора файла доступной станет кнопка \"Отправить на проверку\"; после её нажатия выбранный файл начнёт обрабатываться.",
        "*При больших объёмах исходного файла стоит запастись терпением; обработка может занять больше времени, чем обычно.*",
        "",
        "3. Просмотр результатов проверки",
        "Когда обработка выбранного пользователем файла закончится, станет доступна кнопка \"Просмотр результатов\". При нажатии на неё открывается отдельное окно, в котором представлен отчёт о проверке.",
        "Отчёт можно просмотреть прямо в появившемся окне либо сохранить его на свой компьютер. Сохранение доступно в форматах docx и txt.",
        "",
        "Помимо основного функционала, осуществляющего проверку пользовательского файла, есть несколько дополнительных кнопок, при нажатии на которые вы сможете ознакомиться с дополнительной информацией.",
        "4. Доступ к оригинальному тексту ГОСТ 7.32-2017",
        "При желании ознакомиться с оригинальным текстом ГОСТа для более детального понимания требований необходимо нажать на кнопку \"ГОСТ\" в левом верхнем углу приложения. После нажатия появится диалоговое окно, в котором можно выбрать место для сохранения документа с оригинальным текстом ГОСТ. Документ будет сохранён в формате PNG.",
        "",
        "5. Получение справочной информации",
        "Данное приложение в текущей версии не проверяет пользовательские документы на соответствие абсолютно всем пунктам ГОСТ 7.32-2017. Чтобы ознакомиться с перечнем проверяемых пунктов ГОСТ, необходимо нажать на иконку с информацией в правом верхнем углу приложения.",
        "",
        "6. Получение инструкции по использованию",
        "Для получения подробных инструкций по использованию приложения необходимо нажать на иконку со знаком вопроса в правом верхнем углу приложения. После нажатия вы попадёте на текущую страницу.",
        "*Благодарим за использование нашего приложения! Желаем успешной работы и соответствия всем необходимым стандартам.*",
    ]

    # Определяем соответствие между функциями и метками разделов
    function_sections = {
        'Выбор файла для проверки': 'section1',
        'Проверка документа на соответствие ГОСТ 7.32-2017': 'section2',
        'Просмотр результатов проверки': 'section3',
        'Доступ к оригинальному тексту ГОСТ 7.32-2017': 'section4',
        'Получение справочной информации': 'section5',
        'Получение инструкции по использованию': 'section6',
    }

    # Словарь для хранения тегов функций
    function_tags = {}

    # Вставляем текст с применением стилей
    for paragraph in text_paragraphs:
        # Убираем пробелы в начале и конце
        paragraph = paragraph.strip()
        if not paragraph:
            # Пустая строка для разделения абзацев
            tk_text_widget.insert('end', '\n')
            continue

        # Определяем стиль абзаца
        if re.match(r'^\d+(\.\d+)*', paragraph):  # Начинается с цифры
            tag = 'heading'
        elif re.match(r'^-', paragraph):  # Начинается с дефиса
            tag = 'list_item'
        elif paragraph.startswith("**") and paragraph.endswith("**"):  # Жирный текст
            paragraph = paragraph.strip("**")
            tag = 'bold'
        elif paragraph.startswith("*") and paragraph.endswith("*"):  # Курсив
            paragraph = paragraph.strip("*")
            tag = 'italic'
        else:
            tag = 'normal'

        # Получаем текущий индекс перед вставкой
        start_index = tk_text_widget.index('end-1c')

        # Проверяем, является ли абзац элементом списка в "Основных функциях приложения"
        if paragraph.strip('- ') in function_sections:
            function_name = paragraph.strip('- ')
            tag_name = function_sections[function_name]
            function_tags[function_name] = tag_name

            # Вставляем текст
            tk_text_widget.insert('end', paragraph + '\n\n')

            # Получаем индекс после вставки
            end_index = tk_text_widget.index('end-1c')

            # Применяем тег форматирования
            tk_text_widget.tag_add(tag, start_index, end_index)

            # Добавляем тег для кликабельности
            tk_text_widget.tag_add(tag_name, start_index, end_index)

            # Настраиваем тег как кликабельный (синие и подчеркнутые ссылки)
            tk_text_widget.tag_configure(tag_name, foreground="black", underline=True)

            # Определяем обработчик события клика
            def make_click_handler(name):
                return lambda event: scroll_to_section(name)

            # Привязываем событие клика к тегу
            tk_text_widget.tag_bind(tag_name, '<Button-1>', make_click_handler(tag_name))

        elif paragraph.startswith('1. '):
            # Это раздел "1. Выбор файла для проверки"
            # Вставляем текст
            tk_text_widget.insert('end', paragraph + '\n\n')
            # Получаем индекс после вставки
            end_index = tk_text_widget.index('end-1c')
            # Применяем тег форматирования
            tk_text_widget.tag_add(tag, start_index, end_index)
            # Ставим метку для перехода
            tk_text_widget.mark_set('section1', start_index)

        elif paragraph.startswith('2. '):
            # Раздел "2. Проверка документа"
            tk_text_widget.insert('end', paragraph + '\n\n')
            end_index = tk_text_widget.index('end-1c')
            tk_text_widget.tag_add(tag, start_index, end_index)
            tk_text_widget.mark_set('section2', start_index)

        elif paragraph.startswith('3. '):
            # Раздел "3. Просмотр результатов проверки"
            tk_text_widget.insert('end', paragraph + '\n\n')
            end_index = tk_text_widget.index('end-1c')
            tk_text_widget.tag_add(tag, start_index, end_index)
            tk_text_widget.mark_set('section3', start_index)

        elif paragraph.startswith('4. '):
            # Раздел "4. Доступ к оригинальному тексту ГОСТ 7.32-2017"
            tk_text_widget.insert('end', paragraph + '\n\n')
            end_index = tk_text_widget.index('end-1c')
            tk_text_widget.tag_add(tag, start_index, end_index)
            tk_text_widget.mark_set('section4', start_index)

        elif paragraph.startswith('5. '):
            # Раздел "5. Получение справочной информации"
            tk_text_widget.insert('end', paragraph + '\n\n')
            end_index = tk_text_widget.index('end-1c')
            tk_text_widget.tag_add(tag, start_index, end_index)
            tk_text_widget.mark_set('section5', start_index)

        elif paragraph.startswith('6. '):
            # Раздел "6. Получение инструкции по использованию"
            tk_text_widget.insert('end', paragraph + '\n\n')
            end_index = tk_text_widget.index('end-1c')
            tk_text_widget.tag_add(tag, start_index, end_index)
            tk_text_widget.mark_set('section6', start_index)

        else:
            # Обычная вставка текста
            tk_text_widget.insert('end', paragraph + '\n\n')
            # Получаем индекс после вставки
            end_index = tk_text_widget.index('end-1c')
            # Применяем тег форматирования
            tk_text_widget.tag_add(tag, start_index, end_index)

    # Функция для прокрутки к разделу
    def scroll_to_section(mark_name):
        tk_text_widget.see(mark_name)
        tk_text_widget.tag_remove('highlight', '1.0', 'end')
        tk_text_widget.tag_add('highlight', mark_name, '%s lineend' % mark_name)
        tk_text_widget.tag_configure('highlight', background='LemonChiffon')

    # Настраиваем тег 'highlight' для выделения
    tk_text_widget.tag_configure('highlight', background='LemonChiffon')

    # Делаем текстовое поле только для чтения
    instructions_textbox.configure(state='disabled')

    # Обработка прокрутки колесиком мыши
    def _on_mousewheel(event):
        tk_text_widget.yview_scroll(int(-1 * (event.delta / 120)), "units")

    tk_text_widget.bind("<Enter>", lambda e: tk_text_widget.bind_all('<MouseWheel>', _on_mousewheel))
    tk_text_widget.bind("<Leave>", lambda e: tk_text_widget.unbind_all('<MouseWheel>'))
