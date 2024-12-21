import customtkinter as CTk
from PIL import Image
from info_window import show_info_window
import os
from check_function import check
from tkinter import filedialog
from instructions_window import show_instructions_window
from result_window import show_result_window
from tkinter import messagebox
import shutil
import threading


class App(CTk.CTk):
    def __init__(self):
        super().__init__()

        self.geometry("820x360")
        self.title("Проверка документа на соответствие ГОСТ 7.32-2017")
        self.resizable(False, False)
        self.configure(fg_color="#E6E6FA")

        self.selected_file = None
        self.check_results = None

        # Кнопка "ГОСТ"
        self.gost_but = CTk.CTkButton(master=self, text="ГОСТ", width=150, height=50, fg_color="MediumSlateBlue",
                                      font=("Arial", 26, "bold"), command=self.doc_gost)
        self.gost_but.grid(row=0, column=0, padx=(10, 10))

        # Логотипы
        self.img1 = CTk.CTkImage(dark_image=Image.open("logo1.png"), size=(80, 80))
        self.img1_label = CTk.CTkLabel(master=self, text="", image=self.img1)
        self.img1_label.grid(row=0, column=5, padx=(10, 10), pady=(10, 10))
        self.img1_label.bind("<Button-1>", self.info)

        self.img2 = CTk.CTkImage(dark_image=Image.open("logo2.png"), size=(80, 80))
        self.img2_label = CTk.CTkLabel(master=self, text="", image=self.img2)
        self.img2_label.grid(row=0, column=6, padx=(10, 10), pady=(10, 10))
        self.img2_label.bind("<Button-1>", self.user_instructions)

        # Основной текст
        self.main_label = CTk.CTkLabel(master=self, text="Выберите файл для проверки:",
                                       font=("Arial", 26, "bold"))
        self.main_label.grid(row=1, column=1, columnspan=3, padx=(10, 10))

        # Поле для выбора файла
        self.file_frame = CTk.CTkFrame(master=self, fg_color="transparent")
        self.file_frame.grid(row=2, column=1, columnspan=3, padx=(10, 10), pady=(20, 20), sticky="nsew")

        self.file_entry = CTk.CTkEntry(master=self.file_frame, width=350, height=40, font=("Arial", 20),
                                       state="readonly")
        self.file_entry.grid(row=0, column=0, padx=(15, 20), pady=(0, 0))

        # Горизонтальный скроллбар для поля выбора файла
        self.file_entry_scrollbar = CTk.CTkScrollbar(master=self.file_frame, orientation='horizontal',
                                                     command=self.file_entry._entry.xview)
        self.file_entry_scrollbar.grid(row=1, column=0, padx=(13, 20), sticky='ew', pady=(0, 10))
        self.file_entry._entry.configure(xscrollcommand=self.file_entry_scrollbar.set)

        self.img3 = CTk.CTkImage(dark_image=Image.open("logo3.png"), size=(65, 65))
        self.img3_label = CTk.CTkLabel(master=self.file_frame, text="", image=self.img3)
        self.img3_label.grid(row=0, column=1, sticky="nsew", rowspan=2)
        self.img3_label.bind("<Button-1>", self.user_doc)

        # Кнопки проверки и результатов
        self.send_but = CTk.CTkButton(master=self, text="Отправить на проверку", width=300, height=50,
                                      fg_color="MediumSlateBlue", font=("Arial", 26, "bold"),
                                      command=self.check, state="disabled")
        self.send_but.grid(row=3, column=0, columnspan=3, padx=(50, 10), pady=(10, 10))

        self.res_but = CTk.CTkButton(master=self, text="Просмотр результатов", width=300, height=50,
                                     fg_color="MediumSlateBlue", font=("Arial", 26, "bold"),
                                     command=self.result, state="disabled")
        self.res_but.grid(row=3, column=3, columnspan=3, padx=(10, 40), pady=(10, 10))

        # Добавляем метку статуса
        self.status_label = CTk.CTkLabel(master=self, text="", font=("Arial", 20, "italic"))
        self.status_label.grid(row=4, column=0, columnspan=6, pady=(10, 0))

    def doc_gost(self):
        save_path = filedialog.asksaveasfilename(
            title="Сохранить ГОСТ 7.32-2017",
            defaultextension=".pdf",
            filetypes=(("PDF файлы", "*.pdf"), ("Все файлы", "*.*")),
            initialfile="ГОСТ 7.32-2017.pdf"
        )
        if save_path:
            try:
                gost_file_path = os.path.join(os.getcwd(), "ГОСТ 7.32-2017.pdf")
                shutil.copyfile(gost_file_path, save_path)
                messagebox.showinfo("Успешно", "Файл ГОСТ 7.32-2017 сохранен.")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось сохранить файл ГОСТ: {e}")

    def info(self, event):
        show_info_window(self)

    def user_instructions(self, event):
        show_instructions_window(self)

    def user_doc(self, event):
        filename = filedialog.askopenfilename(title="Выберите файл",
                                              filetypes=(("Word файлы", "*.docx"), ("Все файлы", "*.*")))
        if filename:
            file_name_only = os.path.basename(filename)
            self.selected_file = filename
            self.file_entry.configure(state="normal")
            self.file_entry.delete(0, "end")
            self.file_entry.insert(0, file_name_only)
            self.file_entry.configure(state="readonly")
            self.send_but.configure(state="normal")
        else:
            self.file_entry.configure(state="normal")
            self.file_entry.delete(0, "end")
            self.file_entry.insert(0, "Файл не выбран")
            self.file_entry.configure(state="readonly")
            self.selected_file = None
            self.send_but.configure(state="disabled")
            self.res_but.configure(state="disabled")

    def check(self):
        if self.selected_file:
            # Делаем кнопку недоступной, чтобы предотвратить повторное нажатие
            self.send_but.configure(state="disabled")

            # Показываем сообщение пользователю
            self.status_label.configure(text="Ожидайте, проверка может занять некоторое время...")
            self.update()

            # Функция для выполнения проверки
            def perform_check():
                try:
                    self.check_results = check(self.selected_file)
                    if self.check_results is not None:
                        # Обновляем интерфейс в основном потоке
                        self.after(0, lambda: self.res_but.configure(state="normal"))
                    else:
                        self.after(0, lambda: messagebox.showerror("Ошибка", "Ошибка при проверке файла."))
                except Exception as e:
                    self.after(0, lambda: messagebox.showerror("Ошибка", f"Произошла ошибка: {e}"))
                finally:
                    # Скрываем сообщение после завершения
                    self.after(0, lambda: self.status_label.configure(text=""))

            # Создаём и запускаем поток
            check_thread = threading.Thread(target=perform_check)
            check_thread.start()
        else:
            messagebox.showerror("Ошибка", "Пожалуйста, выберите файл для проверки.")

    # def result(self):
    #     show_result_window(self, self.check_results)

    def result(self):
        # Делаем кнопки и иконку недоступными
        self.send_but.configure(state="disabled")
        self.res_but.configure(state="disabled")
        self.img3_label.unbind("<Button-1>")  # Отключаем привязку клика по иконке

        # Открываем окно с результатами
        try:
            show_result_window(self, self.check_results)
        finally:
            # После закрытия окна восстанавливаем доступ
            self.send_but.configure(state="normal")
            self.res_but.configure(state="normal")
            self.img3_label.bind("<Button-1>", self.user_doc)  # Повторно привязываем клик по иконке

            # Сбрасываем выбранный файл
            self.file_entry.configure(state="normal")
            self.file_entry.delete(0, "end")
            self.file_entry.insert(0, "Файл не выбран")
            self.file_entry.configure(state="readonly")
            self.selected_file = None

            # Делаем кнопку отправки неактивной
            self.send_but.configure(state="disabled")
            self.res_but.configure(state="disabled")


if __name__ == "__main__":
    app = App()
    app.mainloop()
