import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import os
import executing
from tkinter import messagebox
import logging

logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s %(pathname)s:%(lineno)d %(levelname)s [Function: %(funcName)s] %(message)s',
                    datefmt='%d-%b-%y %H:%M:%S')


class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title('m4dOFF')
        # Размеры окна
        window_width = 800
        window_height = 600

        # Позиция окна по центру экрана
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        center_x = int((screen_width - window_width) / 2)
        center_y = int((screen_height - window_height) / 2)

        self.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

        self.docx_files_listbox = tk.Listbox(self, height=10)
        self.docx_files_listbox.grid(row=0, column=0, pady=(10, 0), padx=10, sticky='ew')

        self.controls_frame = tk.Frame(self)
        self.controls_frame.grid(row=0, column=1, pady=(10, 0), padx=10,
                                 sticky='nsew')

        self.select_button = tk.Button(self.controls_frame, text="Выбрать", command=self.select_docx_files)
        self.create_button = tk.Button(self.controls_frame, text="Создать", command=self.select_output_folder)
        self.select_button.pack(side=tk.TOP, fill=tk.X, pady=(5, 0))
        self.create_button.pack(side=tk.TOP, fill=tk.X, pady=(5, 0))
        self.create_button.config(state="disabled")

        self.template_combobox = ttk.Combobox(self.controls_frame, state="readonly",
                                              values=["Шаблон A", "Шаблон B"])
        self.template_combobox.pack(side=tk.TOP, fill=tk.X, pady=(5, 0))
        self.load_templates()

        self.xml_files_listbox = tk.Listbox(self, height=10)
        self.xml_files_listbox.grid(row=1, column=0, pady=10, padx=10, sticky='new', columnspan=2)
        self.xml_files_listbox.bind('<<ListboxSelect>>', self.show_xml_content)

        self.log_textbox = tk.Text(self, height=15)
        self.log_textbox.grid(row=2, column=0, columnspan=2, pady=10, padx=10, sticky='nsew')

        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(0, weight=1)

    def get_files_from_directory(self, directory_path):
        try:
            # Получаем список содержимого в указанном каталоге
            files = [f for f in os.listdir(directory_path) if os.path.isfile(os.path.join(directory_path, f))]
            return files
        except FileNotFoundError:
            print(f"Директория {directory_path} не найдена.")
            return []
        except PermissionError:
            print(f"Нет разрешения на доступ к директории {directory_path}.")
            return []

    def load_templates(self):
        exe = executing.Executing()
        items = []
        for i in self.get_files_from_directory("./templates"):
            if i == "template.py":
                continue
            logging.info(i)
            try:
                temp = exe.get_template(f"{os.getcwd()}{os.sep}templates{os.sep}{i}")
                items.append(temp.get("Template")("get_name").name_template)
            except Exception as ex:
                logging.error(ex)
                messagebox.showerror("Ошибка", f"Произошла ошибка!\n{i}\n{ex}")
        self.template_combobox['values'] = items
        if len(items) > 0:
            self.template_combobox.set(items[0])

    def get_from_name_template(self, name):
        exe = executing.Executing()
        for i in self.get_files_from_directory("./templates"):
            if i == "template.py":
                continue
            logging.info(f"{i} {name}")
            try:
                temp = exe.get_template(f"{os.getcwd()}{os.sep}templates{os.sep}{i}")
                if temp.get("Template")("get_name").name_template == name:
                    return f"{os.getcwd()}{os.sep}templates{os.sep}{i}"
            except Exception as ex:
                messagebox.showerror("Ошибка", f"Произошла ошибка!\n{i}\n{ex}")
                logging.error(ex)

    def select_docx_files(self):
        docx_filenames = filedialog.askopenfilenames(filetypes=[("Word documents", "*.docx")])
        self.docx_files_listbox.delete(0, "end")
        for filename in docx_filenames:
            logging.info(filename)
            self.docx_files_listbox.insert(tk.END, filename)
            self.create_button.config(state="normal")

    def select_output_folder(self):
        folder = filedialog.askdirectory()
        self.xml_files_listbox.delete(0, "end")
        if folder:  # Если пользователь выберет папку и не отменит выбор.
            exe = executing.Executing()
            try:
                for i in exe.load_template(self.get_from_name_template(self.template_combobox.get()), self.docx_files_listbox.get(0, tk.END), folder):
                    logging.info(i)
                    self.xml_files_listbox.insert(tk.END, i)
            except Exception as ex:
                messagebox.showerror("Ошибка", f"Произошла ошибка!\n{ex}")
                logging.error(ex)

    def show_xml_content(self, event):
        selection = event.widget.curselection()
        if selection:
            index = selection[0]
            filename = event.widget.get(index)
            with open(filename, "r", encoding="utf-8") as file:
                content = file.read()
                self.log_textbox.delete(1.0, tk.END)
                self.log_textbox.insert(tk.END, content)


def start_gui():
    app = App()
    app.mainloop()  # Запускаем главный цикл обработки событий
