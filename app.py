import customtkinter as ctk
from tkinter import filedialog, messagebox
import re
import pandas as pd
from docx import Document
import os
from datetime import datetime
from charset_normalizer import from_path
import threading

# Автор ПО: Chernov Igor

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("LinkVideo Converter | Software by Chernov Igor")
        self.geometry("500x400")
        ctk.set_appearance_mode("dark")

        # Заголовок и автор
        self.label_title = ctk.CTkLabel(self, text="LinkVideo Batch Converter", font=("Arial", 20, "bold"))
        self.label_title.pack(pady=(20, 5))
        
        self.label_author = ctk.CTkLabel(self, text="Software by Chernov Igor", font=("Arial", 12, "italic"), text_color="gray")
        self.label_author.pack(pady=(0, 20))

        self.label_desc = ctk.CTkLabel(self, text="Выберите до 10 файлов (.docx, .txt)\nдля объединения в один Excel-отчет")
        self.label_desc.pack(pady=10)

        # Индикатор прогресса
        self.progress_bar = ctk.CTkProgressBar(self, width=400)
        self.progress_bar.pack(pady=20)
        self.progress_bar.set(0)

        self.status_label = ctk.CTkLabel(self, text="Ожидание выбора файлов...", font=("Arial", 11))
        self.status_label.pack(pady=5)

        # Кнопка
        self.btn = ctk.CTkButton(self, text="Выбрать файлы и запустить", command=self.select_files, fg_color="#1f538d")
        self.btn.pack(pady=30)

    def select_files(self):
        file_paths = filedialog.askopenfilenames(filetypes=[("Documents", "*.docx *.txt")])
        if not file_paths:
            return
        
        files_to_process = list(file_paths[:10])
        self.btn.configure(state="disabled")
        self.progress_bar.set(0)
        
        # Запуск в отдельном потоке, чтобы интерфейс не зависал
        thread = threading.Thread(target=self.run_logic, args=(files_to_process,))
        thread.start()

    def run_logic(self, file_paths):
        try:
            all_data = []
            total_files = len(file_paths)
            
            for i, input_path in enumerate(file_paths):
                self.update_status(f"Обработка: {os.path.basename(input_path)}")
                
                ext = os.path.splitext(input_path)[1].lower()
                lines = self.read_docx(input_path) if ext == ".docx" else self.read_txt(input_path)
                
                self.parse_lines(lines, all_data)
                
                # Обновляем прогресс-бар
                self.progress_bar.set((i + 1) / total_files)

            if not all_data:
                self.finish_process(False, "Данные не найдены.")
                return

            df = pd.DataFrame(all_data)
            df = df.sort_values(by="_sort_key").drop(columns=["_sort_key"])

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = os.path.join(os.path.dirname(file_paths[0]), f"Summary_Report_{timestamp}.xlsx")
            
            df.to_excel(output_path, index=False)
            self.finish_process(True, f"Успешно!\nОбъединено файлов: {len(file_paths)}\nЗаписей: {len(all_data)}\nАвтор: Chernov Igor")
            
        except Exception as e:
            self.finish_process(False, f"Ошибка: {str(e)}")

    def read_txt(self, file_path):
        results = from_path(file_path).best()
        return [line.strip() for line in str(results).splitlines()]

    def read_docx(self, file_path):
        doc = Document(file_path)
        return [para.text.strip() for para in doc.paragraphs]

    def parse_lines(self, lines, data_list):
        current_date, current_time, buffer_text = None, None, ""
        date_pattern = re.compile(r"\[(\d{2}\.\d{2}\.\d{4})\s+(\d{1,2}:\d{2})\]")

        for text in lines:
            if not text: continue
            date_match = date_pattern.search(text)
            if date_match:
                if current_date and buffer_text:
                    self.add_to_list(current_date, current_time, buffer_text, data_list)
                current_date, current_time = date_match.groups()
                buffer_text = ""
            else:
                buffer_text = (buffer_text + " " + text).strip() if buffer_text else text
        
        if current_date and buffer_text:
            self.add_to_list(current_date, current_time, buffer_text, data_list)

    def add_to_list(self, date, time, text, data_list):
        parts = re.split(r'\s+[—–-]\s+', text, maxsplit=1)
        if len(parts) == 2:
            try:
                sort_dt = datetime.strptime(f"{date} {time}", "%d.%m.%Y %H:%M")
            except:
                sort_dt = datetime.min
            data_list.append({
                "Дата": date, "Время": time,
                "Адрес": parts[0].strip(), "Комментарий": parts[1].strip(),
                "_sort_key": sort_dt
            })

    def update_status(self, msg):
        self.after(0, lambda: self.status_label.configure(text=msg))

    def finish_process(self, success, msg):
        self.after(0, lambda: self.btn.configure(state="normal"))
        if success:
            self.after(0, lambda: messagebox.showinfo("Готово", msg))
            self.update_status("Завершено успешно")
        else:
            self.after(0, lambda: messagebox.showerror("Ошибка", msg))
            self.update_status("Произошла ошибка")

if __name__ == "__main__":
    app = App()
    app.mainloop()
