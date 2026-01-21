import customtkinter as ctk
from tkinter import filedialog, messagebox
import re
import pandas as pd
from docx import Document
import os
from datetime import datetime

# Автор ПО: Chernov Igor

def read_txt(file_path):
    with open(file_path, 'r', encoding='utf-8') as f:
        return [line.strip() for line in f.readlines()]

def read_docx(file_path):
    doc = Document(file_path)
    return [para.text.strip() for para in doc.paragraphs]

def parse_event(date, time, text, data_list):
    parts = re.split(r'\s+[—–-]\s+', text, maxsplit=1)
    if len(parts) == 2:
        # Создаем объект datetime для последующей сортировки
        try:
            sort_dt = datetime.strptime(f"{date} {time}", "%d.%m.%Y %H:%M")
        except:
            sort_dt = datetime.min
            
        data_list.append({
            "Дата": date, 
            "Время": time,
            "Адрес": parts[0].strip(), 
            "Комментарий": parts[1].strip(),
            "_sort_key": sort_dt
        })

def process_files(file_paths):
    try:
        all_data = []
        for input_path in file_paths:
            ext = os.path.splitext(input_path)[1].lower()
            lines = read_docx(input_path) if ext == ".docx" else read_txt(input_path)
            current_date, current_time, buffer_text = None, None, ""
            date_pattern = re.compile(r"\[(\d{2}\.\d{2}\.\d{4})\s+(\d{1,2}:\d{2})\]")

            for text in lines:
                if not text: continue
                date_match = date_pattern.search(text)
                if date_match:
                    if current_date and buffer_text:
                        parse_event(current_date, current_time, buffer_text, all_data)
                    current_date, current_time = date_match.groups()
                    buffer_text = ""
                else:
                    buffer_text = (buffer_text + " " + text).strip() if buffer_text else text
            
            if current_date and buffer_text:
                parse_event(current_date, current_time, buffer_text, all_data)

        if not all_data: return False, "Данные не найдены."

        # Сортировка по времени и удаление технического ключа
        df = pd.DataFrame(all_data)
        df = df.sort_values(by="_sort_key").drop(columns=["_sort_key"])

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(os.path.dirname(file_paths[0]), f"Summary_Report_{timestamp}.xlsx")
        
        df.to_excel(output_path, index=False)
        return True, f"Объединено файлов: {len(file_paths)}\nЗаписей: {len(all_data)}\nАвтор: Chernov Igor"
    except Exception as e:
        return False, f"Ошибка: {str(e)}"

# Настройка GUI
ctk.set_appearance_mode("dark")
app = ctk.CTk()
app.title("LinkVideo Converter | Software by Chernov Igor")
app.geometry("500x350")

label_title = ctk.CTkLabel(app, text="LinkVideo Batch Converter", font=("Arial", 20, "bold"))
label_title.pack(pady=(20, 5))

label_author = ctk.CTkLabel(app, text="Software by Chernov Igor", font=("Arial", 12, "italic"), text_color="gray")
label_author.pack(pady=(0, 20))

label_desc = ctk.CTkLabel(app, text="Выберите до 10 файлов (.docx, .txt)\nдля объединения в один Excel-отчет")
label_desc.pack(pady=10)

def select_files():
    file_paths = filedialog.askopenfilenames(filetypes=[("Documents", "*.docx *.txt")])
    if not file_paths: return
    if len(file_paths) > 10: file_paths = file_paths[:10]
        
    success, message = process_files(file_paths)
    if success: messagebox.showinfo("Готово", message)
    else: messagebox.showerror("Ошибка", message)

btn = ctk.CTkButton(app, text="Выбрать файлы и запустить", command=select_files, fg_color="#1f538d", hover_color="#14375e")
btn.pack(pady=30)

app.mainloop()
