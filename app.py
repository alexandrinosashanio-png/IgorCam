import customtkinter as ctk
from tkinter import filedialog, messagebox
import re
import pandas as pd
from docx import Document
import os

def read_txt(file_path):
    with open(file_path, 'r', encoding='utf-8') as f:
        return [line.strip() for line in f.readlines()]

def read_docx(file_path):
    doc = Document(file_path)
    return [para.text.strip() for para in doc.paragraphs]

def process_file(input_path):
    try:
        # Определяем расширение и выбираем метод чтения
        ext = os.path.splitext(input_path)[1].lower()
        if ext == ".docx":
            lines = read_docx(input_path)
        elif ext == ".txt":
            lines = read_txt(input_path)
        else:
            return False, "Неподдерживаемый формат файла."

        data = []
        current_date, current_time = None, None
        buffer_text = ""
        
        date_pattern = re.compile(r"\[(\d{2}\.\d{2}\.\d{4})\s+(\d{1,2}:\d{2})\]")

        for text in lines:
            if not text: continue
            
            date_match = date_pattern.search(text)
            if date_match:
                if current_date and buffer_text:
                    parse_event(current_date, current_time, buffer_text, data)
                current_date = date_match.group(1)
                current_time = date_match.group(2)
                buffer_text = ""
            else:
                buffer_text = (buffer_text + " " + text).strip() if buffer_text else text

        if current_date and buffer_text:
            parse_event(current_date, current_time, buffer_text, data)
        
        if not data:
            return False, "Данные не найдены. Проверьте формат текста."

        df = pd.DataFrame(data)
        output_path = os.path.splitext(input_path)[0] + ".xlsx"
        df.to_excel(output_path, index=False)
        return True, f"Успешно! Обработано записей: {len(data)}"
    except Exception as e:
        return False, f"Ошибка: {str(e)}"

def parse_event(date, time, text, data_list):
    parts = re.split(r'\s+[—–-]\s+', text, maxsplit=1)
    if len(parts) == 2:
        data_list.append({
            "Дата": date, "Время": time,
            "Адрес": parts[0].strip(), "Комментарий": parts[1].strip()
        })

# GUI
ctk.set_appearance_mode("dark")
app = ctk.CTk()
app.title("Universal Log Converter")
app.geometry("450x250")

label = ctk.CTkLabel(app, text="Выберите .docx или .txt файл\nдля конвертации в Excel")
label.pack(pady=20)

def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Documents", "*.docx *.txt")])
    if file_path:
        success, message = process_file(file_path)
        if success: messagebox.showinfo("Успех", message)
        else: messagebox.showerror("Ошибка", message)

btn = ctk.CTkButton(app, text="Выбрать файл", command=select_file)
btn.pack(pady=20)
app.mainloop()
