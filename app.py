import customtkinter as ctk
from tkinter import filedialog, messagebox
import re
import pandas as pd
from docx import Document
import os

def process_file(input_path):
    try:
        doc = Document(input_path)
        data = []
        
        current_date = None
        current_time = None
        buffer_text = ""
        
        # Паттерн для поиска даты и времени: [29.12.2025 0:03] [cite: 1]
        date_pattern = re.compile(r"\[(\d{2}\.\d{2}\.\d{4})\s+(\d{1,2}:\d{2})\]")

        for para in doc.paragraphs:
            text = para.text.strip()
            if not text:
                continue
            
            # Проверяем, есть ли в строке дата (начало нового блока) [cite: 1, 3, 5]
            date_match = date_pattern.search(text)
            
            if date_match:
                # Если в буфере остался текст от предыдущего блока, обрабатываем его
                if current_date and buffer_text:
                    parse_event(current_date, current_time, buffer_text, data)
                
                # Запоминаем новую дату и время [cite: 1]
                current_date = date_match.group(1)
                current_time = date_match.group(2)
                buffer_text = "" # Очищаем буфер для нового адреса
            else:
                # Если даты нет, накапливаем текст (адрес и событие) в буфер 
                if buffer_text:
                    buffer_text += " " + text
                else:
                    buffer_text = text

        # Обрабатываем последний блок после завершения цикла
        if current_date and buffer_text:
            parse_event(current_date, current_time, buffer_text, data)
        
        if not data:
            return False, "Данные не найдены. Проверьте формат документа."

        df = pd.DataFrame(data)
        output_path = os.path.splitext(input_path)[0] + ".xlsx"
        df.to_excel(output_path, index=False)
        return True, f"Успешно! Обработано записей: {len(data)}"
    except Exception as e:
        return False, f"Ошибка: {str(e)}"

def parse_event(date, time, text, data_list):
    # Разделяем текст по длинному тире, короткому тире или дефису [cite: 2, 4, 6]
    parts = re.split(r'\s+[—–-]\s+', text, maxsplit=1)
    if len(parts) == 2:
        data_list.append({
            "Дата": date,
            "Время": time,
            "Адрес": parts[0].strip(),
            "Комментарий": parts[1].strip()
        })

# GUI настройка
ctk.set_appearance_mode("dark")
app = ctk.CTk()
app.title("LinkVideo Log Converter v2.0")
app.geometry("450x250")

label = ctk.CTkLabel(app, text="Конвертер логов LinkVideo\n(исправлен для двухстрочного формата)")
label.pack(pady=20)

btn = ctk.CTkButton(app, text="Выбрать файл .docx", command=lambda: select_file())
btn.pack(pady=20)

def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
    if file_path:
        success, message = process_file(file_path)
        if success: messagebox.showinfo("Успех", message)
        else: messagebox.showerror("Ошибка", message)

app.mainloop()
