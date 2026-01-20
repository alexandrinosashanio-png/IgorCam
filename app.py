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
        # Регулярное выражение для формата: [Дата Время] Адрес — Комментарий
        pattern = re.compile(r"\[(\d{2}\.\d{2}\.\d{4})\s+(\d{1,2}:\d{2})\]\s+(.*?)\s+—\s+(.*)")

        for para in doc.paragraphs:
            text = para.text.strip()
            match = pattern.search(text)
            if match:
                date, time, address, comment = match.groups()
                data.append({
                    "Дата": date,
                    "Время": time,
                    "Адрес": address,
                    "Комментарий": comment
                })
        
        if not data:
            return False, "Данные не найдены. Проверьте формат текста в документе."

        df = pd.DataFrame(data)
        output_path = os.path.splitext(input_path)[0] + ".xlsx"
        df.to_excel(output_path, index=False)
        return True, f"Готово! Таблица сохранена:\n{output_path}"
    except Exception as e:
        return False, f"Ошибка: {str(e)}"

def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
    if file_path:
        success, message = process_file(file_path)
        if success: messagebox.showinfo("Успех", message)
        else: messagebox.showerror("Ошибка", message)

app = ctk.CTk()
app.title("LinkVideo Log Converter")
app.geometry("400x200")
label = ctk.CTkLabel(app, text="Выберите .docx файл для конвертации в таблицу")
label.pack(pady=20)
btn = ctk.CTkButton(app, text="Выбрать файл", command=select_file)
btn.pack(pady=20)
app.mainloop()