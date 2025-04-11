import pandas as pd
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog, messagebox
import os
from PIL import Image, ImageTk
from ctypes import windll
import pygame
import random
import logging

#-------DEBUG-------#
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    filename='app.log',
    filemode='w'
)

windll.shcore.SetProcessDpiAwareness(1)

#-------Большая формула--------#
def process_hours(x):
    if pd.isnull(x):
        return None
    x = round(float(x), 2) if isinstance(x, (int, float, str)) else None
    if x is None:
        return "Проверь"
    if x == 0:
        return 0
    elif 0.01 <= x <= 2.99:
        return x
    elif 3.00 <= x <= 3.99:
        return x - 0.25
    elif 4.00 <= x <= 4.99:
        return x - 0.5
    elif 5.00 <= x <= 6.99:
        return x - 0.75
    elif 7.00 <= x <= 9.99:
        return x - 1
    elif 10.00 <= x <= 10.99:
        return x - 1.25
    elif 11.00 <= x <= 12.99:
        return x - 1.5
    elif 13.00 <= x <= 14.99:
        return x - 1.75
    elif 15.00 <= x <= 15.99:
        return x - 2
    elif 16.00 <= x <= 17.99:
        return x - 2.5
    elif 18.00 <= x <= 18.99:
        return x - 2.75
    elif 19.00 <= x <= 20.99:
        return x - 3
    elif 21.00 <= x <= 22.99:
        return x - 3.25
    elif 23.00 <= x <= 24.99:
        return x - 3.5
    else:
        return "Проверь"

def process_file(file_path):
    # Исполняемый файл
    try:
        df = pd.read_excel(
            file_path,
            sheet_name='Лист1',
            dtype={
                'Начало (дата)': 'object',
                'Конец (дата)': 'object',
                'Начало (время)': 'object',
                'Конец (время)': 'object'
            }
        )

        df['Начало (дата)'] = pd.to_datetime(df['Начало (дата)'], format='%d.%m.%Y', dayfirst=True)
        df['Конец (дата)'] = pd.to_datetime(df['Конец (дата)'], format='%d.%m.%Y', dayfirst=True)
        df['Начало (время)'] = pd.to_datetime(df['Начало (время)'], format='%H:%M:%S').dt.time
        df['Конец (время)'] = pd.to_datetime(df['Конец (время)'], format='%H:%M:%S').dt.time

        required_columns = ['Логин', 'Теги', 'Тип', 'Начало (дата)', 'Начало (время)', 'Конец (дата)', 'Конец (время)', 'Навык']
        df = df[required_columns]

        df['Начало'] = df.apply(lambda row: pd.Timestamp.combine(row['Начало (дата)'], row['Начало (время)']), axis=1)
        df['Конец'] = df.apply(lambda row: pd.Timestamp.combine(row['Конец (дата)'], row['Конец (время)']), axis=1)

        # Смены
        shifts_mask = df['Тип'].isin([
            "Смена. Основная",
            "Смена. Доп",
            "Смена. Отработка",
            "Сегмент смены"
        ])
        shifts_df = df[shifts_mask]
        shifts_df['time_diff'] = shifts_df['Конец'] - shifts_df['Начало']
        shifts_df['hours'] = shifts_df['time_diff'].dt.total_seconds() / 3600
        shifts_df['Часы с перерывами'] = shifts_df['hours'].apply(process_hours)

        # Нарушения
        violations_mask = (
            df['Тип'].str.startswith('Наставничество.') |
            (df['Тип'].str.startswith('ПА.') &
             ~df['Тип'].eq('ПА. Ошибочное нарушение') &
             ~df['Тип'].str.startswith('Отсутствие.'))
        ) | df['Тип'].isin([
            "Нарушение. Не работает",
            "Нарушение. Прогул",
            "Нарушение. Опоздание на смену"
        ])
        violations_df = df[violations_mask]
        logging.debug(f"Отфильтрованы нарушения:\n{violations_df.head()}")

        violations_df = df[violations_mask]
        violations_df['time_diff'] = violations_df['Конец'] - violations_df['Начало']
        violations_df['hours'] = violations_df['time_diff'].dt.total_seconds() / 3600
        violations_df['Часы с перерывами'] = violations_df['hours'].apply(process_hours)
        violations_df['Нарушения'] = violations_df['Часы с перерывами'] * -1
        
        # Соединяем данные
        shifts_agg = shifts_df.groupby(['Логин', 'Теги'])['Часы с перерывами'].sum().reset_index()
        violations_agg = violations_df.groupby(['Логин', 'Теги'])['Нарушения'].sum().reset_index()
        merged = pd.merge(shifts_agg, violations_agg, on=['Логин', 'Теги'], how='outer')
        merged.fillna(0, inplace=True)
        merged['Чистые часы'] = merged['Часы с перерывами'] + merged['Нарушения']
    
        # Сохранение результата
        output_path = os.path.splitext(file_path)[0] + '_processed.xlsx'
        merged.to_excel(output_path, index=False)
        logging.debug(f"Сохранен результат в файл: {output_path}")
        return output_path
    except Exception as e:
        logging.error(f"Произошла ошибка: {str(e)}")
        return f"Error: {str(e)}"

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("YM Monitoring V. 0.10")
        self.root.geometry("650x650")

        icon_path = "icon.ico"
        if os.path.exists(icon_path):
            self.root.iconbitmap(icon_path)
        else:
            logging.warning(f"Файл иконки не найден: {icon_path}")
        
        self.file_path = tk.StringVar()
        self.log_text = tk.Text(root, height=10, width=50, bg='white', relief=tk.FLAT,  highlightthickness=0)
        self.log_text.pack(pady=10)
        self.root.tk.call('tk', 'scaling', 2.0)
        
        frame = tk.Frame(root)        
        frame.pack(pady=10)
        
        tk.Button(frame, text="Выбрать файл", command=self.browse_file).grid(row=0, column=0, padx=5)
        tk.Button(frame, text="Обработать", command=self.process).grid(row=0, column=1, padx=5)
        tk.Button(frame, text="Не нажимать", command=self.gif_and_music).grid(row=0, column=2, padx=5)

        self.gif_label = tk.Label(root)
        self.gif_label.pack(pady=0)
        
        # Загружаем гифку
        self.gif_path = "D:/загрузки/loader.gif" 
        self.gif_frames = self.load_gif(self.gif_path)
        self.current_frame = 0
        self.animation = None

        pygame.mixer.init()
        self.music_path = "tomme.mp3"
    
    # Переменные для хранения данных
        self.merged_data = pd.DataFrame()
        self.tags = []

    def browse_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        self.file_path.set(path)
        self.log_text.insert(tk.END, f"Выбран файл: {path}\n")
        logging.debug(f"Выбран файл: {path}")
    
    def process(self):
        path = self.file_path.get()
        if not path:
            messagebox.showerror("Ошибка", "Выберите файл")
            self.log_text.insert(tk.END, "Ошибка: Выберите файл\n")
            logging.error("Ошибка: Выберите файл")
            return
        result = process_file(path)
        if isinstance(result, str) and result.startswith("Error"):
            self.log_text.insert(tk.END, f"Ошибка: {result}\n")
            messagebox.showerror("Ошибка", result)
            logging.error(f"Ошибка: {result}")
        else:
            self.log_text.insert(tk.END, f"Обработка завершена. Результат: {result}\n")
            messagebox.showinfo("Успех", f"Обработка завершена. Результат: {result}")
            logging.info(f"Обработка завершена. Результат: {result}")
            
    def load_gif(self, path):
        gif = Image.open(path)
        frames = []
        try:
            while True:
                frame = ImageTk.PhotoImage(gif.copy())
                frames.append(frame)
                gif.seek(len(frames))  # Переход к следующему кадру
        except EOFError:
            pass
        return frames   

    def gif_and_music(self):
        if self.animation:
            self.root.after_cancel(self.animation)  # Останавливаем предыдущую анимацию
        self.current_frame = 0
        self.animate_gif()

        # Воспроизведение музыки
        pygame.mixer.music.load(self.music_path)
        pygame.mixer.music.play(-1)


    def animate_gif(self):
        if self.current_frame < len(self.gif_frames):
            self.gif_label.config(image=self.gif_frames[self.current_frame])
            self.gif_label.image = self.gif_frames[self.current_frame]  # Сохраняем ссылку
            self.current_frame += 1
            self.animation = self.root.after(30, self.animate_gif)  # Обновляем каждые 100 мс
        else:
            self.current_frame = 0  # Сбрасываем на начало для повторения
            self.animate_gif()

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()