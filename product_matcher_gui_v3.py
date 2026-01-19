import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from rapidfuzz import process, fuzz
from anyascii import anyascii
import re
import os
import time


class ProductMatcherGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Сопоставление товаров")
        self.root.geometry("600x600")
        
        # Переменные для хранения путей к файлам
        self.site_file_path = tk.StringVar()
        self.erp_file_path = tk.StringVar()
        
        # Переменные для процесса обработки
        self.processing = False
        self.current_index = 0
        self.df_site = None
        self.df_erp = None
        self.choices = None
        self.results = []
        self.matched_count = 0
        self.site_id_col = None
        self.site_name_col = None
        self.erp_id_col = None
        self.erp_name_col = None
        self.threshold = 60
        
        self.create_widgets()
        
    def create_widgets(self):
        # Заголовок
        title_label = tk.Label(self.root, text="Сопоставление товаров по названиям", 
                              font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        
        # Фрейм для выбора файлов
        files_frame = ttk.LabelFrame(self.root, text="Выбор файлов", padding=10)
        files_frame.pack(fill="x", padx=20, pady=10)
        
        # Файл сайта
        site_frame = tk.Frame(files_frame)
        site_frame.pack(fill="x", pady=5)
        
        tk.Label(site_frame, text="Файл каталога сайта:", width=20, anchor="w").pack(side="left")
        tk.Entry(site_frame, textvariable=self.site_file_path, width=40).pack(side="left", padx=5)
        tk.Button(site_frame, text="Обзор...", command=self.select_site_file).pack(side="left")
        
        # Файл ERP
        erp_frame = tk.Frame(files_frame)
        erp_frame.pack(fill="x", pady=5)
        
        tk.Label(erp_frame, text="Файл программы учета:", width=20, anchor="w").pack(side="left")
        tk.Entry(erp_frame, textvariable=self.erp_file_path, width=40).pack(side="left", padx=5)
        tk.Button(erp_frame, text="Обзор...", command=self.select_erp_file).pack(side="left")
        
        # Настройки сопоставления
        settings_frame = ttk.LabelFrame(self.root, text="Настройки сопоставления", padding=10)
        settings_frame.pack(fill="x", padx=20, pady=10)
        
        # Минимальный порог схожести
        threshold_frame = tk.Frame(settings_frame)
        threshold_frame.pack(fill="x", pady=5)
        
        tk.Label(threshold_frame, text="Минимальный порог схожести (%):", width=30, anchor="w").pack(side="left")
        self.threshold_var = tk.IntVar(value=60)
        threshold_scale = tk.Scale(threshold_frame, from_=30, to=95, orient="horizontal", 
                                 variable=self.threshold_var, length=200)
        threshold_scale.pack(side="left", padx=5)
        
        # Кнопки управления
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=10)
        
        self.process_button = tk.Button(button_frame, text="Начать сопоставление", 
                                       command=self.start_matching, bg="#4CAF50", fg="white",
                                       font=("Arial", 12, "bold"), height=2, width=20)
        self.process_button.pack(side="left", padx=5)
        
        self.stop_button = tk.Button(button_frame, text="Остановить", 
                                    command=self.stop_matching, bg="#f44336", fg="white",
                                    font=("Arial", 12, "bold"), height=2, width=15, state="disabled")
        self.stop_button.pack(side="left", padx=5)
        
        # Прогресс бар
        progress_frame = tk.Frame(self.root)
        progress_frame.pack(fill="x", padx=20, pady=5)
        
        self.progress = ttk.Progressbar(progress_frame, mode='determinate')
        self.progress.pack(fill="x", side="left", expand=True)
        
        self.progress_label = tk.Label(progress_frame, text="0%", width=8)
        self.progress_label.pack(side="right", padx=(5, 0))
        
        # Статистика
        stats_frame = tk.Frame(self.root)
        stats_frame.pack(fill="x", padx=20, pady=20)
        
        self.stats_label = tk.Label(stats_frame, text="Готов к работе", anchor="w", font=("Arial", 10))
        self.stats_label.pack(fill="x")
        
        # Статус бар внизу окна
        self.status_bar = tk.Frame(self.root, relief=tk.SUNKEN, bd=1)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.status_label = tk.Label(self.status_bar, text="Готов к работе", anchor=tk.W, padx=10, pady=2)
        self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
    def log(self, message):
        """Обновляет статус в строке статистики"""
        self.stats_label.config(text=message)
        self.root.update_idletasks()
        
    def update_status_bar(self, message):
        """Обновляет статус бар"""
        self.status_label.config(text=message)
        self.root.update_idletasks()
        
    def select_site_file(self):
        filename = filedialog.askopenfilename(
            title="Выберите файл каталога сайта",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.site_file_path.set(filename)
            self.log(f"Выбран файл сайта: {os.path.basename(filename)}")
            self.update_status_bar(f"Выбран файл сайта: {os.path.basename(filename)}")
            
    def select_erp_file(self):
        filename = filedialog.askopenfilename(
            title="Выберите файл программы учета",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.erp_file_path.set(filename)
            self.log(f"Выбран файл программы учета: {os.path.basename(filename)}")
            self.update_status_bar(f"Выбран файл программы учета: {os.path.basename(filename)}")
            
    def clean_name(self, name):
        """Очистка названия товара для сопоставления"""
        if not isinstance(name, str):
            return ""
        # Убираем страны в косых чертах /Беларусь/, /Германия/
        name = re.sub(r'/[^/]+/', ' ', name)
        # Транслитерация для сопоставления RU/EN брендов
        name = anyascii(name)
        # Оставляем буквы и цифры
        name = re.sub(r'[^a-zA-Z0-9\s]', ' ', name).lower()
        return " ".join(name.split())
        
    def detect_columns(self, df, file_type):
        """Автоматическое определение колонок ID и названия"""
        columns = df.columns.tolist()
        
        # Ищем колонку ID
        id_col = None
        for col in columns:
            if any(keyword in col.lower() for keyword in ['id', '_id_', 'код', 'артикул']):
                id_col = col
                break
        
        # Ищем колонку с названием
        name_col = None
        for col in columns:
            if any(keyword in col.lower() for keyword in ['наименование', 'название', 'товар', 'name']):
                name_col = col
                break
        
        # Если не нашли, берем первые две колонки
        if not id_col:
            id_col = columns[0]
        if not name_col:
            name_col = columns[1] if len(columns) > 1 else columns[0]
            
        self.log(f"Файл {file_type}: ID = '{id_col}', Название = '{name_col}'")
        return id_col, name_col
        
    def start_matching(self):
        """Запуск процесса сопоставления"""
        if not self.site_file_path.get() or not self.erp_file_path.get():
            messagebox.showerror("Ошибка", "Пожалуйста, выберите оба файла")
            return
            
        try:
            self.log("Начинаю загрузку файлов...")
            self.update_status_bar("Загрузка файлов...")
            self.root.update_idletasks()
            
            # Загрузка файлов
            self.df_site = pd.read_excel(self.site_file_path.get())
            self.df_erp = pd.read_excel(self.erp_file_path.get())
            
            self.log(f"Загружен файл сайта: {len(self.df_site)} товаров")
            self.log(f"Загружен файл программы учета: {len(self.df_erp)} товаров")
            self.update_status_bar(f"Загружены файлы: сайт {len(self.df_site)} товаров, программа {len(self.df_erp)} товаров")
            
            # Определяем колонки автоматически
            self.site_id_col, self.site_name_col = self.detect_columns(self.df_site, "сайта")
            self.erp_id_col, self.erp_name_col = self.detect_columns(self.df_erp, "программы учета")
            
            # Подготовка данных
            self.log("Подготавливаю данные для сопоставления...")
            self.update_status_bar("Подготовка данных для сопоставления...")
            self.root.update_idletasks()
            
            self.df_site['clean'] = self.df_site[self.site_name_col].apply(self.clean_name)
            self.df_erp['clean'] = self.df_erp[self.erp_name_col].apply(self.clean_name)
            
            self.choices = self.df_erp['clean'].tolist()
            self.results = []
            self.matched_count = 0
            self.current_index = 0
            self.threshold = self.threshold_var.get()
            
            self.log(f"Начинаю сопоставление с порогом {self.threshold}%...")
            self.update_status_bar(f"Начинаю сопоставление с порогом {self.threshold}%...")
            
            # Настройка интерфейса для процесса
            self.processing = True
            self.process_button.config(state="disabled")
            self.stop_button.config(state="normal")
            self.progress['maximum'] = len(self.df_site)
            self.progress['value'] = 0
            
            # Запуск пошаговой обработки
            self.process_next_batch()
            
        except Exception as e:
            self.log(f"Ошибка при загрузке: {str(e)}")
            self.update_status_bar(f"Ошибка: {str(e)}")
            messagebox.showerror("Ошибка", f"Ошибка при загрузке файлов: {str(e)}")
            
    def process_next_batch(self):
        """Обрабатывает следующую порцию товаров"""
        if not self.processing:
            return
            
        batch_size = 5  # Обрабатываем по 5 товаров за раз
        end_index = min(self.current_index + batch_size, len(self.df_site))
        
        # Обрабатываем порцию
        for idx in range(self.current_index, end_index):
            row = self.df_site.iloc[idx]
            
            extract = process.extractOne(
                row['clean'],
                self.choices,
                scorer=fuzz.WRatio
            )
            
            if extract and extract[1] >= self.threshold:
                match_text, score, match_idx = extract
                matched_row = self.df_erp.iloc[match_idx]
                
                self.results.append({
                    'id сайт': row[self.site_id_col],
                    'наименование сайт': row[self.site_name_col],
                    'id программа': matched_row[self.erp_id_col],
                    'наименование программа': matched_row[self.erp_name_col],
                    'схожесть %': round(score, 1)
                })
                self.matched_count += 1
        
        # Обновляем прогресс
        self.current_index = end_index
        progress_percent = (self.current_index / len(self.df_site)) * 100
        self.progress['value'] = self.current_index
        self.progress_label.config(text=f"{progress_percent:.1f}%")
        
        # Обновляем статистику и статус бар
        if self.current_index < len(self.df_site):
            stats_text = f"Обработано: {self.current_index}/{len(self.df_site)} | Найдено совпадений: {self.matched_count}"
            self.stats_label.config(text=stats_text)
            self.update_status_bar(f"Обработка... {progress_percent:.1f}% | Найдено: {self.matched_count} совпадений")
        else:
            stats_text = f"Завершено! Обработано: {self.current_index}/{len(self.df_site)} | Найдено совпадений: {self.matched_count}"
            self.stats_label.config(text=stats_text)
            self.update_status_bar(f"Завершено! Найдено {self.matched_count} совпадений из {len(self.df_site)} товаров")
        
        # Логируем прогресс каждые 200 товаров
        if self.current_index % 200 == 0 or self.current_index == len(self.df_site):
            self.log(f"Обработано {self.current_index}/{len(self.df_site)} товаров ({progress_percent:.1f}%) | Найдено: {self.matched_count}")
        
        # Принудительно обновляем интерфейс
        self.root.update_idletasks()
        self.root.update()
        
        # Проверяем завершение
        if self.current_index >= len(self.df_site):
            self.finish_processing()
        else:
            # Планируем следующую порцию через 10мс
            self.root.after(10, self.process_next_batch)
            
    def stop_matching(self):
        """Остановка процесса сопоставления"""
        self.processing = False
        self.log("Процесс остановлен пользователем")
        self.update_status_bar("Процесс остановлен пользователем")
        self.finish_processing()
        
    def finish_processing(self):
        """Завершение процесса обработки"""
        self.processing = False
        self.process_button.config(state="normal")
        self.stop_button.config(state="disabled")
        
        if self.results:
            try:
                output_file = "результат_сопоставления.xlsx"
                df_final = pd.DataFrame(self.results)
                df_final.to_excel(output_file, index=False)
                
                success_message = (f"Сопоставление завершено!\n"
                                 f"Найдено совпадений: {self.matched_count} из {len(self.df_site)}\n"
                                 f"Результат сохранен в: {output_file}")
                
                self.log(f"Готово! Найдено {self.matched_count} совпадений из {len(self.df_site)} товаров")
                self.log(f"Результат сохранен в файл: {output_file}")
                self.update_status_bar(f"Готово! Результат сохранен в {output_file}")
                
                messagebox.showinfo("Успех", success_message)
                
            except Exception as e:
                error_msg = f"Ошибка при сохранении: {str(e)}"
                self.log(error_msg)
                self.update_status_bar(error_msg)
                messagebox.showerror("Ошибка", error_msg)
        else:
            self.log("Совпадений не найдено. Попробуйте снизить порог схожести.")
            self.update_status_bar("Совпадений не найдено")
            messagebox.showwarning("Внимание", "Совпадений не найдено")


if __name__ == "__main__":
    root = tk.Tk()
    app = ProductMatcherGUI(root)
    root.mainloop()