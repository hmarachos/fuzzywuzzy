import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from rapidfuzz import process, fuzz
from anyascii import anyascii
import re
import os
from threading import Thread


class ProductMatcherGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Сопоставление товаров")
        self.root.geometry("600x500")
        
        # Переменные для хранения путей к файлам
        self.site_file_path = tk.StringVar()
        self.erp_file_path = tk.StringVar()
        
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
        
        # Кнопка запуска
        self.process_button = tk.Button(self.root, text="Начать сопоставление", 
                                       command=self.start_matching, bg="#4CAF50", fg="white",
                                       font=("Arial", 12, "bold"), height=2)
        self.process_button.pack(pady=20)
        
        # Прогресс бар
        self.progress = ttk.Progressbar(self.root, mode='indeterminate')
        self.progress.pack(fill="x", padx=20, pady=5)
        
        # Текстовое поле для логов
        log_frame = ttk.LabelFrame(self.root, text="Процесс выполнения", padding=5)
        log_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        self.log_text = tk.Text(log_frame, height=8, wrap="word")
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
    def select_site_file(self):
        filename = filedialog.askopenfilename(
            title="Выберите файл каталога сайта",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.site_file_path.set(filename)
            self.log(f"Выбран файл сайта: {os.path.basename(filename)}")
            
    def select_erp_file(self):
        filename = filedialog.askopenfilename(
            title="Выберите файл программы учета",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.erp_file_path.set(filename)
            self.log(f"Выбран файл программы учета: {os.path.basename(filename)}")
            
    def log(self, message):
        """Добавляет сообщение в лог"""
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update()
        
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
        
    def start_matching(self):
        """Запуск процесса сопоставления в отдельном потоке"""
        if not self.site_file_path.get() or not self.erp_file_path.get():
            messagebox.showerror("Ошибка", "Пожалуйста, выберите оба файла")
            return
            
        # Запускаем в отдельном потоке чтобы не блокировать GUI
        thread = Thread(target=self.perform_matching)
        thread.daemon = True
        thread.start()
        
    def perform_matching(self):
        """Основная логика сопоставления"""
        try:
            # Блокируем кнопку и запускаем прогресс
            self.process_button.config(state="disabled")
            self.progress.start()
            
            self.log("Начинаю загрузку файлов...")
            
            # Загрузка файлов
            df_site = pd.read_excel(self.site_file_path.get())
            df_erp = pd.read_excel(self.erp_file_path.get())
            
            self.log(f"Загружен файл сайта: {len(df_site)} товаров")
            self.log(f"Загружен файл программы учета: {len(df_erp)} товаров")
            
            # Определяем колонки автоматически
            site_id_col, site_name_col = self.detect_columns(df_site, "сайта")
            erp_id_col, erp_name_col = self.detect_columns(df_erp, "программы учета")
            
            # Подготовка данных
            self.log("Подготавливаю данные для сопоставления...")
            df_site['clean'] = df_site[site_name_col].apply(self.clean_name)
            df_erp['clean'] = df_erp[erp_name_col].apply(self.clean_name)
            
            choices = df_erp['clean'].tolist()
            results = []
            threshold = self.threshold_var.get()
            
            self.log(f"Начинаю сопоставление с порогом {threshold}%...")
            
            # Процесс сопоставления
            total = len(df_site)
            matched_count = 0
            
            for idx, row in df_site.iterrows():
                if idx % 100 == 0:  # Обновляем лог каждые 100 товаров
                    self.log(f"Обработано {idx}/{total} товаров...")
                
                extract = process.extractOne(
                    row['clean'],
                    choices,
                    scorer=fuzz.WRatio
                )
                
                if extract and extract[1] >= threshold:
                    match_text, score, match_idx = extract
                    matched_row = df_erp.iloc[match_idx]
                    
                    results.append({
                        'id сайт': row[site_id_col],
                        'наименование сайт': row[site_name_col],
                        'id программа': matched_row[erp_id_col],
                        'наименование программа': matched_row[erp_name_col],
                        'схожесть %': round(score, 1)
                    })
                    matched_count += 1
            
            # Сохранение результата
            if results:
                output_file = "результат_сопоставления.xlsx"
                df_final = pd.DataFrame(results)
                df_final.to_excel(output_file, index=False)
                
                self.log(f"Готово! Найдено {matched_count} совпадений из {total} товаров")
                self.log(f"Результат сохранен в файл: {output_file}")
                
                messagebox.showinfo("Успех", 
                    f"Сопоставление завершено!\n"
                    f"Найдено совпадений: {matched_count} из {total}\n"
                    f"Результат сохранен в: {output_file}")
            else:
                self.log("Совпадений не найдено. Попробуйте снизить порог схожести.")
                messagebox.showwarning("Внимание", "Совпадений не найдено")
                
        except Exception as e:
            error_msg = f"Ошибка: {str(e)}"
            self.log(error_msg)
            messagebox.showerror("Ошибка", error_msg)
        finally:
            # Разблокируем интерфейс
            self.progress.stop()
            self.process_button.config(state="normal")
            
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


if __name__ == "__main__":
    root = tk.Tk()
    app = ProductMatcherGUI(root)
    root.mainloop()