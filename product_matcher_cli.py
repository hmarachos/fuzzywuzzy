#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Консольная версия программы сопоставления товаров
"""

import pandas as pd
from rapidfuzz import process, fuzz
from anyascii import anyascii
import re
import sys
import os


def clean_name(name):
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


def detect_columns(df, file_type):
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
        
    print(f"Файл {file_type}: ID = '{id_col}', Название = '{name_col}'")
    return id_col, name_col


def main():
    print("=== Программа сопоставления товаров ===\n")
    
    # Ввод путей к файлам
    if len(sys.argv) >= 3:
        site_file = sys.argv[1]
        erp_file = sys.argv[2]
        threshold = int(sys.argv[3]) if len(sys.argv) > 3 else 60
    else:
        site_file = input("Введите путь к файлу каталога сайта: ").strip()
        erp_file = input("Введите путь к файлу программы учета: ").strip()
        threshold = int(input("Введите минимальный порог схожести (30-95, по умолчанию 60): ") or "60")
    
    # Проверка существования файлов
    if not os.path.exists(site_file):
        print(f"Ошибка: Файл {site_file} не найден")
        return
    
    if not os.path.exists(erp_file):
        print(f"Ошибка: Файл {erp_file} не найден")
        return
    
    try:
        print("\nЗагружаю файлы...")
        
        # Загрузка файлов
        df_site = pd.read_excel(site_file)
        df_erp = pd.read_excel(erp_file)
        
        print(f"Загружен файл сайта: {len(df_site)} товаров")
        print(f"Загружен файл программы учета: {len(df_erp)} товаров")
        
        # Определяем колонки автоматически
        site_id_col, site_name_col = detect_columns(df_site, "сайта")
        erp_id_col, erp_name_col = detect_columns(df_erp, "программы учета")
        
        # Подготовка данных
        print("\nПодготавливаю данные для сопоставления...")
        df_site['clean'] = df_site[site_name_col].apply(clean_name)
        df_erp['clean'] = df_erp[erp_name_col].apply(clean_name)
        
        choices = df_erp['clean'].tolist()
        results = []
        
        print(f"Начинаю сопоставление с порогом {threshold}%...")
        
        # Процесс сопоставления
        total = len(df_site)
        matched_count = 0
        
        for idx, row in df_site.iterrows():
            if idx % 500 == 0:  # Обновляем прогресс каждые 500 товаров
                print(f"Обработано {idx}/{total} товаров...")
            
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
            
            print(f"\nГотово! Найдено {matched_count} совпадений из {total} товаров")
            print(f"Результат сохранен в файл: {output_file}")
            
            # Показываем несколько примеров
            print("\nПримеры найденных совпадений:")
            for i, result in enumerate(results[:5]):
                print(f"{i+1}. {result['наименование сайт']} -> {result['наименование программа']} ({result['схожесть %']}%)")
                
        else:
            print("Совпадений не найдено. Попробуйте снизить порог схожести.")
            
    except Exception as e:
        print(f"Ошибка: {str(e)}")


if __name__ == "__main__":
    main()