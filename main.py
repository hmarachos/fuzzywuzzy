import pandas as pd
from rapidfuzz import process, fuzz
from anyascii import anyascii
import re


def clean_name(name):
    if not isinstance(name, str):
        return ""
    # Убираем страны в косых чертах /Беларусь/, /Германия/
    name = re.sub(r'/[^/]+/', ' ', name)
    # Транслитерация для сопоставления RU/EN брендов
    name = anyascii(name)
    # Оставляем буквы и цифры
    name = re.sub(r'[^a-zA-Z0-9\s]', ' ', name).lower()
    return " ".join(name.split())


# 1. Загрузка (проверь названия файлов)
df_site = pd.read_excel('site_catalog.xlsx')
df_erp = pd.read_excel('erp_program.xlsx')

# 2. Подготовка данных (используем точные заголовки с твоих фото)
# Файл сайта: _ID_ , Наименование
# Файл программы: id , наименование

# Чистим названия для поиска
df_site['clean'] = df_site['Наименование'].apply(clean_name)
df_erp['clean'] = df_erp['наименование'].apply(clean_name)

choices = df_erp['clean'].tolist()
results = []

print("Начинаю сопоставление...")

# 3. Цикл поиска
for idx, row in df_site.iterrows():
    # WRatio лучше всего подходит для смеси языков и сокращений
    extract = process.extractOne(
        row['clean'],
        choices,
        scorer=fuzz.WRatio
    )

    if extract:
        match_text, score, match_idx = extract
        matched_row = df_erp.iloc[match_idx]

        results.append({
            'id сайт': row['_ID_'],
            'наименование сайт': row['Наименование'],
            'id программа': matched_row['id'],
            'наименование программа': matched_row['наименование'],
            'score': round(score, 1)
        })

# 4. Сохранение результата
df_final = pd.DataFrame(results)
df_final.to_excel('matching_results.xlsx', index=False)

print(f"Готово! Обработано {len(results)} позиций.")
print("Результат сохранен в matching_results.xlsx")