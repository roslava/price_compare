import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Alignment, Font, Border, Side
import os

# Создаем директории, если они не существуют
os.makedirs('input', exist_ok=True)
os.makedirs('result', exist_ok=True)

# Пути к файлам
input_file_2024 = os.path.join('input', 'Прайс 2024 для страховых.xls')
input_file_2025 = os.path.join('input', 'Прайс 2025 для СТРАХОВЫХ.xls')
output_file = os.path.join('result', 'Прайс 2025 для СТРАХОВЫХ (с изменениями).xlsx')

# Проверяем наличие входных файлов
if not os.path.exists(input_file_2024) or not os.path.exists(input_file_2025):
    print("Ошибка: Поместите файлы прайс-листов в папку 'input':")
    print(f"- {os.path.basename(input_file_2024)}")
    print(f"- {os.path.basename(input_file_2025)}")
    exit(1)

# Загрузка файлов
try:
    df_2024 = pd.read_excel(input_file_2024)
    df_2025 = pd.read_excel(input_file_2025)
    
    # Переименовываем колонки в обоих DataFrame
    column_names = {
        df_2024.columns[0]: '№ услуги',
        df_2024.columns[1]: 'Артикул',
        df_2024.columns[2]: 'Наименование услуги',
        df_2024.columns[-1]: 'Стоимость услуг 2024'
    }
    df_2024 = df_2024.rename(columns=column_names)
    
    column_names = {
        df_2025.columns[0]: '№ услуги',
        df_2025.columns[1]: 'Артикул',
        df_2025.columns[2]: 'Наименование услуги',
        df_2025.columns[-1]: 'Стоимость услуг 2025'
    }
    df_2025 = df_2025.rename(columns=column_names)
    
    # Обновляем имена колонок с ценами
    price_column_2024 = 'Стоимость услуг 2024'
    price_column_2025 = 'Стоимость услуг 2025'
    
    print(f"\nИспользуем колонки:")
    print(f"2024: {price_column_2024}")
    print(f"2025: {price_column_2025}")
    
    # Преобразуем значения цен в числовой формат
    df_2024[price_column_2024] = pd.to_numeric(df_2024[price_column_2024], errors='coerce')
    df_2025[price_column_2025] = pd.to_numeric(df_2025[price_column_2025], errors='coerce')
    
    # Удаляем пустые строки
    # Строка считается пустой, если все значения в ней NaN или пустые строки
    df_2024 = df_2024.dropna(how='all').reset_index(drop=True)
    df_2025 = df_2025.dropna(how='all').reset_index(drop=True)
    
    # Удаляем строки, где все значения - пустые строки
    df_2024 = df_2024[~(df_2024.astype(str) == '').all(axis=1)].reset_index(drop=True)
    df_2025 = df_2025[~(df_2025.astype(str) == '').all(axis=1)].reset_index(drop=True)
    
except Exception as e:
    print(f"Ошибка при загрузке файлов: {e}")
    exit(1)

# Создаем словарь цен 2024 года для быстрого поиска
prices_2024 = dict(zip(df_2024['№ услуги'], df_2024[price_column_2024]))

# Добавляем колонку с ценами 2024 года
df_2025['Стоимость услуг 2024'] = df_2025['№ услуги'].map(prices_2024)

# Вычисляем процентное изменение
def calculate_price_change(row):
    service = row['№ услуги']  # Изменено с row.iloc[0] на row['№ услуги']
    price_2025 = row[price_column_2025]
    price_2024 = prices_2024.get(service, None)
    
    if price_2024 is not None and price_2024 != 0 and pd.notnull(price_2024) and pd.notnull(price_2025):
        change = ((price_2025 - price_2024) / price_2024) * 100
        return round(change, 1)  # Округляем до одного знака после запятой
    return None

# Добавляем новую колонку с процентным изменением
df_2025['Изменение цены %'] = df_2025.apply(calculate_price_change, axis=1)

# Форматируем процентное изменение для отображения
df_2025['Изменение цены % (текст)'] = df_2025['Изменение цены %'].apply(
    lambda x: f"+{x:.1f}%" if pd.notnull(x) and x > 0 else (f"{x:.1f}%" if pd.notnull(x) else "")
)

# Переупорядочиваем колонки
columns = list(df_2025.columns)
price_2025_index = columns.index(price_column_2025)
columns.remove('Стоимость услуг 2024')
columns.insert(price_2025_index, 'Стоимость услуг 2024')
df_2025 = df_2025[columns]

# Сохраняем результат в новый файл Excel
df_2025.to_excel(output_file, index=False)

# Добавляем цветовое форматирование
wb = load_workbook(output_file)
ws = wb.active

# Определяем цвета для форматирования
olive_fill = PatternFill(start_color='E6F2D5', end_color='E6F2D5', fill_type='solid')  # Светлый оливковый
light_blue_fill = PatternFill(start_color='DEEAF6', end_color='DEEAF6', fill_type='solid')  # Светлый синий
yellow_fill = PatternFill(start_color='FFFF99', end_color='FFFF99', fill_type='solid')  # Более мягкий желтый
light_gray_fill = PatternFill(start_color='F5F5F5', end_color='F5F5F5', fill_type='solid')  # Светло-серый
black_fill = PatternFill(start_color='000000', end_color='000000', fill_type='solid')  # Черный
white_fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')  # Белый

# Определяем стиль границ
thin_border = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)

# Определяем белые границы для ячеек вне таблицы
white_border = Border(
    left=Side(style='thin', color='FFFFFF'),
    right=Side(style='thin', color='FFFFFF'),
    top=Side(style='thin', color='FFFFFF'),
    bottom=Side(style='thin', color='FFFFFF')
)

# Находим индексы нужных колонок
price_col_2025 = None
price_col_2024 = None
percent_col = None
percent_text_col = None
service_name_col = None  # Добавляем поиск колонки с наименованием услуги

for idx, cell in enumerate(ws[1], 1):
    if cell.value == price_column_2025:
        price_col_2025 = idx
    elif cell.value == 'Стоимость услуг 2024':
        price_col_2024 = idx
    elif cell.value == 'Изменение цены %':
        percent_col = idx
    elif cell.value == 'Изменение цены % (текст)':
        percent_text_col = idx
    elif cell.value == 'Наименование услуги':
        service_name_col = idx

# Применяем форматирование
if price_col_2025 and percent_col:
    # Получаем размеры таблицы
    max_row = ws.max_row
    max_col = ws.max_column
    
    # Форматируем заголовки колонок (первая строка)
    for col in range(1, max_col + 1):
        header_cell = ws.cell(row=1, column=col)
        header_cell.border = thin_border
        header_cell.font = Font(bold=True)
        header_cell.alignment = Alignment(
            wrap_text=True,
            horizontal='center',
            vertical='center',
            shrink_to_fit=False,
            indent=0
        )
        # Добавляем отступы в текст заголовка
        if header_cell.value:
            header_cell.value = f" {header_cell.value} "
    
    # Устанавливаем автоматическую высоту для строки заголовков
    ws.row_dimensions[1].height = None
    
    # Применяем форматирование к остальным ячейкам
    for row in range(2, max_row + 1):  # Начинаем со второй строки
        # Устанавливаем автоматическую высоту строки
        ws.row_dimensions[row].height = None
        
        for col in range(1, max_col + 1):
            cell = ws.cell(row=row, column=col)
            # Добавляем границы для всех ячеек
            cell.border = thin_border
            
            # Настраиваем форматирование в зависимости от типа колонки
            if col == service_name_col:
                cell.alignment = Alignment(
                    wrap_text=True,
                    vertical='center',
                    horizontal='left',
                    shrink_to_fit=False,
                    indent=0
                )
                # Добавляем пробелы в начало и конец текста для создания отступов
                if cell.value and isinstance(cell.value, str):
                    cell.value = f" {cell.value} "
            else:
                cell.alignment = Alignment(
                    horizontal='center',
                    vertical='center',
                    shrink_to_fit=False,
                    indent=0
                )
                # Добавляем пробелы в начало и конец текста для создания отступов
                if cell.value and isinstance(cell.value, (str, int, float)):
                    cell.value = f" {cell.value} "
            
            # Применяем цветовое форматирование для колонок с процентами
            if col == percent_col or col == percent_text_col:
                try:
                    if col == percent_col:
                        value = float(cell.value) if isinstance(cell.value, (int, float)) else None
                    else:  # percent_text_col
                        # Извлекаем число из текстового представления
                        value = float(cell.value.strip(' %+')) if cell.value and cell.value.strip(' %+-').replace('.', '').isdigit() else None
                    
                    if value is not None:
                        if value > 5:  # Значительное повышение
                            cell.fill = olive_fill
                        elif value < -5:  # Значительное снижение
                            cell.fill = light_blue_fill
                        elif value != 0:  # Небольшое изменение
                            cell.fill = yellow_fill
                except (ValueError, TypeError):
                    pass  # Пропускаем ячейки с некорректными значениями
        
        # Проверяем цену
        price_cell = ws.cell(row=row, column=price_col_2025)
        
        # Если цена отсутствует
        if price_cell.value is None or price_cell.value == '':
            # Проверяем, содержит ли строка информацию только в одной ячейке
            non_empty_cells = 0
            first_non_empty_cell = None
            
            for col in range(1, max_col + 1):
                cell = ws.cell(row=row, column=col)
                if cell.value and str(cell.value).strip():
                    non_empty_cells += 1
                    if first_non_empty_cell is None:
                        first_non_empty_cell = cell
            
            # Если информация только в одной ячейке
            if non_empty_cells == 1 and first_non_empty_cell:
                # Сохраняем значение из непустой ячейки
                cell_value = first_non_empty_cell.value.strip()  # Убираем лишние пробелы
                
                # Очищаем все ячейки в строке
                for col in range(1, max_col + 1):
                    ws.cell(row=row, column=col).value = None
                
                # Объединяем все ячейки в строке в одну
                ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=max_col)
                
                # Форматируем объединенную ячейку
                merged_cell = ws.cell(row=row, column=1)
                merged_cell.value = f" {cell_value} "  # Добавляем отступы
                merged_cell.fill = black_fill
                merged_cell.font = Font(bold=True, color='FFFFFF')
                merged_cell.alignment = Alignment(
                    horizontal='center',
                    vertical='center',
                    shrink_to_fit=False,
                    indent=0
                )
                merged_cell.border = thin_border
            else:
                # Для остальных строк без цен - только полужирный шрифт
                for col in range(1, max_col + 1):
                    cell = ws.cell(row=row, column=col)
                    cell.font = Font(bold=True)

# Устанавливаем ширину колонок с учетом отступов
ws.column_dimensions[chr(64 + service_name_col)].width = 65  # Увеличиваем ширину для учета отступов
for col in range(1, max_col + 1):
    if col != service_name_col:
        column_letter = chr(64 + col)
        current_width = ws.column_dimensions[column_letter].width
        if not current_width:
            current_width = 8.43  # Стандартная ширина Excel
        ws.column_dimensions[column_letter].width = current_width + 2  # Добавляем место для отступов

# Добавляем легенду в конец документа
# Сначала добавим пустую строку
last_row = ws.max_row + 2

# Добавляем заголовок легенды
legend_header = ws.cell(row=last_row, column=1, value="Легенда цветового обозначения:")
legend_header.font = Font(bold=True)
legend_header.border = thin_border
legend_header.fill = white_fill
last_row += 2

# Добавляем описание цветов
legend_items = [
    (olive_fill, "Повышение цены более чем на 5%"),
    (light_blue_fill, "Снижение цены более чем на 5%"),
    (yellow_fill, "Изменение цены в пределах ±5%")
]

for fill, description in legend_items:
    # Создаем ячейку с цветом
    color_cell = ws.cell(row=last_row, column=1, value="")
    color_cell.fill = fill
    color_cell.border = thin_border
    
    # Добавляем описание
    desc_cell = ws.cell(row=last_row, column=2, value=description)
    desc_cell.alignment = Alignment(vertical='center')
    desc_cell.border = thin_border
    desc_cell.fill = white_fill
    
    last_row += 1

# Устанавливаем ширину колонок для легенды
ws.column_dimensions['A'].width = 15
ws.column_dimensions['B'].width = 40

# Собираем статистику
changes = df_2025['Изменение цены %'].dropna()

# Расчет общего процентного изменения
total_price_2024 = sum(price for price in prices_2024.values() if pd.notnull(price))
total_price_2025 = df_2025[price_column_2025].sum()
total_price_change = ((total_price_2025 - total_price_2024) / total_price_2024) * 100

stats = {
    'Среднее изменение': round(changes.mean(), 1),
    'Максимальное повышение': round(changes.max(), 1),
    'Максимальное снижение': round(changes.min(), 1),
    'Количество повышений': (changes > 0).sum(),
    'Количество снижений': (changes < 0).sum(),
    'Без изменений': (changes == 0).sum(),
    'Новые услуги': df_2025['Изменение цены %'].isna().sum(),
    'Общее изменение всех цен': round(total_price_change, 1)
}

# Добавляем общее изменение цен в Excel файл
last_row += 2  # Добавляем пустую строку после легенды

# Добавляем заголовок для общего изменения
total_change_header = ws.cell(row=last_row, column=1, value="Общее изменение всех цен:")
total_change_header.font = Font(bold=True)
total_change_header.border = thin_border
total_change_header.fill = white_fill

# Добавляем значение общего изменения
total_change_value = ws.cell(row=last_row, column=2, 
    value=f"{'+' if total_price_change > 0 else ''}{total_price_change:.1f}%")
total_change_value.alignment = Alignment(horizontal='left')
total_change_value.border = thin_border

# Применяем цветовое форматирование для общего изменения
if total_price_change > 5:
    total_change_value.fill = olive_fill
elif total_price_change < -5:
    total_change_value.fill = light_blue_fill
elif total_price_change != 0:
    total_change_value.fill = yellow_fill
else:
    total_change_value.fill = white_fill

# Сохраняем файл с форматированием
wb.save(output_file)

# Выводим статистику
print("\nСтатистика изменения цен:")
for key, value in stats.items():
    if key in ['Среднее изменение', 'Максимальное повышение', 'Максимальное снижение', 'Общее изменение всех цен']:
        print(f"{key}: {'+' if value > 0 else ''}{value:.1f}%")
    else:
        print(f"{key}: {value}")

print(f"\nАнализ завершен. Результат сохранен в файл '{output_file}'") 