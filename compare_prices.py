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
    # Читаем Excel файлы с помощью xlrd для старых .xls файлов
    df_2024 = pd.read_excel(
        input_file_2024,
        engine='xlrd',  # Используем xlrd для .xls файлов
        dtype={3: str},  # Читаем колонку с ценами как текст
        na_values=['', ' '],  # Считаем пустые строки и пробелы как NaN
        keep_default_na=True
    )
    df_2025 = pd.read_excel(
        input_file_2025,
        engine='xlrd',  # Используем xlrd для .xls файлов
        dtype={3: str},  # Читаем колонку с ценами как текст
        na_values=['', ' '],  # Считаем пустые строки и пробелы как NaN
        keep_default_na=True
    )
    
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
    def clean_price(value):
        if pd.isna(value):
            return value
        try:
            # Преобразуем в строку и очищаем
            value_str = str(value)
            # Проверяем, является ли значение числом в научной нотации
            if 'e' in value_str.lower():
                return float(value_str)
            # Удаляем все пробелы (в начале, в конце и в середине)
            value_str = ''.join(value_str.split())
            # Удаляем апостроф в начале, если он есть
            value_str = value_str.lstrip("'")
            # Заменяем запятые на точки
            value_str = value_str.replace(',', '.')
            # Пробуем преобразовать в число
            if value_str:
                num = float(value_str)
                # Округляем до 2 знаков после запятой для устранения ошибок округления
                return round(num, 2)
            return None
        except (ValueError, TypeError):
            # Выводим проблемное значение для отладки
            print(f"Не удалось преобразовать значение: '{value}' (тип: {type(value)})")
            return None

    # Применяем очистку к колонкам с ценами и выводим уникальные значения для отладки
    print("\nУникальные значения в колонке цен 2024:")
    print(df_2024[price_column_2024].unique())
    
    print("\nУникальные значения в колонке цен 2025:")
    print(df_2025[price_column_2025].unique())
    
    # Применяем очистку к колонкам с ценами
    df_2024[price_column_2024] = df_2024[price_column_2024].apply(clean_price)
    df_2025[price_column_2025] = df_2025[price_column_2025].apply(clean_price)
    
    # Выводим результаты преобразования для проверки
    print("\nПосле преобразования:")
    print("Уникальные значения в колонке цен 2024:")
    print(df_2024[price_column_2024].unique())
    
    print("\nУникальные значения в колонке цен 2025:")
    print(df_2025[price_column_2025].unique())
    
    # Удаляем пустые строки
    # Строка считается пустой, если все значения в ней NaN или пустые строки
    df_2024 = df_2024.dropna(how='all').reset_index(drop=True)
    df_2025 = df_2025.dropna(how='all').reset_index(drop=True)
    
    # Удаляем строки, где все значения - пустые строки
    df_2024 = df_2024[~(df_2024.astype(str) == '').all(axis=1)].reset_index(drop=True)
    df_2025 = df_2025[~(df_2025.astype(str) == '').all(axis=1)].reset_index(drop=True)
    
    # Выводим статистику по количеству позиций
    print("\nСтатистика по количеству позиций:")
    print(f"Количество позиций в прайсе 2024: {len(df_2024)}")
    print(f"Количество позиций в прайсе 2025: {len(df_2025)}")
    
    # Подсчитываем уникальные артикулы (конвертируем в строки для корректного сравнения)
    articles_2024 = set(df_2024['Артикул'].astype(str).dropna())
    articles_2025 = set(df_2025['Артикул'].astype(str).dropna())
    
    print(f"\nКоличество уникальных артикулов:")
    print(f"Прайс 2024: {len(articles_2024)}")
    print(f"Прайс 2025: {len(articles_2025)}")
    
    # Анализируем разницу по артикулам
    missing_in_2025 = articles_2024 - articles_2025
    new_in_2025 = articles_2025 - articles_2024
    
    print(f"\nАнализ изменений по артикулам:")
    print(f"Услуг, отсутствующих в прайсе 2025: {len(missing_in_2025)}")
    if len(missing_in_2025) > 0:
        print("\nСписок артикулов, отсутствующих в прайсе 2025:")
        for article in sorted(missing_in_2025):
            # Конвертируем артикул в строку для поиска
            service = df_2024[df_2024['Артикул'].astype(str) == article]['Наименование услуги'].iloc[0]
            price = df_2024[df_2024['Артикул'].astype(str) == article][price_column_2024].iloc[0]
            print(f"Артикул: {article}, Цена 2024: {price} руб., Услуга: {service}")
    
    print(f"\nНовых услуг в прайсе 2025: {len(new_in_2025)}")
    if len(new_in_2025) > 0:
        print("\nСписок новых артикулов в прайсе 2025:")
        for article in sorted(new_in_2025):
            # Конвертируем артикул в строку для поиска
            service = df_2025[df_2025['Артикул'].astype(str) == article]['Наименование услуги'].iloc[0]
            price = df_2025[df_2025['Артикул'].astype(str) == article][price_column_2025].iloc[0]
            print(f"Артикул: {article}, Цена 2025: {price} руб., Услуга: {service}")
    
    # Сохраняем список удаленных позиций для добавления в конец документа
    removed_services = []
    if len(missing_in_2025) > 0:
        for article in sorted(missing_in_2025):
            service_num = df_2024[df_2024['Артикул'].astype(str) == article]['№ услуги'].iloc[0]
            service = df_2024[df_2024['Артикул'].astype(str) == article]['Наименование услуги'].iloc[0]
            price = df_2024[df_2024['Артикул'].astype(str) == article][price_column_2024].iloc[0]
            removed_services.append((service_num, article, price, service))
        
        # Сортируем список по номеру услуги
        # Преобразуем номер услуги в число для корректной сортировки
        removed_services.sort(key=lambda x: float(str(x[0]).replace(',', '.')) if str(x[0]).replace(',', '.').replace('.', '').isdigit() else float('inf'))

except Exception as e:
    print(f"Ошибка при загрузке файлов: {e}")
    exit(1)

# Создаем словарь цен 2024 года для быстрого поиска по артикулу (конвертируем в строки)
prices_2024 = dict(zip(df_2024['Артикул'].astype(str), df_2024[price_column_2024]))

# Добавляем колонку с ценами 2024 года, используя артикул для сопоставления
df_2025['Стоимость услуг 2024'] = df_2025['Артикул'].astype(str).map(prices_2024)

# Вычисляем процентное изменение
def calculate_price_change(row):
    article = row['Артикул']
    price_2025 = row[price_column_2025]
    price_2024 = prices_2024.get(article, None)
    
    if price_2024 is not None and price_2024 != 0 and pd.notnull(price_2024) and pd.notnull(price_2025):
        change = ((price_2025 - price_2024) / price_2024) * 100
        return round(change, 1)
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
                # Форматируем ячейки с ценами
                if col in [price_col_2024, price_col_2025]:
                    cell.number_format = '0.00'
                # Добавляем пробелы в начало и конец текста для создания отступов (кроме цен)
                elif cell.value and isinstance(cell.value, (str, int, float)):
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

# После легенды добавляем список удаленных позиций
if removed_services:
    last_row += 3  # Добавляем три пустые строки после легенды
    
    # Добавляем заголовок для списка удаленных позиций
    header = ws.cell(row=last_row, column=1, value="Список позиций, отсутствующих в прайсе 2025 года:")
    header.font = Font(bold=True)
    header.border = thin_border
    header.fill = white_fill
    ws.merge_cells(start_row=last_row, start_column=1, end_row=last_row, end_column=4)  # Увеличиваем до 4 колонок
    
    last_row += 2  # Пропускаем строку после заголовка
    
    # Добавляем заголовки колонок в новом порядке
    headers = ['№ услуги', 'Артикул', 'Наименование услуги', 'Цена 2024']
    for col, header_text in enumerate(headers, 1):
        cell = ws.cell(row=last_row, column=col, value=header_text)
        cell.font = Font(bold=True)
        cell.border = thin_border
        cell.fill = light_gray_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Добавляем данные в новом порядке
    for service_num, article, price, service in removed_services:
        last_row += 1
        # № услуги
        ws.cell(row=last_row, column=1, value=service_num).border = thin_border
        # Артикул
        ws.cell(row=last_row, column=2, value=article).border = thin_border
        # Наименование услуги
        service_cell = ws.cell(row=last_row, column=3, value=service)
        service_cell.border = thin_border
        service_cell.alignment = Alignment(horizontal='left', vertical='center')
        # Цена
        price_cell = ws.cell(row=last_row, column=4, value=price)
        price_cell.border = thin_border
        price_cell.number_format = '0.00'  # Формат с двумя десятичными знаками
    
    # Устанавливаем ширину колонок для списка удаленных позиций
    ws.column_dimensions['A'].width = 15  # Для номера услуги
    ws.column_dimensions['B'].width = 15  # Для артикула
    ws.column_dimensions['C'].width = 65  # Для наименования услуги
    ws.column_dimensions['D'].width = 15  # Для цены

# Вычисляем общую разницу цен (только для позиций, которые есть в обоих прайсах)
total_2024 = 0
total_2025 = 0

# Используем только те позиции, которые есть в обоих прайсах (исключаем строки с NaN в ценах)
df_common = df_2025[
    (df_2025['Стоимость услуг 2024'].notna()) & 
    (df_2025[price_column_2025].notna()) &
    (df_2025['Стоимость услуг 2024'] > 0) &
    (df_2025[price_column_2025] > 0)
]

# Считаем общие суммы
total_2024 = df_common['Стоимость услуг 2024'].sum()
total_2025 = df_common[price_column_2025].sum()

# Выводим информацию о количестве позиций для проверки
print(f"\nКоличество позиций для расчета общего изменения: {len(df_common)}")

# Вычисляем процент изменения
total_change_percent = ((total_2025 - total_2024) / total_2024 * 100) if total_2024 > 0 else 0

# Добавляем информацию об общем изменении цен в конец документа
last_row += 3  # Добавляем три пустые строки перед итогами

header = ws.cell(row=last_row, column=1, value="Общее изменение цен (без учета исключенных позиций):")
header.font = Font(bold=True)
header.border = thin_border
header.fill = white_fill
ws.merge_cells(start_row=last_row, start_column=1, end_row=last_row, end_column=2)

# Добавляем значение изменения
change_value = ws.cell(row=last_row, column=3, value=f"{'+' if total_change_percent > 0 else ''}{total_change_percent:.1f}%")
change_value.border = thin_border
change_value.alignment = Alignment(horizontal='left', vertical='center')

# Применяем цветовое форматирование
if total_change_percent > 5:
    change_value.fill = olive_fill
elif total_change_percent < -5:
    change_value.fill = light_blue_fill
elif total_change_percent != 0:
    change_value.fill = yellow_fill
else:
    change_value.fill = white_fill

# Добавляем дополнительную информацию
ws.cell(row=last_row + 1, column=1, value=f"Общая сумма цен 2024: {total_2024:,.2f} руб.").border = thin_border
ws.cell(row=last_row + 2, column=1, value=f"Общая сумма цен 2025: {total_2025:,.2f} руб.").border = thin_border

# Сохраняем файл с форматированием
wb.save(output_file)

# Выводим статистику в консоль
print(f"\nОбщее изменение цен (без учета исключенных позиций): {'+' if total_change_percent > 0 else ''}{total_change_percent:.1f}%")
print(f"Общая сумма цен 2024: {total_2024:,.2f} руб.")
print(f"Общая сумма цен 2025: {total_2025:,.2f} руб.")

print(f"\nАнализ завершен. Результат сохранен в файл '{output_file}'") 