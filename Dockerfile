FROM python:3.11-slim

WORKDIR /app

# Установка необходимых пакетов
RUN pip install pandas openpyxl xlrd>=2.0.1

# Создаем директории
RUN mkdir -p input output

# Копирование файлов
COPY compare_prices.py .

# Запуск скрипта
CMD ["python", "compare_prices.py"] 