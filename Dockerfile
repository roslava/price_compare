FROM python:3.11-slim

WORKDIR /app

# Установка необходимых пакетов
RUN pip install pandas openpyxl xlrd>=2.0.1

# Создаем директории
RUN mkdir -p input/old input/new result

# Копирование файлов
COPY compare_prices.py .

# Скрипт для обработки файлов перед запуском основного скрипта
COPY entrypoint.sh /app/entrypoint.sh
RUN chmod +x /app/entrypoint.sh

ENTRYPOINT ["/app/entrypoint.sh"] 