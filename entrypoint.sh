#!/bin/bash
set -e

# Функция для очистки имен файлов
clean_filenames() {
    dir=$1
    echo "Проверка и исправление имен файлов в директории $dir..."
    for file in "$dir"/*; do
        if [[ "$file" == *":Zone.Identifier"* ]]; then
            newname=$(echo "$file" | sed 's/:Zone.Identifier//g')
            echo "Переименование: $file -> $newname"
            mv "$file" "$newname"
        fi
    done
}

# Очистка имен файлов в директориях input/old и input/new
clean_filenames "/app/input/old"
clean_filenames "/app/input/new"

# Вывод списка файлов для проверки
echo "Файлы в директории input/old:"
ls -la /app/input/old

echo "Файлы в директории input/new:"
ls -la /app/input/new

# Запуск основного скрипта
echo "Запуск скрипта сравнения цен..."
python compare_prices.py
