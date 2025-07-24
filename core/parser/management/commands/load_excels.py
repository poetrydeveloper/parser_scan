import os
import re
from datetime import datetime
from django.core.management.base import BaseCommand
from openpyxl import load_workbook
from django.core.files import File
from termcolor import cprint

from parser.models import ExcelFile, Product, Invoice

INPUT_DIR = 'parser/input'
FILENAME_PATTERN = re.compile(
    r'^(?P<number>\d+?)_(?P<date>\d{2}-\d{2}-\d{4})_(?P<page>\d+)\.xlsx$'
)

def safe_float(value):
    """Безопасное преобразование в float с заменой запятых"""
    if isinstance(value, str):
        value = value.replace(',', '.').strip()
    try:
        return float(value) if value else 0.0
    except ValueError:
        return 0.0

class Command(BaseCommand):
    help = "Загружает и парсит Excel-файлы из папки input (только первые 4 колонки)"

    def handle(self, *args, **kwargs):
        if not os.path.exists(INPUT_DIR):
            os.makedirs(INPUT_DIR, exist_ok=True)
            cprint(f"Создана папка {INPUT_DIR}", 'yellow')

        files = [f for f in os.listdir(INPUT_DIR) if f.endswith('.xlsx')]
        if not files:
            cprint("Нет .xlsx файлов в папке input", 'yellow')
            return

        for filename in files:
            filepath = os.path.join(INPUT_DIR, filename)
            cprint(f"\nОбработка файла: {filename}", 'cyan', attrs=['bold'])

            # Проверка имени файла
            match = FILENAME_PATTERN.match(filename)
            if not match:
                cprint(" ⛔ Имя файла не соответствует шаблону! Пропускаем...", 'red')
                continue

            # Извлечение номера и даты
            number = match.group('number')
            date_str = match.group('date')
            try:
                date = datetime.strptime(date_str, "%d-%m-%Y").date()
            except Exception as e:
                cprint(f" ⛔ Неверная дата в имени файла: {e}", 'red')
                continue

            # Создание записи о файле
            with open(filepath, 'rb') as f:
                excel_file = ExcelFile.objects.create(file=File(f, name=filename))

            invoice, _ = Invoice.objects.get_or_create(number=number, defaults={'date': date})
            excel_file.invoice = invoice
            excel_file.save()

            # Обработка Excel
            try:
                wb = load_workbook(excel_file.file.path, data_only=True)
                ws = wb.active

                success_count = 0
                error_count = 0

                for row_index, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
                    if not any(row[:4]):  # Пропускаем пустые строки
                        continue

                    try:
                        # Обрабатываем только первые 4 колонки
                        name = str(row[0]).strip() if row[0] else "Без названия"
                        quantity = safe_float(row[2])
                        price = safe_float(row[3])

                        Product.objects.create(
                            invoice=invoice,
                            excel_file=excel_file,
                            name=name,
                            quantity=quantity,
                            price=price
                        )
                        cprint(f" ✅ Строка {row_index}: {name[:50]}...", 'green')
                        success_count += 1

                    except Exception as e:
                        cprint(f" ❌ Ошибка в строке {row_index}: {e}", 'red')
                        cprint(f"    👉 Данные: {row[:4]}", 'yellow')
                        error_count += 1

                excel_file.processed = True
                excel_file.save()
                cprint(f"✔ Файл обработан. Успешно: {success_count}, Ошибок: {error_count}", 'blue')

            except Exception as e:
                cprint(f"🔥 Критическая ошибка при обработке файла: {e}", 'red')
                excel_file.delete()  # Удаляем файл если не смогли обработать