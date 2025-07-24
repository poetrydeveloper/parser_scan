import os
import re
import math  # 👈 Новый импорт
from datetime import datetime
from django.core.management.base import BaseCommand
from openpyxl import load_workbook
from django.core.files import File
from django.db import transaction
from termcolor import cprint

from parser.models import ExcelFile, Product, Invoice

INPUT_DIR = 'parser/input'
FILENAME_PATTERN = re.compile(
    r'^(?P<number>\d+?)_(?P<date>\d{2}-\d{2}-\d{4})_(?P<page>\d+)\.xlsx$'
)


# --- Старые функции (без изменений) ---
def validate_header_row(row):
    """Проверяет, является ли строка заголовком с номерами колонок"""
    expected_headers = ['1', '2', '3', '4']
    return all(str(cell).strip() == expected_headers[i] for i, cell in enumerate(row[:4]) if cell)


def strict_float_conversion(value, row_index, field_name):
    """Строгая конвертация в float с генерацией исключений"""
    if value is None:
        raise ValueError(f"Пустое значение в поле {field_name} (строка {row_index})")

    original_value = str(value).strip()
    if not original_value:
        raise ValueError(f"Пустое значение в поле {field_name} (строка {row_index})")

    try:
        cleaned_value = original_value.replace(',', '.').replace(' ', '')
        return float(cleaned_value)
    except ValueError:
        raise ValueError(
            f"Невозможно преобразовать '{original_value}' в число "
            f"(строка {row_index}, поле {field_name})"
        )


# --- Новая функция проверки цены/стоимости ---
def validate_price_quantity_total(row, row_index):
    """Проверяет соответствие: цена = стоимость / количество"""
    try:
        quantity = strict_float_conversion(row[2], row_index, "Количество")
        price = strict_float_conversion(row[3], row_index, "Цена")
        total = strict_float_conversion(row[4], row_index, "Стоимость")

        calculated_price = total / quantity if quantity != 0 else 0
        if not math.isclose(price, calculated_price, rel_tol=1e-4):
            raise ValueError(
                f"Несоответствие цены и стоимости (строка {row_index}): "
                f"цена={price}, но {total}/{quantity}≈{calculated_price:.2f}"
            )
    except (ValueError, ZeroDivisionError) as e:
        raise ValueError(f"Ошибка проверки цены/стоимости: {e}")


class Command(BaseCommand):
    help = "Загружает и парсит Excel-файлы с использованием транзакций и bulk_create"

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

            try:
                with transaction.atomic():
                    match = FILENAME_PATTERN.match(filename)
                    if not match:
                        raise ValueError("Имя файла не соответствует шаблону!")

                    number = match.group('number')
                    date_str = match.group('date')
                    try:
                        date = datetime.strptime(date_str, "%d-%m-%Y").date()
                    except ValueError:
                        raise ValueError(f"Неверный формат даты: {date_str}")

                    invoice, _ = Invoice.objects.get_or_create(
                        number=number,
                        defaults={'date': date}
                    )

                    excel_file = ExcelFile(invoice=invoice, processed=False)
                    with open(filepath, 'rb') as f:
                        excel_file.file.save(filename, File(f), save=True)

                    wb = load_workbook(excel_file.file.path, data_only=True)
                    ws = wb.active

                    products_to_create = []
                    validation_errors = []

                    for row_index, row in enumerate(ws.iter_rows(min_row=1, values_only=True), start=1):
                        if not any(row[:4]):
                            continue

                        if validate_header_row(row):
                            cprint(f" ⚠️ Пропущена строка с номерами колонок (строка {row_index})", 'yellow')
                            continue

                        try:
                            name = str(row[0]).strip() if row[0] else None
                            if not name:
                                raise ValueError(f"Пустое наименование товара (строка {row_index})")

                            quantity = strict_float_conversion(row[2], row_index, "Количество")
                            price = strict_float_conversion(row[3], row_index, "Цена")

                            # --- НОВАЯ ПРОВЕРКА (добавлена без изменения старого кода) ---
                            if len(row) > 4 and row[4]:  # Проверяем, есть ли столбец "Стоимость"
                                validate_price_quantity_total(row, row_index)
                            # --- Конец новой проверки ---

                            products_to_create.append(
                                Product(
                                    invoice=invoice,
                                    excel_file=excel_file,
                                    name=name,
                                    quantity=quantity,
                                    price=price
                                )
                            )
                            cprint(f" ✅ Строка {row_index}: {name[:50]}...", 'green')

                        except Exception as e:
                            cprint(f"❌ ОШИБКА ВАЛИДАЦИИ (строка {row_index}): {e}", 'red')
                            cprint(f"    Содержимое строки: {row[:5]}", 'yellow')
                            validation_errors.append(f"Строка {row_index}: {e}")
                            continue

                    if validation_errors:
                        raise ValueError(
                            f"Найдены ошибки в {len(validation_errors)} строках:\n" +
                            "\n".join(validation_errors[:5]) +
                            ("\n..." if len(validation_errors) > 5 else "")
                        )

                    if not products_to_create:
                        raise ValueError("В файле не найдено данных для импорта.")

                    Product.objects.bulk_create(products_to_create)
                    excel_file.processed = True
                    excel_file.save()

                    cprint(f"✔ Файл успешно обработан. Добавлено записей: {len(products_to_create)}", 'green')

            except Exception as e:
                cprint(f"\n🔥 КРИТИЧЕСКАЯ ОШИБКА: {e}", 'red')
                cprint("⏹️ Парсинг файла остановлен. Все изменения для этого файла были отменены.", 'red')