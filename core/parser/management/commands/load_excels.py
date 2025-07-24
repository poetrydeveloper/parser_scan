import os
import re
import math
from datetime import datetime
from django.core.management.base import BaseCommand
from openpyxl import load_workbook
from django.core.files import File
from django.db import transaction
from termcolor import cprint

from parser.models import ExcelFile, Product, Invoice, TTN

INPUT_DIR = 'parser/input'
FILENAME_PATTERN = re.compile(
    r'^(?P<number>\d+?)_(?P<date>\d{2}-\d{2}-\d{4})_(?P<page>\d+)\.xlsx$'
)


def validate_header_row(row):
    expected_headers = ['1', '2', '3', '4']
    return all(str(cell).strip() == expected_headers[i] for i, cell in enumerate(row[:4]) if cell)


def strict_float_conversion(value, row_index, field_name):
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


def validate_price_quantity_total(row, row_index):
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
    help = "Загружает и парсит Excel-файлы с группировкой по ТТН"

    def handle(self, *args, **options):
        if not os.path.exists(INPUT_DIR):
            os.makedirs(INPUT_DIR, exist_ok=True)
            cprint(f"Создана папка {INPUT_DIR}", 'yellow')

        files = [f for f in os.listdir(INPUT_DIR) if f.endswith('.xlsx')]
        if not files:
            cprint("Нет .xlsx файлов в папке input", 'yellow')
            return

        ttn_data = {}

        for filename in files:
            filepath = os.path.join(INPUT_DIR, filename)
            cprint(f"\nОбработка файла: {filename}", 'cyan', attrs=['bold'])

            try:
                with transaction.atomic():
                    match = FILENAME_PATTERN.match(filename)
                    if not match:
                        raise ValueError("Имя файла не соответствует шаблону!")

                    ttn_number = match.group('number')
                    date_str = match.group('date')
                    page_number = match.group('page')

                    try:
                        date = datetime.strptime(date_str, "%d-%m-%Y").date()
                    except ValueError:
                        raise ValueError(f"Неверный формат даты: {date_str}")

                    # Создаем или получаем ТТН
                    ttn, created = TTN.objects.get_or_create(
                        number=ttn_number,
                        defaults={
                            'date': date,
                            'status': 'in_progress'
                        }
                    )

                    if created:
                        cprint(f"➕ Создана новая ТТН: {ttn_number}", 'green')
                    else:
                        cprint(f"↻ Обновляется существующая ТТН: {ttn_number}", 'blue')

                    # Создаем Invoice с привязкой к ТТН
                    invoice, _ = Invoice.objects.get_or_create(
                        number=f"{ttn_number}_{page_number}",
                        defaults={
                            'date': date,
                            'ttn': ttn
                        }
                    )

                    # Создаем ExcelFile с привязкой к ТТН
                    excel_file = ExcelFile(
                        invoice=invoice,
                        ttn=ttn,
                        processed=False,
                        page_number=page_number
                    )
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

                            if len(row) > 4 and row[4]:
                                validate_price_quantity_total(row, row_index)

                            products_to_create.append(
                                Product(
                                    invoice=invoice,
                                    excel_file=excel_file,
                                    ttn=ttn,
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

                    # Обновляем статистику по ТТН
                    if ttn_number not in ttn_data:
                        ttn_data[ttn_number] = {
                            'total_products': 0,
                            'files': set()
                        }

                    ttn_data[ttn_number]['total_products'] += len(products_to_create)
                    ttn_data[ttn_number]['files'].add(filename)

                    excel_file.processed = True
                    excel_file.save()

                    cprint(f"✔ Файл успешно обработан. Добавлено товаров: {len(products_to_create)}", 'green')

            except Exception as e:
                cprint(f"\n🔥 ОШИБКА: {e}", 'red')
                continue

        # Финальное обновление ТТН
        for ttn_number, data in ttn_data.items():
            try:
                ttn = TTN.objects.get(number=ttn_number)
                ttn.total_products = data['total_products']
                ttn.processed_files = len(data['files'])
                ttn.status = 'completed'
                ttn.save()
                cprint(f"✅ ТТН {ttn_number} завершена. Товаров: {data['total_products']}, файлов: {len(data['files'])}",
                       'green')
            except Exception as e:
                cprint(f"⚠️ Ошибка обновления ТТН {ttn_number}: {e}", 'yellow')