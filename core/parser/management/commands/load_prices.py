# parser/management/commands/load_prices.py
import os
import re
from datetime import datetime
from django.core.management.base import BaseCommand
from openpyxl import load_workbook
from django.db import transaction
from termcolor import cprint
from parser.models import Price

INPUT_DIR = 'parser/input/input_prices'
DATE_PATTERN = re.compile(r'price_(\d{2}-\d{2}-\d{4})\.xls$')

class Command(BaseCommand):
    help = "Загружает прайс-листы из папки input_prices"

    def handle(self, *args, **options):
        if not os.path.exists(INPUT_DIR):
            os.makedirs(INPUT_DIR, exist_ok=True)
            cprint(f"Создана папка {INPUT_DIR}", 'yellow')
            return

        files = [f for f in os.listdir(INPUT_DIR) if f.endswith('.xls') or f.endswith('.xlsx')]
        if not files:
            cprint("Нет файлов прайсов в папке input_prices", 'yellow')
            return

        for filename in files:
            filepath = os.path.join(INPUT_DIR, filename)
            cprint(f"\nОбработка файла: {filename}", 'cyan', attrs=['bold'])

            try:
                with transaction.atomic():
                    # Извлекаем дату из имени файла
                    date_match = DATE_PATTERN.search(filename)
                    if not date_match:
                        cprint(f"⚠️ Имя файла {filename} не содержит дату в формате DD-MM-YYYY", 'yellow')
                        continue

                    price_date = datetime.strptime(date_match.group(1), "%d-%m-%Y").date()

                    wb = load_workbook(filepath, data_only=True)
                    ws = wb.active

                    created_count = 0
                    skipped_count = 0
                    errors = []

                    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
                        if not any(row):
                            continue

                        try:
                            code = str(row[0]).strip() if row[0] else None
                            if not code:
                                raise ValueError("Пустой код товара")

                            # Создаем или обновляем запись
                            price, created = Price.objects.get_or_create(
                                code=code,
                                defaults={
                                    'type': str(row[1]).strip() if row[1] else '',
                                    'article': str(row[2]).strip() if row[2] else '',
                                    'name': str(row[3]).strip() if row[3] else '',
                                    'price1': float(str(row[4]).replace(',', '.')) if row[4] else 0,
                                    'price2': float(str(row[5]).replace(',', '.')) if row[5] else 0,
                                    'stock': int(row[6]) if row[6] else 0,
                                    'quantity': int(row[7]) if row[7] else 0,
                                    'price_date': price_date
                                }
                            )

                            if created:
                                created_count += 1
                                cprint(f"✅ Добавлен: {code} - {price.name[:30]}...", 'green')
                            else:
                                skipped_count += 1
                                cprint(f"↻ Пропущен (уже существует): {code}", 'blue')

                        except Exception as e:
                            error_msg = f"❌ Ошибка в строке {row_idx}: {e}\n    Данные: {row[:9]}"
                            cprint(error_msg, 'red')
                            errors.append(error_msg)
                            continue

                    cprint(
                        f"\nИтоги по файлу {filename}:\n"
                        f"Добавлено: {created_count}\n"
                        f"Пропущено: {skipped_count}\n"
                        f"Ошибок: {len(errors)}",
                        'cyan'
                    )

                    if errors:
                        cprint("\nПоследние ошибки:", 'yellow')
                        for error in errors[-3:]:
                            cprint(error, 'red')

            except Exception as e:
                cprint(f"\n🔥 КРИТИЧЕСКАЯ ОШИБКА: {e}", 'red')
                continue