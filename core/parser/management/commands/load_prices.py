# parser/management/commands/load_prices.py
import os
import re
from django.core.management.base import BaseCommand
import xlrd
from openpyxl import load_workbook
from django.db import transaction
from termcolor import cprint
from parser.models import Price

INPUT_DIR = 'parser/input/input_prices'
FILENAME_PATTERN = re.compile(r'.*\.(xls|xlsx)$')


class Command(BaseCommand):
    help = "Загружает прайс-листы из папки input_prices"

    def clean_stock_value(self, value):
        """Очищает и форматирует значение остатка для текстового хранения"""
        if value is None:
            return ""

        if isinstance(value, (int, float)):
            return str(int(value)) if value == int(value) else str(value)

        if isinstance(value, str):
            value = value.strip()
            value = re.sub(r'\s+', ' ', value)
            return value

        return str(value)

    def handle(self, *args, **options):
        if not os.path.exists(INPUT_DIR):
            os.makedirs(INPUT_DIR, exist_ok=True)
            cprint(f"Создана папка {INPUT_DIR}", 'yellow')
            return

        files = [f for f in os.listdir(INPUT_DIR)
                 if (f.endswith('.xls') or f.endswith('.xlsx'))
                 and not f.startswith('~$')]

        if not files:
            cprint("Нет файлов прайсов в папке input_prices", 'yellow')
            return

        for filename in sorted(files):
            filepath = os.path.join(INPUT_DIR, filename)
            cprint(f"\nОбработка файла: {filename}", 'cyan', attrs=['bold'])

            try:
                with transaction.atomic():
                    file_ext = os.path.splitext(filename)[1].lower()

                    if file_ext == '.xlsx':
                        try:
                            wb = load_workbook(filepath, data_only=True)
                            sheet = wb.active
                            rows = sheet.iter_rows(min_row=2, values_only=True)
                        except Exception as e:
                            cprint(f"❌ Ошибка чтения XLSX файла: {e}", 'red')
                            continue

                    elif file_ext == '.xls':
                        try:
                            wb = xlrd.open_workbook(filepath)
                            sheet = wb.sheet_by_index(0)
                            rows = (sheet.row_values(row_idx) for row_idx in range(1, sheet.nrows))
                        except Exception as e:
                            cprint(f"❌ Ошибка чтения XLS файла: {e}", 'red')
                            continue

                    new_count = 0
                    exists_count = 0
                    errors = []

                    for row_idx, row in enumerate(rows, start=2):
                        if not any(row):
                            continue

                        try:
                            code = str(row[0]).strip() if row[0] else None
                            article = str(row[2]).strip() if len(row) > 2 and row[2] else ''

                            if not code:
                                raise ValueError("Пустой код товара")

                            # Проверяем существование записи с таким code и article
                            exists = Price.objects.filter(code=code, article=article).exists()

                            if exists:
                                exists_count += 1
                                cprint(f"⏩ Пропуск (уже существует): {code} - {article}", 'blue')
                                continue

                            price_data = {
                                'type': str(row[1]).strip() if len(row) > 1 and row[1] else '',
                                'article': article,
                                'name': str(row[3]).strip() if len(row) > 3 and row[3] else '',
                                'price1': float(str(row[4]).replace(',', '.')) if len(row) > 4 and row[4] else 0,
                                'price2': float(str(row[5]).replace(',', '.')) if len(row) > 5 and row[5] else 0,
                                'stock': self.clean_stock_value(row[6]) if len(row) > 6 else "",
                                'quantity': int(float(row[7])) if len(row) > 7 and row[7] else 0,
                                'price_clear': float(str(row[8]).replace(',', '.')) if len(row) > 8 and row[8] else 0
                            }

                            Price.objects.create(**price_data)
                            new_count += 1
                            cprint(f"✅ Добавлен: {code} - {article}", 'green')

                        except Exception as e:
                            error_msg = f"❌ Ошибка в строке {row_idx}: {e}\n    Данные: {row[:9]}"
                            cprint(error_msg, 'red')
                            errors.append(error_msg)
                            continue

                    cprint(
                        f"\nИтоги по файлу {filename}:\n"
                        f"Новых добавлено: {new_count}\n"
                        f"Пропущено (уже существует): {exists_count}\n"
                        f"Ошибок: {len(errors)}",
                        'cyan'
                    )

                    if errors:
                        cprint("\nПоследние ошибки:", 'yellow')
                        for error in errors[:3]:
                            cprint(error, 'red')

            except Exception as e:
                cprint(f"\n🔥 КРИТИЧЕСКАЯ ОШИБКА: {e}", 'red')
                continue

        cprint("\nОбработка всех файлов завершена!", 'green', attrs=['bold'])