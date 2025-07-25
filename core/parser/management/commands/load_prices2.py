# parser/management/commands/load_prices.py
import math
import os
import re
from django.core.management.base import BaseCommand
from django.db import transaction
from termcolor import cprint
from parser.models import Price
import pandas as pd  # pip install pandas openpyxl xlrd

INPUT_DIR = 'parser/input/input_prices'
FILENAME_PATTERN = re.compile(r'.*\.(xls|xlsx)$')

class Command(BaseCommand):
    help = "Загружает прайс-листы из папки input_prices"

    @staticmethod
    def clean_stock_value(value):
        if value is None:
            return ""
        if isinstance(value, (int, float)):
            return str(int(value)) if value == int(value) else str(value)
        return str(value).strip()

    @staticmethod
    def safe_float_convert(value):
        if value is None:
            return 0.0

        if isinstance(value, float) and math.isnan(value):
            print(f"⚠️ Пропущен NaN в значении: {value}")
            return 0.0

        try:
            return float(str(value).replace(',', '.').strip())
        except (ValueError, TypeError):
            return 0.0

    @staticmethod
    def safe_int_convert(value):
        if value is None or (isinstance(value, float) and math.isnan(value)):
            return 0
        try:
            return int(float(str(value).replace(',', '.').strip()))
        except (ValueError, TypeError):
            return 0

    def handle(self, *args, **options):
        if not os.path.exists(INPUT_DIR):
            os.makedirs(INPUT_DIR, exist_ok=True)
            cprint(f"Создана папка {INPUT_DIR}", 'yellow')
            return

        files = [
            f for f in os.listdir(INPUT_DIR)
            if f.lower().endswith(('.xls', '.xlsx')) and not f.startswith('~$')
        ]

        if not files:
            cprint("Нет файлов прайсов в папке input_prices", 'yellow')
            return

        existing_codes = set(Price.objects.values_list('code', flat=True))

        for filename in sorted(files):
            filepath = os.path.join(INPUT_DIR, filename)
            cprint(f"\nОбработка файла: {filename}", 'cyan', attrs=['bold'])

            try:
                try:
                    df = pd.read_excel(filepath)
                    rows = df.values.tolist()
                except Exception as e:
                    cprint(f"❌ Ошибка чтения файла: {e}", 'red')
                    continue

                stats = {'new': 0, 'exists': 0, 'errors': []}
                new_prices = []
                seen_codes = set()

                for row_idx, row in enumerate(rows, start=2):
                    if not any(row):
                        continue

                    try:
                        code = str(row[0]).strip() if len(row) > 0 and row[0] else None
                        if not code:
                            stats['errors'].append(f"Строка {row_idx}: отсутствует код")
                            continue

                        if code in existing_codes:
                            stats['exists'] += 1
                            cprint(f"⏩ Пропуск: код {code} уже существует", 'blue')
                            continue

                        if code in seen_codes:
                            stats['errors'].append(f"Строка {row_idx}: дублирующийся код в файле ({code})")
                            cprint(f"⚠️ Пропуск дублирующегося кода в файле: {code}", 'yellow')
                            continue

                        seen_codes.add(code)

                        article = str(row[2]).strip() if len(row) > 2 and row[2] else None

                        price_data = {
                            'type': str(row[1]).strip() if len(row) > 1 and row[1] else '',
                            'name': str(row[3]).strip() if len(row) > 3 and row[3] else '',
                            'price1': self.safe_float_convert(row[4]) if len(row) > 4 else 0,
                            'price2': self.safe_float_convert(row[5]) if len(row) > 5 else 0,
                            'stock': self.clean_stock_value(row[6]) if len(row) > 6 else "",
                            'quantity': self.safe_int_convert(row[7]) if len(row) > 7 else 0,
                            'price_clear': self.safe_float_convert(row[8]) if len(row) > 8 else 0
                        }

                        new_prices.append(Price(
                            code=code,
                            article=article,
                            **price_data
                        ))
                        existing_codes.add(code)
                        stats['new'] += 1
                        cprint(f"✅ Добавлен: {code} (артикул: {article or 'нет'})", 'green')

                    except Exception as e:
                        stats['errors'].append(f"Строка {row_idx}: {str(e)}")
                        cprint(f"❌ Ошибка в строке {row_idx}: {e}", 'red')

                if new_prices:
                    try:
                        with transaction.atomic():
                            Price.objects.bulk_create(new_prices, batch_size=500)
                    except Exception as e:
                        cprint(f"🔥 Ошибка при сохранении в БД: {e}", 'red')
                        continue

                cprint(
                    f"\nИтоги по файлу {filename}:\n"
                    f"Добавлено новых: {stats['new']}\n"
                    f"Пропущено существующих: {stats['exists']}\n"
                    f"Ошибок: {len(stats['errors'])}",
                    'cyan'
                )

                if stats['errors']:
                    cprint("\nПоследние ошибки:", 'yellow')
                    for err in stats['errors'][:5]:
                        cprint(f"• {err}", 'red')

            except Exception as e:
                cprint(f"\n🔥 Критическая ошибка файла {filename}: {e}", 'red')

        cprint("\nОбработка всех файлов завершена!", 'green', attrs=['bold'])
