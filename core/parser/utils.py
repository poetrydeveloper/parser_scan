# parser/utils.py
from openpyxl import load_workbook
from .models import ExcelFile, ScanData, Invoice  # Добавьте этот импорт


def parse_excel_file(excel_file_id, invoice_number=None):
    try:
        # Получаем объект файла
        excel_file = ExcelFile.objects.get(id=excel_file_id)

        # Создаем или получаем накладную
        if invoice_number:
            invoice, created = Invoice.objects.get_or_create(number=invoice_number)
        else:
            invoice = Invoice.objects.create(number=f"auto_{excel_file.id}")

        # Связываем файл с накладной
        excel_file.invoice = invoice
        excel_file.save()

        # Открываем Excel-файл
        wb = load_workbook(excel_file.file.path)
        ws = wb.active

        # Обрабатываем каждую строку
        for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            try:
                # Парсим данные (адаптируйте под вашу структуру)
                raw_code = str(row[0]) if len(row) > 0 else ""
                raw_article = str(row[1]) if len(row) > 1 else ""
                raw_name = str(row[2]) if len(row) > 2 else ""

                # Создаем запись
                ScanData.objects.create(
                    excel_file=excel_file,
                    invoice=invoice,
                    code=int(raw_code) if raw_code.isdigit() else 0,
                    article=raw_article,
                    clean_name=raw_name.split(';')[0].strip(),
                    unit=str(row[3]) if len(row) > 3 else "",
                    quantity=float(row[4]) if len(row) > 4 else 0,
                    price=float(row[5]) if len(row) > 5 else 0,
                    original_data=str(row)
                )
            except Exception as e:
                print(f"Ошибка в строке {row_num}: {e}")
                continue

        # Помечаем файл как обработанный
        excel_file.processed = True
        excel_file.save()
        return True

    except Exception as e:
        print(f"Ошибка обработки файла: {e}")
        return False