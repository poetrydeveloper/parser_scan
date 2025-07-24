import re
from django.core.management.base import BaseCommand
from django.db import transaction
from termcolor import cprint
from difflib import SequenceMatcher
from parser.models import TTN, Product, Price, FinalSample


class Command(BaseCommand):
    help = "Обрабатывает TTN с поиском по коду и частичным совпадением артикула"

    def parse_product_name(self, name):
        """Разбирает наименование продукта на код, артикул и название"""
        try:
            parts = name.split(';', 1)
            main_part = parts[0].strip()
            match = re.match(r'^(\d+)\s+([^\s]+)\s+(.+)$', main_part)
            if match:
                return {
                    'code': match.group(1),
                    'article': match.group(2),
                    'name': match.group(3).strip()
                }
            return None
        except Exception:
            return None

    def article_similarity(self, a, b):
        """Вычисляет схожесть артикулов (0-1)"""
        return SequenceMatcher(None, a.lower(), b.lower()).ratio()

    def find_best_price_match(self, code, article):
        """Ищет лучший вариант в прайсе по коду и схожести артикула"""
        prices = Price.objects.filter(code=code)
        if not prices:
            return None

        best_match = None
        best_score = 0

        for price in prices:
            score = self.article_similarity(price.article, article)
            if score > best_score:
                best_score = score
                best_match = price

        return best_match if best_score >= 0.6 else None  # Порог схожести 60%

    def handle(self, *args, **options):
        ttn_number = input("Введите номер TTN для обработки: ").strip()

        try:
            ttn = TTN.objects.get(number=ttn_number)
        except TTN.DoesNotExist:
            cprint(f"❌ TTN с номером {ttn_number} не найдена", 'red')
            return

        products = Product.objects.filter(ttn=ttn).order_by('id')
        if not products:
            cprint(f"ℹ️ Для TTN {ttn_number} нет товаров", 'yellow')
            return

        cprint(f"\nНачинаем обработку {len(products)} товаров...", 'cyan')

        with transaction.atomic():
            for product in products:
                parsed = self.parse_product_name(product.name)
                if not parsed:
                    FinalSample.objects.create(
                        ttn_number=ttn_number,
                        product_name=product.name,
                        product_quantity=product.quantity,
                        product_price=product.price,
                        match_status='none'
                    )
                    cprint(f"❌ Не удалось разобрать название: {product.name[:50]}...", 'red')
                    continue

                # Поиск лучшего совпадения
                price_match = self.find_best_price_match(parsed['code'], parsed['article'])

                if price_match:
                    similarity = self.article_similarity(price_match.article, parsed['article'])
                    FinalSample.objects.create(
                        ttn_number=ttn_number,
                        price_code=price_match.code,
                        price_type=price_match.type,
                        price_article=price_match.article,
                        price_name=price_match.name,
                        price1=price_match.price1,
                        price2=price_match.price2,
                        price_clear=price_match.price_clear,
                        product_name=product.name,
                        product_quantity=product.quantity,
                        product_price=product.price,
                        match_status='full' if similarity >= 0.9 else 'partial'
                    )
                    status = "✅ Точное" if similarity >= 0.9 else "⚠️ Частичное"
                    cprint(
                        f"{status} совпадение: {parsed['code']} ({similarity:.0%}) {price_match.article} ≈ {parsed['article']}",
                        'green')
                else:
                    FinalSample.objects.create(
                        ttn_number=ttn_number,
                        product_name=product.name,
                        product_quantity=product.quantity,
                        product_price=product.price,
                        match_status='none'
                    )
                    cprint(f"❌ Нет совпадения для: {parsed['code']} {parsed['article']}", 'yellow')

        cprint(f"\nОбработка TTN {ttn_number} завершена!", 'cyan', attrs=['bold'])