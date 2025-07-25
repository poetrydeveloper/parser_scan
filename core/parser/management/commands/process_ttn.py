import os
import re
import logging
from logging.handlers import RotatingFileHandler

from django.core.management.base import BaseCommand
from django.db import transaction
from termcolor import cprint
from difflib import SequenceMatcher
from parser.models import TTN, Product, Price, FinalSample

os.makedirs('parser/logs', exist_ok=True)

# Настройка логирования с правильной кодировкой
log_handler = RotatingFileHandler(
    'parser/logs/ttn_processing.log',
    maxBytes=5*1024*1024,  # 5 MB
    backupCount=3,
    encoding='utf-8'
)
logging.basicConfig(
    handlers=[log_handler],
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)


class Command(BaseCommand):
    help = "Обрабатывает TTN с улучшенным поиском и логированием"

    def parse_product_name(self, name):
        """Улучшенный парсер названия продукта"""
        try:
            clean_name = re.sub(r';.*$', '', name).strip()

            match = re.match(r'^(\d+)\s+([^\s]+)\s+(.+)$', clean_name)
            if match:
                return {
                    'code': match.group(1),
                    'article': match.group(2),
                    'name': match.group(3).strip()
                }

            match = re.match(r'^(\d+)\s+(.+?)\s+([^\s]+)$', clean_name)
            if match:
                return {
                    'code': match.group(1),
                    'article': match.group(3),
                    'name': match.group(2).strip()
                }

            match = re.match(r'^(\d+)\s+(.+)$', clean_name)
            if match:
                return {
                    'code': match.group(1),
                    'article': '',
                    'name': match.group(2).strip()
                }

            return None
        except Exception as e:
            logger.error(f"Ошибка разбора имени: {name} - {str(e)}")
            return None

    def article_similarity(self, a, b):
        """Улучшенное сравнение артикулов"""
        if not a or not b:
            return 0

        a_clean = re.sub(r'[^a-zA-Z0-9]', '', a).lower()
        b_clean = re.sub(r'[^a-zA-Z0-9]', '', b).lower()

        if a_clean == b_clean:
            return 1.0

        if a_clean in b_clean or b_clean in a_clean:
            return 0.9

        return SequenceMatcher(None, a_clean, b_clean).ratio()

    def text_name_similarity(self, name1, name2):
        """Сравнение названий по словам и совпавшим символам"""
        def normalize(text):
            return re.sub(r'[^\w\s]', '', text.lower()).split()

        words1 = normalize(name1)
        words2 = normalize(name2)

        matches = []
        for w1 in words1:
            for w2 in words2:
                sim = SequenceMatcher(None, w1, w2).ratio()
                if sim >= 0.5:
                    matches.append((w1, w2, sim))

        return matches

    def find_price_matches(self, code, article):
        """Поиск всех возможных совпадений в прайсе с логированием"""
        prices = Price.objects.filter(code=code)
        matches = []

        for price in prices:
            similarity = self.article_similarity(price.article, article)
            if similarity >= 0.5:
                matches.append({
                    'price': price,
                    'similarity': similarity,
                    'details': f"{price.code} {price.article} ({price.name[:30]}...)"
                })

        if prices.exists() and not matches:
            logger.debug(f"Для кода {code} найдены в прайсе, но нет подходящих артикулов:")
            for p in prices:
                logger.debug(f"- {p.article} (ID: {p.id})")

        return matches

    def handle(self, *args, **options):
        ttn_number = input("Введите номер TTN для обработки: ").strip()
        logger.info(f"Начало обработки TTN {ttn_number}")

        try:
            ttn = TTN.objects.get(number=ttn_number)
        except TTN.DoesNotExist:
            error_msg = f"TTN с номером {ttn_number} не найдена"
            cprint(f"❌ {error_msg}", 'red')
            logger.error(error_msg)
            return

        products = Product.objects.filter(ttn=ttn).order_by('id')
        if not products:
            error_msg = f"Для TTN {ttn_number} нет товаров"
            cprint(f"ℹ️ {error_msg}", 'yellow')
            logger.warning(error_msg)
            return

        cprint(f"\nНачинаем обработку {len(products)} товаров...", 'cyan')
        logger.info(f"Найдено {len(products)} товаров для обработки")

        with transaction.atomic():
            for idx, product in enumerate(products, 1):
                log_prefix = f"[{idx}/{len(products)}]"
                logger.info(f"{log_prefix} Обработка: {product.name[:100]}...")

                parsed = self.parse_product_name(product.name)
                if not parsed:
                    FinalSample.objects.create(
                        ttn_number=ttn_number,
                        product_name=product.name,
                        product_quantity=product.quantity,
                        product_price=product.price,
                        product_full_price=product.full_price,
                        match_status='none'
                    )
                    error_msg = f"{log_prefix} Не удалось разобрать название"
                    cprint(f"❌ {error_msg}", 'red')
                    logger.error(f"{error_msg}: {product.name[:200]}")
                    continue

                logger.debug(f"Разобрано: код={parsed['code']}, артикул={parsed['article']}, название={parsed['name'][:50]}...")

                matches = self.find_price_matches(parsed['code'], parsed['article'])

                if matches:
                    best_match = max(matches, key=lambda x: x['similarity'])
                    similarity = best_match['similarity']
                    price_match = best_match['price']

                    status = 'full' if similarity >= 0.85 else 'partial'
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
                        product_full_price=product.full_price,
                        match_status=status
                    )
                    log_msg = f"{log_prefix} Совпадение ({similarity:.0%}): {parsed['code']} | Продукт: '{parsed['article']}' ≈ Прайс: '{price_match.article}'"
                    if status == 'full':
                        cprint(f"✅ {log_msg}", 'green')
                    else:
                        cprint(f"⚠️ {log_msg}", 'yellow')
                    logger.info(log_msg)

                else:
                    prices = Price.objects.filter(code=parsed['code'])
                    best_text_match = None
                    best_match_info = []
                    max_matches = 0

                    for price in prices:
                        word_matches = self.text_name_similarity(parsed['name'], price.name)
                        if len(word_matches) > max_matches:
                            max_matches = len(word_matches)
                            best_text_match = price
                            best_match_info = word_matches

                    if best_text_match and max_matches >= 2:
                        FinalSample.objects.create(
                            ttn_number=ttn_number,
                            price_code=best_text_match.code,
                            price_type=best_text_match.type,
                            price_article=best_text_match.article,
                            price_name=best_text_match.name,
                            price1=best_text_match.price1,
                            price2=best_text_match.price2,
                            price_clear=best_text_match.price_clear,
                            product_name=product.name,
                            product_quantity=product.quantity,
                            product_price=product.price,
                            product_full_price=product.full_price,
                            match_status='textual'
                        )
                        log_msg = f"{log_prefix} 🔍 Доп. совпадение по тексту: найдено {max_matches} совпавших слов."
                        for w1, w2, sim in best_match_info:
                            log_msg += f"\n   \"{w1}\" ≈ \"{w2}\" ({sim:.0%})"
                        cprint(log_msg, 'blue')
                        logger.info(log_msg)
                    else:
                        FinalSample.objects.create(
                            ttn_number=ttn_number,
                            product_name=product.name,
                            product_quantity=product.quantity,
                            product_price=product.price,
                            match_status='none'
                        )
                        log_msg = f"{log_prefix} ❌ Нет совпадений даже по тексту для: {parsed['code']} {parsed['article']}"
                        cprint(log_msg, 'red')
                        logger.warning(log_msg)

        logger.info(f"Обработка TTN {ttn_number} завершена")
        cprint(f"\nОбработка TTN {ttn_number} завершена!", 'cyan', attrs=['bold'])
