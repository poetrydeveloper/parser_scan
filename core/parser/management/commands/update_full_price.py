from django.core.management.base import BaseCommand
from parser.models import Product

class Command(BaseCommand):
    help = "Обновляет поле full_price для всех товаров"

    def handle(self, *args, **options):
        updated = 0
        for product in Product.objects.all():
            if product.quantity is not None and product.price is not None:
                product.full_price = round(product.quantity * product.price, 2)
                product.save(update_fields=["full_price"])
                updated += 1
        self.stdout.write(self.style.SUCCESS(f"Обновлено {updated} товаров."))
