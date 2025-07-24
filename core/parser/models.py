# parser/models.py
from django.db import models


class Invoice(models.Model):
    number = models.CharField("Номер ТТН", max_length=50, unique=True)
    date = models.DateField("Дата накладной")
    created_at = models.DateTimeField("Создано", auto_now_add=True)

    def __str__(self):
        return f"{self.number} от {self.date}"


class ExcelFile(models.Model):
    file = models.FileField("Файл Excel", upload_to='uploads/')
    uploaded_at = models.DateTimeField("Загружен", auto_now_add=True)
    processed = models.BooleanField("Обработан", default=False)
    invoice = models.ForeignKey(
        Invoice,
        verbose_name="Накладная",
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='files'
    )

    def __str__(self):
        return self.file.name


class Product(models.Model):
    invoice = models.ForeignKey(
        Invoice,
        verbose_name="Накладная",
        on_delete=models.CASCADE,
        related_name='products'
    )
    excel_file = models.ForeignKey(
        ExcelFile,
        verbose_name="Исходный файл",
        on_delete=models.CASCADE,
        related_name='products'
    )
    name = models.TextField("Наименование товара")
    quantity = models.FloatField("Количество")
    price = models.FloatField("Цена, руб.коп.")
    total = models.FloatField("Сумма", blank=True, null=True)

    created_at = models.DateTimeField("Создано", auto_now_add=True)

    def save(self, *args, **kwargs):
        # Автоматически рассчитываем сумму при сохранении
        self.total = round(self.quantity * self.price, 2)
        super().save(*args, **kwargs)

    def __str__(self):
        return f"{self.name[:50]} ({self.quantity} шт. x {self.price} руб.)"