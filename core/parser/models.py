from django.db import models


class TTN(models.Model):
    TTN_STATUS_CHOICES = [
        ('in_progress', 'В обработке'),
        ('completed', 'Завершена'),
        ('canceled', 'Отменена')
    ]

    number = models.CharField("Номер ТТН", max_length=50, unique=True)
    date = models.DateField("Дата ТТН")
    status = models.CharField(
        "Статус",
        max_length=20,
        choices=TTN_STATUS_CHOICES,
        default='in_progress'
    )
    total_products = models.PositiveIntegerField("Всего товаров", default=0)
    processed_files = models.PositiveIntegerField("Обработано файлов", default=0)
    created_at = models.DateTimeField("Создано", auto_now_add=True)
    updated_at = models.DateTimeField("Обновлено", auto_now=True)

    class Meta:
        verbose_name = "ТТН"
        verbose_name_plural = "ТТН"
        ordering = ['-date', 'number']

    def __str__(self):
        return f"ТТН №{self.number} от {self.date.strftime('%d.%m.%Y')}"


class Invoice(models.Model):
    number = models.CharField("Номер накладной", max_length=50, unique=True)
    date = models.DateField("Дата накладной")
    ttn = models.ForeignKey(
        TTN,
        verbose_name="ТТН",
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='invoices'
    )
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
    ttn = models.ForeignKey(
        TTN,
        verbose_name="ТТН",
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='excel_files'
    )
    page_number = models.PositiveIntegerField("Номер страницы", null=True, blank=True)

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
    ttn = models.ForeignKey(
        TTN,
        verbose_name="ТТН",
        on_delete=models.CASCADE,
        related_name='products',
        null=True,
        blank=True
    )
    name = models.TextField("Наименование товара")
    quantity = models.FloatField("Количество")
    price = models.FloatField("Цена, руб.коп.")
    total = models.FloatField("Сумма", blank=True, null=True)
    created_at = models.DateTimeField("Создано", auto_now_add=True)

    def save(self, *args, **kwargs):
        self.total = round(self.quantity * self.price, 2)
        super().save(*args, **kwargs)

    def __str__(self):
        return f"{self.name[:50]} ({self.quantity} шт. x {self.price} руб.)"