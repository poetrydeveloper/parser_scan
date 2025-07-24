from django.contrib import admin
from django.utils.html import format_html
from django.urls import reverse
from .models import Invoice, ExcelFile, Product, TTN, Price


class ProductInline(admin.TabularInline):
    model = Product
    extra = 0
    readonly_fields = ('name', 'quantity', 'price', 'total', 'invoice_link', 'excel_file_link')
    fields = ('name', 'quantity', 'price', 'total', 'invoice_link', 'excel_file_link')
    can_delete = False
    show_change_link = True

    def invoice_link(self, obj):
        url = reverse('admin:parser_invoice_change', args=[obj.invoice.id])
        return format_html('<a href="{}">{}</a>', url, obj.invoice)
    invoice_link.short_description = 'Накладная'

    def excel_file_link(self, obj):
        url = reverse('admin:parser_excelfile_change', args=[obj.excel_file.id])
        return format_html('<a href="{}">{}</a>', url, obj.excel_file.file.name.split('/')[-1])
    excel_file_link.short_description = 'Файл'




@admin.register(TTN)
class TTNAdmin(admin.ModelAdmin):
    list_display = ('number', 'date', 'status', 'products_count', 'files_count', 'updated_at')
    search_fields = ('number',)
    list_filter = ('status', 'date')
    readonly_fields = ('created_at', 'updated_at', 'products_link', 'files_link')
    actions = ['mark_as_completed']
    inlines = [ProductInline]

    def products_count(self, obj):
        return obj.products.count()
    products_count.short_description = 'Товаров'

    def files_count(self, obj):
        return obj.excel_files.count()
    files_count.short_description = 'Файлов'

    def products_link(self, obj):
        count = obj.products.count()
        url = reverse('admin:parser_product_changelist') + f'?ttn__id__exact={obj.id}'
        return format_html('<a href="{}">{} товаров</a>', url, count)
    products_link.short_description = 'Список товаров'

    def files_link(self, obj):
        count = obj.excel_files.count()
        url = reverse('admin:parser_excelfile_changelist') + f'?ttn__id__exact={obj.id}'
        return format_html('<a href="{}">{} файлов</a>', url, count)
    files_link.short_description = 'Список файлов'

    def mark_as_completed(self, request, queryset):
        queryset.update(status='completed')
    mark_as_completed.short_description = "Пометить как завершенные"

@admin.register(Invoice)
class InvoiceAdmin(admin.ModelAdmin):
    list_display = ('number', 'date', 'ttn_link', 'products_count', 'total_sum')
    search_fields = ('number', 'ttn__number')
    list_filter = ('date', 'ttn')
    readonly_fields = ('created_at',)

    def ttn_link(self, obj):
        if obj.ttn:
            url = reverse('admin:parser_ttn_change', args=[obj.ttn.id])
            return format_html('<a href="{}">{}</a>', url, obj.ttn)
        return "-"
    ttn_link.short_description = 'ТТН'

    def products_count(self, obj):
        return obj.products.count()
    products_count.short_description = 'Товаров'

    def total_sum(self, obj):
        total = sum(p.total for p in obj.products.all() if p.total)
        return f"{total:.2f} руб." if total else "0.00 руб."
    total_sum.short_description = 'Общая сумма'

@admin.register(ExcelFile)
class ExcelFileAdmin(admin.ModelAdmin):
    list_display = ('file_name', 'uploaded_at', 'processed', 'ttn_link', 'products_link')
    list_filter = ('processed', 'uploaded_at', 'ttn')
    readonly_fields = ('uploaded_at', 'products_link', 'ttn_link')
    actions = ['mark_as_processed']

    def file_name(self, obj):
        return obj.file.name.split('/')[-1]
    file_name.short_description = 'Файл'

    def ttn_link(self, obj):
        if obj.ttn:
            url = reverse('admin:parser_ttn_change', args=[obj.ttn.id])
            return format_html('<a href="{}">{}</a>', url, obj.ttn)
        return "-"
    ttn_link.short_description = 'ТТН'

    def products_link(self, obj):
        count = obj.products.count()
        url = reverse('admin:parser_product_changelist') + f'?excel_file__id__exact={obj.id}'
        return format_html('<a href="{}">{} товаров</a>', url, count)
    products_link.short_description = 'Товары'

    def mark_as_processed(self, request, queryset):
        queryset.update(processed=True)
    mark_as_processed.short_description = "Пометить как обработанные"

@admin.register(Product)
class ProductAdmin(admin.ModelAdmin):
    list_display = ('name', 'quantity', 'price', 'total', 'invoice_link', 'ttn_link')
    search_fields = ('name', 'invoice__number', 'ttn__number')
    list_filter = ('invoice', 'ttn')
    readonly_fields = ('total', 'created_at', 'invoice_link', 'ttn_link')
    list_per_page = 50

    def invoice_link(self, obj):
        url = reverse('admin:parser_invoice_change', args=[obj.invoice.id])
        return format_html('<a href="{}">{}</a>', url, obj.invoice)
    invoice_link.short_description = 'Накладная'

    def ttn_link(self, obj):
        if obj.ttn:
            url = reverse('admin:parser_ttn_change', args=[obj.ttn.id])
            return format_html('<a href="{}">{}</a>', url, obj.ttn)
        return "-"
    ttn_link.short_description = 'ТТН'


@admin.register(Price)
class PriceAdmin(admin.ModelAdmin):
    list_display = (
        'code',
        'type',
        'article',
        'short_name',
        'price1',
        'price2',
        'formatted_stock',
        'quantity',
        'price_clear',
        'created_at'
    )
    list_display_links = ('code', 'short_name')
    search_fields = ('code', 'article', 'name', 'type')
    list_filter = ('type', 'created_at')
    list_per_page = 50
    ordering = ('code',)
    readonly_fields = ('created_at', 'updated_at')

    fieldsets = (
        ('Основная информация', {
            'fields': ('code', 'type', 'article', 'name')
        }),
        ('Цены и остатки', {
            'fields': ('price1', 'price2', 'price_clear', 'stock', 'quantity')
        }),
        ('Системные данные', {
            'fields': ('created_at', 'updated_at'),
            'classes': ('collapse',)
        })
    )

    def short_name(self, obj):
        return obj.name[:60] + '...' if len(obj.name) > 60 else obj.name

    short_name.short_description = 'Наименование'

    def formatted_stock(self, obj):
        return obj.stock if obj.stock else "-"

    formatted_stock.short_description = 'Остаток'

    def price1(self, obj):
        return f"{obj.price1:.2f} ₽"

    price1.short_description = 'Цена 1'

    def price2(self, obj):
        return f"{obj.price2:.2f} ₽" if obj.price2 else "-"

    price2.short_description = 'Цена 2'

    def price_clear(self, obj):
        return f"{obj.price_clear:.2f} ₽"

    price_clear.short_description = 'Цена за ед.'

    def quantity(self, obj):
        return obj.quantity if obj.quantity else "-"

    quantity.short_description = 'Кол-во'