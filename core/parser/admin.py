# parser/admin.py
from django.contrib import admin
from django.utils.html import format_html
from django.urls import reverse
from .models import Invoice, ExcelFile, Product

@admin.register(Invoice)
class InvoiceAdmin(admin.ModelAdmin):
    list_display = ('number', 'date', 'products_count', 'total_sum')
    search_fields = ('number',)
    list_filter = ('date',)
    readonly_fields = ('created_at',)

    def products_count(self, obj):
        return obj.products.count()
    products_count.short_description = 'Товаров'

    def total_sum(self, obj):
        total = sum(p.total for p in obj.products.all() if p.total)
        return f"{total:.2f} руб." if total else "0.00 руб."
    total_sum.short_description = 'Общая сумма'

@admin.register(ExcelFile)
class ExcelFileAdmin(admin.ModelAdmin):
    list_display = ('file_name', 'uploaded_at', 'processed', 'products_link')
    list_filter = ('processed', 'uploaded_at')
    readonly_fields = ('uploaded_at', 'products_link')
    actions = ['mark_as_processed']

    def file_name(self, obj):
        return obj.file.name.split('/')[-1]
    file_name.short_description = 'Файл'

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
    list_display = ('name', 'quantity', 'price', 'total', 'invoice_link')
    search_fields = ('name', 'invoice__number')
    list_filter = ('invoice',)
    readonly_fields = ('total', 'created_at', 'invoice_link')
    list_per_page = 50

    def invoice_link(self, obj):
        url = reverse('admin:parser_invoice_change', args=[obj.invoice.id])
        return format_html('<a href="{}">{}</a>', url, obj.invoice)
    invoice_link.short_description = 'Накладная'