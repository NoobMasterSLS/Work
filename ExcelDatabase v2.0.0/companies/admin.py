from django.contrib import admin
from .models import Company, Employee

@admin.register(Company)
class CompanyAdmin(admin.ModelAdmin):
    list_display = ('id', 'name')
    search_fields = ('name',)

@admin.register(Employee)
class EmployeeAdmin(admin.ModelAdmin):
    list_display = ('id', 'company', 'last_name', 'first_name', 'position')
    list_filter = ('company',)
    search_fields = ('last_name', 'first_name')