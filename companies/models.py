from django.db import models

class Company(models.Model):
    name = models.CharField(max_length=255, unique=True, verbose_name="Название")
    extra_data = models.JSONField(default=dict, blank=True, verbose_name="Дополнительные реквизиты")

    def __str__(self):
        return self.name

    class Meta:
        verbose_name = "Компания"
        verbose_name_plural = "Компании"


class Employee(models.Model):
    company = models.ForeignKey(Company, on_delete=models.CASCADE, related_name='employees', verbose_name="Компания")
    last_name = models.CharField(max_length=100, blank=True, verbose_name="Фамилия")
    first_name = models.CharField(max_length=100, blank=True, verbose_name="Имя")
    middle_name = models.CharField(max_length=100, blank=True, verbose_name="Отчество")
    birth_date = models.CharField(max_length=20, blank=True, verbose_name="Дата рождения")
    hire_date = models.CharField(max_length=20, blank=True, verbose_name="Дата трудоустройства")
    phone = models.CharField(max_length=50, blank=True, verbose_name="Телефон")
    position = models.CharField(max_length=255, blank=True, verbose_name="Должность")
    salary = models.CharField(max_length=50, blank=True, verbose_name="Оклад")

    def __str__(self):
        return f"{self.last_name} {self.first_name}"

    class Meta:
        verbose_name = "Сотрудник"
        verbose_name_plural = "Сотрудники"