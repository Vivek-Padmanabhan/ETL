################################################################################
#   Imports                                                                    #
################################################################################

from django.db import models

################################################################################
#   Models                                                                     #
################################################################################

class Customer(models.Model):
    ''''''
    id = models.AutoField(primary_key=True)
    name = models.CharField(max_length=500)
    mobile = models.BigIntegerField()
    address = models.CharField(max_length=100)
    city = models.CharField(max_length=100)
    state = models.CharField(max_length=100)
    pincode = models.BigIntegerField()

class Product(models.Model):
    ''''''
    id = models.AutoField(primary_key=True)
    sku = models.CharField(max_length=100)
    size = models.CharField(max_length=100)
    color = models.CharField(max_length=100, null=True)

class Package(models.Model):
    ''''''
    id = models.AutoField(primary_key=True)
    sku = models.CharField(max_length=100)
    products = models.ManyToManyField(Product)

class Orders(models.Model):
    ''''''
    id = models.AutoField(primary_key=True)
    order_id = models.CharField(max_length=100)
    sub_order_id = models.CharField(max_length=100)
    package = models.ForeignKey(Package, blank=False, null=True, on_delete=models.SET_NULL)
    customer = models.ForeignKey(Customer, blank=False, null=True, on_delete=models.SET_NULL)
    quantity = models.IntegerField()
    order_date = models.DateField()

class Invoice(models.Model):
    ''''''
    id = models.AutoField(primary_key=True)
    invoice_id = models.CharField(max_length=100)
    order = models.ForeignKey(Orders, blank=False, null=True, on_delete=models.SET_NULL)
    mrp = models.DecimalField(max_digits=11, decimal_places=2)
    selling_price = models.DecimalField(max_digits=11, decimal_places=2)
    tax_amount = models.DecimalField(max_digits=11, decimal_places=2)
    invoice_date = models.DateField()
