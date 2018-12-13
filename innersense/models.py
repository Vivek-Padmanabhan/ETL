################################################################################
#   Imports                                                                    #
################################################################################

from django.db import models

################################################################################
#   Models                                                                     #
################################################################################

class Customer(models.Model):
    '''
        Schema to store Customer Data.
        Each Customer is uniquely identified by mobile field.
    '''
    id = models.AutoField(primary_key=True)
    name = models.CharField(max_length=500)
    mobile = models.BigIntegerField()
    address = models.CharField(max_length=100)
    city = models.CharField(max_length=100)
    state = models.CharField(max_length=100)
    pincode = models.BigIntegerField()

class Product(models.Model):
    '''
        Schema to store Product Data.
        Each Product is uniquely identified by a combination of sku, size and color.
    '''
    id = models.AutoField(primary_key=True)
    sku = models.CharField(max_length=100)
    size = models.CharField(max_length=100)
    color = models.CharField(max_length=100, null=True)

class Package(models.Model):
    '''
        Schema to store Package Data.
        Each Package can contain one or more Products.
    '''
    id = models.AutoField(primary_key=True)
    sku = models.CharField(max_length=100)
    products = models.ManyToManyField(Product)

class Orders(models.Model):
    '''
        Schema to store Order Data.
        Each Order corresponds to a single Package.
    '''
    id = models.AutoField(primary_key=True)
    order_id = models.CharField(max_length=100)
    sub_order_id = models.CharField(max_length=100)
    package = models.ForeignKey(Package, blank=False, null=True, on_delete=models.SET_NULL)
    customer = models.ForeignKey(Customer, blank=False, null=True, on_delete=models.SET_NULL)
    quantity = models.IntegerField()
    order_date = models.DateField()

class Invoice(models.Model):
    '''
        Schema to store Invoice Data.
        Each Invoice corresponds to a single Order.
    '''
    id = models.AutoField(primary_key=True)
    invoice_id = models.CharField(max_length=100)
    order = models.ForeignKey(Orders, blank=False, null=True, on_delete=models.SET_NULL)
    mrp = models.DecimalField(max_digits=11, decimal_places=2)
    selling_price = models.DecimalField(max_digits=11, decimal_places=2)
    tax_amount = models.DecimalField(max_digits=11, decimal_places=2)
    invoice_date = models.DateField()
