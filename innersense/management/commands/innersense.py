################################################################################
#   Imports                                                                    #
################################################################################

import sys
import traceback

import re
import xlrd
import datetime

from django.core.management.base import BaseCommand, CommandError
from django.core.exceptions import ObjectDoesNotExist

from innersense.models import Customer, Package, Product, Orders, Invoice

################################################################################
#   Global Variables                                                           #
################################################################################

size_chart = {'32B': 'S',
              '32C': 'M',
              '34B': 'M',
              '34C': 'L',
              '34D': 'L',
              '36B': 'L',
              '36C': 'XL',
              '36D': 'XL',
              '38B': 'XL',
              '38C': 'XXL',
              '38D': 'XXL',
              '40B': 'XXL',
              '40C': 'XXXL'}

################################################################################
#   ETL Class                                                                  #
################################################################################

class Command(BaseCommand):
    '''
    '''

################################################################################
#   Public Functions                                                           #
################################################################################

    def add_arguments(self, parser):
        '''
        '''
        parser.add_argument('file')

    def handle(self, *args, **options):
        '''
        '''
        file = options['file']

        dict_list = self._extract(file)

        dict_list = self._transform(dict_list)

        self._load(dict_list)

################################################################################
#   Extract Functions                                                          #
################################################################################

    def _extract(self, file):
        '''
        '''
        book = xlrd.open_workbook(file)
        sheet = book.sheet_by_index(0)

        keys = [sheet.cell(0, col_index).value for col_index in range(sheet.ncols)]

        dict_list = []
        for row_index in range(1, sheet.nrows):
            d = {keys[col_index]: sheet.cell(row_index, col_index).value
                 for col_index in range(sheet.ncols)}
            dict_list.append(d)

        return dict_list

################################################################################
#   Transform Functions                                                        #
################################################################################

    def _validate_sku(self, sku):
        '''
        '''
        type = sku[:-3]
        code = sku[-3:]

        if type == 'ISP':
            if code == '003' or code == '004':
                return 'IMP' + code

        if type == 'IMP':
            if code != '003' or code != '004':
                return 'ISP' + code

        return sku

    def _format_package_one(self, sku):
        '''
        ''' 
        sku_list = sku.split('-')
        sku = sku_list[0].split('_')

        combo_one = (self._validate_sku(sku[0][:-1]), sku[0][-1], sku_list[1].replace(' ', ''))
        combo_two = (self._validate_sku(sku[1][:-1]), sku[1][-1], sku_list[1].replace(' ', ''))

        return combo_one, combo_two

    def _format_package_two(self, sku):
        '''
        '''
        sku_list = sku.split('-')
        sku = sku_list[0].split('_')

        combo_one = (self._validate_sku(sku[0][:-1]), sku[0][-1], sku_list[1].replace(' ', ''))
        combo_two = (self._validate_sku(sku[0][:-2] + sku[1][:-1]), sku[1][-1], sku_list[1].replace(' ', ''))

        return combo_one, combo_two

    def _format_package_three(self, sku):
        '''
        '''
        sku_list = sku.split('-')
        sku = sku_list[0].split('_')

        combo_one = (self._validate_sku(sku[0][:-1]), sku[0][-1], sku_list[1].replace(' ', ''))
        combo_two = (self._validate_sku(sku[0][:-3] + sku[1]), sku[0][-1], sku_list[1].replace(' ', ''))
        combo_three = (self._validate_sku(sku[0][:-3] + sku[2]), sku[0][-1], sku_list[1].replace(' ', ''))

        return combo_one, combo_two, combo_three

    def _format_package_four(self, sku):
        '''
        '''
        sku_list = sku.split('-')
        sku = sku_list[0].split('_')

        combo_one = (self._validate_sku(sku[0]), sku[2][-1], sku_list[1].replace(' ', ''))
        combo_two = (self._validate_sku(sku[0][:-3] + '0' + sku[1]), sku[2][-1], sku_list[1].replace(' ', ''))
        combo_three = (self._validate_sku(sku[0][:-3] + '00' + sku[2][0]), sku[2][-1], sku_list[1].replace(' ', ''))

        return combo_one, combo_two, combo_three

    def _format_product_color(self, sku):
        '''
        '''
        sku_list = sku.split('-')
        sku = sku_list[0]
        if sku.startswith('ISBP'):
            bra_size = sku_list[1].replace(' ', '')
            panty_size = size_chart.get(bra_size)
            return ('ISB' + sku[-4:-1], sku[-1], bra_size), ('ISP' + sku[-4:-1], sku[-1], panty_size)
        else:
            return self._validate_sku(sku[:-1]), sku[-1], sku_list[1].replace(' ', '')

    def _format_product(self, sku):
        '''
        '''
        sku_list = sku.split('-')

        return self._validate_sku(sku_list[0]), None, sku_list[1].replace(' ', '')

    def _transform(self, dict_list):
        '''
        '''
        for obj in dict_list:
            # Order
            #
            obj['suborder_num'] = obj['suborder_num'].replace('`', '')
            obj['order_date'] = datetime.datetime.strptime(obj['order_date'], '`%Y-%m-%d')

            # Invoice
            #
            obj['invoice_date'] = datetime.datetime.strptime(obj['invoice_date'], '`%Y-%m-%d')

            try:
                obj['mrp'] = float(obj['mrp'])
            except ValueError:
                obj['mrp'] = float(0.0)

            try:
                obj['selling_price'] = float(obj['selling_price'])
            except ValueError:
                obj['selling_price'] = float(0.0)

            try:        
                obj['tax_amount'] = float(obj['tax_amount'])
            except ValueError:
                obj['tax_amount'] = float(0.0)

            # SKUs
            #
            sku = obj['sku']

            # Combos
            #
            SKU_PATTERN = re.compile('\D\D\D\d\d\d\D_\D\D\D\d\d\d\D-\d\d\D')
            match = SKU_PATTERN.search(sku)
            if match:
                obj[sku] = self._format_package_one(sku)
                continue

            SKU_PATTERN = re.compile('\D\D\D\d\d\d\D_\d\D-\d\d\D')
            match = SKU_PATTERN.search(sku)
            if match:
                obj[sku] = self._format_package_two(sku)
                continue

            SKU_PATTERN = re.compile('\D\D\D\d\d\d\D_\d\d_\d\d-\D')
            match = SKU_PATTERN.search(sku)
            if match:
                obj[sku] = self._format_package_three(sku)
                continue

            SKU_PATTERN = re.compile('\D\D\D\d\d\d\_\d\d_\d\D-\D')
            match = SKU_PATTERN.search(sku)
            if match:
                obj[sku] = self._format_package_four(sku)
                continue

            # Bras
            #
            SKU_PATTERN = re.compile('\D\D\D\d\d\d\D-\d\d\D')
            match = SKU_PATTERN.search(sku)
            if match:
                obj[sku] = self._format_product_color(sku)
                continue

            SKU_PATTERN = re.compile('\D\D\D\d\d\d-\d\d\D')
            match = SKU_PATTERN.search(sku)
            if match:
                obj[sku] = self._format_product(sku)
                continue

            # Panties
            #
            SKU_PATTERN = re.compile('\D\D\D\d\d\d\D-\D')
            match = SKU_PATTERN.search(sku)
            if match:
                obj[sku] = self._format_product_color(sku)
                continue

            SKU_PATTERN = re.compile('\D\D\D\d\d\d-\D')
            match = SKU_PATTERN.search(sku)
            if match:
                obj[sku] = self._format_product(sku)
                continue

            obj[sku] = None

        return dict_list

################################################################################
#   Load Functions                                                             #
################################################################################

    def _upsert_package(self, sku):
        '''
        '''
        try:
            # Update
            #
            package = Package.objects.get(sku=sku)
            return package
        except ObjectDoesNotExist:
            # Create
            #
            package = Package(sku=sku)
            package.save()
            return package

    def _upsert_product(self, product):
        '''
        '''
        try:
            # Update
            #
            product = Product.objects.get(sku=product[0], color=product[1], size=product[2])
            return product
        except ObjectDoesNotExist:
            # Create
            #
            product = Product(sku=product[0], color=product[1], size=product[2])
            product.save()
            return product

    def _upsert_sku(self, sku):
        '''
        '''
        if sku:
            if type(sku[0]) == tuple:
                for value in sku:
                    return self._upsert_sku(value)
            else:
                yield self._upsert_product(sku)

    def _upsert_customer(self, obj):
        '''
        '''
        try:
            # Update
            #
            customer = Customer.objects.get(mobile=obj['mobile_no'])
            return customer
        except ObjectDoesNotExist:
            # Create
            #
            customer = Customer(name=obj['customer_name'], mobile=obj['mobile_no'], address=obj['address_line_1'], city=obj['city'], state=obj['state'], pincode=obj['pin_code'])
            customer.save()
            return customer

    def _upsert_order(self, obj, package, customer):
        '''
        '''
        try:
            # Update
            #
            order = Orders.objects.get(sub_order_id=obj['suborder_num'])
            return order
        except ObjectDoesNotExist:
            # Create
            #
            order = Orders(order_id=obj['reference_code'], sub_order_id=obj['suborder_num'], package=package, customer=customer, quantity=obj['quantity'], order_date=obj['order_date'])
            order.save()
            return order

    def _upsert_invoice(self, obj, order):
        '''
        '''
        try:
            # Update
            #
            invoice = Invoice.objects.get(order=order)
            return invoice
        except ObjectDoesNotExist:
            # Create
            #
            invoice = Invoice(invoice_id=obj['reference_invoice_num'], order=order, mrp=obj['mrp'], selling_price=obj['selling_price'], tax_amount=obj['tax_amount'], invoice_date=obj['invoice_date'])
            invoice.save()
            return invoice

    def _load(self, dict_list):
        '''
        '''
        for obj in dict_list:
            try:
                # Package
                #
                sku = obj['sku']
                package = self._upsert_package(sku)

                # Product
                #
                for product in self._upsert_sku(obj[sku]):
                    package.products.add(product)

                # Customer
                #
                customer = self._upsert_customer(obj)

                # Order
                #
                order = self._upsert_order(obj, package, customer)

                # Invoice
                #
                invoice = self._upsert_invoice(obj, order)
            except:
                traceback.print_exc(file=sys.stdout)
                continue
