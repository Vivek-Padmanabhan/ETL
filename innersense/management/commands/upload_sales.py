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

panty_size = ('S', 'M', 'L', 'XL', 'XXL', 'XXXL')

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

color_chart = {'IMB001': {'A': 'Baby Pink',
                          'B': 'Navy Blue'},
               'ISB002': {None: 'Baby Pink'},
               'IMB003': {'A': 'Royal Blue',
                          'B': 'Peacock Green',
                          'C': 'Black',
                          'D': 'Skin'},
               'IMB004': {'A': 'Skin',
                          'B': 'Lavender',
                          'C': 'Pink Floral Print',
                          'D': 'Purple Floral Print',
                          'E': 'Black'},
               'IMB005': {'A': 'Skin',
                          'B': 'Baby Pink',
                          'D': 'Lavender',
                          'E': 'Navy Blue'},
               'IMB006': {'A': 'Royal Blue',
                          'B': 'Peacock Green',
                          'D': 'Black',
                          'E': 'Skin'},
               'IMB007': {'A': 'Pink Floral Print',
                          'B': 'Purple Floral Print',
                          'C': 'Aqua Blue',
                          'D': 'Lace Print'},
               'IMB008': {'A': 'Aqua Print'},
               'IMB009': {'B': 'Skin',
                          'C': 'Black'},
               'ISB012': {'B': 'Lavender Color'},
               'ISB017': {None: 'Maroon',
                          'A': 'Royal Blue'},
               'ISB018': {'A': 'Black',
                          'B': 'Pink Flowy Line',
                          'C': 'Green Matrix',
                          'D': 'Blue Matrix'},
               'ISB019': {'A': 'Pink Flowy Lines'},
               'ISB020': {None: 'Ocean Green'},
               'ISB021': {None: 'Maroon Print'},
               'ISB026': {None: 'Green Grunge Print',
                          'B': 'Green Grunge Print'},
               'ISB036': {'A': 'Pink Flowy Lines'},
               'ISB037': {'A': 'Green Flowy Lines'},
               'ISB040': {None: 'Pink Floral Print'},
               'ISB041': {None: 'Purple Floral Print'},
               'ISB042': {None: 'Lavender'},
               'ISB046': {None: 'Green Flowy Line'},
               'ISB047': {None: 'Black Lace'},
               'ISB051': {None: 'Black Lace'},
               'ISB052': {None: 'Purple Floral Print'},
               'ISB053': {None: 'Blue Grunge'},
               'ISB054': {None: 'Green'},
               'ISB055': {None: 'Black'},
               'ISB056': {None: 'Jungle Print'},
               'ISB057': {None: 'Fuschia'},
               'ISB058': {None: 'Purple'},
               'ISB060': {None: 'Green'},
               'ISB061': {None: 'Black'},
               'ISB062': {None: 'Skin'},
               'ISB065': {None: 'Baby Pink'},
               'ISB066': {None: 'Green Flowy Line'},
               'ISB067': {None: 'Baby Pink'},
               'ISB068': {None: 'Skin'},
               'ISB069': {None: 'Black'},
               'ISB071': {None: 'Floral Fresh Print'},
               'ISB073': {None: 'Boho Aztec Print'},
               'ISB074': {None: 'Floral Fiesta Print'},
               'ISB075': {None: 'Blue Water Color'},
               'ISB076': {None: 'Brown Feather'},
               'ISB080': {None: 'Pink Spring Print'},
               'ISB081': {None: 'Jungle Print'},
               'ISB082': {None: 'Boho Blue Print'},
               'ISB083': {None: 'Red Print'},
               'ISB084': {None: 'Green'},
               'ISB085': {None: 'Navy Blue'},
               'ISB086': {None: 'Red Print'},
               'ISB087': {None: 'Lace Print'},
               'ISB088': {None: 'Brown Feather Print'},
               'ISB089': {None: 'Floral Peach Print'},
               'ISB090': {None: 'Boho Blue Print'},
               'ISB091': {None: 'Boho Blue Print'},
               'ISB092': {None: 'Purple'},
               'ISB093': {None: 'Navy Blue'},
               'ISB094': {None: 'Aqua Green Floral'},
               'ISB095': {None: 'Skin'},
               'ISB096': {None: 'Beige Print'},
               'ISB097': {None: 'Black'},
               'ISB098': {None: 'Royal Blue'},
               'ISB100': {None: 'Black'},
               'ISB102': {None: 'Popping Skin'},
               'IMP003': {'A': 'Royal Blue',
                          'B': 'Peacock Green'},
               'IMP004': {'A': 'Skin',
                          'B': 'Lavender'},
               'ISP012': {'A': 'Peacock Green',
                          'B': 'Lavender'},
               'ISP015': {None: 'Coral'},
               'ISP016': {None: 'Pink Matrix'},
               'ISP017': {'A': 'Royal Blue'},
               'ISP018': {'C': 'Green Matrix'},
               'ISP019': {'A': 'Pink Flowy Lines',
                          'B': 'Green  Flowy Lines'},
               'ISP020': {None: 'Coral Print'},
               'ISP022': {None: 'Aqua Matrix Print'},
               'ISP024': {None: 'Coral Print'},
               'ISP025': {None: 'Black'},
               'ISP026': {'A': 'Blue Grunge Print',
                          'B': 'Green Grunge Print'},
               'ISP027': {None: 'Black'},
               'ISP028': {None: 'Green Matrix Print'},
               'ISP029': {None: 'Blue Matrix Print'},
               'ISP030': {None: 'Pink Floral Print'},
               'ISP031': {None: 'Purple Floral Print'},
               'ISP032': {None: 'Purple Floral Print'},
               'ISP033': {None: 'Pink Floral Print'},
               'ISP034': {None: 'Aqua',
                          'A': 'Blue Matrix'},
               'ISP035': {None: 'Green Matrix'},
               'ISP036': {None: 'Pink Flowy Lines'},
               'ISP037': {None: 'Green Flowy Lines'},
               'ISP055': {None: 'Black'},
               'ISP056': {None: 'Jungle Print'},
               'ISP065': {None: 'Baby Pink'}}

################################################################################
#   ETL Class                                                                  #
################################################################################

class Command(BaseCommand):
    '''
        Excel ETL Script
    '''

################################################################################
#   Public Functions                                                           #
################################################################################

    def add_arguments(self, parser):
        '''
            Take the excel file as an argument.
        '''
        parser.add_argument('file')
        parser.add_argument('--dry-run')

    def handle(self, *args, **options):
        '''
            Extract
            Transform
            Load
        '''
        file = options['file']

        dict_list = self._extract(file)

        dict_list = self._transform(dict_list)

        '''
        sku_list = []
        for obj in dict_list:
            sku_list.append(obj['sku'])

        for sku in set(sku_list):
            for obj in dict_list:
                if obj['sku'] == sku:
                    self.stdout.write(sku + ' :: ' + str(obj[sku]))
                    break
        '''
        self._load(dict_list)

################################################################################
#   Extract Functions                                                          #
################################################################################

    def _split_list(self, dict_list):
        '''
        '''
        order_list = []

        for obj in dict_list:
            if obj['description'] == 'Shipped':
                order_list.append(obj)

        return order_list

    def _extract(self, file):
        '''
            Read Excel file, convert to a list of dicts.
        '''
        book = xlrd.open_workbook(file)
        sheet = book.sheet_by_index(0)

        keys = [sheet.cell(0, col_index).value for col_index in range(sheet.ncols)]

        dict_list = []
        for row_index in range(1, sheet.nrows):
            d = {keys[col_index]: sheet.cell(row_index, col_index).value
                 for col_index in range(sheet.ncols)}
            dict_list.append(d)

        return self._split_list(dict_list)

################################################################################
#   Transform Functions                                                        #
################################################################################

    def _validate_sku(self, sku):
        '''
            Split ISP from IMP.
            Both are mutually exclusive.
        '''
        type = sku[:-3]
        code = sku[-3:]

        if type == 'ISP':
            if code in ('003', '004'):
                return 'IMP' + code

        if type == 'IMP':
            if code not in ('003', '004'):
                return 'ISP' + code

        return sku

    def _format_package_zero(self, sku):
        '''
            Transformation for format :: \D\D\D\d\d\d\D_\D\D\D\d\d\d--\D+
        '''
        sku_list = sku.split('--')
        sku = sku_list[0].split('_')

        combo_one = (self._validate_sku(sku[0][:-1]), sku[0][-1], sku_list[1].replace(' ', ''))
        combo_two = (self._validate_sku(sku[1]), None, sku_list[1].replace(' ', ''))

        return combo_one, combo_two

    def _format_package_one(self, sku):
        '''
            Transformation for format :: \D\D\D\d\d\d_\D\D\D\d\d\d-\D
        '''
        sku_list = sku.split('-')
        sku = sku_list[0].split('_')

        combo_one = (self._validate_sku(sku[0]), None, sku_list[1].replace(' ', ''))
        combo_two = (self._validate_sku(sku[1]), None, sku_list[1].replace(' ', ''))

        return combo_one, combo_two

    def _format_package_two(self, sku):
        '''
            Transformation for format :: \D\D\D\d\d\d\D_\D\D\D\d\d\d\D-\d\d\D & \D\D\D\d\d\d\D_\D\D\D\d\d\d\D-\D
        ''' 
        sku_list = sku.split('-')
        sku = sku_list[0].split('_')

        combo_one = (self._validate_sku(sku[0][:-1]), sku[0][-1], sku_list[1].replace(' ', ''))
        combo_two = (self._validate_sku(sku[1][:-1]), sku[1][-1], sku_list[1].replace(' ', ''))

        return combo_one, combo_two

    def _format_package_three(self, sku):
        '''
            Transformation for format :: \D\D\D\d\d\d\D_\d\D-\d\d\D
        '''
        sku_list = sku.split('-')
        sku = sku_list[0].split('_')

        combo_one = (self._validate_sku(sku[0][:-1]), sku[0][-1], sku_list[1].replace(' ', ''))
        combo_two = (self._validate_sku(sku[0][:-2] + sku[1][:-1]), sku[1][-1], sku_list[1].replace(' ', ''))

        return combo_one, combo_two

    def _format_package_four(self, sku):
        ''' 
            Transformation for format :: \D\D\D\d\d\d_\d\d-\d\d\D
        '''
        sku_list = sku.split('-')
        sku = sku_list[0].split('_')

        combo_one = (self._validate_sku(sku[0]), None, sku_list[1].replace(' ', ''))
        combo_two = (self._validate_sku(sku[0][:-2] + sku[1]), None, sku_list[1].replace(' ', ''))

        return combo_one, combo_two

    def _format_package_five(self, sku):
        '''
            Transformation for format :: \D\D\D\d\d\d_\d\d-\D+
        '''
        sku_list = sku.split('-')
        sku = sku_list[0].split('_')

        combo_one = (self._validate_sku(sku[0]), None, sku_list[1].replace(' ', ''))
        combo_two = (self._validate_sku(sku[0][:-2] + sku[1]), None, sku_list[1].replace(' ', ''))

        return combo_one, combo_two

    def _format_package_six(self, sku):
        '''
            Transformation for format :: \D\D\D\d\d\d\D_\d\D-\D+
        '''
        sku_list = sku.split('-')
        sku = sku_list[0].split('_')

        combo_one = (self._validate_sku(sku[0][:-1]), sku[0][-1], sku_list[1].replace(' ', ''))
        combo_two = (self._validate_sku(sku[0][:-2] + sku[1][:-1]), sku[1][-1], sku_list[1].replace(' ', ''))

        return combo_one, combo_two

    def _format_package_seven(self, sku):
        '''
            Transformation for format :: \D\D\D\d\d\d\D_\d\d_\d\d-\D
        '''
        sku_list = sku.split('-')
        sku = sku_list[0].split('_')

        combo_one = (self._validate_sku(sku[0][:-1]), sku[0][-1], sku_list[1].replace(' ', ''))
        combo_two = (self._validate_sku(sku[0][:-3] + sku[1]), None, sku_list[1].replace(' ', ''))
        combo_three = (self._validate_sku(sku[0][:-3] + sku[2]), None, sku_list[1].replace(' ', ''))

        return combo_one, combo_two, combo_three

    def _format_package_eight(self, sku):
        '''
            Transformation for format :: \D\D\D\d\d\d\_\d\d_\d\D-\D
        '''
        sku_list = sku.split('-')
        sku = sku_list[0].split('_')

        combo_one = (self._validate_sku(sku[0]), None, sku_list[1].replace(' ', ''))
        combo_two = (self._validate_sku(sku[0][:-2] + sku[1]), None, sku_list[1].replace(' ', ''))
        combo_three = (self._validate_sku(sku[0][:-2] + '0' + sku[2][0]), sku[2][-1], sku_list[1].replace(' ', ''))

        return combo_one, combo_two, combo_three

    def _format_product_new(self, sku):
        '''
            Transformation for format :: \D\D\D\d\d\d\D-\D+-\d\d\D
        '''
        sku_list = sku.split('-')

        return self._validate_sku(sku_list[0]), sku_list[1], sku_list[2]

    def _format_product_color(self, sku):
        '''
            Transformation for format :: \D\D\D\d\d\d\D-\d\d\D
        '''
        sku_list = sku.split('-')
        sku = sku_list[0]
        if sku.startswith('ISBP'):
            bra_size = sku_list[1].replace(' ', '')
            panty_size = size_chart.get(bra_size)
            return (self._validate_sku('ISB' + sku[-4:-1]), sku[-1], bra_size), (self._validate_sku('ISP' + sku[-4:-1]), sku[-1], panty_size)
        else:
            return self._validate_sku(sku[:-1]), sku[-1], sku_list[1].replace(' ', '')

    def _format_product(self, sku):
        '''
            Transformation for format :: \D\D\D\d\d\d-\d\d\D
        '''
        sku_list = sku.split('-')
        sku = sku_list[0]
        if sku.startswith('ISBP'):
            bra_size = sku_list[1].replace(' ', '')
            panty_size = size_chart.get(bra_size)
            return (self._validate_sku('ISB' + sku[-3:]), None, bra_size), (self._validate_sku('ISP' + sku[-3:]), None, panty_size)
        else:
            return self._validate_sku(sku_list[0]), None, sku_list[1].replace(' ', '')

    def _replace_color(self, sku):
        '''
            Replace Code with Color
        '''
        sku_id = sku[0]
        sku_code = sku[1]

        sku_color_obj = color_chart.get(sku_id)
        if sku_color_obj:
            sku_color = sku_color_obj.get(sku_code)
            if sku_color:
                return (sku_id, sku_color, sku[2])
        return sku

    def _replace(self, dict_list):
        '''
            Replace Code with Color
        '''
        for obj in dict_list:
            sku = obj[obj['sku']]
            if type(sku[0]) == tuple:
                sku_list = []
                for value in sku:
                    sku_list.append(self._replace_color(value))
                obj[obj['sku']] = tuple(sku_list)
            else:
                obj[obj['sku']] = self._replace_color(sku)

    def _transform(self, dict_list):
        '''
            Transform Combo SKUs into individual Product SKUs.
            Convert Date String to Datetime Object.
            Handle cast to float ValueErrors.
            Remove special characters.
        '''
        sku_list = []

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

            SKU_PATTERN = re.compile('\D\D\D\d\d\d\D-\D+-\d\d\D')
            match = SKU_PATTERN.search(sku)
            if match:
                obj[sku] = self._format_product_new(sku)
                sku_list.append(obj)
                continue

            SKU_PATTERN = re.compile('\D\D\D\d\d\d-\D+-\d\d\D')
            match = SKU_PATTERN.search(sku)
            if match:
                obj[sku] = self._format_product_new(sku)
                sku_list.append(obj)
                continue

            # Combos
            #
            SKU_PATTERN = re.compile('\D\D\D\d\d\d\D_\D\D\D\d\d\d--\D+')
            match = SKU_PATTERN.search(sku)
            if match:
                obj[sku] = self._format_package_zero(sku)
                sku_list.append(obj)
                continue

            SKU_PATTERN = re.compile('\D\D\D\d\d\d_\D\D\D\d\d\d-\D+')
            match = SKU_PATTERN.search(sku)
            if match:
                obj[sku] = self._format_package_one(sku)
                sku_list.append(obj)
                continue

            SKU_PATTERN = re.compile('\D\D\D\d\d\d\D_\D\D\D\d\d\d\D-\D+')
            match = SKU_PATTERN.search(sku)
            if match:
                obj[sku] = self._format_package_two(sku)
                sku_list.append(obj)
                continue

            SKU_PATTERN = re.compile('\D\D\D\d\d\d\D_\D\D\D\d\d\d\D-\d\d\D')
            match = SKU_PATTERN.search(sku)
            if match:
                obj[sku] = self._format_package_two(sku)
                sku_list.append(obj)
                continue

            SKU_PATTERN = re.compile('\D\D\D\d\d\d\D_\d\D-\d\d\D')
            match = SKU_PATTERN.search(sku)
            if match:
                obj[sku] = self._format_package_three(sku)
                sku_list.append(obj)
                continue

            SKU_PATTERN = re.compile('\D\D\D\d\d\d_\d\d-\d\d\D')
            match = SKU_PATTERN.search(sku)
            if match:
                obj[sku] = self._format_package_four(sku)
                sku_list.append(obj)
                continue

            SKU_PATTERN = re.compile('\D\D\D\d\d\d_\d\d-\D+')
            match = SKU_PATTERN.search(sku)
            if match:
                obj[sku] = self._format_package_five(sku)
                sku_list.append(obj)
                continue

            SKU_PATTERN = re.compile('\D\D\D\d\d\d\D_\d\D-\D+')
            match = SKU_PATTERN.search(sku)
            if match:
                obj[sku] = self._format_package_six(sku)
                sku_list.append(obj)
                continue

            SKU_PATTERN = re.compile('\D\D\D\d\d\d\D_\d\d_\d\d-\D+')
            match = SKU_PATTERN.search(sku)
            if match:
                obj[sku] = self._format_package_seven(sku)
                sku_list.append(obj)
                continue

            SKU_PATTERN = re.compile('\D\D\D\d\d\d_\d\d_\d\D-\D+')
            match = SKU_PATTERN.search(sku)
            if match:
                obj[sku] = self._format_package_eight(sku)
                sku_list.append(obj)
                continue

            # Bras
            #
            SKU_PATTERN = re.compile('\D\D\D\d\d\d\D-\d\d\D')
            match = SKU_PATTERN.search(sku)
            if match:
                obj[sku] = self._format_product_color(sku)
                sku_list.append(obj)
                continue

            SKU_PATTERN = re.compile('\D\D\D\d\d\d-\d\d\D')
            match = SKU_PATTERN.search(sku)
            if match:
                obj[sku] = self._format_product(sku)
                sku_list.append(obj)
                continue

            # Panties
            #
            SKU_PATTERN = re.compile('\D\D\D\d\d\d\D-\D+')
            match = SKU_PATTERN.search(sku)
            if match:
                obj[sku] = self._format_product_color(sku)
                sku_list.append(obj)
                continue

            SKU_PATTERN = re.compile('\D\D\D\d\d\d-\D+')
            match = SKU_PATTERN.search(sku)
            if match:
                obj[sku] = self._format_product(sku)
                sku_list.append(obj)
                continue

        self._replace(sku_list)

        return sku_list

################################################################################
#   Load Functions                                                             #
################################################################################

    def _upsert_package(self, sku):
        '''
            Create or Update Package Schema.
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
            Create or Update Product Schema.
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
            Deflate Combos into individual Products.
        '''
        if sku:
            if type(sku[0]) == tuple:
                for value in sku:
                    yield self._upsert_product(value)
            else:
                yield self._upsert_product(sku)

    def _upsert_customer(self, obj):
        '''
            Create or Update Customer Schema.
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
            Create or Update Order Schema.
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
            Create or Update Invoice Schema.
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
            Load Function.
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
