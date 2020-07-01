import requests
import json
import sys
import random
from bs4 import BeautifulSoup
import time
import threading
import pprint
from xlwt import Workbook

# workbook is created
wb = Workbook()

# add_sheet is used to create sheet.
sheet1 = wb.add_sheet('sheet1')

sheet1.write(0, 1, '제품번호')
sheet1.write(0, 2, '제품 SKU')
sheet1.write(0, 3, '옵션명')
sheet1.write(0, 4, '이미지(300x300)')
sheet1.write(0, 5, '이미지(900x900)')
sheet1.write(0, 6, '판매가')
sheet1.write(0, 7, '정가')
sheet1.write(0, 8, '세일가')
sheet1.write(0, 9, '수량')
sheet1.write(0, 10, '판매가')
wb.save('xlwt example.xls')


# global variables / settings
FRAGRANCE_API_ROOT = 'https://www.fragrancenet.com/fragrances'


def LimitProduct():
    x=0
    while x < 10:
        if 'https://www.fragrancenet.com/fragrances?page=x' != 'https://www.fragrancenet.com/fragrances?page=x+1':
            return LimitProduct
            return False

'''
simple function for comparing strings
strings match if they are the same, excluding case
'''

def inputCompare(x, y):
    if len(x) > 0 and len(y) > 0:
        return x.lower() == y.lower()
    return False

# determine if Node contains product detail data
def nodeHasDetailData(node):
    return node.string is not None and node.string.find('var variant_id') != -1



'''
Parses the product options details data from a script node as JSON
'''
def parseProductOptionsDetails(node):
    # convert node to string
    text = node.string

    # find starting location of details
    start = text.find('sku_map')

    # find end location of details
    end = text.find('has_reviews')
    
    # get the value between start and end, and strip out all unneeded whitespace
    sku_map = text[start + 10: end].strip()

    # if value has a trailing comma, remove it
    if sku_map[-1:] == ',':
        sku_map = sku_map[:-1]

    # return value as JSON
    return json.loads(sku_map)


'''
Maps product values 
'''
def mapProductDetails(group_id, options):
    products = []
    for key, value in options.items():
        products.append({
            'id': group_id, # the id used to group variants of the same product i.e. product options
            'sku': key,
            'name': value.get('SIZE_default'),
            '''
            'brand': value.get(),
            '''
            'img': value.get('img'),
            'zoom_img': value.get('zoom_img'),
            'list_price': value.get('price_int'),
            'retail_price': value.get('retail_price_int'),
            'sale_price': value.get('discount_price_int'),
            'quantity': value.get('quantity'),
        })
    return products



'''
Reads command line inputs and runs associated functions
'''
def commandLineQuerier():
    COMMAND_LINE_ACTIVE = True
    while COMMAND_LINE_ACTIVE:
        value = input('\nSelect Operation:\n::$ ')
        if inputCompare(value, 'fragrances'):
            fetchItems()
        elif inputCompare(value, 'quit') or inputCompare(value, 'exit' or inputCompare(value, 'x')):
            COMMAND_LINE_ACTIVE = False
        else:
            print('\nUnknown Command: please enter only valid commands.\nEnter (help) for more details.')



'''
Fetches and returns option details for given product
'''
def fetchDetails(index, url):
    # get product detail page data
    response = requests.get(url)

    # parse website data
    soup = BeautifulSoup(response.text, 'lxml')

    # select script node with detail data
    node = list(filter(nodeHasDetailData, soup.find_all('script')))[0]

    # parse json product options data from node
    options = parseProductOptionsDetails(node)

    # maps details for each product into list item
    details = mapProductDetails(index, options)

    # return list of product details
    return details



'''
Fetches and prints all fragrance product data details
'''
def fetchItems():
    try:
        # get website data
        response = requests.get(FRAGRANCE_API_ROOT)

        # parse website data
        soup = BeautifulSoup(response.text, 'html.parser')

        # select product items
        items = soup.select('.resultItem > section > a')

        # collect product details
        products = []
        for index, item in enumerate(items):
            # add product details to list
            products += fetchDetails(index, item['href'])
        
        # print all product details
        pprinter = pprint.PrettyPrinter(depth=4)
        pprinter.pprint(products)

    except Exception as error:
        print('Exception occurred\n', error)



# start querying cli
commandLineQuerier()
sys.exit
