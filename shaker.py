import requests
import json
import sys
import random
from bs4 import BeautifulSoup
import time
import threading
import pprint
import xlsxwriter

# global variables / settings
FRAGRANCE_API_ROOT = 'https://www.fragrancenet.com/fragrances'
row = 0
col = 0

# Create workbook
workbook = xlsxwriter.Workbook('제품정보.xlsx')
worksheet = workbook.add_worksheet()

# Initialize document with titles
titles = ['아디디', 'SKU', '제품명', '브랜드', '이미지 300', '이미지 900', 'List 가격', 'Retail 가격', '세일 가격', '재고', '성별']
for index, title in enumerate(titles):
    worksheet.write(row, index, title)


# Iterate over the data and write it out row by row
def writeToFile(content):
    global row
    keys = ['id', 'sku', 'name', 'brand_designer', 'img', 'zoom_img', 'list_price', 'retail_price', 'sale_price', 'stock_warn', 'gender']
    for value in content:
        row += 1
        for index, key in enumerate(keys):
            worksheet.write(row, index, value.get(key))


def endOfProductsReached(items, previous_items):
    return items is not None and previous_items is not None and items[0] == previous_items[0]
        

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
def mapProductDetails(group_id, options, group):
    products = []
    brand_designer = ''
    brand = group.get('brand')
    designer = group.get('designer')
    gender = group.get('gender')
    if brand is not None and designer is not None:
        brand_designer = brand + ' by ' + designer
    elif brand is not None:
        brand_designer = brand
    else:
        brand_designer = designer
    for key, value in options.items():

        stockWarn = value.get("stock_warn")
        if stockWarn == 0:
            stockWarn = 'enough quantity'

        products.append({
            'id': group_id, # the id used to group variants of the same product i.e. product options
            'sku': key,
            'name': value.get('SIZE_default'),
            'brand_designer': brand_designer,
            'img': value.get('img'),
            'zoom_img': value.get('zoom_img'),
            'list_price': value.get('price_int'),
            'retail_price': value.get('retail_price_int'),
            'sale_price': value.get('discount_price_int'),
            'stock_warn': stockWarn,
            'gender': gender,
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

    node = None
    node_2 = None

    # select script node with detail data
    scripts = list(filter(nodeHasDetailData, soup.find_all('script')))
    if len(scripts) > 0:
        node = scripts[0]
    else:
        return None

    # select script node with group detail data
    scripts_2 = list(soup.find_all('script', type='optimize-js'))
    if len(scripts) > 0:
        node_2 = scripts_2[::-1][0]
    else:
        return None

    # parse json product options data from node
    options = parseProductOptionsDetails(node)

    # parse json newproduct options data from node
    group = parseProductGroupDetails(node_2)

    # maps details for each product into list item
    details = mapProductDetails(index, options, group)

    # return list of product details
    return details


def parseProductGroupDetails(node):

    # convert node to string
    text = node.string

    # find starting location of details
    start = text.find('productView_id')

    # find end location of details
    end = text.find('"gender"')

    # get the value between start and end, and strip out all unneeded whitespace
    productView = text[start + 17: end + 13].strip()

    # return value as JSON
    return json.loads(productView)


'''
Fetches and prints all fragrance product data details
'''
def fetchItems():
    x = 0
    total = 0
    previous_items = [None]
    while(True):
        try:
            # get website data
            response = requests.get(FRAGRANCE_API_ROOT + '?page=' + str(x))

            # parse website data
            soup = BeautifulSoup(response.text, 'html.parser')

            # select product items
            items = soup.select('.resultItem > section > a')

            # collect product details
            products = []
            for index, item in enumerate(items):
                # add product details to list
                details = fetchDetails(index + total, item['href'])
                if details is not None:
                    products += details
                    total += len(items)

            if endOfProductsReached(items, previous_items):
                break
            else:
                writeToFile(products)
            previous_items = items
            x += 1

        except Exception as error:
            print('Exception occurred\n', error)

    workbook.close()


# start querying cli
commandLineQuerier()
sys.exit