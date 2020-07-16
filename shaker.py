import requests
import json
import sys
import random
from bs4 import BeautifulSoup
import time
import threading
import pprint
import xlsxwriter

# Create workbook
workbook = xlsxwriter.Workbook('Productinfo.xlsx')
worksheet = workbook.add_worksheet()
row = 0
col = 0

# Iternate over the data and write it out row by row
def writeToFile(x):
    for value in x:
        worksheet.write(row, col, value.id)
        worksheet.write(row, col, value.sku)
        worksheet.write(row, col, + 1, value.name)
        worksheet.write(row, col, + 2, value.img)
        worksheet.write(row, col, + 3, value.zoom_img)
        worksheet.write(row, col, + 4, value.list_price)
        worksheet.write(row, col, + 5, value.retail_price)
        worksheet.write(row, col, + 6, value.sale_price)
        worksheet.write(row, col, + 7, value.quantity)
        row += 1
    workbook.close()


# global variables / settings
FRAGRANCE_API_ROOT = 'https://www.fragrancenet.com/fragrances'

def LimitProduct(items, previous_items):
     items[0] = previous_items[0]
 
        

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
    x=1
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
                products += fetchDetails(index, item['href'])
            
            if LimitProduct(items, previous_items):
                break
            else: 
                # print all product details
                pprinter = pprint.PrettyPrinter(depth=4)
                pprinter.pprint(products)

            previous_items = items
            x+=1


        except Exception as error:
            print('Exception occurred\n', error)



# start querying cli
commandLineQuerier()
sys.exit