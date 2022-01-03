from bs4 import BeautifulSoup
from urllib.request import urlopen
import ssl
import re
import openpyxl as xl

wb = xl.load_workbook(filename="master.xlsx")

try:
    _create_unverified_https_context = ssl._create_unverified_context
except AttributeError:
    # Legacy Python that doesn't verify HTTPS certificates by default
    pass
else:
    # Handle target environment that doesn't support HTTPS verification
    ssl._create_default_https_context = _create_unverified_https_context

def print_rows(sheet):
    for row in sheet.iter_rows(values_only=True):
        print(row)

def clear_row(sheet, row):
    sheet["A" + str(row)] = None
    sheet["B" + str(row)] = None
    sheet["C" + str(row)] = None
    
def clear_sheet(sheet):
    idx = 2
    while True:
        if sheet["A" + str(idx)].value == None:
            return
        clear_row(sheet, idx)
        idx += 1

# Retrieve all  
def fill_shs():
    sheet = wb['SHS']
    titles = []
    prices = []
    lengths = []
    retrieve_values(titles, prices, lengths, 'https://handysteel.com.au/hollow-section-square-hollow-section-duragal-shs?p=', 1)

    if len(titles) != len(prices) or len(titles) != len(lengths):
        raise Exception
    clear_sheet(sheet)
    for cell_num in range(2, len(titles) + 2):
        sheet["A" + str(cell_num)] = titles[cell_num - 2]
        sheet["B" + str(cell_num)] = prices[cell_num - 2]
        sheet["C" + str(cell_num)] = lengths[cell_num - 2]
    print_rows(sheet)
    wb.save(filename="master.xlsx")

def retrieve_values(titles, prices, lengths, page, p_num):
    response = urlopen(page + str(p_num))

    soup = BeautifulSoup(response.read(), 'html.parser')    
    for list_item in soup.find('ul', {'id': 'product_list1'}):
        for line in list_item.find_all('input', {'name': 'main_price'}):
            price = float(re.findall(r"(?<=value\=\"\$).[^\"]*", str(line))[0])
            if p_num > 1 and price == prices[0]:
                return
            prices.append(price)
            print(price)
        for title in list_item.find_all('a', class_="product_img_link"):
            t_string = re.findall(r"(?<=title\=\").[^\"]*", str(title))[0]
            titles.append(t_string)
            print(t_string)
        for length in list_item.find_all('input', {'name': 'max_len'}):
            value = float(re.findall(r"(?<=value\=\").[^\"]*", str(length))[0])
            lengths.append(value)
            print(value)
    retrieve_values(titles, prices, lengths, page, p_num + 1)

fill_shs()

