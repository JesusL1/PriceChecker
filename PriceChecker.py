from bs4 import BeautifulSoup
import requests
import time
import re
import sys
import tldextract 
import ExcelEditor
import PriceCheckerGUI as PCG


def ValidateWebsite(webLink, excelPrice, action):
    """ Adds the item to the excel sheet if the item is new, otherwise update's the item's price
        Args:
        webLink: the url link of the product's store page/website
        excelPrice: the price that will be saved in the excel sheet and used as comparison for a deal
        action: tells the parser functions which action to perform
        """
    time.sleep(2)  # sleep before each request in order to not overwhelm the servers with rapid requests
    info = tldextract.extract(webLink)
    domain = info.domain  # gets the domain of the URL in order to run the appropriate function 
    func = functionList[domain]  # func holds the parsing function that matches the domain from the functionList
    if domain in functionList and action == 'Check':
        func(webLink, excelPrice, action)
    elif domain in functionList:
        func(webLink, excelPrice, action='Add')
    else:
        PCG.InsertText((domain + 'NOT IN functionList'), color=0)
   
def CheckPrices():
    """ Checks the prices of all the items in the excel sheet for a better deal
        Passes every product from the excel sheet into the ValidateWebsite function with the 'Check' action
    """
    PCG.InsertText('\nCHECKING PRICES...', color=0)
    for row in EE.sheet.iter_rows(min_row=2, max_col=4, max_row=EE.sheet.max_row, values_only=True):
        excel_productPrice = row[1]
        excel_websiteLink = row[2]
        ValidateWebsite(excel_websiteLink, excel_productPrice, 'Check')
    PCG.InsertText('FINISHED CHECKING PRICES...', color=0)

def BestPrice(productName, excelPrice, webLink, currentPrice):
    if currentPrice > excelPrice:
        result = '[FOUND] ' + productName + ' WORSE PRICE: (BEST_PRICE = ' + str(excelPrice) +') ' + '(CURRENT_PRICE = ' + str(currentPrice)+')'
        PCG.InsertText(result, color=1)
    elif currentPrice == excelPrice:
        result = '[FOUND] ' + productName + ' WITH EQUAL PRICE: (BEST_PRICE = ' + str(excelPrice) +') ' + '(CURRENT_PRICE = ' + str(currentPrice)+')'
        PCG.InsertText(result, color=2)
    elif currentPrice < excelPrice:
        result = '[FOUND] ' + productName + ' WITH BETTER PRICE: (OLD_BEST_PRICE = ' + str(excelPrice) +') ' + '(CURRENT_PRICE = ' + str(currentPrice)+')'
        PCG.InsertText(result, color=3)
        EE.UpdateExcel(webLink, currentPrice)
        EE.SendEmail(EE.emailUsername, EE.emailUsername, productName + ': ' + str(currentPrice), productName + ' is on sale for: $' + str(currentPrice) +' on the following website: ' + webLink)
        PCG.InsertText('EMAIL SENT!', color=0)

""" The following Parse functions vary slightly since every website is built differently.
Each Parse function uses requests to connect to the product's website and then the html is parsed using
Beautiful Soup for the product's name and price. The parse info is passed to the AddToExcel function if 
we're adding the item for the first time and will run the BestPrice function to determine whether 
it's a better deal than the price saved in the excel sheet. """
def Parse_93Brand(webLink, excelPrice, action):
    header_info = {'user-agent':'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.120 Safari/537.36'}
    r = requests.get(webLink, headers=header_info)
    try:
        soup = BeautifulSoup(r.text, 'lxml')
        productName_temp1 = soup.find('div', class_='product-title')
        productName_temp2 =  productName_temp1.find('h1', itemprop='name')
        productName = productName_temp2.text.strip()
        match1 = soup.find('div', class_='price')
        match2 = match1.find('span', class_='money')
        currentPrice = match2.text
        currentPrice = re.findall(r"[-+]?\d*\.\d+|\d+", currentPrice) #removes the $ sign, if the product has a sale price, it will take both prices and put it into an array where index 0 = sale price
        currentPrice = float(currentPrice[0])
    except AttributeError:
        print('ATTRIBUTE ERROR FOUND ON: ', webLink)
    except requests.exceptions.RequestException as e:
        print(e)
    if action == 'Add':
        EE.AddToExcel(productName, excelPrice, webLink)
        PCG.InsertText(('ADDED: ' + productName +' at $' + str(currentPrice)), color=4)
    BestPrice(productName, excelPrice, webLink, currentPrice)

def Parse_Adidas(webLink, excelPrice, action):
    header_info = {'user-agent':'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.120 Safari/537.36'}
    r = requests.get(webLink, headers=header_info)
    try:
        soup = BeautifulSoup(r.text, 'lxml')
        productName = soup.find('h1', class_='product_title')
        productName = productName.text
        productColor = soup.find('div', class_='color_text___mgoYV')
        productColor = productColor.text
        productName = productName + ' (Color: ' + productColor + ')'
        match1 = soup.find('div', class_='gl-price')
        currentPrice = match1.text
        currentPrice = re.findall(r"[-+]?\d*\.\d+|\d+", currentPrice) #removes the $ sign, if the product has a sale price, it will take both prices and put it into an array where index 0 = sale price
        currentPrice = float(currentPrice[0])
    except AttributeError:
        print('ATTRIBUTE ERROR FOUND ON: ', webLink)
    except requests.exceptions.RequestException as e:
        print(e)
    if action == 'Add':
        EE.AddToExcel(productName, excelPrice, webLink)
        PCG.InsertText(('ADDED: ' + productName +' at $' + str(excelPrice)), color=4)
    BestPrice(productName, excelPrice, webLink, currentPrice)

def Parse_BananaRepublic(webLink, excelPrice, action):
    r = requests.get(webLink)
    try:
        soup = BeautifulSoup(r.text, 'lxml')
        productName_temp =  soup.find('h1', class_='product-title__text')
        productName = productName_temp.text
        match1 = soup.find('h2', class_='product-price--pdp')
        currentPrice = match1.text
        currentPrice = re.findall(r"[-+]?\d*\.\d+|\d+", currentPrice) #removes the $ sign, if the product has a sale price, it will take both prices and put it into an array where index 0 = sale price
        currentPrice = float(currentPrice[0])
    except AttributeError:
        print('ATTRIBUTE ERROR FOUND ON: ', webLink)
    except requests.exceptions.RequestException as e:
        print(e)
    if action == 'Add':
        EE.AddToExcel(productName, excelPrice, webLink)
        PCG.InsertText(('ADDED: ' + productName +' at $' + str(excelPrice)), color=4)
    BestPrice(productName, excelPrice, webLink, currentPrice)
        
def Parse_FightersMarket(webLink, excelPrice, action):
    header_info = {'user-agent':'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.120 Safari/537.36'}
    r = requests.get(webLink, headers=header_info)
    try:
        soup = BeautifulSoup(r.text, 'lxml')
        productName_temp1 = soup.find('div', class_='product-title')
        productName_temp2 =  productName_temp1.find('h1', itemprop='name')
        productName = productName_temp2.text
        match1 = soup.find('div', class_='price_money')
        match2 = match1.find('span', class_='money')
        currentPrice = match2.text
        currentPrice = re.findall(r"[-+]?\d*\.\d+|\d+", currentPrice) #removes the $ sign, if the product has a sale price, it will take both prices and put it into an array where index 0 = sale price
        currentPrice = float(currentPrice[0])
    except AttributeError:
        print('ATTRIBUTE ERROR FOUND ON: ', webLink)
    except requests.exceptions.RequestException as e:
        print(e)
    if action == 'Add':
        EE.AddToExcel(productName, excelPrice, webLink)
        PCG.InsertText(('ADDED: ' + productName +' at $' + str(excelPrice)), color=4)
    BestPrice(productName, excelPrice, webLink, currentPrice)

def Parse_Microcenter(webLink, excelPrice, action):
    header_info = {'user-agent':'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.120 Safari/537.36'}
    r = requests.get(webLink,headers=header_info)
    try:
        soup = BeautifulSoup(r.text, 'lxml')
        match1 = soup.find('div', id='details')
        match2 = [item['data-brand'] for item in match1.find_all() if "data-brand" in item.attrs] #item brand
        match3 = [item['data-name'] for item in match1.find_all() if "data-name" in item.attrs] #item name
        productName = match2[0] + ' ' + match3[0]
        match1 = soup.find('span', id='pricing')
        currentPrice = match1.text
        currentPrice = re.findall(r"[-+]?\d*\.\d+|\d+", currentPrice) #removes the USD after the price (just want the value)
        currentPrice = float(currentPrice[0])
    except AttributeError:
        print('ATTRIBUTE ERROR FOUND ON: ', webLink)
    except requests.exceptions.RequestException as e:
        print(e)
    if action == 'Add':
        EE.AddToExcel(productName, excelPrice, webLink)
        PCG.InsertText(('ADDED: ' + productName +' at $' + str(excelPrice)), color=4)
    BestPrice(productName, excelPrice, webLink, currentPrice)

def Parse_Reebok(webLink, excelPrice, action):
    header_info = {'user-agent':'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.120 Safari/537.36'}
    r = requests.get(webLink,headers=header_info)
    try:
        soup = BeautifulSoup(r.text, 'lxml')
        match1 = soup.find('h1', class_='product_information_title___2rG9M')
        productName = match1.text
        match1 = soup.find('span', class_='gl-price__value')
        currentPrice = match1.text
        currentPrice = re.findall(r"[-+]?\d*\.\d+|\d+", currentPrice) #removes the USD after the price (just want the value)
        currentPrice = float(currentPrice[0])
    except AttributeError:
        print('ATTRIBUTE ERROR FOUND ON: ', webLink)
    except requests.exceptions.RequestException as e:
        print(e)
    if action == 'Add':
        EE.AddToExcel(productName, excelPrice, webLink)
        PCG.InsertText(('ADDED: ' + productName +' at $' + str(excelPrice)), color=4)
    BestPrice(productName, excelPrice, webLink, currentPrice)

""" functionList holds all the available website domains and their respective parsing functions. """
functionList = {'93brand':Parse_93Brand, 'adidas':Parse_Adidas, 'gap':Parse_BananaRepublic, 'fightersmarket':Parse_FightersMarket, 'microcenter':Parse_Microcenter, 'reebok':Parse_Reebok} 

EE = ExcelEditor.ExcelEditor('Book1.xlsx', 'Sheet', 2, 4, 'tom@example.com', 'password123')  # Create an ExcelEditor object


if __name__ == "__main__":
    """ Running the python file without any arguments will start the application 
    however running the program with an additional argument (doesn't matter what it is) will 
    start the program's CheckPrices() function. This is so that I can use the Windows Task Scheduler to 
    check for prices at a given time and so that the user doesn't have to wait for CheckPrices() to conclude
    at the beginning of each startup.
    """
    if len(sys.argv) == 1:
        PCG.root.mainloop()
    elif len(sys.argv) == 2:
        CheckPrices()
        PCG.root.mainloop()
    else:
        pass
