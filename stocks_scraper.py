import requests
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver

import os
import time
from time import sleep
from datetime import date

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException


def get_text(html):
    soup = BeautifulSoup(html, 'lxml')
    rows= soup.find_all('tbody', attrs={'class':"mi-table__tbody"})
    return rows

def get_data(rows):
    data=[]
    for row in rows:
        cols = row.find_all('td')
        cols = [ele.text.strip() for ele in cols]
        data.append([ele for ele in cols])
    return data

def get_columns(data, company, stock, sector, last_price, percent_change, last_update):
    for col in range(6):
        for col_num in range(col,500, 6):
            try:
                if col == 0:
                    company.append(data[0][col_num])
                elif col == 1:
                    stock.append(data[0][col_num])
                elif col == 2:
                    sector.append(data[0][col_num])
                elif col == 3:
                    last_price.append(data[0][col_num])
                elif col == 4:
                    percent_change.append(data[0][col_num])
                elif col == 5:
                    last_update.append(data[0][col_num])
            except:
                break
    return company, stock, sector, last_price, percent_change, last_update


if __name__ == "__main__":

    resp = requests.get('https://www.mubasher.info/countries/sa/companies')
    resp.raise_for_status()

    EXE_PATH = 'C:/webdrivers/chromedriver.exe'

    browser = webdriver.Chrome(executable_path=EXE_PATH)
    browser.get('https://www.mubasher.info/countries/sa/companies')

    delay = 3 # seconds
    try:
        element = WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.LINK_TEXT, """التالي""")))
        print("Page is ready!")
    except TimeoutException:
        print("Loading took too much time!\nPlease retry again")

    company= []
    stock=[]
    sector=[]
    last_price=[]
    percent_change=[]
    last_update=[]
    i=1
    while True:
        sleep(2)
        data = []
        html = browser.page_source
        rows= get_text(html)
        table_data= get_data(rows)
        company, stock, sector, last_price, percent_change, last_update= \
        get_columns(table_data, company, stock, sector, last_price, percent_change, last_update)
        sleep(3)
        print(str(i) + '.'*i*3)
        i+=1
        try:
            element = browser.find_element_by_link_text("""التالي""")
            browser.execute_script("arguments[0].click();", element)
        except:
            break

    browser.close()
    browser.quit()
    
    data_dict= {'الشركة': company,'السهم': stock, 'القطاع':sector, 'آخر سعر':last_price, 'نسبة التغيير':percent_change, 'آخر تحديث':last_update}
    df = pd.DataFrame.from_dict(data_dict, orient='index').T
    
    t = time.localtime()
    folder_dir= date.today().strftime("%Y/%m")
    filename = time.strftime('%d-%b-%Y', t)
    
    if os.path.isdir(folder_dir):
        df.to_excel(folder_dir + '/' + filename + '.xlsx')
        
    else:
        os.makedirs(folder_dir)
        df.to_excel(folder_dir + '/' + filename + '.xlsx')
    print('............................Done............................')
