from bs4 import BeautifulSoup
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By

driver = webdriver.Chrome("C:\\Users\\Ephraim.Sun\\PycharmProjects\\CME GROUP\\chromedriver.exe")
driver.get("https://www.cmegroup.com/markets/energy/crude-oil/light-sweet-crude.quotes.html#")

y = driver.find_element_by_css_selector('[class="cmeButton cmeButtonPrimary btn primary').click()
x = driver.find_element_by_css_selector('[class="btn primary load-all').click()

soup = BeautifulSoup(driver.page_source,'html.parser')
tables = soup.find('table')

src = driver.page_source # gets the html source of the page
parser = BeautifulSoup(src,"html.parser") # initialize the parser and parse the source "src"

# Month=""
# Options=""
# Chart=""
# Last=""
# Change=""
# Prior=""
# Settle=""
# Open=""
# High=""
# Low=""
# Volume=""
# Updated=""


list_of_rows = []
for row in tables.findAll('tr')[1:]:
    list_of_cells = []
    for cell in row.findAll('td'):
        text = cell.text.replace("&nbsp;", "")
        list_of_cells.append(text)
    list_of_rows.append(list_of_cells)


'''Create DataFrame'''
df = pd.DataFrame(list_of_rows, columns= ['MONTH', 'OPTIONS', 'CHART', 'LAST', 'CHANGE', 'PRIOR SETTLE', 'OPEN', 'HIGH', 'LOW', 'VOLUME', 'UPDATED'])


df.to_csv('CME.csv', index=False)

driver.close()
