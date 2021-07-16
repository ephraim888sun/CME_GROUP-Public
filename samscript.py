from bs4 import BeautifulSoup
import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

#
# options = Options()
# options.page_load_strategy = 'normal'
# driver = webdriver.Chrome(options=options)
driver = webdriver.Chrome("C:\\Users\\Ephraim.Sun\\PycharmProjects\\CME GROUP\\chromedriver.exe")
driver.get("https://www.cmegroup.com/markets/energy/crude-oil/light-sweet-crude.quotes.html#")

# driver.find_element_by_xpath('//*[@id="productTabData"]/div/div/div/div/div/div/div[5]/div/div/div/div[1]/div[1]/table/tbody').click()



# try:
#   path.click()
#   # add to list of clickable elements
# except WebDriverException:
#   print("Element is not clickable")


soup = BeautifulSoup(driver.page_source,'html.parser')
tables = soup.find('table')
# print(tables)

src = driver.page_source # gets the html source of the page
parser = BeautifulSoup(src,"html.parser") # initialize the parser and parse the source "src"



# print(tables)

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
# print (list_of_rows)
# print(" ")
# print(tables)

print("\n\nPANDAS DATAFRAME\n")
df = pd.DataFrame(list_of_rows, columns= ['MONTH', 'OPTIONS', 'CHART', 'LAST', 'CHANGE', 'PRIOR SETTLE', 'OPEN', 'HIGH', 'LOW', 'VOLUME', 'UPDATED'])
print(len(df.index))

df.to_csv('CME.csv', index=False)

driver.close()
