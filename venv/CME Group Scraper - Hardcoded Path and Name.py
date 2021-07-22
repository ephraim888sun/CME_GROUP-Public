from selenium import webdriver
import selenium.webdriver.support.ui as ui
import tkinter as tk
import selenium.common.exceptions as sce
from tkinter.filedialog import askdirectory
import xlsxwriter as xl
import datetime

#https://stackoverflow.com/questions/16927354/how-can-i-make-selenium-python-wait-for-the-user-to-login-before-continuing-to-r

path = r"C:\chromedriver_win32\chromedriver.exe"
#driver = webdriver.Chrome(path)
table = []
monthtable = []
editedmonthtable = []
editedtable = []
x = 0
date = '{:%m-%d-%Y}'.format(datetime.datetime.today())

username = 'sokudjeto'
#fullname = f'C:\\Users\{username}\Documents\Projects\Ad-Hoc Request\Logan T\lightsweetcrudequotes_{date}.xlsx'
#fullname = 'test'
text = ''
sites = ["https://www.cmegroup.com/markets/energy/crude-oil/light-sweet-crude.quotes.html",
         "https://www.cmegroup.com/trading/energy/natural-gas/natural-gas_quotes_globex.html",
         "https://www.cmegroup.com/trading/energy/petrochemicals/mont-belvieu-natural-gasoline"
         "-5-decimal-opis-swap_quotes_globex.html",
         "https://www.cmegroup.com/trading/energy/petrochemicals/mont-belvieu-iso-butane"
         "-5-decimal-opis-swap-futures.html",
         "https://www.cmegroup.com/trading/energy/petrochemicals/mont-belvieu-normal-butane-5-decimals-swap.html",
         "https://www.cmegroup.com/trading/energy/petrochemicals/mont-belvieu-propane"
         "-5-decimals-swap_quotes_globex.html",
         "https://www.cmegroup.com/trading/energy/petrochemicals/mont-belvieu-ethane-opis-5-decimals-swap.html"]

# pull data from website
with webdriver.Chrome(path) as driver:
    for site in range(len(sites)):
    #for site in range(3):
        try:
            #driver.get("https://www.cmegroup.com/markets/energy/crude-oil/light-sweet-crude.quotes.html")
            driver.get(f'{sites[site]}')
            if site == 0:
                wait = ui.WebDriverWait(driver, 10)  # timeout after 10 seconds
                results = wait.until(lambda driver: driver.find_element_by_id("pardotCookieButton")) # accept cookies on pop so data can be pulled
                results.click()
            title = driver.find_element_by_xpath('// *[ @ id = "main-content"] / div / div[2] / div / h1').text
            fullname = f'C:\\Users\{username}\Documents\Projects\Ad-Hoc Request\Logan T\{title}quotes_{date}.xlsx'
            wait2 = ui.WebDriverWait(driver, 5)  # timeout after 5 seconds
            results2 = wait2.until(lambda driver: driver.find_element_by_xpath(
                '//*[@id="productTabData"]/div/div/div/div/div/div/div[5]/div/div/div/div[2]/div[2]/button'))
            results2.click()
                # load all available data in page
            #pull & format main table
            contents = driver.find_element_by_xpath(
                '//*[@id="productTabData"]/div/div/div/div/div/div/div[5]/div/div/div/div[1]/div/table').text
                # pull unedited data into array
            for item in contents:
                if item != " ": #data currently viewed only be individual characters, merges data into whole words and skips blanks
                    text += item
                else:
                    table.append(text)
                    text = ''
            if title == 'Mont Belvieu Iso-Butane (OPIS)':
                for thing in range(1, len(table)): #further clean up data formatting, skips OPTIONS & CHART columns that aren't needed
                    newtext = table[thing]
                    if '\n' in newtext: #scrub newlines from text
                        if 'PRIOR' in newtext: # keeps PRIOR SETTLE item as a single item instead of breaking it up
                            edit = newtext.replace('\n', ' ')
                            editedtable.append(edit)
                        else:
                            newnew = newtext.split('\n') # breaks up text into multiple items whenever a \n is encountered
                            for item in range(len(newnew)):
                                if newnew[item] != 'OPT' and newnew[item] != '':
                                    if "(" in newnew[item]: #keeps data for CHANGE column in a single item
                                        changeloc = len(editedtable)-1
                                        changetext = editedtable[changeloc]
                                        editedtable[changeloc] += newnew[item]
                                    else:
                                        editedtable.append(newnew[item])
                    else:
                        editedtable.append(newtext)
                # pull and format month header table
                months = driver.find_element_by_xpath(
                    '//*[@id="productTabData"]/div/div/div/div/div/div/div[5]/div/div/div/div[1]/div[2]/table/tbody').text
                for month in months:
                    if month != " ":  # data currently viewed only be individual characters, merges data into whole words and skips blanks
                        text += month
                    else:
                        monthtable.append(text)
                        text = ''
                for mthing in range(len(monthtable)): #further clean up data formatting
                    if '\n' in monthtable[mthing]: #scrub newlines from text
                        newnewmonth = monthtable[mthing].split('\n\n') # breaks up text into multiple items whenever a \n\n is encountered, \n\n marks the start of items that need to be grouped into one
                        for mitem in range(len(newnewmonth)):
                            if '\n' in newnewmonth[mitem]:
                                editedmonthtable.append(newnewmonth[mitem].split('\n'))
                            else:
                                editedmonthtable.append(newnewmonth[mitem])
                    else:
                        editedmonthtable.append(monthtable[mthing])
                for emthing in range(len(editedmonthtable)):
                    #print(len(editedmonthtable[emthing]))
                    if len(editedmonthtable[emthing]) == 2:
                        emtextedit = f'{editedmonthtable[emthing - 1]} {editedmonthtable[emthing][0]}'
                        editedmonthtable[emthing - 1] = emtextedit
                        editedmonthtable[emthing] = ''  #program shits itself when using 'pop' function here, change items to remove to empyt strings and skip when writing to excel

                try: # formats UPDATED items into single item and removes items no longer needed in array after formatting
                    for t in range(len(editedtable)):
                        if ":" in editedtable[t]:
                            merge = f'{editedtable[t]} {editedtable[t + 1]} {editedtable[t + 2]}' \
                                f' {editedtable[t + 3]} {editedtable[t + 4]}'
                            editedtable[t] = merge
                            for poparray in range(4):
                                editedtable.pop(t + 1)
                except IndexError:
                    pass
            else:
                for thing in range(2, len(
                        table)):  # further clean up data formatting, skips OPTIONS & CHART columns that aren't needed
                    newtext = table[thing]
                    if '\n' in newtext:  # scrub newlines from text
                        if 'PRIOR' in newtext:  # keeps PRIOR SETTLE item as a single item instead of breaking it up
                            edit = newtext.replace('\n', ' ')
                            editedtable.append(edit)
                        else:
                            newnew = newtext.split('\n')
                            # breaks up text into multiple items whenever a \n is encountered
                            for item in range(len(newnew)):
                                if newnew[item] != 'OPT' and newnew[item] != '':
                                    if "(" in newnew[item]:  # keeps data for CHANGE column in a single item
                                        changeloc = len(editedtable) - 1
                                        changetext = editedtable[changeloc]
                                        editedtable[changeloc] += newnew[item]
                                    else:
                                        editedtable.append(newnew[item])
                    else:
                        editedtable.append(newtext)
                # pull and format month header table
                months = driver.find_element_by_xpath(
                    '//*[@id="productTabData"]/div/div/div/div/div/div/div[5]'
                    '/div/div/div/div[1]/div[2]/table/tbody').text
                for month in months:
                    if month != " ":  # data currently viewed only be individual characters, merges data into whole words and skips blanks
                        text += month
                    else:
                        monthtable.append(text)
                        text = ''
                if title == "Mont Belvieu Natural Gasoline (OPIS)":
                    for mthing in range(len(monthtable)):  # further clean up data formatting
                        if '\n' in monthtable[mthing]:  # scrub newlines from text
                            if 'OPT' in monthtable[mthing]:
                                altered = monthtable[mthing].replace("\n", '')
                                editedmonthtable.append(altered)
                            else:
                                newnewmonth = monthtable[mthing].split('\n\n')
                                # breaks up text into multiple items whenever a \n\n is encountered, \n\n marks the start of items that need to be grouped into one
                                for mitem in range(len(newnewmonth)):
                                    if '\n' in newnewmonth[mitem]:
                                        editedmonthtable.append(newnewmonth[mitem].split('\n'))
                                    else:
                                        editedmonthtable.append(newnewmonth[mitem])
                        else:
                            editedmonthtable.append(monthtable[mthing])
                else:
                    for mthing in range(len(monthtable)):  # further clean up data formatting
                        if '\n' in monthtable[mthing]:  # scrub newlines from text
                            newnewmonth = monthtable[mthing].split('\n\n')
                            # breaks up text into multiple items whenever a \n\n is encountered, \n\n marks the start of items that need to be grouped into one
                            for mitem in range(len(newnewmonth)):
                                if '\n' in newnewmonth[mitem]:
                                    editedmonthtable.append(newnewmonth[mitem].split('\n'))
                                else:
                                    editedmonthtable.append(newnewmonth[mitem])
                        else:
                            editedmonthtable.append(monthtable[mthing])
                for emthing in range(len(editedmonthtable)):
                    if len(editedmonthtable[emthing]) == 2:
                        emtextedit = f'{editedmonthtable[emthing - 1]} {editedmonthtable[emthing][0]}'
                        editedmonthtable[emthing - 1] = emtextedit
                        editedmonthtable[emthing] = ''
                        # program shits itself when using 'pop' function here, change items to remove to empyt strings and skip when writing to excel
                try:  # formats UPDATED items into single item and removes items no longer needed in array after formatting
                    for t in range(len(editedtable)):
                        if ":" in editedtable[t]:
                            merge = f'{editedtable[t]} {editedtable[t + 1]} {editedtable[t + 2]}' \
                                f' {editedtable[t + 3]} {editedtable[t + 4]}'
                            editedtable[t] = merge
                            for poparray in range(4):
                                editedtable.pop(t + 1)
                except IndexError:
                    pass
        except IndexError:
            pass
        newspreadsheet = rf"{fullname}"
        wb = xl.Workbook(newspreadsheet)
        ws = wb.add_worksheet()
        row = 0
        mrow = 1
        column = 0

        ws.write(row, column, "MONTH")
        column += 1

        for h in range(len(editedmonthtable)):
            if editedmonthtable[h] == '':
                pass
            else:
                ws.write(mrow, 0, editedmonthtable[h])
                mrow += 1
        for i in range(0, len(editedtable)):
            if i % 7 == 0 and i != 0 and row == 0:
                ws.write(row, column, editedtable[i])
                column = 1
                row += 1
            elif row != 0 and ":" in editedtable[i]:
                ws.write(row, column, editedtable[i])
                column = 1
                row += 1
            else:
                ws.write(row, column, editedtable[i])
                column += 1
        wb.close()
        table.clear()
        monthtable.clear()
        editedmonthtable.clear()
        editedtable.clear()