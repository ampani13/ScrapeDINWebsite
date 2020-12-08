from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
import xlrd
import xlwt
from xlutils.copy import copy

loc = ("C:\\Users\\kanishk\\Desktop\\ML\\MosaikIntern-Assignment\\DIN Values -Test1.xls")

wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

sheet1 = xlwt.Workbook()

def writer():
    read_book = xlrd.open_workbook(loc)
    work_book = copy(read_book)

    sheet1 = work_book.get_sheet(0)

    for i,e in enumerate(resultset):
        sheet1.write(i,1,e)
    work_book.save(loc)


print(sheet.nrows)
resultset = []
for i in range(sheet.nrows):
    data = sheet.row_values(i)
    browser = webdriver.Chrome("C:\\Users\\kanishk\\Downloads\\chromedriver")
    browser.get("http://www.mca.gov.in/mcafoportal/showVerifyDIN.do")
    delay = 3
    
    try:
        myElem = WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.ID, 'IdOfMyElement')))
        print ("Page is ready!")
    except TimeoutException:
        print ("Loading took too much time!")

    search = browser.find_element_by_id('DIN')
    search.send_keys(data)
    delay = 2
    search.submit()
    dfname = browser.find_element_by_xpath('//*[@id="directorFullName"]').get_attribute('value')
    resultset.append(dfname)
    browser.quit()
    writer()



'''
import pandas as pd
df = pd.read_excel("C:\\Users\\kanishk\\Desktop\\ML\\MosaikIntern-Assignment\\DIN Values -Test1.xls")

print(df.loc[0])

browser = webdriver.Chrome("C:\\Users\\kanishk\\Downloads\\chromedriver")
browser.get("http://www.mca.gov.in/mcafoportal/showVerifyDIN.do ")
delay = 3
try:
    myElem = WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.ID, 'IdOfMyElement')))
    print ("Page is ready!")
except TimeoutException:
    print ("Loading took too much time!")
search = browser.find_element_by_id('DIN')
search.send_keys(data)
delay = 2
search.submit()
dfname = browser.find_element_by_xpath('//*[@id="directorFullName"]').get_attribute('value')
'''


