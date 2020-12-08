# -*- coding: utf-8 -*-
"""
Created on Wed May 15 12:01:21 2019

@author: ampani
"""

import os
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
import time
from selenium.webdriver.common.by import By
import csv
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import openpyxl
import xlrd

def get_driver(browser:str = "firefox") -> webdriver:
    current_dir = os.path.dirname(os.path.abspath(__file__))
    #print(current_dir)
    if browser == "chrome":
        driver = webdriver.Chrome(executable_path=current_dir + "\\chromedriver.exe")
    else:
        driver = webdriver.Firefox(executable_path=current_dir + "\\geckodriver.exe")
    # Maximum Wait time to find elements in page (in order to let them load first)
    driver.implicitly_wait(10)

    return driver

def scarpe_din(din_no :str):
    driver = get_driver("chrome")
    link = 'http://www.mca.gov.in/mcafoportal/showVerifyDIN.do'
    driver.get(link)
    driver.maximize_window()
    try:
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, 'container')))
        print ("Page is ready!")
    except TimeoutException:
        print ("Loading took too much time!")
    din = driver.find_element_by_xpath('//*[@id="DIN"]')
    din.send_keys(din_no)
    time.sleep(5)
    submit = driver.find_element_by_xpath('//*[@id="verifyDIN_0"]')
    time.sleep(2)
    submit.click()
    time.sleep(2)
    dfname = driver.find_element_by_xpath('//*[@id="directorFullName"]').get_attribute('value')
    time.sleep(2)
    print(dfname)
    driver.quit()

if __name__ == "__main__":
    filepath = ("DIN Values -Test1.xls")
    book = xlrd.open_workbook(filepath)
    sheet = book.sheet_by_index(0)

    for i in range(sheet.nrows): 
        search_term = sheet.cell(i,0).value
        scarpe_din(search_term)
    #search_term = input("DIN/DPIN please? ")
    
    print('done...')