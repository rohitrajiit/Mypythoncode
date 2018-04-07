# -*- coding: utf-8 -*-
"""
Created on Sun Apr  1 23:28:18 2018

@author: rohit raj
downloads scripwise price volume data
"""
import sys
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import os
import json

path =r'C:\Users\rohit\AppData\Local\Google\Chrome\User Data\Default\Preferences'

with open(path) as f:
    pref = json.load(f)

path = pref['savefile']['default_directory']

driver = webdriver.Chrome()
driver.get('https://www.bseindia.com/markets/equity/EQReports/StockPrcHistori.aspx?scripcode=512289&flag=sp&Submit=G')
try:
    fromdate = driver.find_element_by_id('ctl00_ContentPlaceHolder1_txtFromDate')
    todate = driver.find_element_by_id('ctl00_ContentPlaceHolder1_txtToDate')
    scrip = driver.find_element_by_id('ctl00_ContentPlaceHolder1_GetQuote1_smartSearch')
    fromdate.clear()
    fromdate.send_keys('12/03/2018')
    driver.execute_script("document.getElementById('ctl00_ContentPlaceHolder1_txtToDate').value = '16/03/2018'")
    scrip.clear()
    scrip.send_keys('RCOM')
    time.sleep(10)
    scrip.send_keys(Keys.RETURN)
    submitt= driver.find_element_by_id("ctl00_ContentPlaceHolder1_btnSubmit")
    submitt.click()
    download = driver.find_element_by_id("ctl00_ContentPlaceHolder1_btnDownload1")
    download.click()
    scripname = driver.find_element_by_id("ctl00_ContentPlaceHolder1_lblScripCodeValue").text
    name = scripname + '.csv'
    while True:
        if name in os.listdir(path):
            break
        time.sleep(1)
except:
    print(sys.exc_info()[0])
    
    
finally:
    driver.close()
