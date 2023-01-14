# -*- coding: utf-8 -*-
"""
Created on Sun Jan  8 19:50:56 2023

@author: ignac
"""

# -*- coding: utf-8 -*-
from selenium import webdriver
from undetected_chromedriver import Chrome
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

from openpyxl import load_workbook




workbook = load_workbook(filename="Prueba.xlsx")
sheet = workbook.active

browser = Chrome()
options = webdriver.ChromeOptions() 

n=1
while sheet['A'+str(n)].value:
    rut=sheet['A'+str(n)]


    browser.get('https://fabianvillena.cl/rut-a-edad-fecha-de-nacimiento.html')
        
    WebDriverWait(browser, 5)\
        .until(EC.element_to_be_clickable((By.XPATH, '/html/body/section/div/div/div/div/div/div/input')))\
        .send_keys(rut.value)
        
    WebDriverWait(browser, 5)\
        .until(EC.element_to_be_clickable((By.XPATH,'/html/body/section/div/div/div/div/div/button')))\
        .click()
        
    nacioen=browser.find_element(by=By.XPATH, value=('/html/body/section/div/div/div/div/p/strong'))
    nacioen=nacioen.text
    
    sheet['H'+str(n)]= nacioen
    n=n+1
    time.sleep(4)

workbook.save(filename='Prueba3.xlsx')