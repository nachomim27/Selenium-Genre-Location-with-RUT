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


    browser.get('https://www.nombrerutyfirma.com')

    WebDriverWait(browser, 5)\
        .until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[1]/ul/li[2]/a')))\
        .click()
        
    WebDriverWait(browser, 5)\
        .until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[1]/div/div[2]/form/div/input')))\
        .send_keys(rut.value)
        
    WebDriverWait(browser, 5)\
        .until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[1]/div/div[1]/div/div[2]/form/div/span/button/i')))\
        .click()

    WebDriverWait(browser, 5)\
        .until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[2]/div/table/tbody/tr/td[3]')))
        
    sexo=browser.find_element(by=By.XPATH, value=('/html/body/div[2]/div/table/tbody/tr/td[3]'))
    sexo=sexo.text
    comuna=browser.find_element(by=By.XPATH, value=('/html/body/div[2]/div/table/tbody/tr/td[5]'))
    comuna=comuna.text
    sheet['F'+str(n)]= sexo
    sheet['G'+str(n)]= comuna
    n=n+1
    time.sleep(1)

workbook.save(filename='Prueba4.xlsx')