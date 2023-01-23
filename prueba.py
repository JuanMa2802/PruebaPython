# -*- coding: utf-8 -*-
"""
Created on Fri Nov  4 14:24:40 2022

@author: Juan Manuel Gaviria
"""

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from time import sleep
from selenium.webdriver.common.by import By
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import Select
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.action_chains import ActionChains
from datetime import datetime
import datetime
from time import localtime, strptime
from datetime import timedelta
from datetime import date, datetime, timedelta
from dateutil.relativedelta import relativedelta
from dateutil.relativedelta import relativedelta, MO
from time import localtime, strftime
import time
global pagereports
import pandas as pd
import time
import glob
import os
from os.path import exists
import sys
import openpyxl
import warnings
import re
from datetime import datetime
import numpy as np
# from datetime import date
import calendar
from office365.sharepoint.files.file import File
from office365.sharepoint.listitems.listitem import ListItem
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from shareplum import Office365, Site
from shareplum.site import Version
import pyautogui
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.drawing.image import Image

warnings.filterwarnings("ignore")


opc = webdriver.ChromeOptions()

opc.add_argument("--no-sandbox")
opc.add_argument("--disable-dev-shm-usage")
opc.add_argument("--disable-gpu")
opc.add_argument("--disable-blink-features=AutomationControlled")
opc.add_argument("--start-maximized")
# opc.add_argument("--window-size=1920x1080")
opc.add_argument("--enable-features=NetworkService,NetworkServiceInProcess")
opc.add_argument("--ignore-certificate-errors")
opc.add_argument("--allow-running-insecure-content")
opc.add_argument("--disable-notifications")
opc.add_argument("--disable-blink-features")
# opc.add_argument("--incognito")
opc.add_argument('--no-proxy-server')
opc.add_argument("--proxy-server='direct://'")
opc.add_argument("--proxy-bypass-list=*")
opc.add_argument('--disable-dev-shm-usage')
opc.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.0.0 Safari/537.36")


driver = webdriver.Chrome(ChromeDriverManager().install(),chrome_options=opc)
action = ActionChains(driver)


driver.get('https://www1.sincoerp.com/SincoArconsa/V3/Marco/Login.aspx')
time.sleep(5)

empresa_dic = {
    'AC': '/html/body/section/section[2]/div[2]/section[1]/select/option[2]'
}


def IngresarTexto(xpath, texto):
    WebDriverWait(driver, 5)\
        .until(EC.element_to_be_clickable((By.XPATH,xpath,)))\
        .send_keys(texto)
        
def Click(xpath):
    WebDriverWait(driver, 7)\
        .until(EC.element_to_be_clickable((By.XPATH, xpath)))\
        .click()
        
def LimpiarCampos(xpath):
    WebDriverWait(driver, 10)\
        .until(EC.element_to_be_clickable((By.XPATH,xpath,)))\
        .clear()

site_url ='https://arquitecturayconstrucciones.sharepoint.com/sites/UNIDOSSOMOSMS9/SGD/Vales'
ctx = ClientContext(site_url).with_credentials(UserCredential("automations@arconsa.com.co", "Lambda63074"))


#Inicio de sesión
IngresarTexto('//*[@id="txtUsuario"]', 'automate3')
time.sleep(2)
IngresarTexto('//*[@id="txtContrasena"]', 'Lambda2022*')
time.sleep(2)
Click('/html/body/section/section[2]/div[2]/div[2]/button')

#Selección empresa
Click('//*[@id="ddlEmpresa"]')
time.sleep(2)
Click(empresa_dic['AC'])
time.sleep(3)
Click('//*[@id="btnIngresar"]')

#FE/Recepción/Conciliación de documentos
Click('/html/body/section[1]/aside[1]/nav/div[3]')
time.sleep(3)
Click('/html/body/section[1]/section/section[4]/nav/nav[1]/nav/div')
time.sleep(3)
Click('/html/body/section[1]/section/section[4]/nav/nav[1]/nav[2]/div[4]/div[2]')

time.sleep(2)
driver.quit()