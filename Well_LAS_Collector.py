from typing import Counter
from selenium import webdriver
import pandas as pd
import pyautogui, re, os
from time import sleep, time
from os import close, path, write
import pyperclip as pc
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
import xlwings as xw
from random import randint as rnd
import zipfile
from xlwings.utils import exception

xlSheet, fileRange = 'WellIndex', 'fileNumbers'

if True:
    wDir = 'D:\\UND\\Data collection\\'
    fireProfile = 'C:/Users/Lotfi/AppData/Roaming/Mozilla/Firefox/Profiles/vsh67dar.default-release/'
else:
    wDir = 'D:\\NDIC Database\\'
    fireProfile = 'D:/Lotfi Allam/Projects/Python/NDIC Data Compilation/vsh67dar.default-release/'

missingWells = []

xlapp = xw.App(visible=False)
bk = xw.books.open(os.path.join(wDir,'Deviated & Horizontal Wells.xlsx'))
fileNoList = bk.sheets[xlSheet].range(fileRange).options(numbers=lambda x: str(int(x))).value
bk.close()

fp = webdriver.FirefoxProfile(fireProfile)
browser = webdriver.Firefox(firefox_profile=fp)
browser.implicitly_wait(5)

browser.get('https://www.dmr.nd.gov/oilgas/')
browser.find_element_by_id('premiumServiceBtn').click()
browser.find_element_by_id('scoutTicketBtn').click()
pyautogui.write('undgeology')
pyautogui.press("tab")
pyautogui.write('ND#beautiful18')
pyautogui.press("enter")

def wellBrowser(fileID):
    global browser
    fileElem = browser.find_element_by_name('FileNumber')
    fileElem.clear()
    fileElem.send_keys(fileID)
    browser.find_element_by_id('B1').click()

def dataCollector():
    global browser
    pyautogui.keyDown("ctrl")
    pyautogui.press("a")
    pyautogui.press("c")
    pyautogui.keyUp("ctrl")
    content = pc.paste()
    return content

def verifyPage(browser):
    content = browser.page_source
    tries = 0
    while 'Enter File Number' not in content:
        try:
            content = browser.page_source
        except:
            sleep(0.5)
            tries += 1

        if tries == 20:
            raise TimeoutError('Ran out of tries')
    return content

def verifyWell(fileNumber):
    global browser
    start = time()
    while 'NDIC File No: <b>{}</b>'.format(fileNumber) not in browser.page_source:
        if time() - start > 120:
            raise TimeoutError("Too much time spent")
        else:
            pass

for fileNumber, county in fileNoList:    
    wellBrowser(fileNumber)
    verifyWell(fileNumber)
    if '.las' not in browser.page_source:
        continue

    lasDONE = False
    counter = 0
    
    while not lasDONE:
        lasLinks = browser.find_elements_by_partial_link_text('.las')[counter:]
        if not lasLinks:
            break
        else:
            for lasLink in lasLinks:
                lasDONE = lasLink == lasLinks[-1]
                lasName = lasLink.text
                lasLink.click()
                counter += 1
                
                if 'WELL' in browser.page_source and 'COMP' in browser.page_source:
                    content = dataCollector()
                    browser.back()
                    path = os.path.join(wDir,'Well Data', county,str(fileNumber),lasName)
                    file = open(path, 'w+', encoding="utf-8")
                    content.replace('\ufffd','')
                    file.write(content)
                    file.close()
                    break
                else:
                    zipLAS = os.listdir('D:\\UND\\Data collection\\Well logs in LAS\\')[0]
                    dlPath = 'D:\\UND\\Data collection\\Well logs in LAS\\' + zipLAS 
                    path = os.path.join(wDir,county,str(fileNumber),str(fileNumber) + '-' + lasName)
                    with zipfile.ZipFile(dlPath, 'r') as zip_ref:
                        zip_ref.extractall(path)
                    os.remove(dlPath)

            