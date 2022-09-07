from typing import Counter
from selenium import webdriver
import pandas as pd
import pyautogui, re, os
from time import sleep
from os import path
import pyperclip as pc
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
import xlwings as xw
from random import randint as rnd

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

def survParser(source, wellPath):
    if True:
        headerPattern = '<thead><tr><th>(.*)</th><th>(.*)</th><th>(.*)</th><th>(.*)</th><th>(.*)</th><th>(.*)</th><th>(.*)</th><th>(.*)</th><th>(.*)</th><th>(.*)</th><th>(.*)</th><th>(.*)</th><th>(.*)</th></tr></thead>'
        bodyPattern = 'XX(.*)XX(.*)XXX(.*)XXX(.*)XX(.*)XX(.*)XX(.*)XX(.*)XX(.*)XX(.*)XX(.*)XX(.*)XX(.*)XX'
        Y = re.compile('<.*?>')
        contentX = re.sub(Y,'X',source)

    RegEx1 = re.compile(headerPattern)
    RegEx2 = re.compile(bodyPattern)

    header = list(re.findall(RegEx1,source)[0])
    body = re.findall(RegEx2,contentX)
    for i, line in enumerate(body):
        body[i] = ', '.join(line)

    header = ', '.join(header)
    body = '\n'.join(body)
    table = header + '\n' + body

    csvFile = open(os.path.join(wellPath, 'Well Survey Data.csv'),'w+')
    csvFile.write(table)
    csvFile.close()

def verifyWell(browser):
    content = browser.page_source
    tries = 0
    while 'API No' not in content:
        try:
            content = browser.page_source
        except:
            sleep(0.5)
            tries += 1

        if tries == 20:
            raise TimeoutError('Ran out of tries')
    return content


for FILENUMBER, COUNTY in fileNoList:
    try:
        wellBrowser(FILENUMBER)
        main_window_handle = browser.current_window_handle

        browser.find_element_by_name('Image9').click()
        for handle in browser.window_handles:
                if handle != main_window_handle:
                    survey_window_handle = handle
                    break

        browser.switch_to.window(survey_window_handle)
        
        content = browser.page_source
        tries = 0
        while 'API No' not in content:
            try:
                content = browser.page_source
            except:
                sleep(1)
                tries += 1

            if tries == 6:
                raise TimeoutError('Ran out of tries')

        wellPath = os.path.join(wDir, 'Well Data', COUNTY, str(FILENUMBER))

        survParser(content, wellPath)

        browser.close()
        browser.switch_to.window(main_window_handle)
    except:
        missingWells.append([FILENUMBER, COUNTY])