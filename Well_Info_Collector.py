from selenium import webdriver
import pandas as pd
import pyautogui, re, os
from time import sleep
from os import path
import pyperclip as pc
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
import xlwings as xw
from random import randint as rnd

if True:
    wDir = 'D:\\UND\\Data collection\\'
    xlSheet, fileRange = 'MCKENZIE', 'fileNos'
    fireProfile = 'C:/Users/Lotfi/AppData/Roaming/Mozilla/Firefox/Profiles/vsh67dar.default-release/'
else:

    wDir = 'D:\\NDIC Database\\'
    xlSheet, fileRange = 'WellIndex', 'fileNumbers'
    fireProfile = 'D:/Lotfi Allam/Projects/Python/NDIC Data Compilation/vsh67dar.default-release/'

completedWells = []
headerFrame = pd.DataFrame(columns=('County', 'Wellbore type', 'Latitude', 
    'Longitude', 'KB (ft)', 'Total Depth (ft)', 'Field', 'Start Production', 
    'Multi-Well Pad', 'Perfs', 'Cum Oil', 'Cum Gas', 'Cum Water', 'LAS File', 'Sonic Log'))

xlapp = xw.App(visible=False)
bk = xw.books.open(os.path.join(wDir,'Well Index.xlsx'))
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

def dataParser(content):
    Matches = []
    RegExList = [
        re.compile("NDIC File No: (\d+)"),
        re.compile("County: (\w+)"),
        re.compile("Wellbore type: (.*)\n"),
        re.compile("Latitude: (-?\d+\.\d+)"),
        re.compile("Longitude: (-?\d+\.\d+)"),
        re.compile("Elevation\(s\).*?\s+?(\d+ KB)|(\d+ GL)|(\d+ DF)|(\d+) Total"),
        re.compile("Total Depth: (\d+)"),
        re.compile("Field: (\w+)"),
        re.compile("SKIP (DATA)"),
        re.compile("   NDIC File No: (\d+)   Well"),
        re.compile("Pool: (\w+\s?\w+\s?\w+)     (Perfs: (\d+-\d+\s?\w*)     )?Comp"),
        re.compile("Cum Oil: (\d+)"),
        re.compile("Cum MCF Gas: (\d+)"),
        re.compile("Cum Water: (\d+)"),
        re.compile("\S+.las"),
        re.compile("SOS.las"),
        ]

    for index, regex in enumerate(RegExList):
        datum = re.findall(regex,content)
        if not datum:
            if index == 14 or index == 15:
                datum = 'No'
            else:
                datum = 'N/A'
        else:
            if index == 3 or index == 4:
                datum = float(datum[0])

            elif index == 5:
                datum = int(next(s for s in datum[0] if s)[:-3])

            elif index == 9:
                datum = ', '.join(datum)

            elif index == 10:
                output = ''
                for index2, completion in enumerate(datum):

                    if not completion[2]:
                        interval = 'N/A'
                    else:
                        interval = completion[2]

                    if index2 == len(datum) - 1:
                        output += completion[0] + ':' + interval
                    else:
                        output += completion[0] + ':' + interval + ', '
                datum = output
            
            elif index == 11 or index == 12 or index == 13:
                cumulProd = 0
                for Prod in datum:
                    cumulProd += int(Prod)
                datum = cumulProd
            elif index == 14 or index == 15:
                datum = 'Yes'
            else:
                try:
                    datum = int(datum[0])
                except ValueError:
                    datum = datum[0]

        Matches.append(datum)
        
    return Matches

def prodParser(content, wellPath):
    RegEx = re.compile('\n(\w.*\t\w+\t\w+\t\w+.*)+')
    prodTable = re.findall(RegEx,content)
    csvFile = open(os.path.join(wellPath, 'Production Data.csv'),'w+')
    csvFile.write('\n'.join(prodTable).replace('\t',','))
    csvFile.close()

for fileNumber in fileNoList:
    wellBrowser(fileNumber)
    wellData = dataCollector()
    parsedData = dataParser(dataCollector())
    headerFrame.loc[parsedData[0]] = parsedData[1:]

    wellPath = os.path.join(wDir, 'Well Data', parsedData[1], str(fileNumber)) 
    try:
        os.makedirs(wellPath)
    except FileExistsError:
        pass
    
    if 'production history for this well' in wellData:
        browser.find_element_by_name('Image27').click()
        pyautogui.press("tab")
        prodData = dataCollector()
        browser.back()
        prodParser(prodData, wellPath)

    

bk2 = xw.books.open(os.path.join(wDir,'Well Index.xlsx'))
bk2.sheets['Well Information'].range('A1').value = headerFrame
bk2.sheets['Well Information'].range('A1').value = 'NDIC File No'

if fileNoList == headerFrame.index.values.tolist():
    print('Parsing successful. All wells parsed.')
else:
    print('Parsing unsuccessful. Missing wells percentage: %s.'.format(headerFrame.index.values.tolist()/fileNoList))

bk2.save()
bk2.close()