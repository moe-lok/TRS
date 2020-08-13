from bs4 import BeautifulSoup
import requests
import urllib.request
import re
import xml.etree.ElementTree as ET

from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys

URL1 = "http://wwmfg.analog.com/wwmfg/apps/TRS/SpecIndex.cfm"
URL2 = "http://wwmfg.analog.com/wwmfg/apps/TRS/DrawForm.cfm"
URL3 = "http://wwmfg.analog.com/userfiles/wwmfg/apps/TRS/"

trsNumber = ''
procId = []
progId = ''
projFol = ''
fixts = [[],[]]
notes = []


def extractXML(docNumber):
    global procId
    global progId
    global projFol
    global fixts
    global notes

    print("getting from XML...")

    driver = webdriver.Chrome("C:/chromedriver_win32/chromedriver.exe")
    driver.get(URL1)
    usernameInp = driver.find_element_by_name("username")
    usernameInp.clear()
    usernameInp.send_keys("Mlokman")

    passInp = driver.find_element_by_name("password")
    passInp.clear()
    passInp.send_keys("413C@Batuuban")
    passInp.send_keys(Keys.RETURN)

    select = Select(driver.find_element_by_name('FieldName'))

    # select by visible text
    select.select_by_visible_text('SpecNumber')

    trsnumberInp = driver.find_element_by_name("FieldNameString")
    trsnumberInp.clear()
    trsnumberInp.send_keys(docNumber)
    trsnumberInp.send_keys(Keys.RETURN)

    driver.find_element_by_xpath("/html/body/div[2]/form/table/tbody/tr/td/table/tbody/tr[4]/td[1]/a").click()
    driver.find_element_by_xpath("//*[@id=\"adi-sse-wwmfg-header-app-bar\"]/span/span[15]/a").click()

    mrkup = driver.find_element_by_xpath("//*[@id=\"webkit-xml-viewer-source-xml\"]/TRSForm")
    soup = BeautifulSoup(mrkup.get_attribute("innerHTML"), 'lxml')
    tree = ET.ElementTree(ET.fromstring(str(soup)))
    root = tree.getroot()

    for child in root.findall('./body/general/products/product'):
        procId.append(child.find('finishedgoodspartnumber').text)

    str1 = root.find('./body/testrequirements/verifytestsetup/specialinstructions').text.split()
    progId = str1[0]
    projFol = str1[4]


    for child in root.findall('./body/testflows/configurations/configuration/fixtures/'):
        fixts[0].append(child.find('hardware').text)
        fixts[1].append(child.find('spec').text)

    str2 = root.find('./body/testflows/configurations/configuration/notes').text
    notes = str2.splitlines()

def queryPromisParam(procId):
    print("query Promis... "+procId)



def main():
    inp = input("pls enter spec number: ")
    docNumber = str(inp[3:])

    extractXML(docNumber)
    print("trs number: "+ trsNumber)
    print("procedure Id: ")
    print(procId)
    print("progId: "+ progId)
    print("projFol: "+ projFol)
    print("fixtures: ")

    for idx,fix in enumerate(fixts[0]):
        print(fixts[0][idx]+" | "+fixts[1][idx])

    print("notes: ")

    for note in notes:
        print(note)

    #queryPromisParam("LTM4675IY#PBF-T0")


if __name__ == "__main__":
    main()