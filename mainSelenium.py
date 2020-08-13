from bs4 import BeautifulSoup
import requests
import urllib.request
import re
import xml.etree.ElementTree as ET
import httplib2
import cx_Oracle
import time

from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options

URL1 = "http://wwmfg.analog.com/wwmfg/apps/TRS/SpecIndex.cfm"
URL2 = "http://wwmfg.analog.com/wwmfg/apps/TRS/DrawForm.cfm"
URL3 = "http://wwmfg.analog.com/userfiles/wwmfg/apps/TRS/"

trsNumber = 'TRS024374C'
procId = []
progId = ''
projFol = ''
fixts = [[],[]]
notes = []

cx_Oracle.init_oracle_client(lib_dir=r"C:\instantclient-basiclite-windows.x64-19.6.0.0.0dbru\instantclient_19_6")


def extractXML(docNumber):
    global procId
    global progId
    global projFol
    global fixts
    global notes
    global trsNumber

    print("Launch Chromedriver...")

    options = Options()
    options.headless = True
    driver = webdriver.Chrome("C:/chromedriver_win32/chromedriver.exe", options=options)
    driver.get(URL1)
    print("logging in...")
    usernameInp = driver.find_element_by_name("username")
    usernameInp.clear()
    usernameInp.send_keys("Mlokman")

    passInp = driver.find_element_by_name("password")
    passInp.clear()
    passInp.send_keys("413C@Batuuban")
    passInp.send_keys(Keys.RETURN)

    print("enter trsnumber...")
    select = Select(driver.find_element_by_name('FieldName'))

    # select by visible text
    select.select_by_visible_text('SpecNumber')

    trsnumberInp = driver.find_element_by_name("FieldNameString")
    trsnumberInp.clear()
    trsnumberInp.send_keys(docNumber)
    trsnumberInp.send_keys(Keys.RETURN)

    print("click trsnumber...")
    trsLink = driver.find_element_by_xpath("/html/body/div[2]/form/table/tbody/tr/td/table/tbody/tr[4]/td[1]/a")
    # TODO; get trs number with revision
    trsNumber = trsLink.text
    trsLink.click()

    print("click XML...")
    driver.find_element_by_xpath("//*[@id=\"adi-sse-wwmfg-header-app-bar\"]/span/span[15]/a").click()

    print("extract XML...")
    mrkup = driver.find_element_by_tag_name("TRSForm")
    # //*[@id=\"webkit-xml-viewer-source-xml\"]/TRSForm
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

    HTTPURL = "ADP2PROM1.AD.ANALOG.COM:10736"
    TP_Function = "/PROQUERY_PPARINFO"

    httpobj = httplib2.Http()
    header = {"Content-Type": "text/plain;charset=UTF-8"}
    params = r"\USERID p2lokman\PWD analog1234\PRCDID " + procId + r"\FROM PPAR.PARAMETERS\SHOW PARMNAME\SHOW PARMVAL\END\\"
    #params = r"\USERID p2lokman\PWD analog1234\PRCDID " + procId + r"\FROM PPAR.PARAMETERS\SHOW PARMNAME\SHOW PARMVAL\END\\"
    uri = "http://" + HTTPURL + TP_Function
    content = httpobj.request(uri, "POST", headers=header, body=params)
    lotinfo = content[1].decode("utf-8")
    sliced = lotinfo.split("\\")

    del sliced[-1]
    del sliced[0]
    del sliced[0]

    splitList = [sliced[i::2] for i in range(2)]

    return splitList

def getProcActiveVer(procId):
    print("getting procedure active version on database...")

    dsn_tns = cx_Oracle.makedsn('ADGTVMODS6.ad.analog.com', '1526', service_name='p2pll')
    conn = cx_Oracle.connect(user=r'long', password='PASSWD', dsn=dsn_tns)

    c = conn.cursor()
    c.execute("""
        SELECT
            plldba.prcd.prcdname,
            plldba.prcd.prcdversion,
            plldba.prcd.activekey,
            plldba.prcd.ecn,
            plldba.prcd.prodstatus,
            plldba.prcd.engowner,
            to_char(plldba.prcd.createdate, 'DD-MON-YYYY HH24:Mi:SS') AS createddate,
            to_char(plldba.prcd.activedate, 'DD-MON-YYYY HH24:Mi:SS') AS activedate
        FROM
            plldba.prcd
        WHERE
            plldba.prcd.prcdname = '"""+procId+"""'
            and
            plldba.prcd.activekey like 'ALTM%'""") # use triple quotes if you want to spread your query across multiple lines

    tmpLst = list(c)
    procIDver = tmpLst[0][0]+tmpLst[0][1]

    return procIDver
    conn.close()


def compareParam(ppl):
    print("compare parameter...")
    print(ppl)

    """for idx,s in enumerate(ppl[0]):
        print(s,ppl[1][idx])"""

    # TODO: compare TRS number $TRS
    if (trsNumber != ppl[1][ppl[0].index('$TRS')]):
        print("trs number not same")
        print("change "+ ppl[1][ppl[0].index('$TRS')] + " to " + trsNumber)

    # TODO: compare program id $TSCLS1P1
    # TODO: compare project folder $TSCLS1N1
    # TODO: compare fixture
    """
    $TSCLS1H1E1 PERFBRD: L-65180
    $TSCLS1H1E2 CNTCR: PGC-0020
    $TSCLS1H1E3 CORDEV: PGR-0053
    $TSCLS1H1E4 C/ACTR: PGA-0033
    $TSCLS1H1E5 HNDLRITF: LTC-00282
    $TSCLS1H1E6 CONVKIT: PGK-0052
    """
    # TODO: compare notes
    """
    $MCREF1 04-10-26420 REV 0
    $MCREF2 04-10-26421 REV B
    $MCREF3 04-10-26422 REV B
    """
    # TODO: compare pidref $PIDREF 04-04-5430 REV A
    # TODO: compare owner


def main():
    #inp = input("pls enter spec number: ")
    #docNumber = str(inp[3:])

    """extractXML(docNumber)
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
        print(note)"""

    promisParamList = queryPromisParam(getProcActiveVer("LTM2884IY#PBF-T0"))
    compareParam(promisParamList)


if __name__ == "__main__":
    main()