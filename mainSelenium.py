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

trsNumber = ''
procIds = []
progId = ''
projFol = ''
fixts = [[],[]]
notes = []

cx_Oracle.init_oracle_client(lib_dir=r"\instantclient-basiclite-windows.x64-19.6.0.0.0dbru\instantclient_19_6")


def extractXML(docNumber):
    global procIds
    global progId
    global projFol
    global fixts
    global notes
    global trsNumber

    print("Launch Chromedriver...")

    options = Options()
    options.headless = True
    driver = webdriver.Chrome("/chromedriver_win32/chromedriver.exe", options=options)
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
    # get trs number with revision
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
        procIds.append(child.find('finishedgoodspartnumber').text)

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

    try:
        # compare TRS number $TRS
        if (trsNumber != ppl[1][ppl[0].index('$TRS')]):
            print("\ntrs number not same $TRS")
            print("change "+ ppl[1][ppl[0].index('$TRS')] + " to " + trsNumber)
    except ValueError as e:
        print("\n$TRS does not exist in Promis")
        print(str(e))

    try:
        # compare program id $TSCLS1P1
        if (progId.upper() != ppl[1][ppl[0].index('$TSCLS1P1')]):
            print("\nprogram id not same $TSCLS1P1")
            print("change "+ ppl[1][ppl[0].index('$TSCLS1P1')] + " to " + progId.upper())
    except ValueError as e:
        print("\n$TSCLS1P1 program id does not exist in Promis")
        print(str(e))

    try:
        # compare project folder $TSCLS1N1
        if (projFol != ppl[1][ppl[0].index('$TSCLS1N1')]):
            print("\nproject folder not same $TSCLS1N1")
            print("change "+ ppl[1][ppl[0].index('$TSCLS1N1')] + " to " + projFol)
    except ValueError as e:
        print("\n$TSCLS1N1 project folder does not exist in Promis")
        print(str(e))

    # compare fixture
    """
    $TSCLS1H1E1 PERFBRD: L-65180
    $TSCLS1H1E2 CNTCR: PGC-0020
    $TSCLS1H1E3 CORDEV: PGR-0053
    $TSCLS1H1E4 C/ACTR: PGA-0033
    $TSCLS1H1E5 HNDLRITF: LTC-00282
    $TSCLS1H1E6 CONVKIT: PGK-0052
    """

    for idx,fix in enumerate(fixts[0]):

        try:
            val = ppl[1][ppl[0].index('$TSCLS1H1E'+str(idx+1))].split()[1]

            if (fixts[1][idx] != val):
                print("\nfixture not same $TSCLS1H1E"+str(idx+1))
                print("change "+ val + " to " + fixts[1][idx])

        except ValueError as e:
            print("\n$TSCLS1H1E"+str(idx+1)+" fixture does not exist in Promis")
            print(str(e))


    # compare pidref $PIDREF 04-04-5430 REV A
    temp = notes[1].split(" ",1)[1].split()
    del temp[-1]
    pidref = " ".join(temp)

    try:
        if (pidref.upper() != ppl[1][ppl[0].index('$PIDREF')]):
            print("\nPIDREF not same $PIDREF")
            print("change "+ ppl[1][ppl[0].index('$PIDREF')] + " to " + pidref)

    except ValueError as e:
        print("\n$PIDREF does not exist in Promis")
        print(str(e))

    newNotes = notes.copy()

    del newNotes[0]
    del newNotes[0]

    # compare notes
    """
    $MCREF1 04-10-26420 REV 0
    $MCREF2 04-10-26421 REV B
    $MCREF3 04-10-26422 REV B
    """

    for idx,rawnote in enumerate(newNotes):
        temp1 = rawnote.split(" ",1)[1].split()
        del temp1[-1]
        note = " ".join(temp1)

        try:
            if (note.upper() != ppl[1][ppl[0].index('$MCREF'+str(idx+1))]):
                print("\nnote not same $MCREF"+str(idx+1))
                print("change "+ ppl[1][ppl[0].index('$MCREF'+str(idx+1))] + " to " + note)

        except ValueError as e:
            print("\n$MCREF"+str(idx+1)+" does not exist in Promis")
            print(str(e))


    try:
        # compare owner
        if ("P2LOKMAN" != ppl[1][ppl[0].index('$OWNER')]):
            print("\nowner not P2LOKMAN $OWNER")
            print("change "+ ppl[1][ppl[0].index('$OWNER')] + " to " + "P2LOKMAN")

    except ValueError as e:
        print("\n$OWNER does not exist in Promis")
        print(str(e))


def main():
    inp = input("pls enter spec number: ")
    docNumber = str(inp[3:])


    extractXML(docNumber)
    print("trs number: "+ trsNumber)
    print("procedure Id: ")
    print(procIds)
    print("progId: "+ progId)
    print("projFol: "+ projFol)
    print("fixtures: ")

    for idx,fix in enumerate(fixts[0]):
        print(fixts[0][idx]+" | "+fixts[1][idx])

    print("notes: ")

    for note in notes:
        print(note)


    for procId in procIds:
        print("\n############# "+ procId+"-T0")
        promisParamList = queryPromisParam(getProcActiveVer(procId+"-T0"))
        compareParam(promisParamList)


if __name__ == "__main__":
    main()
    input("\npress any key to close...")