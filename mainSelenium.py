from bs4 import BeautifulSoup
import xml.etree.ElementTree as ET
import httplib2
import cx_Oracle
import time
import pyodbc
import difflib
import re

from pywinauto.findwindows import ElementNotFoundError
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from pywinauto.application import Application
from webdriver_manager.chrome import ChromeDriverManager

URL1 = "http://wwmfg.analog.com/wwmfg/apps/TRS/SpecIndex.cfm"
URL2 = "http://wwmfg.analog.com/wwmfg/apps/TRS/DrawForm.cfm"
URL3 = "http://wwmfg.analog.com/userfiles/wwmfg/apps/TRS/"

trsNumber = ''
procIds = []
progId = ''
projFol = ''
fixts = [[], []]
notes = []
driver = None
changes = []
procID = None
comment_notes = ''
fixt_len = 0
deletes = []
adds = []
corrs = []


# cx_Oracle.init_oracle_client(lib_dir=r"\instantclient-basiclite-windows.x64-19.6.0.0.0dbru\instantclient_19_6")

def loginTrs():
    global driver

    print("Launch Chromedriver...")

    options = Options()
    options.headless = True
    driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
    driver.get(URL1)
    print("logging in...")
    usernameInp = driver.find_element_by_name("username")
    usernameInp.clear()
    usernameInp.send_keys("Mlokman")

    passInp = driver.find_element_by_name("password")
    passInp.clear()
    passInp.send_keys("413C@Batuuban1")
    passInp.send_keys(Keys.RETURN)


def extractXML(docNumber):
    global procIds
    global progId
    global projFol
    global fixts
    global notes
    global trsNumber
    global comment_notes
    global corrs

    trsNumber = ''
    procIds = []
    progId = ''
    projFol = ''
    fixts = [[], []]
    notes = []
    corrs = []
    comment_notes = ''
    element_found = True

    print("enter trsnumber...")
    select = Select(driver.find_element_by_name('FieldName'))

    # select by visible text
    select.select_by_visible_text('SpecNumber')

    trsnumberInp = driver.find_element_by_name("FieldNameString")
    trsnumberInp.clear()
    trsnumberInp.send_keys(docNumber)
    trsnumberInp.send_keys(Keys.RETURN)

    print("click trsnumber...")

    try:
        trsLink = driver.find_element_by_xpath("/html/body/div[2]/form/table/tbody/tr/td/table/tbody/tr[4]/td[1]/a")
    except NoSuchElementException:
        print("### TRS does not exist")
        element_found = False
        return element_found

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

    str1 = root.find('./body/testrequirements/verifytestsetup/specialinstructions').text.split("|")
    progId = str1[0].strip()
    projFol = str1[1].split(":")[1].strip()

    for child in root.findall('./body/testflows/configurations/configuration/fixtures/'):
        fixts[0].append(child.find('hardware').text)
        fixts[1].append(child.find('spec').text)

    str2 = root.find('./body/testflows/configurations/configuration/notes').text
    notes = str2.splitlines()
    notes = list(filter(None, notes))  # remove empty string from list

    for child in root.findall('./body/testrequirements/specialbinningbinning/specialbinning/specialinstructions'):
        spIns = child.text
        try:
            if "SL:" in spIns:
                SLstr = re.search('SL:(.*?)\|', spIns).group(1).strip()
                corrs.append(["$SL", SLstr])
            if "CRC:" in spIns:
                CRCstr = re.search('CRC:(.*?)\|', spIns).group(1).strip()
                corrs.append(["$CRC", CRCstr])
            if "SLCORR:" in spIns:
                SLCORRstr = re.search('SLCORR:(.*?)\|', spIns).group(1).strip()
                corrs.append(["$SLCORR", SLCORRstr])
            if "CRCCORR:" in spIns:
                CRCCORRstr = re.search('CRCCORR:(.*?)\|', spIns).group(1).strip()
                corrs.append(["$CRCCORR", CRCCORRstr])
            if "GENERICCORR:" in spIns:
                GENERICCORRstr = re.search('GENERICCORR:(.*)', spIns).group(1).strip()
                corrs.append(["$GENERICCORR", GENERICCORRstr])
        except AttributeError:
            pass

    element_found = True

    try:
        comment_notes = root.find('./body/notesandattachments/notes').text
    except AttributeError:
        print("\nTRS XML Version is broken\n")
        element_found = False

    driver.execute_script("window.history.go(-1)")
    driver.execute_script("window.history.go(-1)")

    return element_found


def queryPromisParam(procId):
    print("query Promis... " + procId)

    global procID
    procID = procId

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
            MAX(plldba.prcd.prcdversion)
        FROM
            plldba.prcd
        WHERE
            plldba.prcd.prcdname = '""" + procId + """'
            GROUP BY plldba.prcd.prcdname""")  # use triple quotes if you want to spread your query across multiple lines

    tmpLst = list(c)

    try:
        procIDver = tmpLst[0][0] + tmpLst[0][1]
        conn.close()
        return procIDver

    except IndexError:
        print("procedure does not exist...")
        conn.close()
        return None


def compareParam(ppl):
    print("compare parameter...")

    global changes
    global fixt_len
    global deletes
    global adds
    changes = []
    fixt_len = 0
    deletes = []
    adds = []

    try:
        # compare TRS number $TRS
        if trsNumber != ppl[1][ppl[0].index('$TRS')]:
            print("\ntrs number not same $TRS")
            print("change " + ppl[1][ppl[0].index('$TRS')] + " to " + trsNumber)
            changes.append(['$TRS', trsNumber])
    except ValueError as e:
        print("\n$TRS does not exist in Promis :" + trsNumber)
        print(str(e))

    try:
        # compare program id $TSCLS1P1
        if progId.upper() != ppl[1][ppl[0].index('$TSCLS1P1')]:
            print("\nprogram id not same $TSCLS1P1")
            print("change " + ppl[1][ppl[0].index('$TSCLS1P1')] + " to " + progId.upper())
            changes.append(['$TSCLS1P1', progId.upper()])
    except ValueError as e:
        print("\n$TSCLS1P1 program id does not exist in Promis :" + progId.upper())
        print(str(e))

    try:
        # compare project folder $TSCLS1N1
        if projFol != ppl[1][ppl[0].index('$TSCLS1N1')]:
            print("\nproject folder not same $TSCLS1N1")
            print("change " + ppl[1][ppl[0].index('$TSCLS1N1')] + " to " + projFol)
            changes.append(['$TSCLS1N1', projFol])
    except ValueError as e:
        print("\n$TSCLS1N1 project folder does not exist in Promis :" + projFol)
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
    # if contain Performance board = PERFBRD
    # Contactor = CNTCR
    # Verifier/SUV Code = CORDEV
    # Handler Interface = HNDLRITF
    # Manual Actuator = C/ACTR
    # Handler Kit = CONVKIT
    # last should always have SOAK TIME: 240 SECS

    TRS_fixt = [[], []]

    for idx, fix in enumerate(fixts[0]):
        if "PERFORMANCE" in fix.upper():
            TRS_fixt[0].append("PERFBRD")
        if "CONTACTOR" in fix.upper():
            TRS_fixt[0].append("CNTCR")
        if "VERIFIER" in fix.upper():
            TRS_fixt[0].append("CORDEV")
        if "INTERFACE" in fix.upper():
            TRS_fixt[0].append("HNDLRITF")
        if "ACTUATOR" in fix.upper():
            TRS_fixt[0].append("C/ACTR")
        if "KIT" in fix.upper():
            TRS_fixt[0].append("CONVKIT")

        TRS_fixt[1].append(fixts[1][idx])

    for idx, trs_hard in enumerate(TRS_fixt[0]):

        fullVal = ""
        try:
            fullVal = ppl[1][ppl[0].index('$TSCLS1H1E' + str(idx + 1))]
            val = ppl[1][ppl[0].index('$TSCLS1H1E' + str(idx + 1))].split(':')[1].strip()
            paramKey = ppl[1][ppl[0].index('$TSCLS1H1E' + str(idx + 1))].split(':')[0]
            trs_spec = TRS_fixt[1][idx].strip()

            if trs_hard != paramKey or trs_spec != val:
                print("\nfixture not same $TSCLS1H1E" + str(idx + 1))
                print("change " + fullVal + " to " + trs_hard + ": " + trs_spec)
                changes.append(["$TSCLS1H1E" + str(idx + 1), trs_hard + ": " + trs_spec])

        except ValueError as e:
            print("\n$TSCLS1H1E" + str(idx + 1) + " fixture does not exist in Promis :" + TRS_fixt[1][idx])
            print(str(e))
        except IndexError:
            print("\nfixture not properly formatted $TSCLS1H1E" + str(idx + 1))
            print(fullVal + " format is not proper, code can't process")

    for key in ppl[0]:
        if '$TSCLS1H1E' in key:
            fixt_len += 1

    if fixt_len != len(TRS_fixt[0]):
        del_count = fixt_len - len(TRS_fixt[0])

        for i in range(1, del_count):
            idx = len(TRS_fixt[0]) + i
            val = ppl[1][ppl[0].index('$TSCLS1H1E' + str(idx + 1))].split(':')[1].strip()
            paramKey = ppl[1][ppl[0].index('$TSCLS1H1E' + str(idx + 1))].split(':')[0]
            fullVal = ppl[1][ppl[0].index('$TSCLS1H1E' + str(idx))]

            print("\nfixture not same $TSCLS1H1E" + str(idx))
            print("change " + fullVal + " to " + paramKey + ": " + val)
            changes.append(["$TSCLS1H1E" + str(idx), paramKey + ": " + val])

        for i in range(1, del_count):
            del_idx = fixt_len + i - 1
            print("\ndelete extra param $TSCLS1H1E" + str(del_idx))
            deletes.append("$TSCLS1H1E" + str(del_idx))

        # TODO: this is temporary, to delete later
        # deletes.extend(
        #     ["$TEMP_QC4", "$TQCLS4O1", "$TQCLS4P1", "$TQCLS4P1RS", "$TQCLS4T", "$TQCLS4TC1", "$TQCLS4_ACBIN"])

    # compare pidref $PIDREF 04-04-5430 REV A
    # notes[1] is like PIDREF: 04-04-9670 REV A|

    notes1 = notes[1].replace("|", "")
    temp = notes1.split(" ", 1)[1].split()
    pidref = " ".join(temp)

    try:
        if pidref.upper() != ppl[1][ppl[0].index('$PIDREF')]:
            print("\nPIDREF from TRS is " + notes[1])
            print("PIDREF not same $PIDREF")
            print("change " + ppl[1][ppl[0].index('$PIDREF')] + " to " + pidref)
            changes.append(["$PIDREF", pidref])

    except ValueError as e:
        print("\n$PIDREF does not exist in Promis :" + pidref)
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

    for idx, rawnote in enumerate(newNotes):

        rawnote1 = rawnote.replace("|", "")
        temp1 = rawnote1.split(" ", 1)[1].split()
        note = " ".join(temp1)

        try:
            if note.upper() != ppl[1][ppl[0].index('$MCREF' + str(idx + 1))]:
                print("\nMCREF" + str(idx + 1) + " from TRS is " + rawnote)
                print("note not same $MCREF" + str(idx + 1))
                print("change " + ppl[1][ppl[0].index('$MCREF' + str(idx + 1))] + " to " + note)
                changes.append(['$MCREF' + str(idx + 1), note])

        except ValueError as e:
            print("\n$MCREF" + str(idx + 1) + " does not exist in Promis :" + note)
            adds.append(["$MCREF" + str(idx + 1), note])
            print(str(e))

    for corr in corrs:
        param = corr[0]
        val = corr[1]

        try:
            if val != ppl[1][ppl[0].index(param)]:
                print(f"\n{param} from TRS is {val}")
                print("change " + ppl[1][ppl[0].index(param)] + " to " + val)
                changes.append(corr)

        except ValueError as e:
            print(f"{param} does not exist in Promis :{val}")
            adds.append(corr)

    try:
        # compare owner
        if "P2LOKMAN" != ppl[1][ppl[0].index('$OWNER')]:
            print("\nowner not P2LOKMAN $OWNER")
            print("change " + ppl[1][ppl[0].index('$OWNER')] + " to " + "P2LOKMAN")
            changes.append(['$OWNER', 'P2LOKMAN'])

    except ValueError as e:
        print("\n$OWNER does not exist in Promis")
        print(str(e))

    # TODO: below code is temporary remove once done
    # changes.extend([['$TEMP_QC3', 'SAMPLING'], ['$TQCLS3T', '125'], ['$TQCLS3TC1', 'QA HOT1']])


def updatePromisProductCore(procId):
    print("update promis product core...")
    print(procId)
    p2title1 = "adp2prom1.ad.analog.com - PuTTY"
    p2title2 = "ADP2PROM1.AD.ANALOG.COM - PuTTY"
    # ADP2PROM1.AD.ANALOG.COM
    # adp2prom1.ad.analog.com

    procActive = getProcActiveVer(procId)
    ppl = queryPromisParam(procActive)

    print("procActive: " + procActive)

    productCore = re.split('#', procId)[0]

    print("productCore: " + productCore)

    PECNNumber = "PPN0054858"

    updates = []

    try:
        # compare ProductCore number $$PRODUCTCORE
        if productCore != ppl[1][ppl[0].index('$PRODUCTCORE')]:
            print("\nProdcutCore not same $PRODUCTCORE")
            print("change " + ppl[1][ppl[0].index('$PRODUCTCORE')] + " to " + productCore)
            updates.append(['$PRODUCTCORE', productCore])
    except ValueError as e:
        print("\n$PRODUCTCORE does not exist in Promis :" + productCore)
        print(str(e))

    try:
        # compare ECN number $ECN
        if PECNNumber != ppl[1][ppl[0].index('$ECN')]:
            print("\nECN not same $ECN")
            print("change " + ppl[1][ppl[0].index('$ECN')] + " to " + PECNNumber)
            updates.append(['$ECN', PECNNumber])
    except ValueError as e:
        print("\n$ECN does not exist in Promis :" + PECNNumber)
        print(str(e))

    try:
        # compare owner
        if "P2LOKMAN" != ppl[1][ppl[0].index('$OWNER')]:
            print("\nowner not P2LOKMAN $OWNER")
            print("change " + ppl[1][ppl[0].index('$OWNER')] + " to " + "P2LOKMAN")
            updates.append(['$OWNER', 'P2LOKMAN'])

    except ValueError as e:
        print("\n$OWNER does not exist in Promis")
        print(str(e))

    print("updates")
    print(updates)

    if updates and True if input("\nAuto update? [Y/N]...")[0].upper() == "Y" else False:

        try:
            app = Application(backend="uia").connect(title=p2title1)
        except ElementNotFoundError:
            app = Application(backend="uia").connect(title=p2title2)

        dialog = app.window()

        temp = procActive.split(".")

        procIDver = temp[0] + "." + str(int(temp[1]) + 1).zfill(2)

        dialog.type_keys("^z")  # Ctrl+z
        dialog.type_keys("q ma proc create{ENTER}", with_spaces=True)
        time.sleep(0.5)
        dialog.type_keys(procId + "{ENTER}")
        time.sleep(0.5)
        dialog.type_keys("y{ENTER}")
        time.sleep(0.5)
        dialog.type_keys("y{ENTER}")
        time.sleep(0.5)
        dialog.type_keys(procId + "{ENTER}")
        time.sleep(0.5)
        dialog.type_keys("y{ENTER}")
        time.sleep(0.5)
        dialog.type_keys("param{ENTER}")
        time.sleep(0.5)
        dialog.type_keys("m{ENTER}")

        for update in updates:
            dialog.type_keys(update[0] + "{ENTER}")
            time.sleep(0.5)
            dialog.type_keys("{ENTER}")
            time.sleep(0.5)
            dialog.type_keys(update[1] + "{ENTER}", with_spaces=True)
            time.sleep(0.5)

        dialog.type_keys("END{ENTER}")
        time.sleep(0.5)
        dialog.type_keys("{ENTER}")
        time.sleep(0.5)

        dialog.type_keys("q ma eng ass proc{ENTER}", with_spaces=True)
        time.sleep(0.5)

        # procedure id version
        dialog.type_keys(procIDver + "{ENTER}")
        time.sleep(0.5)
        # trsnumber
        dialog.type_keys(PECNNumber + "{ENTER}")
        time.sleep(0.5)
        # end
        dialog.type_keys("END{ENTER}")
        time.sleep(0.5)

        # set nopage
        dialog.type_keys("set nopage{ENTER}", with_spaces=True)
        time.sleep(0.5)
        # q ma proc free
        dialog.type_keys("q ma proc free{ENTER}", with_spaces=True)
        time.sleep(0.5)
        # procedure id version
        dialog.type_keys(procIDver + "{ENTER}")
        time.sleep(0.5)
        # y
        dialog.type_keys("y{ENTER}")
        time.sleep(0.5)
        # y
        dialog.type_keys("y{ENTER}")
        time.sleep(0.5)
        # make
        dialog.type_keys("make{ENTER}")
        time.sleep(0.5)
        # procedure id version
        dialog.type_keys(procIDver + "{ENTER}")
        time.sleep(0.5)
        dialog.type_keys("set page{ENTER}", with_spaces=True)


def updatePromisParam(procId):
    print("update promis param...")
    print(procId)
    p2title1 = "adp2prom1.ad.analog.com - PuTTY"
    p2title2 = "ADP2PROM1.AD.ANALOG.COM - PuTTY"
    # ADP2PROM1.AD.ANALOG.COM
    # adp2prom1.ad.analog.com

    try:
        app = Application(backend="uia").connect(title=p2title1)
    except ElementNotFoundError:
        app = Application(backend="uia").connect(title=p2title2)

    dialog = app.window()

    temp = procID.split(".")

    procIDver = temp[0] + "." + str(int(temp[1]) + 1).zfill(2)

    print(changes)
    dialog.type_keys("^z")  # Ctrl+z
    dialog.type_keys("q ma proc create{ENTER}", with_spaces=True)
    time.sleep(0.5)
    dialog.type_keys(procId + "{ENTER}")
    time.sleep(0.5)
    dialog.type_keys("y{ENTER}")
    time.sleep(0.5)
    dialog.type_keys("y{ENTER}")
    time.sleep(0.5)
    dialog.type_keys(procId + "{ENTER}")
    time.sleep(0.5)
    dialog.type_keys("y{ENTER}")
    time.sleep(0.5)
    dialog.type_keys("param{ENTER}")
    time.sleep(0.5)
    dialog.type_keys("m{ENTER}")

    for change in changes:
        dialog.type_keys(change[0] + "{ENTER}")
        time.sleep(0.5)
        dialog.type_keys("{ENTER}")
        time.sleep(0.5)
        dialog.type_keys(change[1] + "{ENTER}", with_spaces=True)
        time.sleep(0.5)

    dialog.type_keys("END{ENTER}")
    time.sleep(0.5)
    dialog.type_keys("{ENTER}")
    time.sleep(0.5)

    for delete in deletes:
        dialog.type_keys("param{ENTER}")
        time.sleep(0.2)
        dialog.type_keys("d{ENTER}")
        time.sleep(0.2)
        dialog.type_keys(delete + "{ENTER}")
        time.sleep(0.5)
        dialog.type_keys("y{ENTER}")
        time.sleep(0.5)
        dialog.type_keys("END{ENTER}")
        time.sleep(0.2)
        dialog.type_keys("{ENTER}")
        time.sleep(0.2)

    for add in adds:
        dialog.type_keys("param{ENTER}")
        time.sleep(0.2)
        dialog.type_keys("a{ENTER}")
        time.sleep(0.2)
        dialog.type_keys(add[0] + "{ENTER}")
        time.sleep(0.5)
        dialog.type_keys("STRING{ENTER}")
        time.sleep(0.5)
        dialog.type_keys(add[1] + "{ENTER}", with_spaces=True)
        time.sleep(0.5)
        dialog.type_keys("END{ENTER}")
        time.sleep(0.2)
        dialog.type_keys("{ENTER}")
        time.sleep(0.2)

    dialog.type_keys("q ma eng ass proc{ENTER}", with_spaces=True)
    time.sleep(0.5)

    # procedure id version
    dialog.type_keys(procIDver + "{ENTER}")
    time.sleep(0.5)
    # trsnumber
    dialog.type_keys(trsNumber + "{ENTER}")
    time.sleep(0.5)
    # end
    dialog.type_keys("END{ENTER}")
    time.sleep(0.5)
    # set nopage
    dialog.type_keys("set nopage{ENTER}", with_spaces=True)
    time.sleep(0.5)
    # q ma proc free
    dialog.type_keys("q ma proc free{ENTER}", with_spaces=True)
    time.sleep(0.5)
    # procedure id version
    dialog.type_keys(procIDver + "{ENTER}")
    time.sleep(0.5)
    # y
    dialog.type_keys("y{ENTER}")
    time.sleep(0.5)
    # y
    dialog.type_keys("y{ENTER}")
    time.sleep(0.5)
    # make
    dialog.type_keys("make{ENTER}")
    time.sleep(0.5)
    # procedure id version
    dialog.type_keys(procIDver + "{ENTER}")
    time.sleep(0.5)
    dialog.type_keys("set page{ENTER}", with_spaces=True)


def insert_into_PIDComments(product_core, cmnt_notes, part_type, slflow):
    # get pidref
    notes1 = notes[1].replace("|", "")
    temp = notes1.split(" ", 1)[1].split()
    pidref = " ".join(temp)

    sqlconn = pyodbc.connect('Driver={SQL Server};'
                             'Server=adpgsql1\\adpgsql;'
                             'Database=MIPS;'
                             'UID=webUser;'
                             'PWD=Adipg!234567890;')

    msSql_insert_query = f"\
                                INSERT INTO [dbo].[PIDComments]\
                                           ([ProductCore]\
                                           ,[PidReference]\
                                           ,[CommentsNotes]\
                                           ,[PartType]\
                                           ,[SLFLOW]\
                                           ,[ECN]\
                                           ,[CreatedDate])\
                                     VALUES\
                                           ('{product_core}'\
                                           ,'{pidref}'\
                                           ,'{cmnt_notes}'\
                                           ,'{part_type}'\
                                           ,'{slflow}'\
                                           ,'{trsNumber}'\
                                           ,GETDATE()\
                                           );"

    cursor2 = sqlconn.cursor()
    cursor2.execute(msSql_insert_query)
    sqlconn.commit()
    records = cursor2.rowcount
    print(records, " rows inserted")
    cursor2.close()


def compare_comment_notes(procId):
    print("\ncomparing comment notes from database...")

    procActive = getProcActiveVer(procId)
    ppl = queryPromisParam(procActive)

    product_core = procId.split("#")[0]
    part_type = 'T'

    try:
        part_type = ppl[1][ppl[0].index('$PARTTYPE')]
    except ValueError as e:
        print("$PARTTYPE does not exist in in Promis")

    try:
        slflow = re.search('#(.*)PBF', procId).group(1)
    except AttributeError:
        print("procedure id is not in #PBF format: " + procId + "\nskipping this procedure...")
        return

    slflow = slflow if slflow else "NA"

    sqlconn = pyodbc.connect('Driver={SQL Server};'
                             'Server=adpgsql1\\adpgsql;'
                             'Database=MIPS;'
                             'UID=webUser;'
                             'PWD=Adipg!234567890;')

    sql_query = f"SELECT " \
                f"[ProductCore]," \
                f"[CommentsNotes]" \
                f"FROM [MIPS].[dbo].[PIDComments]" \
                f"WHERE ProductCore = '{product_core}'" \
                f"AND PartType = '{part_type}'" \
                f"AND SLFLOW = '{slflow}';"

    cursor = sqlconn.cursor()
    cursor.execute(sql_query)

    row = cursor.fetchone()
    cursor.close()

    cmnt_notes = comment_notes.replace('<br>', '\n')

    if row:

        com_note1 = row.CommentsNotes.splitlines()
        com_note2 = cmnt_notes.splitlines()

        d = difflib.Differ()

        both_are_same = True

        for idx, line in enumerate(com_note1):
            if line != com_note2[idx]:
                both_are_same = False
                diff = d.compare(com_note1, com_note2)
                print('\n'.join(diff))
                break

        if not both_are_same:
            print("\nComment notes are not the same...\n")
            print("product core: " + product_core + "\nPartType:" + part_type + "\nslflow: " + slflow)

            if True if input("\nAuto update comment? [Y/N]...")[0].upper() == "Y" else False:
                # update comment in database
                msSql_update_query = f"\
                    UPDATE [MIPS].[dbo].[PIDComments] \
                    SET CommentsNotes = '{cmnt_notes}' \
                    WHERE ProductCore = '{product_core}' \
                    AND PartType = '{part_type}' \
                    AND SLFLOW = '{slflow}';"

                cursor1 = sqlconn.cursor()
                cursor1.execute(msSql_update_query)
                sqlconn.commit()
                records = cursor1.rowcount
                print(records, " rows updated")
                cursor1.close()

        else:
            print("\nComment notes are same, please proceed...")

    else:
        print("\nthere are no comments notes in database\n")
        # insert new PIDComments
        print("product core: " + product_core + "\nPartType:" + part_type + "\nslflow: " + slflow)
        if True if input("\nAuto insert? [Y/N]...")[0].upper() == "Y" else False:
            insert_into_PIDComments(product_core, cmnt_notes, part_type, slflow)

    sqlconn.close()


def main():
    loginTrs()

    cont = True

    while cont:

        inp = input("pls enter spec number: ")
        docNumber = str(inp[3:])

        if extractXML(docNumber):
            print("trs number: " + trsNumber)
            print("procedure Id: ")
            print(procIds)
            print("progId: " + progId)
            print("projFol: " + projFol)
            print("fixtures: ")

            for idx, fix in enumerate(fixts[0]):
                print(fixts[0][idx] + " | " + fixts[1][idx])

            print("notes: ")

            for note in notes:
                print(note)

            for procId in procIds:
                print("\n############# " + procId + "-T0")

                procActVer = getProcActiveVer(procId + "-T0")

                if procActVer:

                    promisParamList = queryPromisParam(procActVer)
                    compareParam(promisParamList)

                    # update promis parameter
                    if changes or deletes or adds:
                        print("\nparam to update...")
                        print(changes)

                        if deletes:
                            print("\nparam to delete...")
                            print(deletes)

                        if adds:
                            print("\nparam to add...")
                            print(adds)

                        while (res := input("\nAuto update? [Y/N]...").lower()) not in {"y", "n"}: pass

                        if res == 'y':
                            updatePromisParam(procId + "-T0")
                    else:
                        print("there are no changes to make...")

                else:
                    print("fail to get procedure active version...")

                # check for Sp inst and auto update
                if comment_notes:
                    compare_comment_notes(procId)
                    compare_comment_notes(procId + "-T0")
                else:
                    print("no comment notes")

        cont = True if input("\nContinue? [Y/N]...")[0].upper() == "Y" else False


def updateProcedureProductCore():
    partNames = []

    print("paste partName here, ENTER if done: ")

    while True:  # taking multiple line input for Lot Id
        line = input()
        if line:
            partNames.append(line)
        else:
            break

    for partName in partNames:
        updatePromisProductCore(partName)


if __name__ == "__main__":
    main()
