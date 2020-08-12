from bs4 import BeautifulSoup
import requests
import re
import xml.etree.ElementTree as ET

URL1 = "http://wwmfg.analog.com/wwmfg/apps/TRS/SpecIndex.cfm"
URL2 = "http://wwmfg.analog.com/wwmfg/apps/TRS/DrawForm.cfm"
URL3 = "http://wwmfg.analog.com/userfiles/wwmfg/apps/TRS/"

trsNumber = ''
procId = ''
progId = ''
projFol = ''
fixts = [[],[]]
notes = []


def trsSearch1():
    global trsNumber
    print("getDocId...")

    payload = {
      'Status': 'All',
      'PlanningSite': 'All',
      'TestSite': 'All',
      'FieldName': 'SpecNumber',
      'FieldNameString': '024026',
      'COMMAND': ''
    }
    headers = {
      'Host': 'wwmfg.analog.com',
      'Connection': 'keep-alive',
      'Content-Length': '93',
      'Cache-Control': 'max-age=0',
      'Upgrade-Insecure-Requests': '1',
      'Origin': 'http://wwmfg.analog.com',
      'Content-Type': 'application/x-www-form-urlencoded',
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.105 Safari/537.36',
      'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
      'Referer': 'http://wwmfg.analog.com/wwmfg/apps/TRS/SpecIndexAction.cfm',
      'Accept-Encoding': 'gzip, deflate',
      'Accept-Language': 'en-US,en;q=0.9',
      'Cookie': 'WWMFGSESSION=637326059812533495; CFID=7972988; CFTOKEN=10229470',
      'Postman-Token': '11ddfa9a-8d47-4b01-a1db-7553dc165259'
    }

    params = {
      'FlipOrder':'',
      'CFID':'7972988',
      'CFTOKEN':'10229470'}

    page = requests.post(URL1, data=payload, headers=headers, params=params)
    soup = BeautifulSoup(page.content, 'html.parser')

    mrkup = soup.select('a[href^="./DrawForm.cfm?DocId="]') # find the desired element (DrawForm)
    str = mrkup[0]['href'] # get the href only
    trsNumber = mrkup[0].string # get trs number

    docId = str[str.find("DocId=")+6:] # get doc id substring

    return docId

def trsSearch2(docId):
    global procId
    global progId
    print("scraping...")

    payload = {'DocId':docId}

    headers = {
      'Host': 'wwmfg.analog.com',
      'Connection': 'keep-alive',
      'Upgrade-Insecure-Requests': '1',
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.105 Safari/537.36',
      'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
      'Referer': 'http://wwmfg.analog.com/wwmfg/apps/TRS/SpecIndex.cfm?FlipOrder=&CFID=7972988&CFTOKEN=10229470',
      'Accept-Encoding': 'gzip, deflate',
      'Accept-Language': 'en-US,en;q=0.9',
      'Cookie': 'WWMFGSESSION=637326059812533495; CFID=7972988; CFTOKEN=10229470'
    }

    # search for procedure id
    page1 = requests.get(URL2, params=payload, headers=headers)
    soup1 = BeautifulSoup(page1.content, 'html.parser')
    mrkup = soup1.select_one("body > div:nth-of-type(2) > form > table > tr > td > table > tr:nth-of-type(8) > td > table > tr:nth-of-type(3) > td:nth-of-type(4)")
    procId = mrkup.string.strip()

    # search for program id
    mrkup1 = soup1.find(string=re.compile("Special Instructions:")).next_sibling.b
    temp = mrkup1.string.split()
    progId = temp[0]

    # search for fixture
    table_body = soup1.find(string=re.compile("Fixture")).parent.parent.parent.parent
    rows = table_body.find_all('tr')
    print(rows[0].prettify())


def extractXML(trsNumber):
    global procId
    global progId
    global projFol
    global fixts
    global notes

    print("getting from XML...")

    headers = {
      'Host': 'wwmfg.analog.com',
      'Connection': 'keep-alive',
      'Pragma': 'no-cache',
      'Cache-Control': 'no-cache',
      'Upgrade-Insecure-Requests': '1',
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.105 Safari/537.36',
      'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
      'Referer': 'http://wwmfg.analog.com/wwmfg/apps/TRS/DrawForm.cfm?DocId=82054',
      'Accept-Encoding': 'gzip, deflate',
      'Accept-Language': 'en-US,en;q=0.9',
      'Cookie': 'WWMFGSESSION=637326059812533495; CFID=7972988; CFTOKEN=10229470'
    }

    response = requests.request("GET", URL3+trsNumber+".xml", headers=headers)
    soup = BeautifulSoup(response.content, 'lxml')
    tree = ET.ElementTree(ET.fromstring(str(soup)))
    root = tree.getroot()


    for child in root.findall('./body/trsform/general/products/product'):
        procId = child.find('finishedgoodspartnumber').text

    str1 = root.find('./body/trsform/testrequirements/verifytestsetup/specialinstructions').text.split()
    progId = str1[0]
    projFol = str1[4]


    for child in root.findall('./body/trsform/testflows/configurations/configuration/fixtures/'):
        fixts[0].append(child.find('hardware').text)
        fixts[1].append(child.find('spec').text)

    str2 = root.find('./body/trsform/testflows/configurations/configuration/notes').text
    notes = str2.splitlines()

def main():
    docId = trsSearch1()
    #trsSearch2(docId)
    extractXML(trsNumber)
    print("trs number: "+ trsNumber)
    print("procedure Id: "+ procId)
    print("progId: "+ progId)
    print("projFol: "+ projFol)
    print("fixtures: ")

    for idx,fix in enumerate(fixts[0]):
        print(fixts[0][idx]+" | "+fixts[1][idx])

    print("notes: ")

    for note in notes:
        print(note)


if __name__ == "__main__":
    main()