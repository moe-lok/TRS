from bs4 import BeautifulSoup
import requests
import urllib.request
import re
import xml.etree.ElementTree as ET

URL1 = "http://wwmfg.analog.com/wwmfg/apps/TRS/SpecIndex.cfm"
URL2 = "http://wwmfg.analog.com/wwmfg/apps/TRS/DrawForm.cfm"
URL3 = "http://wwmfg.analog.com/userfiles/wwmfg/apps/TRS/"

trsNumber = ''
procId = []
progId = ''
projFol = ''
fixts = [[],[]]
notes = []


def trsSearch1(docNumber):
    global trsNumber
    print("getDocId...")

    payload = {
      'Status': 'All',
      'PlanningSite': 'All',
      'TestSite': 'All',
      'FieldName': 'SpecNumber',
      'FieldNameString': docNumber,
      'COMMAND': ''
    }
    headers = {
      'Host': 'wwmfg.analog.com',
      'Connection': 'keep-alive',
      'Content-Length': '93',
      'Cache-Control': 'no-cache',
      'Upgrade-Insecure-Requests': '1',
      'Origin': 'http://wwmfg.analog.com',
      'Content-Type': 'application/x-www-form-urlencoded',
      'User-Agent': 'PostmanRuntime/7.26.2',
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
    print("##### doc id")
    print(docId)

    return docId


def extractXML(trsNum, docId):
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
      'Referer': 'http://wwmfg.analog.com/wwmfg/apps/TRS/DrawForm.cfm?DocId='+docId,
      'Accept-Encoding': 'gzip, deflate',
      'Accept-Language': 'en-US,en;q=0.9',
      'Cookie': 'WWMFGSESSION=637326059812533495; CFID=7972988; CFTOKEN=10229470'
    }

    response = requests.request("GET", URL3+trsNum+".xml", headers=headers)
    soup = BeautifulSoup(response.content, 'lxml')
    tree = ET.ElementTree(ET.fromstring(str(soup)))
    root = tree.getroot()


    for child in root.findall('./body/trsform/general/products/product'):
        procId.append(child.find('finishedgoodspartnumber').text)

    str1 = root.find('./body/trsform/testrequirements/verifytestsetup/specialinstructions').text.split()
    progId = str1[0]
    projFol = str1[4]


    for child in root.findall('./body/trsform/testflows/configurations/configuration/fixtures/'):
        fixts[0].append(child.find('hardware').text)
        fixts[1].append(child.find('spec').text)

    str2 = root.find('./body/trsform/testflows/configurations/configuration/notes').text
    notes = str2.splitlines()

def queryPromisParam(procId):
    print("query Promis... "+procId)



def main():
    inp = input("pls enter spec number: ")
    docNumber = str(inp[3:])

    docId = trsSearch1(docNumber)
    #trsSearch2(docId)
    extractXML(trsNumber, docId)
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