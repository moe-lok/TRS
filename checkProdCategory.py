import cx_Oracle

def getPartVersion(partName):
    dsn_tns = cx_Oracle.makedsn('ADGTVMODS6.ad.analog.com', '1526', service_name='p2pll')
    conn = cx_Oracle.connect(user=r'long', password='PASSWD', dsn=dsn_tns)

    c = conn.cursor()

    c.execute("""
        SELECT
            plldba.part.partname,
            plldba.part.partversion
        FROM
            plldba.part
        WHERE
            plldba.part.partname = '"""+partName+"""'
            and
            plldba.part.activekey like 'ALTM%'""") # use triple quotes if you want to spread your query across multiple lines

    tmpLst = list(c)
    conn.close()

    if not tmpLst:
        print("partname "+partName+" does not exist")
        return None

    return tmpLst[0][1]


def getProdCat(partName, partVersion):
    dsn_tns = cx_Oracle.makedsn('ADGTVMODS6.ad.analog.com', '1526', service_name='p2pll')
    conn = cx_Oracle.connect(user=r'long', password='PASSWD', dsn=dsn_tns)

    c = conn.cursor()

    c.execute("""
        SELECT
            plldba.catg.partprcdname,
            plldba.catg.partprcdversion,
            plldba.catg.category,
            plldba.catg.catgnumber
        FROM
            plldba.catg
        WHERE
            plldba.catg.partprcdname = '"""+partName+"""'
            AND plldba.catg.partprcdversion = '"""+partVersion+"""'""") # use triple quotes if you want to spread your query across multiple lines

    tmpLst = list(c)
    conn.close()

    finalList = [[],[],[]]
    for tmp in tmpLst:
        finalList[0].append(tmp[0])
        finalList[1].append(tmp[2])
        finalList[2].append(tmp[3])

    return finalList

def compareCat(partId, catgList):

    if partId[1] in catgList[1]:
        print(partId[0]+" contains "+partId[1])
    else:
        print("Missing "+partId[1]+" in "+partId[0])

def main():
    print("main")

    #partName = "LTM8027MPV#PBF-T0"
    #partVersion = getPartVersion("LTM8027MPV#PBF-T0")
    #getProdCat(partName, partVersion)

    partId = []

    while True:  # taking multiple line input for eqpId
        line = input()
        if line:
            partId.append(line)
        else:
            break

    print(partId)
    print(len(partId))

    for p in partId:
        partVersion = getPartVersion(p.split()[0])

        if not partVersion:
            continue

        catgList = getProdCat(p.split()[0],partVersion)
        compareCat(p.split(), catgList)



if __name__ == "__main__":
    main()