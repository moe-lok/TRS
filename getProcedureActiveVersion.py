import cx_Oracle

procList = []


def main():
    print("paste list of procedure here...")

    while True:  # taking multiple line input for tap stds
        line = input()
        if line:
            procList.append(line)
        else:
            break

    print("total procedure is " + str(len(procList)))

    for proc in procList:
        procActive = getProcActiveVer(proc + "-T0")
        print(procActive if procActive else "NA")

def getProcActiveVer(procId):

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


if __name__ == "__main__":
    main()