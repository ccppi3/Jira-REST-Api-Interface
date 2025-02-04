import pdf
from pop3 import err as err

pdf.setDebugLevel(err.NONE)
tables = pdf.Tables("tmp.pdf")
tables.selectPage(0)
listTable = tables.setTableNames(["Tabelle 1","NEUEINTRITT","Arbeitsplatzwechsel","NEUEINTRITTE"])
for table in listTable:
    tables.selectTableByObj(table)
    tables.defRows(["Abteilung vorher","Name","A"])
    tables.parseTable()


    print("table 1")
    for x,i in enumerate(tables.getObjectsFromTable()):
        print(i)
        print("all attr:",dir(i))
        if "Abteilung vorher" in dir(i):
            print("getattr:",getattr(i,"Abteilung vorher"))

        try:
            print("getattr:",getattr(i,"Abteilung vorher"))
        except:
            print("attr not available")



