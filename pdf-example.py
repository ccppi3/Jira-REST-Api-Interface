import pdf
from pop3 import err as err

tables = pdf.Tables("test4.pdf")
pdf.setDebugLevel(err.ERROR)
tables.selectPage(0)
listTable = tables.setTableNames(["Tabelle 1","NEUEINTRITT","Arbeitsplatzwechsel","NEUEINTRITTE"])
for table in listTable:
    tables.selectTableByObj(table)
    tables.defRows(["b","Name","A"])
    tables.parseTable()


    print("table 1")
    for i in tables.getObjectsFromTable():
        print(i)

