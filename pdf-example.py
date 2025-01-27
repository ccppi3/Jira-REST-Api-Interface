import pdf

tables = pdf.Tables("tmp.pdf")
tables.selectPage(0)
listTable = tables.setTableNames(["Tabelle 1","NEUEINTRITT","Arbeitsplatzwechsel","NEUEINTRITTE"])
for table in listTable:
    tables.selectTableByObj(table)
    tables.defRows(["b","Name","A"])
    tables.parseTable()


    print("table 1")
    for i in tables.getObjectsFromTable():
        print(i)

