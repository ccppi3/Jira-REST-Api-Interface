import pdf

tables = pdf.Tables("downloads\\Arbeitsplatzeinteilung KW 02 06.01.2025.pdf")
tables.selectPage(0)
listTable = tables.setTableNames(["Tabelle 1","NEUEINTRITT","Arbeitsplatzwechsel","NEUEINTRITTE"])

for table in listTable:
    print("detected rows:",pdf.detectTableRows(tables.pages.selected,table))

for table in listTable:
    input()
    tables.selectTableByObj(table)
    rowList = pdf.detectTableRows(tables.pages.selected,table)
    print("rowList:",rowList)
    tables.defRows(rowList)
    tables.parseTable()


    print("table name:",table.getName())
    for i in tables.getObjectsFromTable():
        print(i)

