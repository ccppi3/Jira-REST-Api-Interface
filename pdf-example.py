import pdf

tables = pdf.Tables("downloads\\Arbeitsplatzeinteilung KW 10 03.03.2025.pdf")
tables.selectPage(0)
listTable = tables.setTableNames(["Tabelle 1","NEUEINTRITT","Arbeitsplatzwechsel","NEUEINTRITTE"])

for table in listTable:
    print("detected rows:",pdf.detectTableRows(tables.pages.selected,table))
input()
for table in listTable:
    tables.selectTableByObj(table)
    rowList = pdf.detectTableRows(tables.pages.selected,table)
    tables.defRows(rowList)
    tables.parseTable()


    print("table name:",table.getName())
    for i in tables.getObjectsFromTable():
        print(i)

