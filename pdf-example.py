import pdf

tables = pdf.Tables("tmp.pdf")
tables.selectPage(0)
listTable = tables.setTableNames(["Tabelle 1","NEUEINTRITT","Arbeitsplatzwechsel","NEUEINTRITTE","Wiedereintritt"])
#listTable = tables.setTableNames(["Wiedereintritt"])

for table in listTable:
    print("detected rows:",pdf.detectTableRows(tables.pages.selected,table))

for table in listTable:
    tables.selectTableByObj(table)
    rowList = pdf.detectTableRows(tables.pages.selected,table)
    print("rowList:",rowList)
    tables.defRows(rowList)
    tables.parseTable()


    print("table name:",table.getName())
    for i in tables.getObjectsFromTable():
        print(i)



def render(page):
    page.getPixmap()

