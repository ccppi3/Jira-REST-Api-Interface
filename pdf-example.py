import pdf

tables = pdf.Tables("tmp.pdf")
tables.selectPage(2)
listTable = tables.setTableNames(["Tabelle 1","NEUEINTRITT","Arbeitsplatzwechsel","Wiedereintritt"])
#listTable = tables.setTableNames(["Wiedereintritt"])

#for table in listTable:
#    tables.selectTableByObj(table)
#    print("detected rows:",pdf.detectTableRows(tables.pages.selected,table))

for table in listTable:
    tables.selectTableByObj(table)
    rowList = pdf.detectTableRows(tables.pages.selected,table)
    print("rowList:",rowList)
    tables.defRows(rowList)
    tables.parseTable()


    print("Result: \ntable name:",table.getName())
    for i in tables.getObjectsFromTable():
        print(i)
#    input()



def render(page):
    page.getPixmap()

