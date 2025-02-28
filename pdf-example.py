import pdf

tables = pdf.Tables("downloads\\Arbeitsplatzeinteilung KW 10 03.03.2025.pdf")
tables.selectPage(0)
listTable = tables.setTableNames(["Tabelle 1","NEUEINTRITT","Arbeitsplatzwechsel","NEUEINTRITTE"])

for table in listTable:
    print("detected rows:",pdf.detectTableRows(tables.pages.selected,table))
input()
for table in listTable:
    tables.selectTableByObj(table)
    tables.defRows(["Vorname","zel","Name","Kür-","Kürzel","Abteilung","Abteilung vorher","Abteilung neu","Abteilung Neu","Platz-Nr."])
    tables.parseTable()


    print("table name:",table.getName())
    for i in tables.getObjectsFromTable():
        print(i)

