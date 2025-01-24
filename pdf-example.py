import pdf

tables = pdf.Tables("test2.pdf")
tables.selectPage(0)
tables.setTableNames(["Tabelle 1","NEUEINTRITT","Arbeitsplatzwechsel","NEUEINTRITTE"])
tables.selectTable(0)
tables.defRows(["b","Name","A"])
tables.parseTable()


print("table 1")
for i in tables.getObjectsFromTable():
    print(i)

