import pdf

tables = pdf.Tables("test2.pdf")
tables.selectPage(0)
tables.selectTable(0)
tables.defRows(["Kürzel","Name","Abteilung"])
tables.parseTable()


print("table 1")
for i in tables.getObjectsFromTable():
    print(i)

tables.selectPage(0)
tables.selectTable(1)
tables.defRows(["Kürzel","Name","Abteilung"])
tables.parseTable()


print("table 2")
for i in tables.getObjectsFromTable():
    print(i)
