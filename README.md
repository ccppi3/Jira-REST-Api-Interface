# Jira-REST-Api-Interface
An Interface to automaticaly create Jira Tickets via their REST-Api.

# Main Use Case
Main Purpose is to create a Ticket whenever there is an onboarding of a new employee.
The Payload gets made from data about the employee we get in an email. 

# pdf library
I have written an abstraction layer for the usecase where there are Tables in a PDF file to parse the table data out and convert them into a python object.
## Example
```
import pdf

tables = pdf.Tables("test2.pdf")
tables.selectPage(0)
tables.setTableNames(["Tabelle 1","Tabelle 2","Tabelle 3","Tabelle 4"])
tables.selectTable(0)
tables.defRows(["b","Name","A"])
tables.parseTable()


print("table 1")
for i in tables.getObjectsFromTable():
    print(i)
```
