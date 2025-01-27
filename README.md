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
## reference
### tables = pdf.Tables(filename)

opens a pdf file  
returns an object witch holds all the table data

### class Tables:

#### selectPage(pagenumber)
Selects the page to be parsed, only one can be specified.

#### setTableNames(["table1","table2","rable3",...])
Define possible Table Names witch can be treated as reference, or also title of the table.  
This will be used as an entry point for the algorythm to determine table positions and fields.
The list is treated as "OR", or table contains any of the following.

This method also parses the previously selected page for the selected Table names.


This method returns the list of all found tables. which is a reference to self.tables
#### selectTable(indexOfTable)
Needs to be called after the setTableNames, as for why read the reference of given method.  
#### selectTableByObj(tableObject)
The same as selectTable but we will take a table as argument,

This was implemented due to simplify coding, and interworking with a list of tables

This will tell the algorythm which table in the index we will parse

#### defRows([row1,row2,row3,...])
This defines all the possible Row names, which should be parsed if they are in the table.  
They can be in the table, but the dont have to be in the table.  


#### parseTable()
This does the actual parsing.  
The results are stored internally an can be retrived with getObjectsFromTable().

#### getObjectsFromTable()
This returns a list of objects witch are stored in the current selected and already parsed table.


The returned objects contain each field named after the defined rows (via defRows) if they are empty they will contain the python keyword None.
