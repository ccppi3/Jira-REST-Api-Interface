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

# pop3 / mail library

I wrote a little abstraction layer for the mail part of the project.
## log / error handling

### class err(Enum)

This is just an enum for defining the desired verbosity of the debug / log to stdout.

possible values are:
- err.NONE   standard informational output
- err.ERROR  only show errors
- err.DEBUG
- err.INFO
- err.ULTRA  this is very verbose!

related mehtod see setDebugLevel and log in section Methods


## Methods

### setDebugLevel(err)

Defines the debug level of the library, the argument is of type class/enum err see reference for possible values. For exampe:
> pop3.setDebugLevel(err.ERROR)

### log(*kwargs,level=error.INFO)

This is used to print / log to stdout which takes into consideration of the previous set log level with setDebugLevel.  
The keyword argument level sets the assosiated verbosity, and defaults to error.INFO.

### log_json(json_string,level=err.INFO)

Same as log but pretty prints json and similar big data junks.

### setupPOP(host,port,user,password=None)

This makes connection to the server according to the arguments.
if no password is specified, the shell will prompt to input the password.

This also logs the server capabilities, when the DEBUG level is at least err.INFO.

This also logs in as the specified user

> returns a mailbox from the class poplib.POP3_SSL

### parseMail(mailbox, msgNum,filterFrom)

filters mail with message number msgNum in mailbox which accord to the filterFrom argument in the mail header field 'from'.

This method downloads any pdf from a match.

> returns a list of all downloaded files.


### getUidsMail(mailbox)

returns a list of uids of all mails on the server

### getUidsDb(filename="mail.db")

> returns all uids from the db, db name defaults to mail.db.
the db is a sqlite3 database.

### addUidsDb(uidsList,filename="mail.db")

This takes the list and synchronises with the db, if a new entry is found it will be returned in a list which contains all new uids.

> returns list of new uids





