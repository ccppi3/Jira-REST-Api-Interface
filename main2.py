import pop3
import pdf
import os
import copy
from pop3 import log,err

from dotenv import load_dotenv

PDFNAMEFILTER = "Arbeitsplatzeint"
load_dotenv()

pdf.setDebugLevel(err.INFO,_filter="")
pop3.setDebugLevel(err.INFO)

#load data from .env
filterName = os.getenv('FilterName')
password = os.getenv('MailPassword')

#settings
host = os.getenv('Host')
user = os.getenv('Mail')
port = os.getenv('Port') 

mailbox = pop3.setupPOP(host,port,user,password)

uids = pop3.getUidsDb()
uidsMail = pop3.getUidsMail(mailbox)
newAdded = pop3.addUidsDb(uidsMail)

def removeIndexesFromList(indexListRemove,_list):
    for i in sorted(indexListRemove,reverse=True): #remove allways highest order first because else the index moves
        del _list[i]

if len(newAdded)==0:
    log("No new mail, nothing to do",level=err.NONE)

for uid in newAdded:
    iUid = uid.decode().split() #convert byte stream into two integers
    log("type newadded:",type(iUid))
    log("newly added uid:",iUid)

#(response, ['mesgnum uid', ...], octets)

newFileList = []
for uid in newAdded:
    msgNum = uid.decode().split()[0]
    log("msgNum to parse:",msgNum)
    newFileList = newFileList + pop3.parseMail(mailbox,msgNum,filterName)

toBeRemoved = []
for i,file1 in enumerate(newFileList):
    for file2 in range(i+1,len(newFileList)):
        log(file1," -> ",newFileList[file2])
        if file1 == newFileList[file2]:
            toBeRemoved.append(file2)
            log("Drop duble files: ",file1,level=err.ERROR)
log("To be removed: ",toBeRemoved)
removeIndexesFromList(toBeRemoved,newFileList)
toBeRemoved.clear()

#filter out: is a plausible pdf?

for i,fileName in enumerate(newFileList):
    if PDFNAMEFILTER not in fileName:
        toBeRemoved.append(i)
log("To be removed: ",toBeRemoved)
removeIndexesFromList(toBeRemoved,newFileList)

log("NewFileList:",newFileList)

log("\033[37;42m",level=err.NONE) #ansi escape sequence to change color
log("I checked ",len(uidsMail)," uids, processed ",len(newAdded)," new mails and downloaded ",len(newFileList)," files\033[0m",level=err.NONE);
log("\033[0m",level=err.NONE)#ansi reset color

objList = []

for file in newFileList:
    tables = pdf.Tables(file)
    countPage = tables.countPages()
    for pageNr in range(countPage):
        log("Parse page ",pageNr," of file ",file,level=err.NONE)
        tables.selectPage(pageNr)
        listTable = tables.setTableNames(["Tabelle 1","NEUEINTRITT","Arbeitsplatzwechsel","NEUEINTRITTE"])
        for table in listTable:
            tables.selectTableByObj(table)
            tables.defRows(["Vorname","Name","Kürzel","Abteilung","Abteilung vorher","Abteilung neu","Abteilung Neu"])
            tables.parseTable()

            for tbl in tables.getObjectsFromTable():
                objcpy = copy.deepcopy(tbl)
                objList.append(objcpy)

for obj in objList:
    print("ALLDATA: \n",obj)

tablesDbg = pdf.Tables('tests/BPN-SAS2-846EL.pdf')
log("Debug special pdf")
page = tablesDbg.selectPage(12)

objDbgList = []
listTablesDbg = tablesDbg.setTableNames(["General Jumper Settings"])
for table in listTablesDbg:
    tablesDbg.selectTableByObj(table)
    tablesDbg.defRows(["Jumper","Jumper Settings","Note"])
    tablesDbg.parseTable()

    for tbl in tablesDbg.getObjectsFromTable():
        objcpy2 = copy.deepcopy(tbl)
        objDbgList.append(objcpy2)
    
    border = pdf.Border(209,1190-723,270,1190-600,50)
    for rect in pdf.getRectsInRange(page,border):
        log("rects tablesDbg:",pdf.transformRect(page,rect))
    else:
        log("no rects found in tablesDbg")

    for [text,rect] in pdf.getTextInRange(page,border):
        log("text tablesDbg:",pdf.transformRect(page,rect))
    else:
        log("no text found in tablesDbg")
log("found obj:",len(objDbgList))
for obj in objDbgList:
    log("dbg_data: \n",obj)
