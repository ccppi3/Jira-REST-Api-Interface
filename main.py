import pop3
import pdf
import os
import copy
import com
from pop3 import log,err

from dotenv import load_dotenv


def removeIndexesFromList(indexListRemove,_list):
    for i in sorted(indexListRemove,reverse=True): #remove allways highest order first because else the index moves
        try:
            del _list[i]
        except:
            log("indexListRemove:",indexListRemove,"\nlist:",_list," len:",len(_list),"\ni:",i)
class TableData:
    def __init__(self,name,data,fileName,pageNumber,creationDate):
        self.name = name
        self.data = data
        self.fileName = fileName
        self.pageNumber = pageNumber
        self.creationDate = creationDate
        
PDFNAMEFILTER = "Arbeitsplatzeint"
load_dotenv()

pdf.setDebugLevel(err.ULTRA,_filter="")
pop3.setDebugLevel(err.ULTRA)

#load data from .env
filterName = os.getenv('FilterName')
password = os.getenv('MailPassword')

#settings
host = os.getenv('Host')
user = os.getenv('Mail')
port = os.getenv('Port') 

#mailbox = pop3.setupPOP(host,port,user,password)

#uids = pop3.getUidsDb()
#uidsMail = pop3.getUidsMail(mailbox)
#newAdded = pop3.addUidsDb(uidsMail)

newAdded = []
uidsMail = []
log("parsing uids...")
uids = com.getEntryIDDb()
uidsMail = com.getEntryIDMail(filterName)
if uidsMail:
    newAdded = com.addEntryIDDb(uidsMail)



if len(newAdded)==0:
    log("No new mail, nothing to do",level=err.NONE)

log("newadded")
for uid in newAdded:
    log("uid:",uid)
    #iUid = uid.decode().split() #convert byte stream into two integers
    log("type newadded:",type(uid))
    log("newly added uid:",uid)

#(response, ['mesgnum uid', ...], octets)

newFileList = []
for uid in newAdded:
    #msgNum = uid.decode().split()[0]
    log("new uid to parse:",uid)
    #newFileList = newFileList + pop3.parseMail(mailbox,msgNum,filterName)
    for att in com.getAttachements(uid):
        if att.path not in newFileList:
            newFileList.append(att)
            log("Date:",att.creationDate)
        else:
            for i,att2 in enumerate(newFileList):
                if att.path == att2.path:
                    if att.creationDate > att2.creationDate:
                        newFileList.pop(i)
                        newFileList.append(att)
                        break

                
log("newFileList:",newFileList)
log("done parsing, doing some filtering")
toBeRemoved = []
for i in range(0,len(newFileList)-1):
    if i not in toBeRemoved:
        for file2 in range(i+1,len(newFileList)):
            log(newFileList[i].path," ",i," -> ",newFileList[file2].path," ",file2)
            if newFileList[i].path == newFileList[file2].path:
                toBeRemoved.append(file2)
                log("Drop duble files: ",newFileList[i].path,"->",newFileList[file2].path,level=err.ERROR)

log("To be removed: ",toBeRemoved)
removeIndexesFromList(toBeRemoved,newFileList)
toBeRemoved.clear()

#filter out: is a plausible pdf?

#for i,fileName in enumerate(newFileList):
#    if PDFNAMEFILTER not in fileName:
#        toBeRemoved.append(i)
#log("To be removed: ",toBeRemoved)
#removeIndexesFromList(toBeRemoved,newFileList)
for a in newFileList:
    log("Download :",a.path,a.uid)
    com.downloadAttachements(a.uid)

log("\033[37;42m",level=err.NONE) #ansi escape sequence to change color
log("I checked ",len(list(uidsMail))," uids, processed ",len(newAdded)," new mails and downloaded ",len(newFileList)," files\033[0m",level=err.NONE);
log("\033[0m",level=err.NONE)#ansi reset color



tableDataList = []

for file in newFileList:
    tables = pdf.Tables(file.path)
    countPage = tables.countPages()
    for pageNr in range(countPage):
        log("Parse page ",pageNr," of file ",file,level=err.NONE)
        tables.selectPage(pageNr)
        listTable = tables.setTableNames(["Tabelle 1","NEUEINTRITT","Arbeitsplatzwechsel","NEUEINTRITTE"])
        for i,table in enumerate(listTable):
            tables.selectTableByObj(table)
            tables.defRows(["Vorname","Name","Kürzel","Abteilung","Abteilung vorher","Abteilung neu","Abteilung Neu"])
            tables.parseTable()
            
            objList = []
            for tbl in tables.getObjectsFromTable():
                objcpy = copy.deepcopy(tbl) # make a real copy of the data, else the data would be overwritten by the next page, as the data would be parsed as reference
                objList.append(objcpy)
            tableDataList.append(TableData(table.name,objList,table.fileName,table.pageNumber,file.creationDate))

for table in tableDataList:
    filenameList = str(table.fileName).split("\\")
    filename = filenameList[len(filenameList)-1]
    print("--------¦",table.name,"¦-------------¦","filename: ",filename, "pageNr:", table.pageNumber,"Date: ",table.creationDate)
    for entry in table.data:
        print(entry)
    print("-------------------------------------\n")


