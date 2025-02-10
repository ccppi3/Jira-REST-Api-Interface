import pop3
import pdf
import os
import copy
import com
import ticket
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
        
PDFNAMEFILTER = os.getenv('PdfNameFilter')
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

def run():
    newAdded = []
    uidsMail = []
    log("parsing uids...")
    uids = com.getEntryIDDb()
    uidsMail = com.getEntryIDMail(filterName)
    if uidsMail:
        newAdded = com.addEntryIDDb(uidsMail)

    if len(newAdded)==0:
        log("No new mail, nothing to do",level=err.NONE)

    for uid in newAdded:
        log("newadded:",uid," type:",type(uid))

    newFileList = _trimNewAdded(newAdded)

    try:             
        log("newFileList:",newFileList,level=err.ULTRA)
    except:
        log("no new Filelist",level=err.ULTRA)

    log("done parsing, doing some filtering")

    newFileList = _removeDoubles(newFileList)
    #newFilelist = _filterList(newFileList)

    for a in newFileList:
        log("Download :",a.path,a.uid)
        com.downloadAttachements(a.uid)

    log("\033[37;42m",level=err.NONE) #ansi escape sequence to change color
    log("I checked ",len(list(uidsMail))," uids, processed ",len(newAdded)," new mails and downloaded ",len(newFileList)," files\033[0m",level=err.NONE);
    log("\033[0m",level=err.NONE)#ansi reset color

    tableDataList = _runPdfParser(newFileList)
    tablesToTicket(tableDataList)

def _runPdfParser(newFileList): #helper function witch wraps all the parsing calls, and copys the volatile data into a non volatile memory
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
    return tableDataList

def tablesToTicket(tableDataList): #tackes the tabledata and creates a ticket for each table
    for table in tableDataList:
        filenameList = str(table.fileName).split("\\")
        filename = filenameList[len(filenameList)-1]
        print("--------¦",table.name,"¦-------------¦","filename: ",filename, "pageNr:", table.pageNumber,"Date: ",table.creationDate)
        tempObjs = []
        for entry in table.data:
            print(entry)
            tempObjs.append(entry)
        print("-------------------------------------\n")
        if "arbeitsplatzwechsel" in str(table.name).lower():
            ticketTable = ticket.Ticket(tempObjs,filename,"Allpower","arbeitsplatzwechsel")
        elif "neueintritte" in str(table.name).lower():
            ticketTable = ticket.Ticket(tempObjs,filename,"Allpower","neueintritt")
        elif "neueintritt" in str(table.name).lower():
            ticketTable = ticket.Ticket(tempObjs,filename,"Allpower","neueintritte")

        ticketTable.create_ticket()
def _removeDoubles(newFileList):
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
    return newFileList

def _filterList(fileNameList,pdfNameFilter=PDFNAMEFILTER): #filter out: is a plausible pdf?
    for i,fileName in enumerate(fileNameList):
        if pdfNameFilter not in fileName:
            toBeRemoved.append(i)
    log("To be removed: ",toBeRemoved)
    removeIndexesFromList(toBeRemoved,fileNameList)
    return fileNameList

def _trimNewAdded(newAddedList):#get attachements, on double entries, choose the newest and drop the older
    newFileList = []
    for uid in newAddedList:
        log("new uid to parse:",uid)
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
    return newFileList

if __name__=="__main__":
    run()
