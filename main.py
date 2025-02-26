import pop3
import pdf
import os
import copy
import com
import ticket
from pop3 import log,err
import pythoncom

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

#load data from .env
load_dotenv()
filterName = os.getenv('FilterName')
password = os.getenv('MailPassword')
PDFNAMEFILTER = os.getenv('PdfNameFilter')

#settings
host = os.getenv('Host')
user = os.getenv('Mail')
port = os.getenv('Port') 

pdf.setDebugLevel(err.ULTRA,_filter="")
pop3.setDebugLevel(err.ULTRA)

cwd = os.getcwd()
if not os.path.exists(cwd + "/downloads"):
    os.makedirs(cwd + "/downloads")

def run(outlook):
    pythoncom.CoInitialize()

    newAdded = []
    uidsMail = []
    log("parsing uids...")
    yield "parsing uids..."

    uids = com.getEntryIDDb()
    #uidsMail = com.getEntryIDMail(filterName,outlook)
    for ids in com.getEntryIDMail(filterName,outlook):
        uidsMail.append(ids)
    
    log("addEntryIDDB uidsMail",uidsMail,"len:",len(uidsMail))
    newAdded = com.addEntryIDDb(uidsMail)
    if len(newAdded)==0:
        log("No new mail, nothing to do",level=err.NONE)
    for uid in newAdded:
        log("newadded:",uid," type:",type(uid))

    newFileList = _trimNewAdded(newAdded,outlook)
    log("newFileList:",newFileList,level=err.INFO)

    newFileList = _removeDoubles(newFileList)
    #newFilelist = _filterList(newFileList)
    for a in newFileList:
        log("Download :",a.path,a.uid)
        com.downloadAttachements(a.uid,outlook)
        yield "Download..." + str(a.path)
    
    printStatistics(uids,newAdded,newFileList)
    yield "Running PDFParser..."

    tableDataList = []
    for ret in _runPdfParser(newFileList):
        if type(ret) == list:
            tableDataList = ret
        else:
            yield ret

    yield "finished parsing"
    yield tableDataList


def printStatistics(uidsMail,newAdded,newFileList):
    log("\033[37;42m",level=err.NONE) #ansi escape sequence to change color
    log("I checked ",len(list(uidsMail))," uids, processed ",len(newAdded)," new mails and downloaded ",len(newFileList)," files\033[0m",level=err.NONE);
    log("\033[0m",level=err.NONE)#ansi reset color

def _runPdfParser(newFileList): #helper function witch wraps all the parsing calls, and copys the volatile data into a non volatile memory
    tableDataList = []
    for file in newFileList:
        tables = pdf.Tables(file.path)
        countPage = tables.countPages()
        for pageNr in range(countPage):
            log("Parse page ",pageNr," of file ",file,level=err.NONE)
            yield "Parse:" + str(file).split("\\")[len(str(file).split("\\"))-1] + " Page " + str(pageNr)
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
    yield tableDataList

def tablesToTicket(tableDataList,check=True): #tackes the tabledata and creates a ticket for each table
    log("count tables: ",len(tableDataList),level=err.INFO)
    for table in tableDataList:
        log("table: ",table,level=err.INFO)
        filenameList = str(table.fileName).split("\\")
        filename = filenameList[len(filenameList)-1]
        print("--------¦",table.name,"¦-------------¦","filename: ",filename, "pageNr:", table.pageNumber,"Date: ",table.creationDate)
        tempObjs = []
        for entry in table.data:
            print(entry)
            tempObjs.append(entry)
        print("-------------------------------------\n")
        if "arbeitsplatzwechsel" in str(table.name).lower():
            ticketTable = ticket.Ticket(table,tempObjs,filename,"Allpower","arbeitsplatzwechsel")
        elif "neueintritte" in str(table.name).lower():
            ticketTable = ticket.Ticket(table,tempObjs,filename,"Allpower","neueintritt")
        elif "neueintritt" in str(table.name).lower():
            ticketTable = ticket.Ticket(table,tempObjs,filename,"Allpower","neueintritte")
        ticketTable.createTicket(check=False)

def tableToTicket(table,check=True):
        log("table: ",table,level=err.INFO)
        filenameList = str(table.fileName).split("\\")
        filename = filenameList[len(filenameList)-1]
        print("--------¦",table.name,"¦-------------¦","filename: ",filename, "pageNr:", table.pageNumber,"Date: ",table.creationDate)
        tempObjs = []
        for entry in table.data:
            print(entry)
            tempObjs.append(entry)
        print("-------------------------------------\n")
        if "arbeitsplatzwechsel" in str(table.name).lower():
            ticketTable = ticket.Ticket(table,tempObjs,filename,"Allpower","arbeitsplatzwechsel")
        elif "neueintritte" in str(table.name).lower():
            ticketTable = ticket.Ticket(table,tempObjs,filename,"Allpower","neueintritt")
        elif "neueintritt" in str(table.name).lower():
            ticketTable = ticket.Ticket(table,tempObjs,filename,"Allpower","neueintritte")
        for status in ticketTable.createTicket(check=False):
            yield status
    

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

def _trimNewAdded(newAddedList,outlook):#get attachements, on double entries, choose the newest and drop the older
    newFileList = []
    for uid in newAddedList:
        log("new uid to parse:",uid)
        for att in com.getAttachements(uid,outlook):
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
    for ret in run():
        if type(ret) == list: #last yield returns the tableData
            print("is list run ticket")
            tablesToTicket(ret)
        else:#return is status string
            print("[main]",ret)
        

