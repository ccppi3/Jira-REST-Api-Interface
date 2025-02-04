
import pop3
import pdf
import os
import copy
from pop3 import log,err

from dotenv import load_dotenv

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
tablesDbg = pdf.Tables('test4.pdf')
log("Debug special pdf")
page = tablesDbg.selectPage(0)

objDbgList = []
listTablesDbg = tablesDbg.setTableNames(["Arbeitsplatzwechsel","NEUEINTRITT","NEUEINTRITTE"])
for table in listTablesDbg:
    tablesDbg.selectTableByObj(table)
    tablesDbg.defRows(["Vorname","Name","Kürzel","Abteilung","Abteilung vorher","Abteilung neu","Abteilung Neu","Platz-Nr."])
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
    log("dbg_data: \n",obj,err.INFO)
