import pop3
import pdf
import os
from pop3 import log,err

from dotenv import load_dotenv


load_dotenv()

#load data from .env
filterName = os.getenv('FilterName')
password = os.getenv('MailPassword')

#settings
host = "pop.mail.ch"
user = "outlook-bridge.santis@mail.ch"
port = 995 

mailbox = pop3.setupPOP(host,port,user,password)

uids = pop3.getUidsDb()
uidsMail = pop3.getUidsMail(mailbox)
newAdded = pop3.addUidsDb(uidsMail)
#overwrite for testing
#newAdded = uidsMail
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
    newFileList = newFileList + parseMail(msgNum,filterName)
log("NewFileList:",newFileList)

log("\033[37;42m",level=err.NONE) #ansi escape sequence to change color
log("I checked ",len(uidsMail)," uids, processed ",len(newAdded)," new mails and downloaded ",len(newFileList)," files\033[0m",level=err.NONE);
log("\033[0m",level=err.NONE)#ansi reset color
