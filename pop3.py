import poplib
from getpass import getpass
from email.parser import Parser
from email.policy import default as default_policy
from dotenv import load_dotenv
from enum import Enum
import os
import json
import sqlite3

#Setup envirement
load_dotenv()

#load data from .env
filterName = os.getenv('FilterName')
password = os.getenv('MailPassword')

#settings
outfile = "tmp.pdf"
host = "pop.mail.ch"
user = "outlook-bridge.santis@mail.ch"
port = 995 

#heper functions for logging
class err(Enum):
    NONE = 0
    ERROR = 1
    DEBUG = 2
    INFO = 3
    ULTRA = 4
    def __int__(self):
        return self.value
#set DEBUG level
DEBUG = err.INFO #1 error 2 debug 3 all

def log(*kwargs,level=err.INFO):
    if int(DEBUG) >= int(level):
        string = ""
        for arg in kwargs:
            string = string + str(arg) + " "
        print("DBG: ",string)

def log_json(s,level=err.INFO):
    if int(DEBUG) >= int(level):
        print("----",json.dumps(s,indent=2,default=str),"\n")
##end log funtions

def writefile(buffer,filename):
    file = open(filename,"wb")
    file.write(buffer)
    file.close()

def getUidsMail(mailbox):
    mailUidl = mailbox.uidl()
    for uid in mailUidl[1]:
        log("uidl:",uid,level=err.ULTRA)
    return mailUidl[1]

def getUidsDb(filename="mail.db"):
    connection = sqlite3.connect(filename)
    cursor = connection.cursor()

    cursor.execute("""CREATE TABLE IF NOT EXISTS mails (
        date_added TEXT DEFAULT CURRENT_TIMESTAMP,
        uid INTEGER)""")
    connection.commit()

    cursor.execute("select * from mails")

    uidsMails = cursor.fetchall()
    for uid in uidsMails:
        log("uidFromDB: ", uid,level=err.ULTRA)

    connection.close()
    return uidsMails

def addUidsDb(uidsList,filename="mail.db"):
    connection = sqlite3.connect(filename)
    cursor = connection.cursor()
    newAddedList = []

    for uid in uidsList:
        cursor.execute("""SELECT EXISTS(SELECT 1 FROM mails WHERE uid = ?)""",(uid,))
        doesExist = cursor.fetchone()
        log("Return Exist:",doesExist,level=err.ULTRA)
        if doesExist[0] == 0:
            log("addUid to db:",uid.decode())
            cursor.execute("""INSERT INTO mails(uid) VALUES(?)""",(uid,))
            newAddedList.append(uid)

    connection.commit()
    connection.close()
    return newAddedList

def setupPOP(host,port,user,password=None):
    mailbox = poplib.POP3_SSL(host,port)
    log("Server Welcome:",mailbox.getwelcome(),level=err.INFO)

    capa = mailbox.capa()
    log("Server capabilities: ", capa,level=err.INFO)

    resp = mailbox.user(user)
    if not password:
        password = getpass()
    else:
        log("read password from .env")

    resp = mailbox.pass_(password)
    log("response: ",resp)
    return mailbox

def parseMail(msgNum,filterFrom):
    fileList = []
    msg_str = ""
    msg = mailbox.retr(msgNum)
    for line in range(len(msg[1])):
        try:
            msg_str = msg_str + msg[1][line].decode() + "\n"
        except:
            log("Cant decode as a string, skipping for more INFO change DEBUG level uid=",mailbox.uidl(which=msgNum),level=err.ERROR)
            log_json(msg,level=err.ULTRA)
            return fileList

    headers = Parser(policy=default_policy).parsestr(msg_str,headersonly=False)

    if filterFrom in headers['from']:
        log("From:",headers['from'])
        log("Content type:",headers['Content-Type'])
        log("Message: ",headers.get_payload(decode=True))
        for part in headers.walk():
            ctype = part.get_content_type()
            log(ctype)
            if(ctype == "text/plain"):
                log("content pt: ",part.get_content())
        log("Message: ",headers.get_body(),level=err.ULTRA)

        for att in headers.iter_attachments():
            atype = att.get_content_type()
            log(atype,level=3)
            if "application/pdf" in atype:
                outFilename = att.get_filename()
                log("out-name(parsed):", outFilename)
                pdffile = att.get_content() 
                writefile(pdffile,outFilename)
                fileList.append(outFilename)
    return fileList

        #print("------------------\n",msg_str)
    #print("headers:",headers)
    #for key in headers.keys():
    #    print("Key: ",key)


mailbox = setupPOP(host,port,user,password)

maillist = mailbox.list()
log("maillist: ",maillist)

uids = getUidsDb()
uidsMail = getUidsMail(mailbox)
newAdded = addUidsDb(uidsMail)
#overwrite for testing
newAdded = uidsMail

for uid in newAdded:
    iUid = uid.decode().split() #convert byte stream into two integers
    log("type newadded:",type(iUid))
    log("newly added uid:",iUid)

#(response, ['mesgnum uid', ...], octets)

for uid in newAdded:
    msgNum = uid.decode().split()[0]
    log("msgNum to parse:",msgNum)
    newFileList = parseMail(msgNum,filterName)
    log("NewFileList:",newFileList)

#for mail in range(len(maillist[1])): #maillist[0] contains the response, and maillist[1] contains the "octets" / bytestrea etc
#    msg_str = ""
#    log("mail: ",mail,level=err.ULTRA)
#    msg = mailbox.retr(mail+1) #msg[0] is the response msg[1] is the data in the form of a line list#counting begins at 1 not 0


