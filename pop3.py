#Implementatiion of a mail client to get all Pdfs sent from a given sender.
#Author: Jonathan Wyss

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
if __name__ == "__main__":
    load_dotenv()

    #load data from .env
    filterName = os.getenv('FilterName')
    password = os.getenv('MailPassword')

    #settings
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

def setDebugLevel(level):
    global DEBUG
    DEBUG = level
def log(*kwargs,level=err.INFO):
    if int(DEBUG) >= int(level):
        string = ""
        for arg in kwargs:
            string = string + str(arg) + " "
        if int(level) != int(err.NONE):
            print("[",(str(level).split(".")[1]).replace(" ",""),"]",string)
        else:
            print("[INFO] ",string)

def log_json(s,level=err.INFO):
    if int(DEBUG) >= int(level):
        print("----",json.dumps(s,indent=2,default=str))
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
        iUid = uid.decode().split()[1]
        cursor.execute("""SELECT EXISTS(SELECT 1 FROM mails WHERE uid = ?)""",(iUid,))
        doesExist = cursor.fetchone()
        log("Return Exist:",doesExist,level=err.ULTRA)
        if doesExist[0] == 0:
            log("addUid to db:",iUid)
            cursor.execute("""INSERT INTO mails(uid) VALUES(?)""",(iUid,))
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

def parseMail(mailbox,msgNum,filterFrom):
    fileList = []
    msg_str = ""
    msg = mailbox.retr(msgNum)
    log("done retrieving, msg:",msg[0].decode().split(),level=err.ULTRA)
    if int(msg[0].decode().split()[1]) > 1000000:
        log("very long message skipping...",level=err.ERROR)
        return fileList
    for line in range(len(msg[1])):

        try:
            msg_str = msg_str + msg[1][line].decode() + "\n"
        except:
            log("Cant decode as a string, skipping for more INFO change DEBUG level uid=",mailbox.uidl(which=msgNum),level=err.ERROR)
            log_json(msg,level=err.ULTRA)
            return fileList

    log("run parser.parsestr...")
    headers = Parser(policy=default_policy).parsestr(msg_str,headersonly=False)
    log("done parsing\nlength of msgstr: ",len(msg_str))
    log("headers:",headers['from'])
    if filterFrom in headers['from']:
        log("From:",headers['from'])
        log("Content type:",headers['Content-Type'])
        log("Message: ",headers.get_payload(decode=True))
        for part in headers.walk():
            ctype = part.get_content_type()
            log(ctype)
            log("ctype:",ctype,level=err.ULTRA)
            if(ctype == "text/plain"):
                log("content pt: ",part.get_content())
        log("Message: ",headers.get_body(),level=err.ULTRA)

        for att in headers.iter_attachments():
            log("att",att,level=err.ULTRA)
            atype = att.get_content_type()
            log(atype,level=err.INFO)
            if "application/pdf" in atype:
                outFilename = att.get_filename()
                log("out-name(parsed):", outFilename)
                pdffile = att.get_content() 
                writefile(pdffile,outFilename)
                fileList.append(outFilename)
    else:
        log("not matching any filter")
    return fileList

if __name__=="__main__":
    mailbox = setupPOP(host,port,user,password)

    uids = getUidsDb()
    uidsMail = getUidsMail(mailbox)
    newAdded = addUidsDb(uidsMail)
    #overwrite for testing
    #newAdded = uidsMail
    if len(newAdded)==0:
        log("No new mail, nothing to do",level=err.NONE)

    for uid in newAdded:
        iUid = uid.decode().split() #convert byte stream into two integers
        log("type newadded:",type(iUid))
        log("newly added uid:",iUid)
    
    newFileList = []
    for uid in newAdded:
        msgNum = uid.decode().split()[0]
        log("msgNum to parse:",msgNum)
        newFileList = newFileList + parseMail(msgNum,filterName)
    log("NewFileList:",newFileList)
    
    log("\033[37;42m",level=err.NONE) #ansi escape sequence to change color
    log("I checked ",len(uidsMail)," uids, processed ",len(newAdded)," new mails and downloaded ",len(newFileList)," files\033[0m",level=err.NONE);
    log("\033[0m",level=err.NONE)#ansi reset color
