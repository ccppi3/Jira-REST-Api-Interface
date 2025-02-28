#This implementation gets all mails from a running outlook instance witch are from a given sender and downloads the pdfs into the folder ./downloads
#Author: Jonathan Wyss
##Notes
#https://learn.microsoft.com/en-us/office/vba/api/outlook.items
#https://learn.microsoft.com/en-us/office/vba/api/outlook.attachment.saveasfile

import win32com.client
import os
from dotenv import load_dotenv
import sqlite3
from pop3 import err,log
import pythoncom
from gui import getResourcePath
import tempfile


#PATH = os.getcwd() + "\\downloads\\"
PATH = tempfile.gettempdir()
print(PATH)
INBOXNR = 6

load_dotenv(getResourcePath(".env"))

pythoncom.CoInitialize()



filterName = os.getenv("FilterName")
class Entry:
    def __init__(self,uid,path,creationDate):
        self.uid = uid
        self.path = path
        self.creationDate = creationDate
    def __eq__(self,other):
        return self.path == other
    def __str__(self):
        return str(self.path) + " " + str(self.creationDate)

def init():
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    except:
        log("error no connection to outlook",level=err.ERROR)
    return outlook

def getEntryIDDb(filename="mail.db"):
    connection = sqlite3.connect(filename)
    cursor = connection.cursor()

    cursor.execute("""CREATE TABLE IF NOT EXISTS outlook (
        date_added TEXT DEFAULT CURRENT_TIMESTAMP,
        id INTEGER)""")
    connection.commit()

    cursor.execute("select * from outlook")

    entryIDMails = cursor.fetchall()
    for uid in entryIDMails:
        log("EntryIDFromDB: ", uid,level=err.ULTRA)

    connection.close()
    return entryIDMails

def addEntryIDDb(entryIDList,filename="mail.db"):
    connection = sqlite3.connect(filename)
    cursor = connection.cursor()
    newAddedList = []
    for uid in entryIDList:
        #iUid = uid.decode().split()[1]
        log("entryID:",uid)
        cursor.execute("""SELECT EXISTS(SELECT 1 FROM outlook WHERE id = ?)""",(uid,))
        doesExist = cursor.fetchone()
        log("Return Exist:",doesExist,level=err.ULTRA)
        if doesExist[0] == 0:
            log("addUid to db:",uid)
            cursor.execute("""INSERT INTO outlook(id) VALUES(?)""",(uid,))
            newAddedList.append(uid)

    connection.commit()
    connection.close()
    return newAddedList
    
def rmEntryIDDB(entryId,filename="mail.db"):
    connection = sqlite3.connect(filename)
    cursor = connection.cursor()
    cursor.execute("""DELETE FROM outlook WHERE id = ? RETURNING id""",(entryId,))
    log("Remove result:",cursor.fetchone())
    connection.commit()
    connection.close()
    
def rmLastEntryIDDB(filename="mail.db"):
    connection = sqlite3.connect(filename)
    cursor = connection.cursor()
    cursor.execute("""SELECT * from outlook ORDER BY rowid DESC LIMIT 1""")
    lastId = cursor.fetchone()
    connection.close()

    log("latest entry in db:",lastId)
    rmEntryIDDB(lastId[1])



def getFileName(filename):
    dotList = str(filename).split('.')
    ending = dotList[len(dotList)-1]#use the last entry
    return ending

def getMAPI(outlook):
    try:
        outlook = win32com.client.GetActiveObject("Outlook.Application")
    except:
        log("could not get active outlook instance!",level=err.ERROR)
    outlook = outlook.GetNameSpace("MAPI")
    return outlook

def getEntryIDMail(filterName,outlook,inboxnr=INBOXNR):
    outlook = getMAPI(outlook)
    inbox = outlook.GetDefaultFolder(inboxnr)
    print("Folder:",inbox.Name)
    messages = inbox.Items
    log("count messages:",len(messages))
    for m in messages:
        try:
            sender = m.Sender
        except:
            log("no Sender object in message")
        else:
            if str(filterName).strip().lower() in str(m.Sender).strip().lower():
                yield m.EntryID
            else:
                pass
                #log("Sender",m.Sender,"does not match",filterName)

    

def downloadAllAttachements(fileEnding,filterName,outlook,path=PATH,inboxnr=INBOXNR):
    outlook = getMAPI(outlook)
    inbox = outlook.GetDefaultFolder(inboxnr)
    print("Folder:",inbox.Name)

    messages = inbox.Items
    for m in messages:
        try:
            sender = m.Sender
        except:
            pass
        else:
            if str(filterName).lower() in str(m.Sender).lower():
                print("From:",m.Sender,"Title:",m, "EntryID:",m.EntryID)
                for att in m.Attachments:
                    print("Attachements:",att)
                    if getFileName(att) == fileEnding:
                        path = path + str(att)
                        att.SaveAsFile(path)
                        print("download to:",PATH)

def getAttachements(entryID,outlook,fileEnding="pdf",path=PATH,inboxnr=INBOXNR):
    outlook = getMAPI(outlook)
    inbox = outlook.GetDefaultFolder(inboxnr)

    messages = inbox.Items
    for m in messages:
        if m.EntryID == entryID:
            print("From:",m.Sender,"Title:",m, "EntryID:",m.EntryID)
            for att in m.Attachments:
                if getFileName(att) == fileEnding:
                    path = path + str(att)
                    yield Entry(m.EntryID,path,m.CreationTime)

def downloadAttachements(entryID,outlook,fileEnding="pdf",path=PATH,inboxnr=INBOXNR):
    outlook = getMAPI(outlook)
    inbox = outlook.GetDefaultFolder(inboxnr)

    messages = inbox.Items
    for m in messages:
        if m.EntryID == entryID:
            print("From:",m.Sender,"Title:",m, "EntryID:",m.EntryID)
            for att in m.Attachments:
                print("Attachements:",att)
                if getFileName(att) == fileEnding:
                    path = path + str(att)
                    att.SaveAsFile(path)
                    print("download to:",path)
