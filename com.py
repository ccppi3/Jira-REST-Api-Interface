#https://learn.microsoft.com/en-us/office/vba/api/outlook.items
#https://learn.microsoft.com/en-us/office/vba/api/outlook.attachment.saveasfile

import win32com.client
import os
from dotenv import load_dotenv

PATH = os.getcwd() + "\\downloads\\"
print(PATH)
INBOXNR = 6

load_dotenv()

filterName = os.getenv("FilterName")

def getFileName(filename):
    dotList = str(filename).split('.')
    ending = dotList[len(dotList)-1]#use the last entry
    return ending
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

inbox = outlook.GetDefaultFolder(INBOXNR)
print("Folder:",inbox.Name)

messages = inbox.Items

message = messages.GetLast()

for m in messages:
    try:
        sender = m.Sender
    except:
        pass
    else:
        if str(filterName).lower() in str(m.Sender).lower():
            print("From:",m.Sender,"Title:",m)
            for att in m.Attachments:
                print("Attachements:",att)
                if getFileName(att) == "pdf":
                    path = PATH + str(att)
                    att.SaveAsFile(path)
                    print("download to:",PATH)



