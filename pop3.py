import poplib
from getpass import getpass
from email.parser import Parser
from email.policy import default as default_policy
from dotenv import load_dotenv
from enum import Enum
import os
import json


class err(Enum):
    NONE = 0
    ERROR = 1
    DEBUG = 2
    INFO = 3
    def __int__(self):
        return self.value
DEBUG = err.INFO #1 error 2 debug 3 all

def log(*kwargs,level=err.INFO):
    if int(DEBUG) >= int(level):
        string = ""
        for arg in kwargs:
            string = string + str(arg) + " "
            print("DBG: ",string)

def log_json(s):
    if int(DEBUG) > 0:
        print("----",json.dumps(s,indent=2,default=str),"\n")

load_dotenv()

mail_name = os.getenv('Name')

password = os.getenv('Password')

def writefile(buffer,filename):
    file = open(filename,"wb")
    file.write(buffer)
    file.close()

outfile = "tmp.pdf"
host = "pop.mail.ch"
user = "outlook-bridge.santis@mail.ch"
port = 995 

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

maillist = mailbox.list()

log("maillist: ",maillist)


for mail in range(len(maillist[1])): #maillist[0] contains the response, and maillist[1] contains the "octets" / bytestrea etc
    msg_str = ""
    log("mail: ",mail)
    msg = mailbox.retr(mail+1) #msg[0] is the response msg[1] is the data in the form of a line list
    for line in range(len(msg[1])):

        try:
            msg_str = msg_str + msg[1][line].decode() + "\n"
        except:
            log("Cant decode as a string, skipping and dumping the msg",err.ERROR)
            log_json(msg)

    headers = Parser(policy=default_policy).parsestr(msg_str,headersonly=False)

    if mail_name in headers['from']:
        log("From:",headers['from'])
        log("Content type:",headers['Content-Type'])
        log("Message: ",headers.get_payload(decode=True))
        for part in headers.walk():
            ctype = part.get_content_type()
            log(ctype)
            if(ctype == "text/plain"):
                log("content pt: ",part.get_content())
        log("Message: ",headers.get_body())

        for att in headers.iter_attachments():
            atype = att.get_content_type()
            log(atype,level=3)
            if "application/pdf" in atype:
                pdffile = att.get_content() 
                writefile(pdffile,outfile)
        #print("------------------\n",msg_str)
    #print("headers:",headers)
    #for key in headers.keys():
    #    print("Key: ",key)


