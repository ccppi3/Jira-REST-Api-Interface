import poplib
from getpass import getpass
from email.parser import Parser
from email.policy import default as default_policy
from dotenv import load_dotenv
import os


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

print("Server Welcome:",mailbox.getwelcome())

capa = mailbox.capa()

print("Server capabilities: ", capa)

resp = mailbox.user(user)
if not password:
    password = getpass()
else:
    print("read password from .env")

resp = mailbox.pass_(password)
print("response: ",resp)

maillist = mailbox.list()

print("maillist: ",maillist)


for mail in range(len(maillist[1])):
    msg_str = ""
    print("mail: ",mail)
    msg = mailbox.retr(mail+1)
    for line in range(len(msg[1])):
        msg_str = msg_str + msg[1][line].decode() + "\n"

    headers = Parser(policy=default_policy).parsestr(msg_str,headersonly=False)

    if mail_name in headers['from']:
        print("From:",headers['from'])
        print("Content type:",headers['Content-Type'])
        print("Message: ",headers.get_payload(decode=True))
        for part in headers.walk():
            ctype = part.get_content_type()
            print(ctype)
            if(ctype == "text/plain"):
                print("content pt: ",part.get_content())
        print("Message: ",headers.get_body())

        for att in headers.iter_attachments():
            atype = att.get_content_type()
            print(atype)
            if "application/pdf" in atype:
                pdffile = att.get_content() 
                writefile(pdffile,outfile)
        #print("------------------\n",msg_str)
    #print("headers:",headers)
    #for key in headers.keys():
    #    print("Key: ",key)


