import poplib
from getpass import getpass
from email.parser import Parser
from email.policy import default as default_policy
DEBUG = True

def log(*s):
    if DEBUG:
        print(s)

outfile = "tmp.pdf"
host = "pop.mail.ch"
user = "outlook-bridge.santis@mail.ch"
port = 995 


class PdfFromMail:
    def __init__(self,host,port):
        self.host = host
        self.port = port
        self.mailbox = poplib.POP3_SSL(host,port)
        log("Server Welcome:",self.mailbox.getwelcome())
    def writefile(self,buffer,filename):
        file = open(filename,"wb")
        file.write(buffer)
        file.close()
    def capa(self):
        self.capa = self.mailbox.capa()
        log("Server capabilities: ", self.capa)
    def login(self,user,password=None):
        self.user = user
        if password:
            self.password = password
        else:
            password = getpass()

        resp = self.mailbox.user(user)
        resp = self.mailbox.pass_(password)
        log("response: ",resp)

    def getAttachement(self,from_name,content_type,max_mails=25):
        maillist = self.mailbox.list()
        log("maillist: ",maillist)

        #limit max pars mails
        if len(maillist[1]) > max_mails:
            mail_count = max_mails
        else:
            mail_count = len(maillist[1])

        for mail in range(mail_count):
            msg_str = ""
            msg = self.mailbox.retr(mail+1)



            #concat the lines into one big parsable junk
            for line in range(len(msg[1])):
                msg_str = msg_str + msg[1][line].decode() + "\n"

            self.message = Parser(policy=default_policy).parsestr(msg_str,headersonly=False)
            for key in self.message.keys():
                print("Key: ",key)

            if from_name in self.message['from']:
                log("From:",self.message['from'])
                log("Content type:",self.message['Content-Type'])
                for part in self.message.walk():
                    ctype = part.get_content_type()
                    print(ctype)
                    if(ctype == "text/plain"):
                        print("content pt: ",part.get_content())
                log("Message: ",self.message.get_body())


                for att in self.message.iter_attachments():
                    atype = att.get_content_type()
                    print(atype)
                    if content_type in atype:
                        pdffile = att.get_content() 
                        self.writefile(pdffile,outfile)



#.getAttachement("Canan Akbas","application/pdf",max_mails=20)
        #print("------------------\n",msg_str)
    #print("message:",headers)
    #for key in message.keys():
    #    print("Key: ",key)

mail = PdfFromMail(host,port)
mail.login(user)
mail.getAttachement("Canan Akbas","application/pdf",max_mails=20)
