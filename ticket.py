# Application to create Jira-Tickets automaticaly from received E-mails about new employees
# Using: Jira REST-API
# Author(s): Joel Bonini, Jonathan Wyss
# Last Updated: Joel Bonini 11.03.25

# --- Imports ---
import requests, json, os
from requests.auth import HTTPBasicAuth
from dotenv import load_dotenv
import json
import pdf
from pop3 import err as err
import copy

# load .env file
load_dotenv()
import inspect
import com
from pop3 import log,err
from gui import getResourcePath



# Load .env file
APPLOCAL = os.path.join(os.path.expanduser("~"),"AppData","Local")
APPNAME = "Jira-Flow"
APPDIR = os.path.join(APPLOCAL,APPNAME)
APPDIR = APPDIR + "\\"
load_dotenv(APPDIR + "config")

# Gen Rows for table in ticket
def generateRow(_list):
    header = {
            "type": "tableRow",
            "content": []
            }
    for i,arg in enumerate(_list):
        content = { 
            "type": "tableCell",
            "attrs": {},
            "content": [
            {
                "type": "paragraph",
                "content": [
                {
                "type": "text",
                "text": arg
                }
                ]
            }]
        }
        header["content"].append(content)
    return header



# Ticket Class
class Ticket:
    def __init__(self,table, data, fileName, company, ticketType):
        # Api info
        self.tableObj = table
        self.url = os.getenv('IssueUrl')
        self.email = os.getenv('EMail')
        self.token = os.getenv("TOKEN")
        # Request Type keys
        self.requestCustomField = os.getenv('RequestTypeField')
        self.onboardKey = os.getenv('RequestOnBoardKey')
        self.changeKey = os.getenv('RequestChangeKey')
        # Auth
        self.auth = HTTPBasicAuth(self.email, self.token)
        # Variable declarations
        self.ticketType = ticketType.lower()
        self.data = data
        self.fileName = fileName
        self.description = ""
        self.label = [ticketType]
        self.summary = ticketType + " " + fileName[29:-4]
        self.id = ""
        self.table = ""
        self.company = company
        self.tableRows = []

        # Check if input is empty
        if not self.data:
            print("Input Data is empty")

        # Headers so the api knows what format of data to expect and what format the response should be
        self.headers = {
            # Expected request format
            "Accept": "application/json",
            # Response format
            "Content-Type": "application/json",
            # Header for file attachements
            "X-Atlassian-Token": "no-check"
        }


    # Format data in a way to use the payload template easily
    def formatData(self):
        self.tableRows = []
        tableHeaders = []
        for item in inspect.getmembers(self.data[0]):
            if not item[0].startswith('_'):
                tableHeaders.append(str(item[0]))
                log("append header member:",item,level=err.INFO)

        for objct in self.data:
            templist = []
            for item in tableHeaders:
                try:
                    templist.append(getattr(objct,item))
                except:
                    print("there are missing object attributes, assume the list is empty so cancel")
                    return 1
            self.tableRows.append(templist)

        templateTable = {
                    "type": "table",
                    "attrs": {
                        "isNumberColumnEnabled": False,
                        "layout": "center",
                        "width": 900,
                        "displayMode": "default"
                    },
                    "content": [
                        ]
                    }
        templateTable["content"].append(generateRow(tableHeaders))
        for row in self.tableRows:
            templateTable["content"].append(generateRow(row))
        self.table = templateTable
        print("template table:",json.dumps(templateTable))


    # Create the Payload from a Template and input data
    def createPayload(self):
        with open(getResourcePath("TemplatePayload.json"), "r") as file:
            templateData = json.load(file)
        payload = templateData
        print("type template:",type(templateData),"before edit:",json.dumps(templateData))
        if self.company.lower() == "allpower":
            portalKey = "lzbapw" # Portalkey is needed for customfield
            payload["fields"]["project"]["key"] = "LZBAPW"
        else:
            portalKey = "lzbict"
        payload["fields"]["description"]["content"].append(self.table)
        payload["fields"]["summary"] = self.company.capitalize() + " " + self.summary.capitalize()
        payload["fields"]["labels"] = self.label if self.label else []

        if self.ticketType != "arbeitsplatzwechsel":
            payload["fields"][self.requestCustomField] = portalKey + "/" + self.onboardKey  #portalkey / requesttype key
        else:
            payload["fields"][self.requestCustomField] = portalKey + "/" + self.changeKey

        payload2 = json.dumps(payload)
        print("payload2 dump:",payload2)
        print("payload:",payload)
        return payload2

    # Send payload to Jira Api --> Ticket creation
    def createTicket(self,check = True):
        if self.formatData() == 1:
            print("createTicket abort")
            yield 1
        payload = self.createPayload()
        # Debug summary
        print("SUMMARY: ", self.summary)
        # Alternate Check for devs to not spam Jira
        if check == True:
            yes = input("really create ticket? insert (yes):")
        else:
            yes = ""
        if yes == "yes" or check == False:
            # Create Ticket
            yield "Posting request"
            response = requests.post(self.url, data=payload, headers=self.headers, auth=self.auth)
            print(response.status_code)
            print(response.text)
            try:
                self.id = response.json()["id"]
            except:
                yield "no valid response"
            else:
                url = response.json()["self"]
                yield "sending Attachement"
                self.sendAttachment(url)
            print("Response 1: ", response)
            print(response.text)
            yield response.status_code
        else:
            print("abort")
            yield "canceled"
        

    def sendAttachment(self,ticketUrl):
        ticketUrl = ticketUrl + "/attachments"
        print("self.table:",self.tableObj)
        print("post url:",ticketUrl)
        headers = {
                "Accept": "application/json",
                "X-Atlassian-Token": "no-check"
                }
        response = requests.post(
                ticketUrl,
                headers=headers,
                auth=self.auth,
                files = {
                    "file": (self.fileName, open(self.tableObj.fileName, "rb"), "application-type")
                    }
                )
        print("Response 1: ", response)
