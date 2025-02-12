# Application to create Jira-Tickets automaticaly from received E-mails about new employees
# Using: Jira REST-API
# Author(s): Joel Bonini, Jonathan Wyss
# Last Updated: Jonathan Wyss 07/02/25
import requests, json, os
from requests.auth import HTTPBasicAuth
from dotenv import load_dotenv
import json
import pdf
from pop3 import err as err
import copy
import inspect
import pop3
from pop3 import log,err


# load .env file
load_dotenv()

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
        

        
    

class Ticket:
    def __init__(self,table, data, file_name, company, ticketType):
        # Api info
        self.tableObj = table
        self.url = "https://santis.atlassian.net/rest/api/3/issue"
        self.email = "jonathan.wyss@santismail.ch"
        self.token = os.getenv("TOKEN")
        # Auth
        self.auth = HTTPBasicAuth(self.email, self.token)
        # Variavble declarations
        self.ticketType = ticketType.lower()
        self.data = data
        self.file_name = file_name
        self.description = ""
        self.label = [ticketType]
        self.summary = ticketType + " " + file_name[29:-4]
        print(self.summary)
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
    def format_data(self):
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
    def create_payload(self):
        with open("TemplatePayload.json", "r") as file:
            template_data = json.load(file)
        payload = template_data
        print("type template:",type(template_data),"before edit:",json.dumps(template_data))
        if self.company.lower() == "allpower":
            payload["fields"]["project"]["key"] = "LZBAPW"
        payload["fields"]["description"]["content"].append(self.table)
        payload["fields"]["summary"] = self.summary
        payload["fields"]["labels"] = self.label if self.label else []
        if self.ticketType == "arbeitsplatzwechsel":
            payload["fields"]["customfield_10202"] = "Onboard new Employee"
        else:
            payload["fields"]["customfield_10202"] = "Workplace change"

        
        #try:
        payload2 = json.dumps(payload)
        print("payload2 dump:",payload2)
        #except:
        #log("error payload type:",type(payload),"\n content:",payload,level=err.ERROR)
        print("payload:",payload)
        return payload2



    # Send payload to Jira Api --> Ticket creation
    def create_ticket(self,check = True):
        if self.format_data() == 1:
            print("create_ticket abort")
            return 1
        payload = self.create_payload()
        # Create ticket
        print("SUMMARY: ", self.summary)

        if check == True:
            yes = input("really create ticket? insert (yes):")
        else:
            yes = ""
        if yes == "yes" or check == False:
            response = requests.post(self.url, data=payload, headers=self.headers, auth=self.auth)
            print(response.status_code)
            print(response.text)
            self.id = response.json()["id"]
            url = response.json()["self"]
            self.sendAttachment(url)
            print("Response 1: ", response)
            print(response.text)
            return response.status_code
        else:
            print("abort")
            return "canceled"
        

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
                    "file": (self.file_name, open(self.tableObj.fileName,"rb"), "application-type")
                    }
                )
        print("Response 1: ", response)
