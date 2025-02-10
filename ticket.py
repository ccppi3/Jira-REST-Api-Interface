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


# load .env file
load_dotenv()
DONOTSEND = False

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
        

        
    

class TableData:
    def __init__(self,name,data,fileName,pageNumber,creationDate):
        self.name = name
        self.data = data
        self.fileName = fileName
        self.pageNumber = pageNumber
        self.creationDate = creationDate


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
        self.tableHeaders = []
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

        match self.ticketType:
            case "arbeitsplatzwechsel":
                self.tableHeaders = ["Kürzel", "Name", "Vorname", "Abteilung Vorher", "Abteilung Neu"]
                for objct in self.data:
                    try:
                        self.tableRows.append([objct.Kürzel, objct.Name, objct.Vorname, getattr(objct, "Abteilung vorher", ""), getattr(objct, "Abteilung neu", "")])
                    except:
                        print("there are missing object attributes, assume the list is empty so cancel")
                        return 1

            case "neueintritt":
                self.tableHeaders = ["Kürzel", "Name", "Vorname", "Abteilung / Platz-Nr."]
                for objct in self.data:
                    try:
                        self.tableRows.append([objct.Kürzel, objct.Name, objct.Vorname, objct.Abteilung])
                    except:
                        print("there are missing object attributes, assume the list is empty so cancel")
                        return 1

            case "neueintritte":
                self.tableHeaders = ["Kürzel", "Name", "Vorname", "Abteilung", "Platz-Nr."]
                for objct in self.data:
                    try:
                        self.tableRows.append([objct.Kürzel, objct.Name, objct.Vorname, objct.Abteilung, getattr(objct, "Platz-Nr.", "")])
                    except:
                        print("there are missing object attributes, assume the list is empty so cancel")
                        return 1

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
        templateTable["content"].append(generateRow(self.tableHeaders))
        for row in self.tableRows:
            templateTable["content"].append(generateRow(row))
        self.table = templateTable
        print("template table:",json.dumps(templateTable))


    # Create the Payload from a Template and input data
    def create_payload(self):
        with open("TemplatePayload.json", "r") as file:
            template_data = json.load(file)
            payload = template_data
            print("before edit:",json.dumps(template_data,separators = (", "," : ")))

            payload["fields"]["description"]["content"].append(self.table)

            payload["fields"]["labels"] = self.label if self.label else []

        
        print("before dump:",payload)
        payload2 = json.dumps(payload)
        print(payload2)
        return payload2



    # Send payload to Jira Api --> Ticket creation
    def create_ticket(self):
        if self.format_data() == 1:
            print("create_ticket abort")
            return 1
        payload = self.create_payload()
        # Create ticket
        print("SUMMARY: ", self.summary)
        print()
        if DONOTSEND == False:
            print("sending...")
            yes = input("really create ticket? insert (yes):")
            if yes == "yes":
                response = requests.post(self.url, data=payload, headers=self.headers, auth=self.auth)
                print(response.status_code)
                print(response.text)
                self.id = response.json()["id"]
                url = response.json()["self"]
                self.sendAttachment(url)
            else:
                print("abort")
        else:
            print("not sending, to not spam Jira while not in production")

        # Attach PDF to ticket
        #if DONOTSEND == False:
         #   response2 = requests.post(
         #       self.url + self.id + "attachements",
         #       headers=self.headers,
         #       auth=self.auth,
         #       files={"file": (self.file_name, open(self.file_name, "rb"), "application-type")}
          #  )
            print("Response 1: ", response)
            print(response.text)
            print("Response 2: ", response2)

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
        print(response.text)
        




if __name__=="__main__":

    tables = pdf.Tables("Arbeitsplatzeinteilung KW 04 20.01.2025.pdf")
    pdf.setDebugLevel(err.ERROR)
    tables.selectPage(0)
    listTable = tables.setTableNames(["Tabelle 1", "Arbeitsplatzwechsel", "NEUEINTRITT","NEUEINTRITTE"])
    tableDataList = []

    for table in listTable:
        tables.selectTableByObj(table)
        tables.defRows(["Vorname","Name","Kürzel","Abteilung vorher","Abteilung neu","Platz-Nr","Abteilung"])
        tables.parseTable()
        objList = []
        obj_list = tables.getObjectsFromTable()
        for obj in obj_list:
            objcpy = copy.deepcopy(obj)
            objList.append(objcpy)

        tableDataList.append(TableData(table.name, objList, table.fileName, table.pageNumber, "04-03-2020"))
        print(table.name)



        ticket = Ticket(table,objList, "Arbeitsplatzeinteilung KW 04 20.01.2025.pdf", "ALLPOWER", "Neueintritt")
        if ticket.create_ticket() == 1:
            print("ticket creation aborted")
