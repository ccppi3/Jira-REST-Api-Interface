# Application to create Jira-Tickets automaticaly from received E-mails about new employees
# Using: Jira REST-API
# Author(s): Joel Bonini
# Last Updated: Joel Bonini 04/02/25
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

class TableData:
    def __init__(self,name,data,fileName,pageNumber,creationDate):
        self.name = name
        self.data = data
        self.fileName = fileName
        self.pageNumber = pageNumber
        self.creationDate = creationDate


class Ticket:
    def __init__(self, data, file_name, company, ticketType):
        # Api info
        self.url = "https://santis.atlassian.net/rest/api/3/issue"
        self.email = "joel.bonini@santismail.ch"
        self.token = os.getenv("TOKEN")
        # Auth
        self.auth = HTTPBasicAuth(self.email, self.token)
        # Variavble declarations
        self.ticketType = ticketType.lower()
        self.data = data
        self.file_name = file_name
        self.description = ""
        self.label = ["Neueintritt"]
        self.summary = "Testing API"
        self.id = ""
        self.table = ""
        self.company = company
        self.tableHeaders = []
        self.tableRows = []

        # Check if input is empty
        if not self.data:
            print("Input Data is empty")
        else:
            self.format_data()

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
                    self.tableRows.append([objct.Kürzel, objct.Name, objct.Vorname, getattr(objct, "Abteilung vorher", ""), getattr(objct, "Abteilung neu", "")])
            case "neueintritt":
                self.tableHeaders = ["Kürzel", "Name", "Vorname", "Abteilung / Platz-Nr."]
                for objct in self.data:
                    self.tableRows.append([objct.Kürzel, objct.Name, objct.Vorname, objct.Abteilung])
            case "neueintritte":
                self.tableHeaders = ["Kürzel", "Name", "Vorname", "Abteilung", "Platz-Nr."]
                for objct in self.data:
                    self.tableRows.append([objct.Kürzel, objct.Name, objct.Vorname, objct.Abteilung, getattr(objct, "Platz-Nr.", "")])

        self.table = {
            "type": "table",
            "content": [
                           {
                               "type": "tableRow",
                               "content": [{"type": "tableHeader", "content": [{"type": "text", "text": col}]} for col
                                           in self.tableHeaders]
                           }
                       ] + [
                           {
                               "type": "tableRow",
                               "content": [{"type": "tableCell", "content": [{"type": "text", "text": str(cell)}]} for
                                           cell in row]
                           }
                           for row in self.tableRows
                       ]
        }


    # Create the Payload from a Template and input data
    def create_payload(self):
        with open("TemplatePayload.json", "r") as file:
            template_data = json.load(file)
            payload = template_data
            payload["fields"]["description"]["content"].append(self.table)

            payload["fields"]["labels"] = self.label if self.label else []


        payload = json.dumps(payload)
        print(payload)
        return payload



    # Send payload to Jira Api --> Ticket creation
    def create_ticket(self):
        self.format_data()
        payload = self.create_payload()
        # Create ticket
        print("SUMMARY: ", self.summary)
        if DONOTSEND == False:
            print("sending...")
            response = requests.post(self.url, data=payload, headers=self.headers, auth=self.auth)
            print(response.status_code)
            print(response.text)
            self.id = response.json()["id"]
        else:
            print("not sending, to not spam Jira while not in production")

        # Attach PDF to ticket
        if DONOTSEND == False:
            response2 = requests.post(
                self.url + self.id + "attachements",
                headers=self.headers,
                auth=self.auth,
                files={"file": (self.file_name, open(self.file_name, "rb"), "application-type")}
            )
            print("Response 1: ", response)
            print(response.text)
            print("Response 2: ", response2)







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



    ticket = Ticket(objList, "Arbeitsplatzeinteilung KW 04 20.01.2025.pdf", "ALLPOWER", "Neueintritt")
    ticket.create_ticket()
