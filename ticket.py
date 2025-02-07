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



class Ticket:
    def __init__(self, data, file_name, company):
        # Api info
        self.url = "https://santis.atlassian.net/rest/api/3/issue"
        self.email = "joel.bonini@santismail.ch"
        self.token = os.getenv("TOKEN")
        # Auth
        self.auth = HTTPBasicAuth(self.email, self.token)
        # General Info
        self.data = data
        self.file_name = file_name
        self.description = ""
        self.label = []
        self.summary = ""
        self.id = ""
        self.company = company
        self.table_data = [["Kürzel", "Name", "Vorname"]]
        self.table = None


        # Headers so the api knows what format of data to expect and what format the response should be
        self.headers = {
            # Expected request format
            "Accept": "application/json",
            # Response format
            "Content-Type": "application/json"
        }

    # Format data in a way to use the payload template easily
    def format_data(self):
        self.summary = "TEST TICKET"
        for objct in self.data:
            self.table_data.append([objct.Kürzel, objct.Name, objct.Vorname])

        table_content = []
        for i, row in enumerate(self.table_data):
            row_content = []
            for cell in row:
                cell_type = "tableHeader" if i == 0 else "tableCell"
                row_content.append({
                    "type": cell_type,
                    "content": [{"type": "text", "text": cell}]
                })
            table_content.append({"type": "tableRow", "content": row_content})

        return {
            "type": "table",
            "content": table_content
        }



    # Create the Payload from a Template and input data
    def create_payload(self):
        with open("TemplatePayload.json", "r") as file:
            template_data = json.load(file)
            payload = template_data
            payload["fields"]["assignee"]["id"] = None
            payload["fields"]["description"]["content"] = self.table
            payload["fields"]["issuetype"]["id"] = 10002
            payload["fields"]["labels"] = self.label
            payload["fields"]["priority"]["id"] = "3"
            payload["fields"]["summary"] = self.summary

        payload = json.dumps(payload)
        return payload

    # Send payload to Jira Api --> Ticket creation
    def create_ticket(self):
        self.table = self.format_data()
        payload = self.create_payload()
        # Create ticket
        print("SUMMARY: ", self.summary)
        response = requests.post(self.url, data=payload, headers=self.headers, auth=self.auth)
        print(response)
        print(response.text)
        # Get ticket id
        self.id = response.json()["id"]
        # Attach PDF to ticket
        response2 = requests.post(
            f"{self.url}/{self.id}/attachments",
            headers=self.headers,
            auth=self.auth,
            files={"file": (self.file_name, open(self.file_name, "rb"), "application-type")}
        )
        print("Response 2: ", response2)









tables = pdf.Tables("Arbeitsplatzeinteilung KW 04 20.01.2025.pdf")
pdf.setDebugLevel(err.ERROR)
tables.selectPage(0)
listTable = tables.setTableNames(["NEUEINTRITT","NEUEINTRITTE", "Arbeitsplatzwechsel", "ARBEITSPLATZWECHSEL" ])

objlist = []
for table in listTable:
    tables.selectTableByObj(table)
    tables.defRows(["Vorname","Name","Kürzel","Abteilung vorher","Abteilung neu","Platz-Nr","Abteilung"])
    tables.parseTable()

    obj_list = tables.getObjectsFromTable()
    for obj in obj_list:
        objcpy = copy.deepcopy(obj)
        objlist.append(objcpy)



ticket = Ticket(objlist, "Arbeitsplatzeinteilung KW 04 20.01.2025.pdf", "ALLPOWER")
ticket.create_ticket()
