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
DONOTSENND = True


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
            "Content-Type": "application/json"
        }

    # Format data in a way to use the payload template easily
    def format_data(self):
        for obj in self.data:
            print("count obj:",obj)
            print(dir(obj))

        for objct in self.data:
            if "Abteilung vorher" in  dir(objct):
                self.label = ["Arbeitsplatzwechsel"]
                string = self.file_name.split()[3].split(".pdf")[0]
                self.summary = self.company + " Arbeitplatzwechsel " +  string
                print(self.summary)
                self.description += objct.Kürzel + " " + objct.Name + " " + objct.Vorname + " "
                self.description += getattr(objct, "Abteilung vorher") + " ➡️ " + getattr(objct, "Abteilung neu") + "\n"

            else:
                self.label = ["Neueintritt"]
                self.summary = self.company + " Neueintritte " + self.file_name.split()[3]
                if self.company == "SANTIS":
                    self.description += objct.Kürzel + " " + objct.Name + " " + objct.Vorname + " " + getattr(objct, "Platz-Nr") + "\n"
                else:
                    self.description += objct.Kürzel + objct.Name + objct.Vorname + objct.Abteilung

    # Create the Payload from a Template and input data
    def create_payload(self):
        with open("TemplatePayload.json", "r") as file:
            template_data = json.load(file)
            payload = template_data
            payload["fields"]["assignee"]["id"] = None
            payload["fields"]["description"]["content"][0]["content"][0]["text"] = self.description
            payload["fields"]["issuetype"]["id"] = 10002
            payload["fields"]["labels"] = self.label
            payload["fields"]["priority"]["id"] = "3"
            payload["fields"]["summary"] = self.summary

        payload = json.dumps(payload)
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
        else:
            print("not sending, to not spam Jira while in production")

        # Get ticket id
        print(response)
        print(response.text)
        self.id = response.json()["id"]
        # Attach PDF to ticket
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

    objlist = []
    for table in listTable:
        tables.selectTableByObj(table)
        tables.defRows(["Vorname","Name","Kürzel","Abteilung vorher","Abteilung neu","Platz-Nr","Abteilung"])
        tables.parseTable()

        obj_list = tables.getObjectsFromTable()
        for obj in obj_list:
            objcpy = copy.deepcopy(obj)
            objlist.append(objcpy)
            print(obj)
            print(obj_list)


    ticket = Ticket(objlist, "Arbeitsplatzeinteilung KW 04 20.01.2025.pdf", "ALLPOWER")
    ticket.create_ticket()
