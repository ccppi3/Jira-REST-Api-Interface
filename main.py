# Application to create Jira-Tickets automaticaly from received E-mails about new employees
# Using: Jira REST-API
# Author(s): Joel Bonini
# Last Updated: Joel Bonini 14/01/25
import requests, json, os
from requests.auth import HTTPBasicAuth
from dotenv import load_dotenv
from pypdf import PdfReader
import json
import pymupdf

# load .env file
load_dotenv()

# api info
url = "https://santis.atlassian.net/rest/api/3/issue"
email = "joel.bonini@santismail.ch"
token = os.getenv("TOKEN")



# Create the Payload from a Template and input data (List --> 0 = Assignee, 1 = Description, 2 = Issuetype, 3 = Labels, 4 = Priority, 5 = Summary)
def CreatePayload(data):
    with open("TemplatePayload.json", "r") as file:
        TemplateData = json.load(file)
        payload = TemplateData
        payload["fields"]["assignee"]["id"] = data[0]
        payload["fields"]["description"]["content"][0]["content"][0]["text"] = data[1]
        payload["fields"]["issuetype"]["id"] = data[2]
        payload["fields"]["labels"] = data[3]
        payload["fields"]["priority"]["id"] = data[4]
        payload["fields"]["summary"] = data[5]
    
    payload = json.dumps(payload)
    return payload



def ParsePdf():
    reader = PdfReader("Test.pdf")
    text = ""
    print(reader.resolved_objects)

    
    #for x in reader.resolved_objects:
    #    print("res_object: ",x)

    print("root obj: \n",reader.get_object(1),"\n\n") # root node
    #print("tree obj: ",reader.get_object(97)) # tree node

    x = reader.get_object(97)
    childNodesPage1 = reader.get_object(1)["/Pages"]["/Kids"][0]# get all child nodes from page 1 (index0)
    child2 = reader.get_object(1)["/StructTreeRoot"]["/K"][1]["/K"]
    print(json.dumps(child2,indent=2,default=str))
    input()
    for entry in child2:
        print("\nlist entry: ",entry)
        child3 = entry["/K"]
        print("child3: ",child3)

    #for key in child2:
    #    print("childnode[",key,"] ")
    #    print("value: ",patternNodes[key])



    numpages = reader.get_num_pages()
    print("number pages: ",numpages)

    page = reader.pages[0] #page 0 -> 1
    print("page_content: ",page.get_contents())
    
    obj = reader.get_object(44)
    print("obj: ",obj)

    text = page._extract_text(reader.get_object(1206),page.pdf)
    print("text:",text)
    #for i in range(len(reader.pages)):
    #    page = reader.pages[i]
    #    text = page.extract_text(extraction_mode="layout")
    #    print("text: ",text)

    #text = text.splitlines()

#    for i in range(len(text)):
#        if "Arbeitsplatzwechsel" in text[i]:
#            print(text[i])

    



# Headers so the api knows what format of data to expect and what format the response should be
headers = {
    "Accept": "application/json",
    "Content-Type": "application/json"
}



def main():
    data = [None, "Api test ticket", 10002, ["Neueintritt"], "5", "API Test"]
    payload = CreatePayload(data)
    #print(payload)
    # api call
    #response = requests.post(url, data=payload, headers=headers, auth=HTTPBasicAuth(email, token))
    #print(response)
    #print(response.text)
    ParsePdf()



if __name__ == "__main__":
    main()
