# Application to create Jira-Tickets automaticaly from received E-mails about new employees
# Using: Jira REST-API
# Author(s): Joel Bonini
# Last Updated: Joel Bonini 14/01/25
import requests, json, os
from requests.auth import HTTPBasicAuth
from dotenv import load_dotenv

# load .env file
load_dotenv()

# api info
url = "https://santis.atlassian.net/rest/api/3/issue"
email = "joel.bonini@santismail.ch"
token = os.getenv("TOKEN")



# Create the Payload from a Template and input data (List --> 0 = Assignee, 1 = Description, 2 = Issuetype, 3 = Labels, 4 = Priority, 5 = Summary)
def CreatePayload(data):
    with open("C:/Users/joel.bonini/.vscode/Jira/TemplatePayload.json", "r") as file:
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



# Headers so the api knows what format of data to expect and what format the response should be
headers = {
    "Accept": "application/json",
    "Content-Type": "application/json"
}



def main():
    data = [None, "Api test ticket", 10002, ["Neueintritt"], "5", "API Test"]
    payload = CreatePayload(data)
    print(payload)
    # api call
    response = requests.post(url, data=payload, headers=headers, auth=HTTPBasicAuth(email, token))

    print(response)
    print(response.text)



if __name__ == "__main__":
    main()
