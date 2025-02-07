# Gui to check incoming data before creating ticket on Jira
# Author(s): Joel Bonini
# Last edited: Joel Bonini 07.02.2025
import tkinter as tk
from tkinter import ttk
import requests, json, os
from requests.auth import HTTPBasicAuth
from dotenv import load_dotenv
import json
import pdf
from pop3 import err as err
import copy

class TableData:
    def __init__(self,name,data,fileName,pageNumber,creationDate):
        self.name = name
        self.data = data
        self.fileName = fileName
        self.pageNumber = pageNumber
        self.creationDate = creationDate

class App:
    def __init__(self, master, data, ticketType):
        self.master = master
        self.master.title("Jira-Check")
        self.master.iconbitmap("jira.ico")
        self.data = data
        self.ticketType = ticketType.lower()
        self.columns = ()

        self.create_table()
        self.confirm_button = tk.Button(self.master, text="Confirm & Create Ticket")
        self.confirm_button.place(x=5, y=240)

    def create_table(self):
        templist = []
        match self.ticketType:
            case "arbeitsplatzwechsel":
                self.columns = ["Kürzel", "Name", "Vorname", "Abteilung Vorher", "Abteilung Neu"]
                for employee in self.data:
                    for item in self.columns:
                        templist.append(getattr(employee, item))

            case "neueintritt":
                self.columns = ["Kürzel", "Name", "Vorname", "Abteilung"]
                for employee in self.data:
                    for item in self.columns:
                        templist.append(getattr(employee, item))

            case "neueintritte":
                self.columns = ["Kürzel", "Name", "Vorname", "Abteilung", "Platz-Nr."]
                for employee in self.data:
                    for item in self.columns:
                        templist.append(getattr(employee, item))

        self.employee_table = ttk.Treeview(self.master, columns=self.columns, show="headings", height=10)

        for item in self.columns:
            self.employee_table.heading(item, text=item)
            self.employee_table.column(item, anchor=tk.CENTER)

        scrollbar = ttk.Scrollbar(self.master, orient="vertical", command=self.employee_table.yview)
        self.employee_table.configure(yscroll=scrollbar.set)

        self.employee_table.grid(row=2, column=0, columnspan=2, padx=5, pady=(5, 45), sticky="nsew")
        scrollbar.grid(row=2, column=2, sticky="ns", pady=5)

        self.employee_table.insert("", tk.END, values=templist)


if __name__ == "__main__":
    tables = pdf.Tables("Arbeitsplatzeinteilung KW 04 20.01.2025.pdf")
    pdf.setDebugLevel(err.ERROR)
    tables.selectPage(0)
    listTable = tables.setTableNames(["Tabelle 1", "Arbeitsplatzwechsel", "NEUEINTRITT", "NEUEINTRITTE"])
    tableDataList = []

    for table in listTable:
        tables.selectTableByObj(table)
        tables.defRows(["Vorname", "Name", "Kürzel", "Abteilung vorher", "Abteilung neu", "Platz-Nr", "Abteilung"])
        tables.parseTable()
        objList = []
        obj_list = tables.getObjectsFromTable()
        for obj in obj_list:
            objcpy = copy.deepcopy(obj)
            objList.append(objcpy)

        tableDataList.append(TableData(table.name, objList, table.fileName, table.pageNumber, "04-03-2020"))
    root = tk.Tk()
    app = App(root, objList, "neueintritt")
    root.mainloop()
