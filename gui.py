# Gui to check incoming data before creating ticket on Jira
# Author(s): Joel Bonini
# Last edited: Joel Bonini 10.02.2025
import time
import tkinter as tk
from tkinter import ttk
import threading
import main

from pymupdf.mupdf import UCDN_SCRIPT_OLD_UYGHUR

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
        # Initialize variables
        self.master = master
        self.master.resizable(False, False)
        self.master.title("Jira-Check")
        self.master.iconbitmap("jira.ico")
        self.data = data
        self.ticketType = ticketType.lower()
        self.columns = ()
        self.status = ""
        # Create table widget with data
        self.create_table()

        # Buttons
        self.confirm_button = tk.Button(self.master, text="Create Ticket", command=self.make_sure)
        self.confirm_button.grid(row=3, column=0, columnspan=1, padx=5, pady=5, sticky="nsew")

        self.refresh_button = tk.Button(self.master, text="Refresh", command=self.refresh_button_handler)
        self.refresh_button.grid(row=3, column=1, columnspan=2, padx=5, pady=5, sticky="nsew")

        # Display current status
        self.status_label = tk.Label(self.master, text="Status: " + self.status)
        self.status_label.grid(row=1, column=0, padx=5, pady=5, sticky="nsew")

    # Handle confirm button click
    def confirm_button_handler(self):
        # Close confirm window
        self.make_sure_window.destroy()
        # Define threads
        self.loadingThread = threading.Thread(target=self.loading)
        self.postThread = threading.Thread(target=self.test_thread)

        # Start Threads
        self.postThread.start()
        self.loadingThread.start()


    # Handle refresh button click
    def refresh_button_handler(self):
        # Define threads
        self.loadingThread = threading.Thread(target=self.loading)
        self.fetchThread = threading.Thread(target=self.fetchProcess)

        # Start threads
        self.fetchThread.start()
        self.loadingThread.start()

    def fetchProcess(self):
        print("self:",self)
        for ret in main.run():
            if type(ret) == list:
                self.tableListData = ret
            else:
                self.status_label.config(text="Status: " + str(ret))
        self.data = self.tableListData
        self.create_table()

    def test_thread(self):
        for i in range(100):
            time.sleep(0.1)


    # Loading function to disable buttons and display progressbar
    def loading(self):
        # Display a loading bar
        self.status_bar = ttk.Progressbar(mode="indeterminate")
        self.status_bar.grid(row=1, column=1, columnspan=1, padx=5, pady=5, sticky="nsew")
        self.status_bar.start()
        # Disable buttons
        self.confirm_button.config(state=tk.DISABLED)
        self.refresh_button.config(state=tk.DISABLED)
        # Wait until Thread is finished
        try:
            while self.fetchThread.is_alive() or self.postThread.is_alive():
                time.sleep(0.1)
        except:
            pass


        # Reavtivate buttons
        self.confirm_button.config(state=tk.NORMAL)
        self.refresh_button.config(state=tk.NORMAL)
        # Remove loading bar
        self.status_bar.stop()
        self.status_bar.destroy()

    # Confirm Winow
    def make_sure(self):
        # Create window
        self.make_sure_window = tk.Toplevel(self.master)
        self.make_sure_window.title("Are you sure? ")
        self.make_sure_window.iconbitmap("jira.ico")
        self.make_sure_window.resizable(False, False)
        self.make_sure_window.geometry("300x100")

        # Display prompt
        self.make_sure_label = tk.Label(self.make_sure_window, text=f"Are you sure you want to create this Ticket?")
        self.make_sure_label.grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")

        # Confirm button
        yes_button = tk.Button(self.make_sure_window, text="Yes", command=self.confirm_button_handler)
        yes_button.grid(row=1, column=0, columnspan=1, padx=5, pady=5, sticky="nsew")
        # Cancel button
        cancel_button = tk.Button(self.make_sure_window, text="Cancel", command=self.cancel)
        cancel_button.grid(row=1, column=1, columnspan=1, padx=5, pady=5, sticky="nsew")

    # Close window on cancel
    def cancel(self):
        self.make_sure_window.destroy()

    # Function to create the table
    def create_table(self):
        templist = []
        # Destinguish ticket type & create columns accordingly
        match self.ticketType:
            case "arbeitsplatzwechsel":
                self.columns = ["Kürzel", "Name", "Vorname", "Abteilung Vorher", "Abteilung Neu"]
                for employee in self.data:
                    for item in self.columns:
                        try:
                            templist.append(getattr(employee, item))
                        except:
                            templist.append("NotFound")

            case "neueintritt":
                self.columns = ["Kürzel", "Name", "Vorname", "Abteilung"]
                for employee in self.data:
                    for item in self.columns:
                        try:
                            templist.append(getattr(employee, item))
                        except:
                            templist.append("NotFound")

            case "neueintritte":
                self.columns = ["Kürzel", "Name", "Vorname", "Abteilung", "Platz-Nr."]
                for employee in self.data:
                    for item in self.columns:
                        try:
                            templist.append(getattr(employee, item))
                        except:
                            templist.append("NotFound")
        # Create table
        self.employee_table = ttk.Treeview(self.master, columns=self.columns, show="headings", height=10)
        for item in self.columns:
            self.employee_table.heading(item, text=item)
            self.employee_table.column(item, anchor=tk.CENTER)

        # Make scrollbar
        scrollbar = ttk.Scrollbar(self.master, orient="vertical", command=self.employee_table.yview)
        self.employee_table.configure(yscroll=scrollbar.set)
        # Place scrollbar and table
        self.employee_table.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")
        scrollbar.grid(row=2, column=2, sticky="ns", pady=5)
        # Insert table data into table
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
