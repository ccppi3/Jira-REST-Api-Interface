# Gui to check incoming data before creating ticket on Jira
# Author(s): Joel Bonini
# Last edited: Joel Bonini 10.02.2025
import time
import tkinter as tk
from tkinter import ttk
import threading
from functools import partial
import queue
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
        self.master.geometry("850x500")
        self.data = data
        self.ticketType = ticketType.lower()
        self.columns = ()
        self.status = ""
        tabs = []
        self.tables = [] #this contains list of table object with a list of entries
        # Create table widget with data
        #self.create_table()

        self.refresh_button = tk.Button(self.master, text="Refresh", command=self.refresh_button_handler)
        self.refresh_button.pack()

        self.tabControl = ttk.Notebook(self.master)
        self.tabControl.pack(expand=1, fill="both")

        self.status_label = tk.Label(self.master, text="Status: " + self.status)
        self.status_label.pack()

    def init_tabs(self,tables):
        self.tabs = []
        for i in tables:
            tabRef = ttk.Frame(self.tabControl)
            self.tabs.append(tabRef)

        for i,tab in enumerate(self.tabs):
            self.tabControl.add(tab, text="werfjghk")
            confirm_wrapper = partial(self.make_sure,tab)

            # Buttons
            tab.confirm_button = tk.Button(tab, text="Create Ticket", command=confirm_wrapper)
            tab.confirm_button.grid(row=3, column=0, columnspan=1, padx=5, pady=5, sticky="nsew")
            print("tab init:",tab)
            self.create_table(tab,tables[i])





    # Handle confirm button click
    def confirm_button_handler(self,tab):
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
        self.loadingThread = threading.Thread(target=self.loading,args=(self.fetch_thread,))
        self.loadingThread.start()

    def test_thread(self,callback_queue):
        for i in range(100):
            time.sleep(0.1)
        callback_queue.put("Thread finished")
    def fetch_thread(self,callback_queue):
        for ret in main.run():
            if type(ret) == list:
                self.tables = ret
            else:
                print("[main]",ret)
        callback_queue.put("Thread finished")

    # Loading function to diable buttons and display progressbar
    def loading(self,threadFunction):
        print("loading...")
        # Disable buttons
        try:
            for x in self.tabs:
                x.confirm_button.config(state=tk.DISABLED)
        except:
            print("no tabs available")

        self.refresh_button.config(state=tk.DISABLED)
        # Display a loading bar
        self.status_bar = ttk.Progressbar(self.master, mode="indeterminate")
        #self.status_bar.grid(row=1, column=0, padx=5, pady=5, sticky="nsew")
        self.status_bar.pack()
        self.status_bar.start()
        

        callback_queue = queue.Queue()
         
        self.childThread = threading.Thread(target=threadFunction,args=(callback_queue,))
        self.childThread.start()

        self.childThread.join() #wait for the child thread
        
        print("callbackmsg:",callback_queue.get())

        self.init_tabs(self.tables)

        # Reavtivate buttons
        try:
            for x in self.tabs:
                x.confirm_button.config(state=tk.NORMAL)
        except:
            print("no tabs avalaible")

        self.refresh_button.config(state=tk.NORMAL)
        # Remove loading bar
        self.status_bar.stop()
        self.status_bar.destroy()

    # Confirm Winow
    def make_sure(self,tab):
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
        confirm_wrapper = partial(self.confirm_button_handler,tab)
        yes_button = tk.Button(self.make_sure_window, text="Yes", command=confirm_wrapper)
        yes_button.grid(row=1, column=0, columnspan=1, padx=5, pady=5, sticky="nsew")
        # Cancel button
        cancel_button = tk.Button(self.make_sure_window, text="Cancel", command=self.cancel)
        cancel_button.grid(row=1, column=1, columnspan=1, padx=5, pady=5, sticky="nsew")

    # Close window on cancel
    def cancel(self):
        self.make_sure_window.destroy()

    # Function to create the table
    def create_table(self,tab,tables):
        objList = tables.data
        listOfList = []
        print("objList",objList)
        # Destinguish ticket type & create columns accordingly
        if "arbeitsplatzwechsel" in str(tables.name).lower():
            columns = ["Kürzel", "Name", "Vorname", "Abteilung Vorher", "Abteilung Neu"]
            for employee in objList:
                templist = []
                for item in employee.dir():
                    try:
                        templist.append(getattr(employee, item))
                    except:
                        templist.append("NotFound")
                listOfList.append(templist)

        elif "neueintritt" in str(tables.name).lower():
            columns = ["Kürzel", "Name", "Vorname", "Abteilung"]
            for employee in objList:
                templist = []
                for item in employee.dir():
                    try:
                        templist.append(getattr(employee, item))
                    except:
                        templist.append("NotFound")
                listOfList.append(templist)

        elif "neueintritte" in str(tables.name).lower():
            columns = ["Kürzel", "Name", "Vorname", "Abteilung", "Platz-Nr."]
            for employee in objList:
                templist = []
                for item in employee.dir():
                    try:
                        templist.append(getattr(employee, item))
                    except:
                        templist.append("NotFound")
                listOfList.append(templist)
        # Create table
        tables.employee_table = ttk.Treeview(tab, columns=columns, show="headings", height=5)
        for item in columns:
            tables.employee_table.heading(item, text=item)
            tables.employee_table.column(item, anchor=tk.CENTER)
        for lists in listOfList:
            tables.employee_table.insert("", tk.END, values=lists)

        # Make scrollbar
        scrollbar = ttk.Scrollbar(tab, orient="vertical", command=tables.employee_table.yview)
        tables.employee_table.configure(yscroll=scrollbar.set)
        # Place scrollbar and table
        tables.employee_table.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")
        scrollbar.grid(row=2, column=2, sticky="ns", pady=5)
        # Insert table data into table

        print("templist:",templist)



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
