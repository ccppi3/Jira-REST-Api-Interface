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
import inspect
from pop3 import log,err
import pop3

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
        #self.master.resizable(False, False)
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
        self.tables = []
        for i,table in enumerate(tables):
            print("type table.data:",type(table.data))
            print("len vars:",len(vars(tables[i].data[0])))
            if len(tables[i].data) > 0 and len(vars(tables[i].data[0]))>0:

                print("len of data:",len(table.data))
                tabRef = ttk.Frame(self.tabControl)
                tabRef.table = table
                self.tabs.append(tabRef)

        for i,tab in enumerate(self.tabs):
            #check if table is empty if not create a table inside the tab
            if len(tables[i].data) > 0 and len(vars(tables[i].data[0]))>0:
                self.tabControl.add(tab, text=str(tables[i].name))
                confirm_wrapper = partial(self.make_sure,tab)

                # Buttons
                tab.confirm_button = tk.Button(tab, text="Create Ticket", command=confirm_wrapper)
                tab.confirm_button.grid(row=3, column=0, columnspan=1, padx=5, pady=5, sticky="nsew")
                self.create_table(tab,tables[i])


    # Handle confirm button click
    def confirm_button_handler(self,tab):
        # Close confirm window
        self.make_sure_window.destroy()
        self.loadingThread = threading.Thread(target=self.loading,args=(self.post_thread,tab))
        self.loadingThread.start()

    # Handle refresh button click
    def refresh_button_handler(self):
        self.loadingThread = threading.Thread(target=self.loading,args=(self.fetch_thread,None))
        self.loadingThread.start()

    def post_thread(self,callback_queue,tab):
        status = main.tableToTicket(tab.table)
        if type(status) == str:
            if status == "canceled":
                log("User aborted sending",level=err.INFO)
            callback_queue.put("Post Thread finished")
        else:
            
            callback_queue.put("destroy")

    def fetch_thread(self,callback_queue,dummy):
        for ret in main.run():
            if type(ret) == list:
                self.tables = ret
            else:
                self.status_label.config(text = "Status: " + str(ret))
                print("[main]",ret)
        callback_queue.put("fetch Thread finished")

    # Loading function to diable buttons and display progressbar
    def loading(self,threadFunction,tab):
        print("loading...")
        # Disable buttons
        if tab:
            tab.confirm_button.config(state=tk.DISABLED)

        self.refresh_button.config(state=tk.DISABLED)
        # Display a loading bar
        self.status_bar = ttk.Progressbar(self.master, mode="indeterminate")
        #self.status_bar.grid(row=1, column=0, padx=5, pady=5, sticky="nsew")
        self.status_bar.pack()
        self.status_bar.start()

        callback_queue = queue.Queue()
         
        self.childThread = threading.Thread(target=threadFunction,args=(callback_queue,tab))
        self.childThread.start()
        self.childThread.join() #wait for the child thread
        msg = callback_queue.get() 
        print("callbackmsg:",msg)

        self.init_tabs(self.tables)
        # Reavtivate buttons
        if tab:
            tab.confirm_button.config(state=tk.NORMAL)
        if msg == "destroy":
            tab.destroy()

        self.refresh_button.config(state=tk.NORMAL)
        # Remove loading bar
        self.status_bar.stop()
        self.status_bar.destroy()

    # Confirm Winow
    def make_sure(self,tab):
        self.refresh_button.config(state=tk.DISABLED)
        if tab:
            tab.confirm_button.config(state=tk.DISABLED)
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

        columns = []
        for item in inspect.getmembers(objList[0]):
            print("allMembers:",item)
            if not item[0].startswith('_'):
                columns.append(item[0])
                print("member to append:",item)

        for employee in objList:
            templist = []
            for item in columns:
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

        print("templist:",listOfList)



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
