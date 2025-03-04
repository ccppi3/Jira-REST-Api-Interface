# Gui to check incoming data before creating ticket on Jira
# Author(s): Joel Bonini
# Last edited: Joel Bonini 10.02.2025
import time
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import customtkinter
import customtkinter as ck
import CTkToolTip as tt
import threading
from functools import partial
import queue
import main
import inspect
from pop3 import log,err
import pop3
import pythoncom
import com
import os
import sys
import configparser
import webbrowser
import sv_ttk
from pymupdf.mupdf import UCDN_SCRIPT_OLD_UYGHUR
import pdf
from pop3 import err as err
import copy

pythoncom.CoInitialize()

class TableData:
    def __init__(self,name,data,fileName,pageNumber,creationDate):
        self.name = name
        self.data = data
        self.fileName = fileName
        self.pageNumber = pageNumber
        self.creationDate = creationDate

class App:
    def __init__(self, master):
        # Set initial Color Theme
        self.lightMode()
        # Initialize variables
        self.master = master
        #self.master.resizable(False, False)
        self.master.title("JiraFlow")
        self.setupToolBar()
        self.master.iconbitmap(getResourcePath("jira.ico"))
        self.master.geometry("1250x500")
        self.columns = ()
        self.status = ""
        tabs = []
        self.tables = [] #this contains list of table object with a list of entries
        # Create table widget with data
        #self.create_table()

        self.master.protocol("WM_DELETE_WINDOW",self.exit)

        self.refresh_button = ttk.Button(self.master, text="Refresh ⭮", command=self.refresh_button_handler)
        self.refresh_button.pack(pady=10, padx=10)

        self.tabControl = ttk.Notebook(self.master)
        self.tabControl.pack(expand=1, fill="both")

        self.status_label = ttk.Label(self.master, text="Status: " + self.status)
        self.status_label.pack()
    def exit(self):
        #make sure to also close running threads
        os._exit(0)

    def setupToolBar(self):
        menubar = tk.Menu(self.master)
        settingsMenu = tk.Menu(menubar, tearoff=0)
        settingsMenu.add_command(label="Config",command = lambda: Config(self.master))

        helpMenu = tk.Menu(menubar, tearoff=0)
        helpMenu.add_command(label="Credits",command = self.showCredits)
        helpMenu.add_command(label="Help",command = self.showHelp)

        editMenu = tk.Menu(menubar, tearoff=0)
        editMenu.add_command(label="Forget last Mail",command = com.rmLastEntryIDDB)

        appearanceMenu = tk.Menu(menubar, tearoff=0)
        appearanceMenu.add_command(label="Dark Mode", command=self.darkMode)
        appearanceMenu.add_command(label="Light Mode", command=self.lightMode)

        menubar.add_cascade(label="Edit",menu=editMenu)
        menubar.add_cascade(label="Settings",menu=settingsMenu)
        menubar.add_cascade(label="About",menu=helpMenu)
        menubar.add_cascade(label="Appearance",menu=appearanceMenu)

        self.master.config(menu=menubar)

    def darkMode(self):
        sv_ttk.set_theme("dark")
        customtkinter.set_appearance_mode("dark")

    def lightMode(self):
        sv_ttk.set_theme("light")
        customtkinter.set_appearance_mode("light")

    def showCredits(self):
        messagebox.showinfo("Credits","This software was developed by \nJoel Bonini \nand\nJonathan Wyss \nin santis for internal usage")
    def showHelp(self):
        helpWindow = tk.Toplevel(self.master)
        helpWindow.iconbitmap(getResourcePath("jira.ico"))
        helpWindow.title("Help")

        l1 = ttk.Label(helpWindow,text="Please refer to the following:")
        l1.pack(padx=10,pady=10)
        l2 = ttk.Label(helpWindow,text="http://wiki.santisedu.local/books/jiraflow",cursor="hand2",relief='raised',foreground='blue')
        l2.pack(padx=10,pady=10)
        l3 = ttk.Label(helpWindow,text="https://github.com/ccppi3/Jira-REST-Api-Interface",cursor="hand2",relief='raised',foreground='blue')
        l3.pack(padx=10,pady=10)
        l2.bind("<Button-1>",lambda e:self.hyperlinkCallback(l2.cget("text")))
        l3.bind("<Button-1>",lambda e:self.hyperlinkCallback(l3.cget("text")))

    def hyperlinkCallback(self,url):
        webbrowser.open_new_tab(url)

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
                self.tabControl.add(tab, text=str(tables[i].name).capitalize() + " \n " + str(tables[i].creationDate.strftime("%x")))
                confirm_wrapper = partial(self.make_sure,tab)

                # Buttons
                tab.confirm_button = ttk.Button(tab, text="Create Ticket", width=15, command=confirm_wrapper)
                tab.confirm_button.grid(row=3, column=0, columnspan=2, pady=5, padx=5, sticky="W")
                self.create_table(tab,tables[i])

    # Handle confirm button click
    def confirm_button_handler(self,tab):
        # Close confirm window
        self.make_sure_window.destroy()
        self.loadingThread = threading.Thread(target=self.loading,args=(self.post_thread,tab))
        self.loadingThread.start()

    # Handle refresh button click
    def refresh_button_handler(self):
        outlook = com.init()
        pythoncom.CoInitialize()
        
        self.loadingThread = threading.Thread(target=self.loading,args=(self.fetch_thread,outlook))
        self.loadingThread.start()

    def post_thread(self,callback_queue,tab):
        for status in main.tableToTicket(tab.table):
            if type(status) == str:
                self.status_label.config(text = "Status: " + str(status))
                if status == "canceled":
                    log("User aborted sending",level=err.INFO)
                    callback_queue.put("Post Thread finished")
            else: #when returns type is intiger we have finised
                if type(status) == int:
                    if status == 401:
                        messagebox.showerror("Error",str(status) + "\nAccess to Jira was denied, is your jira token valid?")
                    elif status >= 300:
                        messagebox.showerror("Error",str(status) + "\nHttpError")

                callback_queue.put("destroy")

    def fetch_thread(self,callback_queue,outlook):
        pythoncom.CoInitialize()
        for ret in main.run(outlook):
            if type(ret) == list:
                self.tables = ret
            else:
                if "ERROR" in ret:
                    log("fetch_thread:",ret,level=err.ERROR)
                    self.status_label.config(text = ret)
                    if "MAPI" in ret:
                        messagebox.showerror("Error",ret + "\nMaybe outlook is not started?\nPlease make sure Outlook is running at least in the background")
                    else:
                        messagebox.showerror("ERROR",ret)
                else:
                    self.status_label.config(text = "Status: " + str(ret))
                    print("[main]",ret)
        callback_queue.put("fetch Thread finished")

    # Loading function to diable buttons and display progressbar
    def loading(self,threadFunction,tab):
        print("loading...")
        # Disable buttons
        if tab:
            try:
                tab.confirm_button.config(state=tk.DISABLED)
            except:
                pass

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
        try:
            tab.confirm_button.config(state=tk.NORMAL)
        except:
            log("noTabs")
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
        self.make_sure_window.iconbitmap(getResourcePath("jira.ico"))
        self.make_sure_window.resizable(False, False)
        self.make_sure_window.geometry("280x75")


        # Display prompt
        self.make_sure_label = ttk.Label(self.make_sure_window, text=f"Are you sure you want to create this Ticket?")
        self.make_sure_label.grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")

        # Confirm button
        confirm_wrapper = partial(self.confirm_button_handler,tab)
        yes_button = ttk.Button(self.make_sure_window, text="Yes", command=confirm_wrapper)
        yes_button.grid(row=1, column=0, columnspan=1, padx=5, pady=5, sticky="nsew")
        # Cancel button
        cancel_wrapper = partial(self.cancel,tab)
        self.make_sure_window.protocol("WM_DELETE_WINDOW",cancel_wrapper)
        cancel_button = ttk.Button(self.make_sure_window, text="Cancel", command=cancel_wrapper)
        cancel_button.grid(row=1, column=1, columnspan=1, padx=5, pady=5, sticky="nsew")

    # Close window on cancel
    def cancel(self,tab):
        tab.confirm_button.config(state=tk.NORMAL)
        self.refresh_button.config(state=tk.NORMAL)
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
        scrollbar.grid(row=2, column=2, sticky="nsew", pady=5)
        # Configure rows and columns to expand them onto whole screen
        tab.grid_rowconfigure(2, weight=1)  # Expand row 2
        tab.grid_columnconfigure(0, weight=1)  # expand column 0
        tab.grid_columnconfigure(1, weight=1)  # expand column 1.
        # Insert table data into table

        print("templist:",listOfList)

def getResourcePath(relPath):
    try:
        basePath = sys._MEIPASS
    except Exception:
        basePath = os.path.abspath(".")
    return os.path.join(basePath,relPath)

def _dir(_object): #wrapper function to exclude internal objects
    _list=[]
    for obj in dir(_object):
        if not obj.startswith("__"):
            _list.append(obj)
    return _list

#gui configuration window
class Config():
    class tkObjects:
        pass
    def __init__(self,root):
        helpMessage="default help"
        listOfEntries = self._loadConfig()

        self.window = tk.Toplevel(root)
        self.window.iconbitmap(getResourcePath("jira.ico"))
        self.window.title("Config")
        self.window.grab_set()

        helpText = self._loadHelpText()
        
        #we create an object attribute for each config entry (metaprogramming)
        #we add a toolbox for each textbox
        for i,entry in enumerate(listOfEntries):
            tmpObj = ck.CTkLabel(self.window,text=entry[0])
            tmpObj.grid(row=i,column=0)

            tmpObj = ck.CTkTextbox(self.window,height=1,width = 500)
            tmpObj.insert(tk.END,entry[1])
            tmpObj.grid(row=i,column=1)

            for txt in helpText:
                if(txt[0] == entry[0]):
                    tt.CTkToolTip(tmpObj,message=txt[1])

            tmpObj.value = entry[1]
            tmpObj.id = i
            setattr(self.tkObjects,entry[0],tmpObj)

        #we apply batch configs for all elements in page
        self._tkConfigAll()

        for objName in _dir(self.tkObjects):
            obj = getattr(self.tkObjects,objName)

        self.buttonSave = ttk.Button(self.window,text="Save",command=self._saveConfig)
        self.buttonSave.grid(row = i+1,column=0, pady=10, padx=10)


    def _tkConfigAll(self):
        for slave in self.window.grid_slaves():
            slave.grid(padx=10,pady=10)
        
    def _loadConfig(self):
        configList=[]
        self.config = configparser.ConfigParser()
        self.config.optionxform = str
        self.config.read(getResourcePath(".env"))
        try:
            sections=self.config.sections()
        except:
            print("Now sections! You are using an old style .env, please update the config to have a [CONFIG] Section")
            return "Section Error"
        else:
            try:
                configSection = self.config['CONFIG']
            except:
                print("You are using an old style .env, please update the config to have a [CONFIG] Section")
                return "Section Error"
            else:
                for element in configSection:
                    configList.append((element,configSection[element]))
                return configList

    def _loadHelpText(self):
        helpList=[]
        self.helpParser = configparser.ConfigParser()
        self.helpParser.optionxform = str
        self.helpParser.read(getResourcePath("help.txt"))
        try:
            sections = self.helpParser.sections()
        except:
            print("found no parsing keys in the help.txt file, does the file exist?")
            return "error parsing help.txt"
        else:
            for  section in sections:
                try:
                    helpList.append((section,self.helpParser[section]['text']))
                except:
                    print("Error parsing help.txt file, look for syntax errors")
                    return "Parsing error help.txt"
            return helpList

    def _saveConfig(self):
        sections=self.config.sections()
        try:
            configSection = self.config['CONFIG']
        except:
            return "config section not found"

        for objName in _dir(self.tkObjects):
            obj = getattr(self.tkObjects,objName)
            self.config.set('CONFIG',objName,obj.get('1.0',tk.END))
        
        try:
            with open(getResourcePath(".env"),'w') as configFile:
                      self.config.write(configFile)
        except:
            print("ERROR writing .env file!")
            messagebox.showerror("Config","Config has not been written")
            return "error writing file"
        else:
            print("Wrote .env file")
            messagebox.showinfo("Config","Config sucessfull saved")
     


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
