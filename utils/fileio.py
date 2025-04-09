import tkinter
import tkinterdnd2
import re

class fileio:
    def __init__(self):
        self.filename = "None"
        self.filetypes = (
                ("PDF-Files","*.pdf"),
                )
    def select_file(self,label):
        self.filename = tkinter.filedialog.askopenfilename(
                title="Open a FIle to convert",
                filetypes=self.filetypes,
                )
        print("label:",label)
        label['text'] = self.filename

def dropFiles(listbox,data):
    print("pre split:",data)
    filelist = listbox.tk.splitlist(data)#re.split("} | {",data)
    print(filelist)
    for line in filelist:
        if line[len(line)-4:] == ".pdf":
            line=line.replace("{","").replace("}","")
            name = line.split("/")
            name = name[len(name)-1]
            listbox.insert("","end",text=name,value=str(line))
        else:
            print("wrong filetype:",line[len(line)-4:])
    return filelist

def DeleteSelection(listbox):
    items = listbox.selection()
    for item in items:
        listbox.delete(item)

