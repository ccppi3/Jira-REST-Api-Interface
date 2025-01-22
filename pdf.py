##NOTE
##Coordination origin in PDF is bottom left!
##But pymupdf has its origin in the top left
##for reference see: https://pymupdf.readthedocs.io/en/latest/app3.html

##Measurement is in points where 1 point is 1/72 inch

import pymupdf
import json
import itertools

DEBUG = True
PRESEARCH = 5 #The search funtion finds all ocurenses of a string we work arround this issue, bc we need exact matches with a look back of 5 point and reading at this position
endTable = ["NEUEINTRITT","Arbeitsplatzwechsel","NEUEINTRITTE"]
FILENAME = "test2.pdf"
PAGENR = 0
TABLENR = 0

def log(*s):
    if DEBUG:
        print(s)
def log_json(s):
    if DEBUG:
        print("----",json.dumps(s,indent=2,default=str),"\n")

def get_all_keys(data, curr_key=[]):
    #if "text" in curr_key:
    #log("curr_key: ",curr_key)
    if isinstance(data,dict) or isinstance(data,list):
        for key, value in data.items():
            if isinstance(value,dict):
                yield from get_all_keys(value,curr_key + [key])
            elif isinstance(value,list):
                for index in value:
                    log(index)
                    log(type(index))
                    yield from get_all_keys(index,curr_key + [key])
            else:
                yield '.'.join(curr_key + [key])
def key_dump(data,needle):
    log("type of data:",type(data))
    keys = [*get_all_keys(data)]
    for key in keys:
        if needle in key:
            log(key)

class text_obj:
    def __init__(self,string,x,y):
        self.string = string
        self.x = x
        self.y = y
    def __str__(self):
        return self.string + "x: " + str(self.x) + " y: " + str(self.y)

class Border:
    def __init__(self,x,y,x2,y2,tolerance):
        self.x = x
        self.y = y
        self.x2 = x2
        self.y2 = y2
        self.tolerance = tolerance
    def check(self,x,y):
        if x > (self.x - self.tolerance) and x < (self.x2 + self.tolerance) and \
                y > (self.y - self.tolerance) and y < (self.y2 + self.tolerance):
                return True
        return False
    def __str__(self):
        return str(self.x) + "," + str(self.y) + "\n" + str(self.x2) + "," + str(self.y2) + "\n" + str(self.tolerance)

def transformPdfToPymupdf(page,x,y,w,h):
    log("orig coordinates:",x," ",y," ",w," ",h)
    return pymupdf.Rect(x,y,w,h) * ~page.transformation_matrix

def transformRect(page,rect):
    return rect * ~page.transformation_matrix

def getRectsInRange(page,border):
    data_drawings = page.get_drawings()

    for strocke in data_drawings:
        if strocke["items"][0][0]=="re":
            rec = strocke["items"]
            rect = rec[0][1]
            if border.check(rect.x0,rect.y0):
                yield rect

def getEndOfTables(page,fieldName):
    rectEnd = page.search_for(fieldName)
    if rectEnd:
        endy = rectEnd[0].y0-25
        print("Rect ACCESOIRES:",transformRect(page,rectEnd[0]))
        if len(rectEnd) > 1:
            print("WARNING: found more than one ",fieldName,"using the first one")
    else:# if no name found, assume until the end of file
        endy = 1190 
    print("endy:",endy)
    return endy

def searchForTable(page):
    tables = []
    for i in endTable:
      for a in page.search_for(i):
          tables.append(a)
    for x in tables:
        print("tables at: ",x)
    print("len table:",len(tables))
    #tables need to be sorted, it apears to be always the case else we have to implement a table sorter
    #now we stretch each one until the next one, but only in the y axis
    for i,x in enumerate(tables):
        rec =  searchTableEnd(page,x) # if no end is found then remove it from the list
        if not rec:
            tables.pop(i)
        else:
            x.y0 = rec.y0
            x.y1 = rec.y1
    #remove double tables
    for ai,a in enumerate(tables):
        for i in range(ai+1,len(tables)):
            if abs(a.x0 - tables[i].x0) < 1 and abs(a.y0 - tables[i].y0) < 1:
                print("pop table:", tables[i], "in favor of: ",a)
                tables.pop(i)


    return tables

def searchTableEnd(page,table):
    border = Border(table.x0-10,table.y0-10,table.x0,table.y0,1)
    for i in getRectsInRange(page,border):
        if (i.x1 - i.x0) < 3: #is it a line?
            print("Potential border of table:",transformRect(page,i))
            return i
    return False


def searchContentFromRowName(page,rowName,table_border):
    break_loop = False
    name = page.search_for(rowName) #returns list of Rect
    for i in name:
        if table_border.check(i.x0,i.y0): # is the row inside the border/table?
            print("----")
            newrec = pymupdf.Rect(i.x0-PRESEARCH,i.y0,i.x1+PRESEARCH,i.y1)
            real_name = page.get_textbox(newrec).strip()
            print(real_name.strip(),";",transformRect(page,newrec))
            if real_name == rowName:
                border = Border(i.x0,i.y1,i.x1,table_border.y2,5)
                temp_rect = transformPdfToPymupdf(page,i.x0,i.y1,i.x1,table_border.y2)
                print("border: ",temp_rect)
                for rect in getRectsInRange(page,border):
                    string = page.get_textbox(rect)
                    if string.strip():
                        print("string: ",string, "rects:",transformRect(page,rect))
                        for end in endTable:
                            if end in string.strip():
                                print("abort")
                                break
                        else:
                            yield string
                        break_loop = True
            else:
                print("no match")

class Entry:
    def __init__(self,kuerzel,name,vorname,abt_vorher,abt_neu,abt,platz):
        self.kuerzel = kuerzel
        self.name = name
        self.vorname = vorname
        self.abt_vorher = abt_vorher
        self.abt_neu = abt_neu
        self.abt = abt
        self.platz = platz
    def __str__(self):
        return str(self.kuerzel) + str(self.name) + str(self.vorname) + " " + str(self.abt_vorher) + " " + str(self.abt_neu) + " " + str(self.abt) + " " + str(self.platz)


doc = pymupdf.open(FILENAME)

#for page in doc:
#    text = page.get_text().encode("utf8")
#    print(text)

page = doc[PAGENR]
data = json.loads(page.get_text("json",sort=False))


jdata = json.dumps(data,indent=2,default=str)

key_dump(data,"")
#objs = []
#
#for line in data["blocks"]:
#    print("data of root:",line.keys())
#    print("type of 'lines'",type(line["lines"]))
#    text = line["lines"][0]["spans"][0]["text"]
#    x = line["lines"][0]["spans"][0]["origin"][0]
#    y = line["lines"][0]["spans"][0]["origin"][1]
#
#    objs.append(text_obj(text,x,y))
#
#for o in objs:
#    print(o,"\n")

#log_json(page.get_drawings())
## search for rects in defined range


entries = []

kuerzel = []
names = []
vorname = []
abt_vorher = []
abt_neu = []
abt = []
platz = []

tables = searchForTable(page)

for i in tables:
    #if i.y1 - i.y0 < 20:
    #    print("not valid table: ", i)
    print("modified tables:",transformRect(page,i))

table = tables[TABLENR] # select first table, means most toppest / Noth west

table_border = Border(table.x0,table.y0,1000,table.y1,5)

for i in searchContentFromRowName(page,"Kürzel",table_border):
    print(i)
    kuerzel.append(i)
for i in searchContentFromRowName(page,"Name",table_border):
    print(i)
    names.append(i)
for i in searchContentFromRowName(page,"Vorname",table_border):
    print(i)
    vorname.append(i)
for i in searchContentFromRowName(page,"Abteilung vorher",table_border):
    print(i)
    abt_vorher.append(i)
for i in searchContentFromRowName(page,"Abteilung neu",table_border):
    print(i)
    abt_neu.append(i)
for i in searchContentFromRowName(page,"Abteilung",table_border):
    print(i)
    abt.append(i)
for i in searchContentFromRowName(page,"Platz-Nr.",table_border):
    print(i)
    platz.append(i)


for (a,b,c,d,e,f,g) in itertools.zip_longest(kuerzel,names,vorname,abt_vorher,abt_neu,abt,platz):
    entries.append(Entry(a,b,c,d,e,f,g))
print("####\nResults")
for a in entries:
    print("res:",a)


file = open("json_of_pdf.json","w")

file.write(jdata)





