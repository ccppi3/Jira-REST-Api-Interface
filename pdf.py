##NOTE
##Coordination origin in PDF is bottom left!
##But pymupdf has its origin in the top left
##for reference see: https://pymupdf.readthedocs.io/en/latest/app3.html

##Measurement is in points where 1 point is 1/72 inch

import pymupdf
import json
import itertools
from enum import Enum

PRESEARCH = 5 #The search funtion finds all ocurenses of a string we work arround this issue, bc we need exact matches with a look back of 5 point and reading at this position
LOGFILTER = ""
TBLLINE = 3 #sets the width ab when is considered a rect to be a line

def log_json(s):
    if int(DEBUG) > int(err.NONE):
        print("----",json.dumps(s,indent=2,default=str),"\n")
        file = open("debug-json.json","a")
        for x in s:
            file.write(json.dumps(s,indent=2,default=str))
        file.close()

class err(Enum):
    NONE = 0
    ERROR = 1
    DEBUG = 2
    INFO = 3
    ULTRA = 4
    def __int__(self):
        return self.value
class Line(pymupdf.Rect):
    pass
#set DEBUG level
DEBUG = err.INFO #1 error 2 debug 3 all

def setDebugLevel(level,_filter=""):
    global DEBUG
    DEBUG = level
    global LOGFILTER
    LOGFILTER = _filter
    print("set pdf Debug level to:",DEBUG)

def log(*kwargs,level=err.INFO):
    if int(DEBUG) >= int(level):
        string = ""
        for arg in kwargs:
            string = string + str(arg) + " "
        if LOGFILTER in string:
            log.hold = 10
        if LOGFILTER in string or log.hold > 0:
            log.hold = log.hold - 1
            if int(level) != int(err.NONE):
                print("[",(str(level).split(".")[1]).replace(" ",""),"]",string)
            else:
                print("[INFO] ",string)
log.hold = 0

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
        if x > (self.x - self.tolerance) and x < (self.x2) and \
                y > (self.y - self.tolerance) and y < (self.y2+3):
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
    #log_json(data_drawings)
    for strocke in data_drawings:
        if strocke["items"][0][0]=="re":
            rec = strocke["items"]
            rect = rec[0][1]
            if border.check(rect.x0,rect.y0):
                yield rect


def getTextInRange(page,border):
    data_text = page.get_text("json",sort=True)

    text_parsed = json.loads(data_text)
    
    for block in text_parsed["blocks"]:
        if "lines" in block.keys():
            for line in block["lines"]:
                text = line["spans"][0]["text"]
                origin = line["spans"][0]["origin"]
                rectBbox = pymupdf.Rect(origin[0],origin[1],0,0)
                if border.check(rectBbox.x0,rectBbox.y0):
                    log("text:",text)
                    log("rect:",rectBbox)
                    yield text,rectBbox
        else:
            log("no lines key found")
            

    #log_json(text_parsed)


def checkBorderDown(page,rectOrigin):#returns the nearest line downwoards to a given rect
    data_drawings = page.get_drawings()
    nearestRect = pymupdf.Rect()
    for strocke in data_drawings:
        if strocke["items"][0][0]=="re":
            rec = strocke["items"]
            rect = rec[0][1]
            if rect.y1 -rect.y0 < TBLLINE: #is line?
                if rect.y0 > rectOrigin.y1:
                    if rect.x0 <= rectOrigin.x0 and rect.x1 >= rectOrigin.x1:
                        if not nearestRect:
                            nearestRect = rect
                        else:
                            if rect.y0 < nearestRect.y0:
                                nearestRect = rect

        elif strocke["items"][0][0]=="l":
            lin = strocke["items"]
            line = Line(lin[0][1],lin[0][2])
                
            #get more left point
            if line.x0 < line.x1:
                leftXPoint = line.x0
                rightXPoint = line.x1
            else:
                leftXPoint = line.x1
                rightXPoint = line.x0
            if leftXPoint <= rectOrigin.x0 and rightXPoint >= rectOrigin.x1:
                if line.y1 > rectOrigin.y1 and line.y0 > rectOrigin.y1:
                    log("table line:",line)

        else:
            log("other item types:",strocke["items"][0][0])
    return nearestRect


def getEndOfTables(page,fieldName):
    rectEnd = page.search_for(fieldName)
    if rectEnd:
        endy = rectEnd[0].y0-25
        log("Rect ACCESOIRES:",transformRect(page,rectEnd[0]))
        if len(rectEnd) > 1:
            print("WARNING: found more than one ",fieldName,"using the first one")
    else:# if no name found, assume until the end of file
        endy = 1190 
    log("endy:",endy)
    return endy

def searchForTable(page,tableNames):
    tables = []
    for i in tableNames:
      for a in page.search_for(i):
          tables.append(a)
    for x in tables:
        log("tables at: ",transformRect(page,x))
    log("len table:",len(tables))
    for i,table in enumerate(tables):
        rec =  searchTableEnd(page,table) 
        if rec:
            table.y0 = rec.y0
            table.y1 = rec.y1
        else:
            log("No table end found use algo2")
            rec = searchTableDown(page,table)
            if rec:
                table.y0 = rec.y0
                table.y1 = rec.y1
            else:
                log("no table found, giving up")

            #search with other algorythm
    #remove double tables where if first point match remove the one witch seems to be a line
    for ai,a in enumerate(tables):
        for i in range(ai+1,len(tables)):
            if abs(a.x0 - tables[i].x0) < 1 and abs(a.y0 - tables[i].y0) < 1:
                log("pop table:", tables[i], "in favor of: ",a)
                tables.pop(i)
    return tables

def searchTableEnd(page,table): #search the table border by moving to the left
    border = Border(table.x0-10,table.y0-10,table.x0+10,table.y0+10,3)
    for rect in getRectsInRange(page,border):
        if abs(rect.x1 - rect.x0) < 3 and abs(rect.y0 - rect.y1) > 20: #is it a line? and filter out very short lines
            log("Potential border of table:",transformRect(page,rect))
            return rect
    return False
def searchTableDown(page,table):
    biggesty = 0
    smallesty = None
    border = Border(table.x0-10,table.y0-10,table.x0+10,table.y0+50,3)
    for rect in getRectsInRange(page,border):
        if abs(rect.x1 - rect.x0) < 3: #and abs(rect.y0 - rect.y1) > 20: #is it a line? do not filter short lines
            log("Potential border of table:",transformRect(page,rect))
            if(rect.y1 > biggesty):
                biggesty = rect.y1
                log("biggesty[mypfcoordinate]:",biggesty)
            if not smallesty:
                smallesty = rect.y0
            if(rect.y0 < smallesty):
                smallesty = rect.y0
                log("smallest[mypfcoordinate]:",smallesty)
    if rect:
        final_rect = rect
        final_rect.y1 = biggesty + 10
        final_rect.y0 = smallesty
        log("table border calculated:",transformRect(page,final_rect))
        return final_rect
    return False




class Entry(object):
    def __init__(self):
        pass
    def __str__(self):
        string = ""
        for va in vars(self):
            string += va + " : " + str(vars(self)[va]) + " "
        return string 

class Tables:
    class State:
        INITIALIZE = 0
    class Page:
        def __init__(self):
            self.count = 0
    def __init__(self,filename):
        self.doc = pymupdf.open(filename)
        self.pages = self.Page()
        self.countPages()

    def countPages(self):
        self.pages.count = self.doc.page_count
        return self.pages.count
    def selectPage(self,nr):
        self.pages.selected = self.doc[nr]
        return self.pages.selected
    def setTableNames(self,names):
        self.tableNames = names
        self.getTables(self.pages.selected)
        return self.tables
    def getTables(self,page):
        self.tables = searchForTable(page,self.tableNames)
    def selectTable(self,nr):
        self.selected_table = self.tables[nr]
    def selectTableByObj(self,obj):
        self.selected_table = obj
    def defRows(self,rowNameList):
        table = self.selected_table
        table.rowNameList = rowNameList
    def parseTable(self):
        rowCache = []
        table = self.selected_table
        table.entries = []
        table.state = self.State()
        page = self.pages.selected
        table.border = Border(table.x0,table.y0,1000,table.y1,5)
        for init,rowName in enumerate(table.rowNameList):
            if init == 0:
                table.entries.append(Entry())
            for index,content in enumerate(self.searchContentFromRowName(page,rowName,table.border)):
                if(index > len(table.entries)-1):
                    table.entries.append(Entry())
                log("entry: ",table.entries[index],"_")
                setattr(table.entries[index],rowName,content)

    def getObjectsFromTable(self):
        table = self.selected_table
        for x in table.entries:
            yield x

    def searchContentFromRowName(self,page,rowName,table_border):
        break_loop = False
        names = page.search_for(rowName) #returns list of Rect
        for recName in names:
            if table_border.check(recName.x0,recName.y0): # is the row inside the border/table?
                log("----border check of recName succeeded--")
                newrec = pymupdf.Rect(recName.x0-PRESEARCH,recName.y0,recName.x1+PRESEARCH,recName.y1)
                real_name = page.get_textbox(newrec).strip()
                log(real_name,";",rowName,";",transformRect(page,newrec))
                if real_name == rowName:
                    posTableLine = checkBorderDown(page,recName)
                    log(rowName,"posTableLine:",transformRect(page,posTableLine))
                    if posTableLine != False:
                        border = Border(recName.x0-4,posTableLine.y1+2,recName.x1+1,table_border.y2,3)
                    else:
                        border = Border(recName.x0-4,recName.y1+2,recName.x1+1,table_border.y2,3)


                    temp_rect = transformPdfToPymupdf(page,recName.x0-4,recName.y1+2,recName.x1,table_border.y2)
                    log("border: ",temp_rect)
                    #for i,rect in enumerate(getRectsInRange(page,border)):
                    #    string = page.get_textbox(rect)
                    for i,[text,rect] in enumerate(getTextInRange(page,border)):
                        string = text
                        log("rect(",i,"):",transformRect(page,rect))
                        if string.strip():
                            log("string: ",string, "rects:",transformRect(page,rect))
                            for end in self.tableNames:
                                if end in string.strip():
                                    log("abort")
                                    break
                            else:
                                yield string
                            break_loop = True
                else:
                    log("no match")
            else:
                log("rec(rowName) not inside table border")
