#This is an abstraction library on top of pymupdf, it implements algorythms to detect table beginnings and margins inside PDF files
#Author: Jonathan Wyss
#
##NOTE
##Coordination origin in PDF is bottom left!
##But pymupdf has its origin in the top left
##for reference see: https://pymupdf.readthedocs.io/en/latest/app3.html

##Measurement is in points where 1 point is 1/72 inch

import pymupdf
import json
import itertools
from enum import Enum
from itertools import zip_longest

PRESEARCH = 5 #The search funtion finds all ocurenses of a string we work arround this issue, bc we need exact matches with a look back of 5 point and reading at this position
LOGFILTER = ""
TBLLINE = 3 #sets the width ab when is considered a rect to be a line

def log_json(s):
    if int(DEBUG) > int(err.NONE):
        print("----",json.dumps(s,indent=2,default=str),"\n")
        file = open("debug-json.json","a")
        for x in s:
            pass
            #file.write(json.dumps(s,indent=2,default=str))
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

def getRectsInRange(page,border,debug=False):
    data_drawings = page.get_drawings()
    #log_json(data_drawings)
    for strocke in data_drawings:
        if strocke["items"][0][0]=="re":
            rec = strocke["items"]
            rect = rec[0][1]
            if border.check(rect.x0,rect.y0):
                if debug:
                    log("RectInRange:")
                    log_json(strocke)
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

def RectToBorder(rect):
    return Border(rect.x0,rect.y0,rect.x1,rect.y1,1)

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

def searchForTable(page,tableNames,pageNr,fileName):
    tables = []
    for name in tableNames:
      for rec in page.search_for(name):
          tables.append(Tbl(rec,name,fileName,pageNr))
    for x in tables:
        log("tables at: ",transformRect(page,x.rec))
    log("len table:",len(tables))
    for i,table in enumerate(tables):
        rec =  searchTableEnd(page,table) 
        if rec:
            table.rec.y0 = rec.y0
            table.rec.y1 = rec.y1
        else:
            log("No table end found use algo2")
            rec = searchTableDown(page,table)
            if rec:
                table.rec.y0 = rec.y0
                table.rec.y1 = rec.y1
            else:
                log("no table found, giving up")

            #search with other algorythm
    #remove double tables where if first point match remove the one witch seems to be a line
    for ai,a in enumerate(tables):
        for i in range(ai+1,len(tables)):
            if abs(a.rec.x0 - tables[i].rec.x0) < 1 and abs(a.rec.y0 - tables[i].rec.y0) < 1:
                log("pop table:", tables[i], "in favor of: ",a)
                tables.pop(i)
    return tables

def detectTableRows(page,table):
    ty=2
    tx=2
    rowNameList=[]
    fields = []

    border = Border(table.rec.x0-tx,table.rec.y0-ty,table.rec.x0+tx,table.rec.y0+ty,3)
    log("border",border)
    #search for vertical line
    for rect in getRectsInRange(page,border,debug=True):
        if isLine(rect) == "vertical":
            rectStartTable = pymupdf.Rect(rect.x0-tx,rect.y0-ty,rect.x0,rect.y0+ty+1)
            log("rectStartTable:",transformRect(page,rectStartTable))
            xLineTable = nextLineCross(page,rectStartTable,"vertical")
            log("xLineTable:",transformRect(page,xLineTable))
            #borderTitle = Border(xLineTable.x0,xLineTable.y0,xLineTable.x1,xLineTable.y1+10,5)
            rectBottomTable = pymupdf.Rect(xLineTable.x0,xLineTable.y1,xLineTable.x1,table.rec.y1)
            xLineBottomTable = nextLineCross(page,rectBottomTable,"vertical")
            log("xLBottomTable",transformRect(page,xLineBottomTable))
            
            for field in getFieldsInRange(\
                    page,pymupdf.Rect(\
                    xLineTable.x0, xLineTable.y1, xLineTable.x1,table.rec.y1 \
                    )):
                log(transformRect(page,field))
                fields.append(field)

            for fieldRow in getHeader(page,fields,xLineTable):
                log("fieldrow:",transformRect(page,fieldRow))
                for text in getTextInRange(page,RectToBorder(fieldRow)):
                    log("text",text)

    return rowNameList

def getHeader(page,fields,rectTable,thresold=10):
    lowestY = None
    filteredFields = []
    for field in fields:
        if abs(rectSize(rectTable,direction='x') - rectSize(field,direction='x')) > thresold:
            sizeTable = rectSize(rectTable,direction='x')
            sizeField = rectSize(field,direction='x')
            size = sizeTable - sizeField
            log("DataField:",transformRect(page,field),"field size:",sizeField)
            #find lowest row y
            if lowestY == None:
                lowestY = field.y0
            else:
                if field.y0 < lowestY:
                    lowestY = field.y0
            filteredFields.append(field)
    log("lowest row:",lowestY)

    fieldsRow = []
    for field in filteredFields:
        if round(field.y0,1) == round(lowestY,1):
            fieldsRow.append(field)
    return fieldsRow
    


def rectSize(rect,direction="x"):
    if direction=='x':
        return abs(rect.x1 - rect.x0)
    if direction=='y':
        return abs(rect.y1 - rect.y0)

def getFieldsInRange(page,rect,borderWidth=3):
    t=2
    border = Border(rect.x0-t,rect.y0-t,rect.x1+t,rect.y1+borderWidth,1)
    for field in getRectsInRange(page,border):
        if not isLine(field):
            yield field
#searches in a given direction and detects a line witch is 90 degree to it
def nextLineCross(page,rect,direction,borderWidth=3):
    t = 2
    if direction=="vertical":
        border = Border(rect.x0-t,rect.y0-t,rect.x1+t,rect.y1+borderWidth,1)
        for rectNew in getRectsInRange(page,border,debug=False):
            if isLine(rectNew) == "horizontal":
                return rectNew
        else:
            return False
    if direction=="horizontal":
        border = Border(rect.x1-t,rect.y0-t,rect.x1+t,rect.y0+borderWidth,1)
        for rectNew in getRectsInRange(page,border):
            if isLine(rectNew) == "vertical":
                return rectNew
        else:
            return False

def nextConnectedLine(page,rect,direction,borderWidth=3):
    t = 4
    if direction=="vertical":
        border = Border(rect.x0-t,rect.y1-t,rect.x0+t,rect.y1+borderWidth,1)
        for rectNew in getRectsInRange(page,border):
            if isLine(rectNew) == "vertical":
                return rectNew
        else:
            return False
    if direction=="horizontal":
        border = Border(rect.x1-t,rect.y0-t,rect.x1+t,rect.y0+borderWidth,1)
        for rectNew in getRectsInRange(page,border):
            if isLine(rectNew) == "horizontal":
                return rectNew
        else:
            return False


def endHorizontalLine(page,rect,borderWidth=3):
    t=2
    border2 = Border(rect.x0-t,rect.y0-t,rect.x0+t,rect.y0+borderWidth,1)
    for rectNew in getRectsInRange(page,border2):
        if isLine(rectNew) == "horizontal":
            return endHorizontalLine(page,rectNew)
    else:#when loop finished
        return rectNew.x1

            #only follow one horizontal line if there are multiple



def isLine(rect,thresold = 3):
    vertical = False
    horizontal = False
    if abs(rect.x1  - rect.x0) < thresold: 
        vertical = True
    if abs(rect.y1 - rect.y0) < thresold:
        horizontal = True
    if horizontal == True and vertical == True:
        return "dot"
    elif horizontal == True:
        return "horizontal"
    elif vertical == True:
        return "vertical"


def searchTableEnd(page,table_full): #search the table border by moving to the left
    table = table_full.rec
    border = Border(table.x0-10,table.y0-10,table.x0+10,table.y0+10,3)
    for rect in getRectsInRange(page,border):
        if abs(rect.x1 - rect.x0) < 3 and abs(rect.y0 - rect.y1) > 20: #is it a line? and filter out very short lines
            log("Potential border of table:",transformRect(page,rect))
            return rect
    return False

def searchTableDown(page,table_full):
    biggesty = 0
    smallesty = None
    t1 = 10
    table = table_full.rec
    border = Border(table.x0-t1,table.y0-t1,table.x0+t1,table.y0+t1*5,3)
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

class Tbl:
    def __init__(self,rec,name,fileName,pageNumber):
        self.rec = rec
        self.name = name
        self.fileName = fileName
        self.pageNumber = pageNumber
        self.rowNameList = []
    def getName(self):
        return self.name


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
        self.fileName = filename
        self.pages = self.Page()
        self.countPages()

    def countPages(self):
        self.pages.count = self.doc.page_count
        return self.pages.count
    def selectPage(self,nr):
        self.pages.selected = self.doc[nr]
        self.nr = nr
        return self.pages.selected
    def setTableNames(self,names):
        self.tableNames = names
        self.getTables(self.pages.selected)
        return self.tables
    def getTables(self,page):
        self.tables = searchForTable(page,self.tableNames,self.nr,self.fileName)
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
        table.border = Border(table.rec.x0,table.rec.y0,1000,table.rec.y1,5)
        for init,rowName in enumerate(self.selected_table.rowNameList):
            if init == 0:
                table.entries.append(Entry())
            _list = self.searchContentFromRowName(page,rowName,table.border)

            for index,content in enumerate(_list):
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
        listA = []          #list from algo A
        listB = []          #list from algo B
        listC = []          #combined list
        names = page.search_for(rowName) #returns list of Rect
        for recName in names:
            if table_border.check(recName.x0,recName.y0): # is the row inside the border/table?
                log("----border check of succeeded--")
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
                    log("Name border: ",temp_rect)
                    for i,rect in enumerate(getRectsInRange(page,border)):
                        string = page.get_textbox(rect)
                        if string.strip():
                            log("getRectsInRange:",string.strip())
                            listA.append(string.strip())

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
                                listB.append(string.strip())

                    for a,b in zip_longest(listA,listB,fillvalue=None):
                        if a == b and a != None and a != '' and a != ' ':
                            listC.append(a)
                        elif a != b:
                            if a == None:
                                listC.append(b)
                            elif b == None:
                                listC.append(a)
                            else:
                                log("algos give different results!")
                                combined = a + " / " + b
                                listC.append(combined)
                    log(listA,level=err.ULTRA)
                    log(listB,level=err.ULTRA)
                    log(listC,level=err.ULTRA)

                    return listC
                else:
                    log("no match")
            else:
                log("rec(",rowName,") not inside table border")
                log("table_border:",transformPdfToPymupdf(page,table_border.x,table_border.y,table_border.x2,table_border.y2))
        return []
