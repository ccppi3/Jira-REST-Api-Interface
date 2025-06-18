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
import traceback

PRESEARCH = 5 #The search funtion finds all ocurenses of a string we work arround this issue, bc we need exact matches with a look back of 5 point and reading at this position
LOGFILTER = ""
TBLLINE = 3 #sets the width ab when is considered a rect to be a line
MAX_BORDER_WIDTH=1000

def trace():
    for line in traceback.format_stack():
        print("traceback",line)

def log_json(s):
    if int(DEBUG) > int(err.NONE):
        #print("----",json.dumps(s,indent=2,default=str),"\n")
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
        if x > (self.x - self.tolerance) and x < (self.x2+self.tolerance) and \
                y > (self.y - self.tolerance) and y < (self.y2+self.tolerance):
                return True
        return False
    def __str__(self):
        return str(self.x) + "," + str(self.y) + "\n" + str(self.x2) + "," + str(self.y2) + "\n" + str(self.tolerance)

def transformPdfToPymupdf(page,x,y,w,h):
    log("orig coordinates:",x," ",y," ",w," ",h)
    return pymupdf.Rect(x,y,w,h) * ~page.transformation_matrix

def transformRect(page,rect):
    try:
        rect
    except:
        log("no rec to tranform")
        return False
    else:
        return rect * ~page.transformation_matrix

def transformYPoint(page,y):
    return page.mediabox[3]-y

def getRectsInRange(page,border,debug=False):
    data_drawings = page.get_drawings()
    if debug:
        log("[getRectsInRange] ",border)
    #log_json(data_drawings)
    for strocke in data_drawings:
        if strocke["items"][0][0]=="re":
            rec = strocke["items"]
            rect = rec[0][1]
            if border.check(rect.x0,rect.y0):
                if debug:
                    log("RectInRange:")
                    #log_json(strocke)
                yield rect


def getTextInRange(page,border,debug=False,t=2):
    data_text = page.get_text("json",sort=True)

    text_parsed = json.loads(data_text)

    #for a in text_parsed["blocks"]:
    #    if a["type"] ==0: # 0 = text type, 1 = image type
    #        print("textblock: ",json.dumps(a,indent=1))
    
    for block in text_parsed["blocks"]:
        #if "lines" in block.keys():
        if block["type"] == 0: #discard image entries and take all text entry blocks
            for line in block["lines"]:
                text = line["spans"][0]["text"]
                origin = line["spans"][0]["origin"]
                #if text=="Vorname" and origin[0]:
                    #log("text",text)
                    #log("origin",origin)
                    #print("border:",border)
                rectBbox = pymupdf.Rect(origin[0],origin[1],0,0)
                if border.check(rectBbox.x0+t,rectBbox.y0+t):
                    font=line["spans"][0]["font"]
                    size=line["spans"][0]["size"]
                    if debug:
                        log("----",text,"----")
                        log("%%%%",font,"%%%%")
                        #log_json(line)
                    #log("text:",text)
                    #log("rect:",rectBbox)
                    yield text,rectBbox,font,size
        #else:
        #    log("no lines key found")
        #    log("block.keys: ",block.keys())
            #log("type: ", block)
            

    #log_json(text_parsed)

def RectToBorder(rect):
    return Border(rect.x0,rect.y0,rect.x1,rect.y1,0)

def checkBorderDown(page,rectOrigin,jointTolerance=3):#returns the nearest line downwoards to a given rect
    data_drawings = page.get_drawings()
    nearestRect = pymupdf.Rect()
    for strocke in data_drawings:
        if strocke["items"][0][0]=="re":
            rec = strocke["items"]
            rect = rec[0][1]
            if rect.y1 -rect.y0 <= jointTolerance: #is line?
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
    def _sort(rects:list[pymupdf.Rect]):
        for n in range(len(rects)-1,0,-1):
            for i in range(0,n):
                if rectSize(rects[i],"x") < rectSize(rects[i+1],"x"):
                    rects[i],rects[i+1] = rects[i+1],rects[i]
        return rects
    def _sortRightest(rects:list[pymupdf.Rect]):
        for n in range(len(rects)-1,0,-1):
            for i in range(0,n):
                if rects[i].x1 < rects[i+1].x1:
                    rects[i],rects[i+1] = rects[i+1],rects[i]
        return rects
    def checkIsInside(outer:pymupdf.Rect,inner:pymupdf.Rect):
        if (outer.x0 < inner.x0) and (outer.x1 > inner.x1):
            if (outer.y0 < inner.y0) and (outer.y1 > inner.y1):
                return True
        return False
    def searchOuter(page,border,innerRect):
        for outerRect in getRectsInRange(page,border):
            if checkIsInside(outerRect,innerRect):
                return outerRect
        return innerRect
    def getBiggest(page,recTable,yRecSearchPoint):
        _list=[]
        for rec in getRectsInRange(page,Border(recTable.x0,recTable.y0,page.bound().x1,page.bound().y1,30)):
            if rectSize(rec,"x") > 10:
                if abs(round(rec.y1) - yRecSearchPoint)  < 5:
                    print("rec:",transformRect(page,rec))
                    _list.append(rec)
        return _list

    tables = []
    for name in tableNames:
        for rec in page.search_for(name):
            #find x point
            #for rec2 in getRectsInRange(page,Border(rec.x0,rec.y0,rec.x1,rec.y1,2)):
            #    if rec.x0 - rec2.x0 > 0 and rec.x0-rec2.x0 < 50:
            #        rec.x0 = rec2.x0
            #Here we need to find the most plausible x line left
            rec.x0 = getTableLineLeftOfName(page,rec.x0,rec.y0)

            log("xpos for getTableLine",rec.x0)
            y = getTableLine(page,rec.x0,rec.y1,skip=1)
            log("y:",transformYPoint(page,y))
            if y != False:
                rec.y1 = y
            tables.append(Tbl(rec,name,fileName,pageNr))

    for x in tables:
        log("{searchForTable}tables at ",x.name,":",transformRect(page,x.rec))
    log("\tcount tables:",len(tables))
    for i,table in enumerate(tables):
        log("{searchForTable}Current tableName:",table.name)
        log("\t",table.name,"No table end found use algo2")
        rec = searchTableDown(page,table)
        if rec:
            #table.rec.y0 = rec.y0
            table.rec.y1 = rec.y1
            log("\tres algo2:")
            print("\t\trec:",transformRect(page,table.rec))
        else:
            log("\t\tno table found, giving up")
        
        #just implement choose the biggest one near y point

        for line in getHorizontalLines(page,50,5):
            log("TableLine:",transformRect(page,line))
            if abs(line.y0 - table.rec.y0) < 5:
                if line.x1 > table.rec.x1:
                    #table.rec.x0 = line.x0
                    table.rec.x1 = line.x1

        biggest = getBiggest(page,table.rec,table.rec.y0-3)
        biggest = _sort(biggest)
        #for line in biggest:
        #    print("sorted:",line)
        if biggest:
            if biggest[0].x1 > table.rec.x1:
                table.rec.x1 = biggest[0].x1
        
        extensionRecs = []

        for line in getHorizontalLines(page,50,5):
            extensionRecs.append(line)
        extensionRecs = _sortRightest(extensionRecs)
        #for x in extensionRecs:
        #    print("sorted Rightest:",transformRect(page,x))
        if extensionRecs:
            extensionRec = extensionRecs[0]
            if abs(extensionRec.y0-table.rec.y0) < 30:
                if extensionRec.x1 > table.rec.x1:
                    if abs(extensionRec.x0-table.rec.x1) < 5 or abs(extensionRec.x0 -table.rec.x0) < 5:
                        log("extensionRec: ",extensionRec)
                        table.rec.x1 = extensionRec.x1



        log("def table:",transformRect(page,table.rec))
            #search with other algorythm
    #remove double tables where if first point match remove the one witch seems to be a line
    toBeRemoved = []
    if len(tables)>1:
        log("len tables: ",len(tables))
        for ai,a in enumerate(tables):
            for i in range(ai+1,len(tables)):
                log("i: ",i)
                if abs(a.rec.x0 - tables[i].rec.x0) < 1 and abs(a.rec.y0 - tables[i].rec.y0) < 1:
                    log("\tpop table:", tables[i].name, "in favor of: ",a.name)
                    toBeRemoved.append(i)
        toBeRemoved.sort(reverse=True)
        for i,e in enumerate(toBeRemoved):
            tables.pop(i)
    return tables


def detectTableRows(page,table,tx=5,ty=6,bx=5,borderT=3,nextConLineHT=10):
    def probeRect(rect):
        def err(n):
            log(n, "is NULL",err=err.DEBUG)
        if rect.x0 == None:
            err("x0")
        if rect.y0 == None:
            err("y0")
        if rect.x1 == None:
            err("x1")
        if rect.y1 == None:
            err("y1")
    rowNameList=[]
    fields = []
    rects = []

    border = Border(table.rec.x0,table.rec.y0,table.rec.x0+tx,table.rec.y0+ty,borderT)
    log("{detectTableRows}tablerec:",table.rec)

    log("Search for vertical line, count rects:",len(rects))
    for rect in getRectsInRange(page,border,debug=False):
        if isLine(rect) == "vertical":
            hLines = getAllTableHLine(page,table.rec.x0,table.rec.y0,skip=0)
            if hLines != False:
                log("hLines:",rect)
                if len(hLines) > 1:
                    hLineHeader = hLines[0]
                elif len(hLines) == 1:
                    hLineHeader = hLines[0]
                else:
                    hLineHeader = pymupdf.Rect(0,0,0,0)
                    log("No hLineHeader Found!",level=err.ERROR)
                headerRec = pymupdf.Rect(hLineHeader.x0,hLineHeader.y0,table.rec.x1,hLineHeader.y1)
                log("HeaderRec:",transformRect(page,headerRec))

                for field in getFieldsInRange(page,headerRec):
                    log(transformRect(page,field))
                    fields.append(field)

                for fieldRow in getHeader(page,fields,headerRec):
                    fieldRow = pymupdf.Rect(fieldRow.x0,fieldRow.y0,fieldRow.x1,fieldRow.y1+10)
                    log("fieldrow:",transformRect(page,fieldRow))
                    fullText = ""
                    i=0
                    for text,rect,font,size in getTextInRange(page,RectToBorder(fieldRow),t=5,debug=False):
                        log("[fieldRow]Text: ",text)
                        text = text.strip()
                        if text:
                            splitText = cutTextOverRect(page,fieldRow,text,font,size)
                            if not splitText:
                                rowNameList.append(text)
                            else:
                                for txt in splitText:
                                    rowNameList.append(txt)
            else:
                log("No hline header found",level=err.ERROR)


    return rowNameList

def getHeader(page,fields,rectTable,thresold=6):
    lowestY = None
    filteredFields = []
    log("rectTable:",rectTable)
    for field in fields:
        log("{getHeader} Fields:", field)
        if abs(rectSize(rectTable,direction='x') - rectSize(field,direction='x')) > thresold:
            sizeTable = rectSize(rectTable,direction='x')
            sizeField = rectSize(field,direction='x')
            size = sizeTable - sizeField
            log("sizeTable:",sizeTable,"sizeField:",sizeField)
            fullText = ""
            text= ""
            for text,rect,font,size in getTextInRange(page,RectToBorder(field),debug=False):
                fullText = fullText + text

            log("DataField:",transformRect(page,field),"field size:",sizeField,"text: ",text)

            #find lowest row y
            if lowestY == None:
                lowestY = field.y0
            else:
                if field.y0 > lowestY:
                    lowestY = field.y0
            filteredFields.append(field)
    if fields:
        log("lowest row:",lowestY)
        log("lowest row:",transformYPoint(page,lowestY))

        fieldsRow = []
        for field in filteredFields:
            #if round(field.y0,1) == round(lowestY,1):#is it in a row?
            fieldsRow.append(field)
        return fieldsRow
    else:
        return []

def cutTextOverRect(page,rect,text,font,size=9,oversizeThresold=10):
    _len=0
    size = rectSize(rect,direction='x')
    try:
        font = pymupdf.Font(font)
    except:
        log("font",font,"not found, using helv")
        font = pymupdf.Font("helv")

    if len(text.strip().split(" "))>1:
        _len = font.text_length(text)
        if size - _len < -oversizeThresold:
            log("text oversized!",level=err.ERROR)
            log("font:",font)
            log("len of field:",size,"len of word:", _len,"text:",text)
            log("len split",len(text.strip().split(" ")))
            return text.strip().split(" ") 


def rectSize(rect,direction="x"):
    if direction=='x':
        return abs(rect.x1 - rect.x0)
    if direction=='y':
        return abs(rect.y1 - rect.y0)
    return 0

def getFieldsInRange(page,rect,borderWidth=3):
    t=2
    border = Border(rect.x0-t,rect.y0-t,rect.x1+t,rect.y1+borderWidth,1)
    for field in getRectsInRange(page,border):
        if not isLine(field):
            yield field
#searches in a given direction and detects a line witch is 90 degree to it
def nextLineCross(page,rect,direction,borderWidth=3,t=2,beginThreshold=10):
    if direction=="vertical":
        border = Border(rect.x0-t-beginThreshold,rect.y0-t,rect.x1+t,rect.y1+borderWidth,1)
        for rectNew in getRectsInRange(page,border,debug=False):
            if isLine(rectNew) == "horizontal":
                if(rect.y0 != rectNew.y0):
                    return rectNew
        else:
            return False
    if direction=="horizontal":
        border = Border(rect.x0+t,rect.y0-t,rect.x1+t+borderWidth*5,rect.y0+borderWidth,1)
        for rectNew in getRectsInRange(page,border):
            if isLine(rectNew) == "vertical":
                return rectNew
        else:
            return False

def nextConnectedLine(page,rect,direction,borderWidth=3,t=4):
    if direction=="vertical":
        border = Border(rect.x0-t,rect.y1-t,rect.x0+t,rect.y1+borderWidth,1)
        for rectNew in getRectsInRange(page,border):
            if isLine(rectNew) == "vertical":
                return rectNew
        else:
            return False
    if direction=="horizontal":
        border = Border(rect.x1-t,rect.y0-borderWidth,rect.x1+t,rect.y1+borderWidth,1)
        for rectNew in getRectsInRange(page,border):
            log("rects in range:",rectNew)
            if isLine(rectNew) == "horizontal":
                return rectNew
        else:
            log("no next line found, border:",border)
            return False
    else:
        log("direction unkown argument:",direction,level=err.ERROR)


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

def getTableLineLeftOfName(page,xPosName,yPosName,borderWidth=3,jointTolerance=2,skip=0):
    leftLimit = xPosName-borderWidth
    rightLimit = xPosName+borderWidth

    def _filter(drawings,page,xPosName,yPosName):
        listVerticalRects = []
        for drawing in drawings:
            rect:pymupdf.Rect = drawing['items'][0][1]
            if type(rect) == pymupdf.Rect:
                if rect.x0 < xPosName:
                    if abs(rect.y0 - yPosName) < 10:                
                        listVerticalRects.append(rect)
        return listVerticalRects

    def _sort(rects:list[pymupdf.Rect]):
        for n in range(len(rects)-1,0,-1):
            for i in range(0,n):
                if(rects[i].x1 < rects[i+1].x1):
                    rects[i],rects[i+1] = rects[i+1],rects[i]
        return rects

    drawings = page.get_drawings()
    verticalRects = _filter(drawings,page,xPosName,yPosName)
    for rect in verticalRects:
        log("[getTableLineLeftOfName]filtered rects: ",rect)
    verticalRects = _sort(verticalRects)
    for rect in verticalRects:
        log("[getTableLineLeftOfName]sorted rects: ",transformRect(page,rect))

    for rect in verticalRects:
        print("[getTableLine] sorted vRect:",transformRect(page,rect))

    print("connected lines:")
    for rect in verticalRects:
        print("\t  connected vRect:",transformRect(page,rect))
    
    if verticalRects:
        return verticalRects[0].x0
    else:
        return False

def getTableLine(page,xPosTabelLine,yPosStartlineDown,borderWidth=3,jointTolerance=2,skip=0):
    leftLimit = xPosTabelLine-borderWidth
    rightLimit = xPosTabelLine+borderWidth

    def _filter(drawings,page,leftLimit,rightLimit,borderWidth,yPosStartlineDown):
        listVerticalRects = []
        for drawing in drawings:
            rect:pymupdf.Rect = drawing['items'][0][1]
            if type(rect) == pymupdf.Rect:
                if abs(round(rect.x0)-469) < 2:
                    print("\trec:",transformRect(page,rect))
                if(rect.x1-rect.x0)<borderWidth:
                    if(rect.y1-rect.y0)>jointTolerance:
                        if(rect.x0 <rightLimit and rect.x0 > leftLimit):
                            if(rect.y0 > yPosStartlineDown):
                                #print("\t----->BigBox: ",drawing)
                                #print("keys:",drawing.keys())
                                #print("seqno:",drawing['width'])

                                listVerticalRects.append(rect)
        return listVerticalRects

    def _sort(rects:list[pymupdf.Rect]):
        for n in range(len(rects)-1,0,-1):
            for i in range(0,n):
                if(rects[i].y1 > rects[i+1].y1):
                    rects[i],rects[i+1] = rects[i+1],rects[i]
        return rects
    def getConnected(rects,jointTolerance,skip=0):
        _list = []
        try:
            _list.append(rects[0])
        except:
            pass
        for n in range(skip,len(rects)-1):
            if( abs(rects[n+1].y0 - rects[n].y1) <= jointTolerance):
                print("sub:",rects[n].y0 - rects[n+1].y1)
                _list.append(rects[n+1])
            else:
                break
        return _list

    drawings = page.get_drawings()
    verticalRects = _filter(drawings,page,leftLimit,rightLimit,borderWidth,yPosStartlineDown)

    verticalRects = _sort(verticalRects)

    for rect in verticalRects:
        print("[getTableLine] sorted vRect:",transformRect(page,rect))

    verticalRects = getConnected(verticalRects,jointTolerance,skip=skip)
    print("connected lines:")
    for rect in verticalRects:
        print("\t  connected vRect:",transformRect(page,rect))
    
    if verticalRects:
        return verticalRects[len(verticalRects)-1].y1
    else:
        return False






def printHorizontalLine(page):
    drawings = page.get_drawings()
    for drawing in drawings:
        rect:pymupdf.Rect = drawing['items'][0][1]
        if type(rect) == pymupdf.Rect:
            if(rect.x1-rect.x0)>100:
                if(rect.y1-rect.y0)<5:
                    if rect.x0<100:
                        print("BigBox: ",drawing)

def getHorizontalLines(page,minLen=100,lineThreshold=5):
    drawings = page.get_drawings()
    for drawing in drawings:
        rect:pymupdf.Rect = drawing['items'][0][1]
        if type(rect) == pymupdf.Rect:
            if abs(rect.x1-rect.x0)>minLen:
                if abs(rect.y1-rect.y0)<lineThreshold:
                    if rect.x0<minLen:
                        yield rect


def searchTableEnd(page,table_full,t=10,borderT=5,thresholdIsLongLine=20): #search the table border by moving to the left
    table = table_full.rec
    border = Border(table.x0-t,table.y0-t,table.x0+t,table.y0+t,borderT)
    for rect in getRectsInRange(page,border):
        if abs(rect.x1 - rect.x0) < borderT and abs(rect.y0 - rect.y1) > thresholdIsLongLine: #is it a line? and filter out very short lines
            log("{searchTableEnd}(iterating rects in range)Potential border of table:",transformRect(page,rect))
            return rect
    return False

def searchTableDown(page,table_full,t1=10,borderWidth=2,magicY=5,endAppendY=10,magicX=3,borderT=5):# How much lookahead for table end? better to much than to little
    log("{SearchTableDown}")
    biggesty = 0
    smallesty = None
    table = table_full.rec
    border = Border(table.x0-t1,table.y0-t1,table.x0+t1*magicX,table.y0+t1*magicY,borderT)
    for rect in getRectsInRange(page,border):
        if abs(rect.x1 - rect.x0) < borderWidth: #and abs(rect.y0 - rect.y1) > 20: #is it a line? do not filter short lines
            #log("{searchTableDown}Potential border of table:",transformRect(page,rect))
            if(rect.y1 > biggesty):
                biggesty = rect.y1
            #    log("\tbiggesty[mypfcoordinate]:",biggesty)
            if not smallesty:
                smallesty = rect.y0
            if(rect.y0 < smallesty):
                smallesty = rect.y0
             #   log("\tsmallest[mypfcoordinate]:",smallesty)
    try:
        rect
    except:
        log("no rect found")
    else:
        if rect:
            final_rect = rect
            final_rect.y1 = biggesty + endAppendY
            final_rect.y0 = smallesty
            log("\ttable border calculated:",transformRect(page,final_rect))
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
        print("tableNames:",self.tableNames)
    def selectTable(self,nr):
        self.selected_table = self.tables[nr]
    def selectTableByObj(self,obj):
        self.selected_table = obj
    def defRows(self,rowNameList):
        table = self.selected_table
        table.rowNameList = rowNameList
    def parseTable(self):
        table = self.selected_table
        table.entries = []
        table.state = self.State()
        page = self.pages.selected
        table.rec.y1 = getTableLine(page,table.rec.x0,table.rec.y0,skip=1)
        table.border = Border(table.rec.x0,table.rec.y0,table.rec.x1,table.rec.y1,5)

        print("table:",transformRect(page,table.rec))
        print("table:",table.rec)
        print("rowNameList:",self.selected_table.rowNameList)
        for init,rowName in enumerate(self.selected_table.rowNameList):
            if init == 0:
                table.entries.append(Entry())
            _list,realName = self.searchContentFromRowName(page,rowName,table.border)

            for index,content in enumerate(_list):
                if(index > len(table.entries)-1):
                    table.entries.append(Entry())
                log("{PARSE TABLE} entry: ",table.entries[index],"_")
                if(len(rowName) < len(realName)):
                    rowName = realName.replace("\n","-").replace(" ","")
                setattr(table.entries[index],rowName,content)


    def getObjectsFromTable(self):
        table = self.selected_table
        for x in table.entries:
            yield x
    def searchContentFromRowName(self,page,rowName,table_border,tx0=4,ty0=2,tx1=1,borderWidth=3):
        def checkIsInside(outer:pymupdf.Rect,inner:pymupdf.Rect):
            if (outer.x0 < inner.x0) and (outer.x1 > inner.x1):
                if (outer.y0 < inner.y0) and (outer.y1 > inner.y1):
                    return True
            return False
        def searchOuter(page,border,innerRect):
            for outerRect in getRectsInRange(page,border):
                if checkIsInside(outerRect,innerRect):
                    return outerRect
            return innerRect

        listA = []          #list from algo A
        listB = []          #list from algo B
        listC = []          #combined list
        names = page.search_for(rowName) #returns list of Rect
        for recName in names:
            outerRec = searchOuter(page,table_border,recName)
            log("--\nouter:",transformRect(page,outerRec))

            border = Border(outerRec.x0,outerRec.y0,outerRec.x1,outerRec.y1,1)
            if table_border.check(outerRec.x0,outerRec.y0): # is the row inside the border/table?
                real_name = page.get_textbox(outerRec).strip()
                log(real_name,";",rowName,";",transformRect(page,outerRec))

                if real_name == rowName or (real_name in rowName) or (rowName in real_name) :
                    #posTableLine = checkBorderDown(page,recName)
                    borderContent = Border(outerRec.x0,outerRec.y1,outerRec.x1,table_border.y2,1)
                    log("tableborder.y2:",transformYPoint(page,table_border.y2))
                   # for i,rect in enumerate(getRectsInRange(page,borderContent)):
                   #     string = page.get_textbox(rect)
                   #     if string.strip():
                   #         log("getRectsInRange:",string.strip())
                   #         listA.append(string.strip())

                    for i,[text,rect,font,size] in enumerate(getTextInRange(page,borderContent)):
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

                    return listC, real_name
                else:
                    log("no match")
            else:
                log("rec(",rowName,") not inside table border")
                log("table_border:",transformPdfToPymupdf(page,table_border.x,table_border.y,table_border.x2,table_border.y2))
        return [],""

def getAllTableHLine(page,xPosTabelLine,yPosStartlineDown,borderWidth=5,jointTolerance=2,skip=0):
    leftLimit = xPosTabelLine-borderWidth
    rightLimit = xPosTabelLine+borderWidth

    def _filter(drawings,page,leftLimit,rightLimit,borderWidth,yPosStartlineDown):
        listVerticalRects = []
        for drawing in drawings:
            rect:pymupdf.Rect = drawing['items'][0][1]
            if type(rect) == pymupdf.Rect:
                if(rect.x1-rect.x0)<borderWidth:
                    if(rect.y1-rect.y0)>jointTolerance:
                        if(rect.x0 <rightLimit and rect.x0 > leftLimit):
                            if(rect.y0 > yPosStartlineDown):
                                listVerticalRects.append(rect)
        return listVerticalRects

    def _sort(rects:list[pymupdf.Rect]):
        for n in range(len(rects)-1,0,-1):
            for i in range(0,n):
                if(rects[i].y1 > rects[i+1].y1):
                    rects[i],rects[i+1] = rects[i+1],rects[i]
        return rects
    def getConnected(rects,jointTolerance,skip=0):
        _list = []
        try:
            _list.append(rects[0])
        except:
            pass
        for n in range(skip,len(rects)-1):
            if( abs(rects[n+1].y0 - rects[n].y1) <= jointTolerance):
                print("sub:",rects[n].y0 - rects[n+1].y1)
                _list.append(rects[n+1])
            else:
                break
        return _list

    drawings = page.get_drawings()
    verticalRects = _filter(drawings,page,leftLimit,rightLimit,borderWidth,yPosStartlineDown)

    verticalRects = _sort(verticalRects)

    for rect in verticalRects:
        print("\t vRect:",rect)

    verticalRects = getConnected(verticalRects,jointTolerance,skip=skip)
    print("connected lines:")
    for rect in verticalRects:
        print("\t  connected vRect:",transformRect(page,rect))
    
    if verticalRects:
        return verticalRects
    else:
        return False
