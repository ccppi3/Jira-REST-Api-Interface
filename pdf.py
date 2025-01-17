import pymupdf
import json
DEBUG = True

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


doc = pymupdf.open("Test.pdf")

#for page in doc:
#    text = page.get_text().encode("utf8")
#    print(text)

page = doc[0]
data = json.loads(page.get_text("json",sort=False))


jdata = json.dumps(data,indent=2,default=str)

key_dump(data,"text")
objs = []
for line in data["blocks"]:
    text = line["lines"][0]["spans"][0]["text"]
    x = line["lines"][0]["spans"][0]["origin"][0]
    y = line["lines"][0]["spans"][0]["origin"][1]
    objs.append(text_obj(text,x,y))

for o in objs:
    print(o,"\n")

#print("all keys: ",[*get_all_keys(data)])
file = open("json_of_pdf.json","w")

file.write(jdata)





