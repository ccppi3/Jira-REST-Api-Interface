import sys,os
import subprocess
import win32api

def amIExe():
    try:
        basePath = sys._MEIPASS # are we a built with pyinstaller or in an dev env?
        return True
    except:
        return False

def getVersionGit():
    version = subprocess.run(["git","tag"],capture_output = True, text=True)
    print("Git tells version (full):",version.stdout)
    print("Git tells version (cut):",version.stdout)
    return version.stdout

def writeVersionFile(version, file="exeVersionInfo.txt"):
    buffer = []
    with open(file) as f:
        stream = f.readlines()
        for line in stream:
            if "ProductVersion" in line:
                print("line:",line)
                posVersion = line.find("ProductVersion")
                posVersion = posVersion + len("ProductVersion")

                posNbr = line.find("u'",posVersion) + 2

                line = list(line)
                line = line[0:posNbr]
                version = list(version)

                for i,a in enumerate(version):
                    line.append(a)
                    #line[posNbr + i] = a    
                sLine = slice(0,posNbr + i)
                line = line[sLine]
                line = ''.join(line)#create a string out of a list
                line = line + "\')])\n"
                print("line:",line)
            buffer.append(line)
    print("Buffer:",buffer)
    
    with open(file,"w") as f:
        f.writelines(buffer)

def getResourcePath(relPath):
    try:
        basePath = sys._MEIPASS
    except Exception:
        basePath = os.path.abspath(".")
    return os.path.join(basePath,relPath)

def getFileVersionExe(path=getResourcePath("")):
    path = path + "..\\jira-flow.exe"
    print("path to exe:",path)
    try:
        info = win32api.GetFileVersionInfo(path,'\\')
    except Exception as e:
        print("we dont seem to run as an exe, aborting version check\n",e)
    else:
        ls = info['ProductVersionLS']
        ms = info['ProductVersionMS']
        version = ".".join([str(win32api.HIWORD(ms)),str(win32api.LOWORD(ms)),str(win32api.HIWORD(ls))])
        return version

if not amIExe():
    version = getVersionGit()
    writeVersionFile(version)

if "__main__" in __name__:
    version = getVersionGit()
    print("version:",version)
    writeVersionFile(version)
    fileVersion = getFileVersionExe()
    print(fileVersion)


                


