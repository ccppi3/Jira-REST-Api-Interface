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
    version = subprocess.run(["git","describe","--tags"],capture_output = True, text=True)
    print("Git tells version (full):",version.stdout.strip())
    return version.stdout.strip()

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
                sLine = slice(0,posNbr + i + 1)
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
    if amIExe():
        path = path + "..\\jira-flow.exe"
    else: # we run in a dev env
        path = "C:\\Program Files\\Jira-Flow\\jira-flow.exe"
    print("path to exe:",path)
    try:
        info = win32api.GetFileVersionInfo(path,'\\')
    except Exception as e:
        print("we dont seem to run as an exe, aborting version check\n",e)
    else:
        Pls = info['ProductVersionLS']
        Pms = info['ProductVersionMS']
        Fls = info['FileVersionLS']
        Fms = info['FileVersionMS']

        Pversion = ".".join([str(win32api.HIWORD(Pms)),str(win32api.LOWORD(Pms)),str(win32api.HIWORD(Pls)),str(win32api.LOWORD(Pls))])
        Fversion = ".".join([str(win32api.HIWORD(Fms)),str(win32api.LOWORD(Fms)),str(win32api.HIWORD(Fls)),str(win32api.LOWORD(Fls))])
        #P2version = info['ProductVersion']
        print("ProductversionMSLS",Pversion)
        print("Fileversion",Fversion)

        return Fversion

def getVersionsExe(path=getResourcePath(""),info_str="ProductVersion"):
    ver_strings = (
        "Comments",
        "InternalName",
        "ProductName",
        "CompanyName",
        "LegalCopyright",
        "ProductVersion",
        "FileDescription",
        "LegalTrademarks",
        "PrivateBuild",
        "FileVersion",
        "OriginalFilename",
        "SpecialBuild",
    )
    if amIExe():#we run as deployed version
        path = path + "..\\jira-flow.exe"
    else: # we run in a dev env sow we run a test against another exe
        path = "C:\\Program Files\\Jira-Flow\\jira-flow.exe"
    print("path to exe:",path)

    try:
        info = win32api.GetFileVersionInfo(path,'\\VarFileInfo\\Translation')
    except Exception as e:
        print("we dont seem to run as an exe, aborting version check\n",e)
    else:
        for lang,codepage in info:
            for ver_string in ver_strings:
                str_info = f"\\StringFileInfo\\{lang:04X}{codepage:04X}\\{ver_string}"
                print(ver_string, repr(win32api.GetFileVersionInfo(path, str_info)))
                if ver_string == info_str:
                    return repr(win32api.GetFileVersionInfo(path,str_info))

if not amIExe():
    version = getVersionGit()
    if version:
        writeVersionFile(version)
    else:
        print("could not get version from Git, check 'git tags'")

if "__main__" in __name__:
    version = getVersionGit()
    print("version:",version)
    writeVersionFile(version)
    fileVersion = getFileVersionExe()
    print(fileVersion)


                


