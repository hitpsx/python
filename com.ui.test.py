import os
def file_name(file_dir):
    for root, dirs, files in os.walk(file_dir):
        return  files

fileName=file_name("G:\迅雷下载\skin")
for name in fileName:
    if(name.find("LOL")!=-1):
        currentName=name
currentName='"'+currentName+'"'
print(os.system("start "+"G:\迅雷下载\skin\\"+currentName))