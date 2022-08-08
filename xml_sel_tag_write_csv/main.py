from collections import defaultdict
from tkinter import filedialog
import xml.etree.ElementTree as ET
import csv

# file load
def Load():
    filename = filedialog.askopenfilename(initialdir="/", title="Select file",
                                          filetypes=(("PPTX files", "*.pptx"),
                                          ("all files", "*.*")))
    #print(filename)

    return filename

filePath = Load()

tree = ET.parse(filePath)

#찾아보려하는 tag 값
want_tag = "member"

root = tree.getroot().findall(want_tag)

dic = defaultdict(int)

for child in root :
    for data in child :
        tagname = data.tag
        if '}' in data.tag :
            tag_url, tagname = data.tag.split('}') # 만약 tag 안에 url str이 포함인 경우 분리 시 사용

        dic[tagname] +=1

#csv파일 쓰기
f = open('data.csv', "w")
writer = csv.writer(f)

for i in dic :
    temp = []
    temp.append(str(i))
    temp.append(str(dic[i]))

    writer.writerow((temp))
    #write 값 확인
    #print(i, dic[i])

