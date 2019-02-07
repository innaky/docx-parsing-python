#!/usr/bin/python

import os
import re
import zipfile
import xml.dom.minidom

# DOCX files are Zipped XML
document = zipfile.ZipFile("foo.docx")

# Read the XML data
xmldata = xml.dom.minidom.parseString(document.read('word/document.xml')).toprettyxml(indent='  ')

# convert data into unicode string
utf8text = xmldata.encode("utf-8")

# split the data when find 70 or more "-" symbol
splitted = re.split(r'-{70,}<', utf8text)

def write_file(file, input_lst):
    f = open(file, "a")
    for i in range(len(input_lst)):
        f.write(input_lst[i])
        f.write("\n")
    f.close

# The main loop will iterate in all splitted blocks
for i in range(len(splitted)):
    selected = list(map(lambda x: x.strip(), re.findall('>(.+[a-zA-Z ].*)</w:t>', splitted[i])))
    # Check needed when files begins with separator "-------" avoiding empty first block
    if len(selected)>0:
	# "ticker" for filename are all characters before firts " " (space)
	ticker = re.split(r' ', selected[0])[0]
	# "date" for filename are the firts 10 characters after last " " (space)
    	ddmmyy = selected[0].split(" ")[-1][5:15]
        filename = ticker+'-'+ddmmyy+'.txt'
        write_file(filename, selected)
