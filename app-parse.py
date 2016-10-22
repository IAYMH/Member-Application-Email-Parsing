#!/usr/bin/python2.7
import base64
import os
import re

print("Parsing email applications...")

# Setup csv export
fieldList = ["Application No: ", "Title: ", "First Name: ", "Last Name: ", "Educational Institution/Organisation: ", "Street Address: ", "Postcode: ", "City: ", "Country: ", "State: ", "Phone Number: ", "Email Address: ", "Occupation: ", "Comments: "]
csvExport = ["~".join(fieldList) + "~Date joined: "]

# Read file list
fileArray = []
for file in os.listdir("./"):
  if file.endswith(".eml"):
    fileArray.append(file)

# Load each file and compile email body into string
for file in fileArray:
  with open(file, 'r') as eml:
    data = eml.readlines()
    for item in data:
      if item.startswith("Date: "):
        dateArray = item.split( )
        dateString = " ".join(dateArray[1:4])
        #csvExport.append(' '.join(dateString))
      if item.startswith("Content-Transfer-Encoding:"):
        encoding = item.split( )[1]
      if item.startswith("X-Virus-Status:"):
        index = data.index(item)+2
        msgBody = []
        while index < len(data):
          if encoding == "base64":
            msgBody.append(data[index])
            msgBodySixFour = "".join(msgBody)
            msgBodyString = base64.decodestring(msgBodySixFour).replace("\r","").replace("\n\n", "\n")
          else:
            cleanText = data[index].replace("=0D=0A", "\n").replace("=","").replace("\r", "").replace("\n\n","\n")
            if cleanText.endswith("\n"):
              cleanText = cleanText[:-1]
            msgBody.append(cleanText)
            msgBodyString = "".join(msgBody)
          index += 1
        msgBodyString = msgBodyString.replace("\n\n","\n")

# Parse msgBodyString for relevant info
        appLine = ""
        mBS = msgBodyString.replace("http://www.iaymh.org/403/", "Application No: ").replace(".ashx\nThe following details were supplied:","")
        for field in fieldList:
          start = mBS.find(field)+len(field)
          if fieldList.index(field) < len(fieldList)-1:
            if field == "Phone Number: ":
              end = mBS.find(": tel:")
            elif field == "Email Address: ":
              end = mBS.find(": mailto:")
            else:
              nextField = fieldList[fieldList.index(field)+1]
              end = mBS.find("\n" + nextField)
          else:
            end = len(mBS)
          appItem = mBS[start:end]
          appLine += appItem.replace("\n"," ") + "~"

        appLine += dateString
        csvExport.append(appLine)
    eml.close()


# Output to CSV
of = open("newMembers.csv", "w")
for item in csvExport:
  of.write(item)
  of.write("\n")
of.close()