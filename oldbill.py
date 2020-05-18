import PyPDF2
import re
import xlsxwriter
import sys

index = 0

def getBillNo(wordList):
    return wordList[28].replace("Bill", "")

def getName(wordList):
    inFile1=sys.argv[1]
    name = wordList[33].replace(inFile1, "")     #2020 is the present year
    for i in range(34, len(wordList)):
        word = wordList[i]
        if word.find("PARTICULARSAMOUNTFLAT") != -1:
            index = i+2
            name += " "
            name += wordList[i].replace("PARTICULARSAMOUNTFLAT", "")
            return name, index
        else:
            name += " "
            name += wordList[i]
                
def getFlatNo(wordList):
    if wordList[index+1].isdigit():
        flat = wordList[index]
        flat += "-"
        flat += wordList[index+1][0]
        return flat
    else:
        flat = wordList[index][:-2]
        return flat
    
def getAmount(wordList):
    for i in range(len(wordList)):
        if wordList[i] == "DUE":
            amount = wordList[i+1]
            i+=2
            for j in range(i, len(wordList)):
                word = wordList[j]
                if word.find("Secretary") != -1:
                    return amount
                else:
                    amount += word
str1=input("Enter in the format MMM-YYYY:-")
str2="BILL SUMMARY FOR "
object = open('BILL.pdf', 'rb')
reader = PyPDF2.PdfFileReader(object)
workbook = xlsxwriter.Workbook('BILL.xlsx')
worksheet = workbook.add_worksheet() 
worksheet.set_column(1, 1, 30)
worksheet.set_column(3, 6, 12)
n = reader.numPages

column = 0
header= {"PARTH CO-OP HSG SOC LTD"}
for head in header:
    worksheet.write(0,column,head)

column = 0
header= {str2+str1}
for head in header:
    worksheet.write(1,column,head)

column = 0
headings = ["Bill No", "Name", "Flat No", "Amount Due", "Bank", "Cheque No", "Amount"]
for head in headings:
    worksheet.write(2, column, head)
    column += 1

for i in range(n):
    index = 0
    page = reader.getPage(i)
    text = page.extractText()
    if len(text) == 0:
        continue
    wordList = re.sub("[^\w]", " ",  text).split()
    
    bill = getBillNo(wordList)
    name, index = getName(wordList)
    flat = getFlatNo(wordList)
    amount = getAmount(wordList)
    
    column = 0
    dataset = [bill, name, flat, amount]
    for data in dataset:
        worksheet.write(i+3, column, data)
        column += 1

workbook.close()    
