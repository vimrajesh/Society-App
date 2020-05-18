import PyPDF2
import re
import sys
import xlsxwriter

def getBILL(lis):
    word= lis[19]
    idx=word.find("B")
    return word[:idx]

def getNAME(lis):
    word1=lis[23]
    a=word1.find(inp)+8
    name=word1[a:]
    idx=23
    if name.find("PARTICULARSAMOUNTFLAT")>=0:
        name=name.replace("PARTICULARSAMOUNTFLAT","")
        idx=idx+2
        return name,idx
    idx=idx+1
    while True:
        wordd=lis[idx]
        if wordd.find("PARTICULARSAMOUNTFLAT") != -1:
            temp1=wordd
            idx=idx+2
            temp=temp1.replace("PARTICULARSAMOUNTFLAT","")
            name = name + " " + temp
            break
        else:
            temp=wordd
            name = name + " " + temp
            idx=idx+1
    return name,idx    

def getFLAT(lis,idx):
    word=lis[idx]
    length=len(word)-19
    flat=word[:length]
    return flat   

def getAMT(lis,idx):
    temp=0
    for i in range(idx,len(lis)):
        if "DUE"==lis[i]:
            temp=i
            break
    word=lis[temp+1]
    word_neg=lis[temp-1].find("CR")
    k=word.find("Secretary")
    word=word[:k-3]
    
    if word_neg > 0:
        word="-"+ word
    return word

inFile="BILL.pdf"
inp = sys.argv[1]    
PDFfileobj= open(inFile,'rb')
inpx="BILL_SUMMARY_"+inp + ".xlsx"
workbook = xlsxwriter.Workbook(inpx)
worksheet=workbook.add_worksheet(inp)

# bold = workbook.add_format({'bold': True}) 
# money = workbook.add_format({'num_format': '#,##0'})

column = 0
inp1="BILL SUMMARY FOR "+ inp
worksheet.write(0,column,"PARTH CO-OP HSG SOC LTD")
worksheet.write(1,column,inp1)
worksheet.set_column(1, 1, 30)
worksheet.set_column(3, 6, 12)
headings = ["Bill No", "Name", "Flat No", "Amount Due", "Bank", "Cheque No", "Amount"]
for head in headings:
    worksheet.write(2, column, head)
    column += 1 
pdfread=PyPDF2.PdfFileReader(PDFfileobj)
num_pages=pdfread.numPages
column=0
for i in range(num_pages):
    column=0
    pageobj=pdfread.getPage(i)
    lis1=pageobj.extractText()
    if(len(lis1)==0):continue
    lis=re.sub("[^,/.///&()a-zA-Z0-9_-]"," ",lis1).split()
    index=0
    BILL_NO=getBILL(lis)
    NAME,index = getNAME(lis)
    FLAT_NO=getFLAT(lis,index)
    AMT_DUE=getAMT(lis,index+1)
    
    # worksheet.write(i+3,column,BILL_NO,money)
    # worksheet.write(i+3,column+1,NAME)
    # worksheet.write(i+3,column+2,FLAT_NO)
    # worksheet.write(i+3,column+3,AMT_DUE,money)
    data_=[BILL_NO,NAME,FLAT_NO,AMT_DUE]
    for data in data_:
        worksheet.write(i+3,column,data)
        column += 1
    # print(BILL_NO,NAME,FLAT_NO,AMT_DUE)
PDFfileobj.close()      
workbook.close()    