"""
	Description: PDF Scraper
	Input: Link of pdf
"""

import argparse, sys, PyPDF2, os, xlwt, textract, re
from PyPDF2 import PdfFileWriter, PdfFileReader
from xlwt import Workbook 

#func
def fileExist(f):
	if os.path.isfile(f):
		return True
	else:
		return False

def crop(fl):
	left = 0
	top	= 90
	right = 460
	bottom = 100

	fl = open(args.input, 'rb')
	pdf = PdfFileReader(fl, 'rb')
	out = PdfFileWriter()
	for page in pdf.pages:
		page.mediaBox.upperRight = (page.mediaBox.getUpperRight_x() - right, page.mediaBox.getUpperRight_y() - top)
		page.mediaBox.lowerLeft  = (page.mediaBox.getLowerLeft_x()  + left,  page.mediaBox.getLowerLeft_y()  + bottom)
		out.addPage(page)    

	ous = open("crop.pdf", 'wb')
	out.write(ous)
	ous.close()
	
def isPhone(inputString):
	return bool(re.search(r'[0-9]{10,11}', inputString))

#init
argparser = argparse.ArgumentParser()
argparser.add_argument('input',help = 'PDF path')
args = argparser.parse_args()
textTableSplitter = "Total Labels"

if fileExist(args.input):
	print("[{0}] File Found".format(args.input))
else:
	print("[{0}] File not Found".format(args.input))
	sys.exit(1)


crop(args.input)
fullText = textract.process("crop.pdf", method='pdftotext').decode("utf-8").replace("\r","")
fullText = fullText.replace("\n\n\n","\n\n")
fullText = fullText.replace("\n\n","\n")

lines = fullText.split("\n")
records = []
tmp = []
for line in lines:
	if line == "N/A": continue
	tmp.append(line)
	if isPhone(line):
		records.append(tmp)
		tmp = []
		
wb = Workbook() 
sheet1 = wb.add_sheet('Sheet 1') 
valids = []
invalids = []
for record in records:
	if (len)(record) > 8:
		invalids.append(record)
	else:	
		valids.append(record)

c=0
for valid in valids:
	print(valid)
		
	sheet1.write(c, 0, valid[1])
	sheet1.write(c, 1, valid[2])
	sheet1.write(c, 2, valid[3])
	sheet1.write(c, 3, valid[4])
	sheet1.write(c, 4, valid[-1])
	c+=1

tmp = []
for invalid in invalids:
	print(invalid)
		
	sheet1.write(c, 0, invalid[-7])
	sheet1.write(c, 1, invalid[-6])
	sheet1.write(c, 2, invalid[-5])
	sheet1.write(c, 3, invalid[-4])
	sheet1.write(c, 4, invalid[-1])
	tmp.append(invalid[-7])
	c+=1

wb.save('out.xls')

wb = Workbook() 
sheet1 = wb.add_sheet('Sheet 1') 
for c, invalid in enumerate(invalids):
	invalid = "".join(invalid)
	sheet1.write(c, 0, invalid[:invalid.find(tmp[c])])	
	
	
wb.save('nophones.xls')
os.remove("crop.pdf")
