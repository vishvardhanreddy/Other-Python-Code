'''
Created on Aug 6, 2014

@author: vvaka
'''

import sys
import re
import math
import xlrd

    
if len(sys.argv) != 3:
    sys.exit('Usage: python Table5-1 infile outfile')
    
  
inputFile = sys.argv[1]
outputFile = sys.argv[2]
oldNodeName = 'X'
    
try:
    workbook = xlrd.open_workbook(inputFile)
except IOError:
    print "[ERROR] Cound't open file: " + inputFile
    sys.exit()
    
try:
    outFile = open(outputFile,'w')
except IOError:
    print "[ERROR] Cound't write"+outputFile
    sys.exit()
    
worksheetname = workbook.sheet_names() 
worksheet = workbook.sheet_by_name(worksheetname[0])
num_rows = worksheet.nrows - 1
num_cells = worksheet.ncols -1
  
curr_row = -1
    
outFile.write('Node Name,Shelf/Slot/Port,Equipment,CTP Shelf/Slot/Port,CTP Equipment\n')

while curr_row < num_rows :
    curr_row += 1
    row = worksheet.row(curr_row)
    if (row[10].value) and (row[0].value != '#Node'):
        
        if row[4].value:
            columnB = str(int(row[3].value))+'/'+str(int(row[4].value))
        else:
            columnB = str(int(row[3].value))
        if row[6].value and row[8].value:
            outFile.write(str(row[0].value)+','+columnB+','+str(row[6].value)+','+columnB+','+str(row[8].value)+'\n')
            print row[6].value, row[8].value
        elif row[6].value:
            outFile.write(str(row[0].value)+','+columnB+','+str(row[6].value)+','+''+','+''+'\n')
        elif row[8].value:
            outFile.write(str(row[0].value)+','+''+','+''+','+columnB+','+str(row[8].value)+'\n')

        
