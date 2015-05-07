'''
Created on Jul 2, 2014

@author: vvaka
'''
'''
Created on Jul 2, 2014

@author: vvaka
'''

import re, sys, csv, xlrd
#================================================================================

def is_number(s):
  try:
    float(s)
    return True
  except ValueError:
    return False


def main():
    
  if len(sys.argv) != 3:
    sys.exit('Usage: Table21 infile outfile ')
    return None
  
  inputFile = sys.argv[1]
  outputFile = sys.argv[2]
  
  yellow = 0
  red = 0
  oldNodeName = 'X'
  try:
    workbook = xlrd.open_workbook(inputFile)
   
  except IOError:
    print "[ERROR] Cound't open file: " + inputFile
    return None
    
  try:
    outFile = open(outputFile,'w')
  except IOError:
    print "[ERROR] Cound't write"+outputFile
    return False
    
  worksheetname = workbook.sheet_names()
  
  
  worksheet = workbook.sheet_by_name(worksheetname[0])
  num_rows = worksheet.nrows - 1
  num_cells = worksheet.ncols -1
  
  curr_row = -1
  try:
    outFile = open(outputFile,'w')
  except IOError:
    print "[ERROR] Cound't write"+outputFile
    return False
    
  outFile.write('NodeName,Expected Min,Actual Loss,DHI\n')
  while curr_row < num_rows :
    curr_row += 1
    row = worksheet.row(curr_row)
    curr_row += 1
    row = worksheet.row(curr_row)
    nodeName = row[0].value
    minSpan = 'Min Expected Span Loss'
    maxSpan = 'Max Expected Span Loss'
    if (((row[7].value) in minSpan) or ((row[7].value) in maxSpan)) and ((row[7].value) != ''):
      print row[7].value
      outFile.write(row[0].value+','+row[7].value+','+row[14].value+','+row[9].value+'\n')

  outFile.close()
      
  
#================================================================================
if __name__ == "__main__":
  main()       