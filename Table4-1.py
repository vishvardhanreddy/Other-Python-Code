'''
Created on Jul 9, 2014

@author: vvaka
'''
'''
Created on Jul 9, 2014

@author: vvaka
'''

import sys, xlrd, math
#================================================================================

def is_number(s):
  try:
    float(s)
    return True
  except ValueError:
    return False


def main():
    
  if len(sys.argv) != 3:
    sys.exit('Usage: python Table4-1.py infile outfile ')
    return None
  
  inputFile = sys.argv[1]
  outputFile = sys.argv[2]
  
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
  rowCount = 0
  curr_row = -1
  try:
    outFile = open(outputFile,'w')
  except IOError:
    print "[ERROR] Cound't write"+outputFile
    return False
    
  outFile.write('Node Name,Facility,Equipment,Pluggable,High,Low,Tx Power,Tx DHI\n')
  while curr_row < num_rows :
    curr_row += 1
    row = worksheet.row(curr_row)
    curr_row += 1
    row = worksheet.row(curr_row)
    nodeName = row[0].value
    columnD = 'TX'
    columnH = 'Threshold'

    if ((columnD in row[3].value) and (columnH in row[7].value)):
      if  is_number(row[10].value):
        a = row[14].value
        b = row[10].value
      elif is_number(row[9].value):
        a = row[14].value
        b = row[9].value
      c = (math.fabs(float(a)-float(b)))/2
     # print a, b, c  
      TCA = float(a)+(c*0.25)
      TCB = float(b)-(c*0.25)
      
      if (TCA >= TCB):
        High = TCA
        Low = TCB
      else:
        High = TCB
        Low = TCA
      
      if row[14].value < Low:
        DHI = 'GREEN'
        rowCount = rowCount +1 
      elif row[14].value >= High:
        DHI =  'YELLOW'
        rowCount = rowCount +1 
      else:
        DHI = 'RED'
        rowCount = rowCount +1   
      outFile.write(str(row[0].value)+','+str(row[2].value)+','+str(row[3].value)+','+str(row[7].value)+','+str(High)+','+str(Low)+','+str(row[14].value)+','+DHI+'\n')
  if rowCount == 0:
    outFile.write('N/A,N/A,N/A,N/A,N/A,N/A,N/A,N/A\n') 
  outFile.close()
  #================================================================================
if __name__ == "__main__":
  main()    