'''
Created on Jul 9, 2014

@author: vvaka
'''

import sys, re, math, xlrd
#================================================================================

def is_number(s):
  try:
    float(s)
    return True
  except ValueError:
    return False


def main():
    
  if len(sys.argv) != 3:
    sys.exit('Usage: Table2-4 infile outfile ')
    return None
  
  inputFile = sys.argv[1]

  outputFile = sys.argv[2]
  yellow = 0
  red = 0
  oldNodeName = 'X'
  greenDHI = 0
  yellowDHI = 0
  redDHI = 0
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
    
  outFile.write('Node Name,Side,Expected Min,Expected Max,Actual Loss,DHI\n')
  while curr_row < num_rows :
    curr_row += 1
    row = worksheet.row(curr_row)
    
    if ('Min Expected Span Loss' in (row[7].value)) and (float(row[9].value != 'N/A')):
      aValue = float(row[9].value)
      ActualLoss = row[14].value
      ActualValue = re.findall("\d+.\d+",str(row[14].value))
      
    elif ('Max Expected Span Loss' in (row[7].value)) and (float(row[9].value != 'N/A')):
      bValue = float(row[9].value)     
      

     
      #a = filter(str.isdigit, str(row[14].value)) 
   
      C = bValue - aValue
      ThresholdCenterA = (aValue+(C*0.25))
      ThresholdCenterB = (bValue-(C*0.25))
      if ActualValue:
        if  ActualValue[0] < ThresholdCenterA:
          DHI = 'GREEN'
          greenDHI = greenDHI + 1
        elif ActualValue[0] >= ThresholdCenterB:
          DHI = 'YELLOW'
          yellowDHI = yellowDHI + 1
        else:
          DHI = 'RED'
          redDHI = redDHI + 1
      else:
        DHI = 'N/A'
      outFile.write(row[0].value+','+''+','+str(aValue)+','+str(bValue)+','+ActualLoss+','+DHI+'\n')
      aValue = 'N/A'
      ActualLoss = 'N/A'
  
  print redDHI
  print yellowDHI
  print greenDHI
  totalNo = redDHI+yellowDHI+greenDHI
  total = (redDHI*1)+(yellowDHI*3)+(5*greenDHI) 
  print total
  average = total/float(totalNo)
  print average    
  outFile.close()     

  
#================================================================================
if __name__ == "__main__":
  main()