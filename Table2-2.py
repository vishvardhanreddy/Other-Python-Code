
'''
Created on Jul 2, 2014

@author: vvaka
'''
#This script will read data from WDM_ANS Excel Analysis and output Node Name, Shelf/Slot/Port, Equipment, 
#Wavelength, Origin, Parameter, CTP Value, Card Reading and DHI

# DHI is determined as the difference between CTP Value and Card reading 
# Difference of less than 0db is considered Red and less than 1db considered to be Yellow


import re, sys, csv, xlrd, math
#================================================================================

# This function checks if the number is passed is number or character, if number returns True, else it will
# return False
def is_number(s):
  try:
    float(s)
    return True
  except ValueError:
    return False


def main():

#Input statement requires to include both input file name and output CSV file name      
  if len(sys.argv) != 3:
    sys.exit('Usage: Table21 infile outfile ')
    return None
  
  inputFile = sys.argv[1]
  # First argument is considered to be Input file name
  outputFile = sys.argv[2]
  # Second argument is considered to be Output file name
  
 
  try:
    workbook = xlrd.open_workbook(inputFile)
   # Opening input file
  except IOError:
    print "[ERROR] Cound't open file: " + inputFile
    return None
    
  try:
    outFile = open(outputFile,'w')
    # Open output file
  except IOError:
    print "[ERROR] Cound't write"+outputFile
    return False
  
  greenDHI = 0
  yellowDHI = 0
  redDHI = 0  
  worksheetname = workbook.sheet_names()
# Reading all the worksheet names  
  
  worksheet = workbook.sheet_by_name(worksheetname[0])
  num_rows = worksheet.nrows - 1
  num_cells = worksheet.ncols -1
  
  curr_row = -1
  try:
    outFile = open(outputFile,'w')
  except IOError:
    print "[ERROR] Cound't write"+outputFile
    return False
    
  outFile.write('Node Name,Shelf/Slot/Port,Equipment(Type),Wave length(Lambda),Origin(Flag),Parameter,ANS Value,CTP Value,DHI\n')
  while curr_row < num_rows :
    curr_row += 1
    row = worksheet.row(curr_row)

    if (is_number(row[13].value)) and (row[13].value != 0):
# Column N values in file are DHI values between Card level readings and CTP are captured in row[13] 
      DHI = math.fabs(float(row[13].value))
      if DHI == 0.0:
        greenDHI = greenDHI + 1
      if (DHI != 0.0):
        if DHI > 1:
          DHIColor = 'RED'
          redDHI = redDHI + 1
        elif DHI <= 1:
          DHIColor = 'YELLOW' 
          yellowDHI = yellowDHI + 1
        port = int(row[8].value)
        port = str(port)
        print port 
        row2 = str(int(row[1].value))+'/'+str(int(row[2].value))
        outFile.write(row[0].value+','+row2+'/'+port+','+row[4].value+','+row[5].value+','+row[6].value+','+row[7].value+','+row[9].value+','+row[12].value+','+DHIColor+'\n')
  
  print redDHI
  print yellowDHI
  print greenDHI
  totalNo = redDHI+yellowDHI+greenDHI
  total = (redDHI*1)+(yellowDHI*3)+(5*greenDHI) 
  print total
  average = total/float(totalNo)
  print average
  outFile.close() 
  outFile.close()     

  
#================================================================================
if __name__ == "__main__":
  main()