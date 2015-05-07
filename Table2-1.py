#This script will read data from WDM_ANS Excel Analysis and output Node Name, Shelf/Slot/Port, Equipment, 
#Wavelength, Origin, Parameter, ANs Value, Card Reading and DHI

# DHI is determined as the difference between ANS Value and Card reading 
# Difference of less than 1db is considered to Yellow and more than 1db considered to be Red



import re, sys, csv, xlrd, math, xlsxwriter
#================================================================================

def is_number(s):
# This function checks if the number is passed is number or character, if number returns True, else it will
# return False

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
  greenDHI = 0
  yellowDHI = 0
  redDHI = 0
  oldNodeName = 'X'
  DHITotal = 0
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
    
  worksheetname = workbook.sheet_names()
# Reading all the worksheet names  
  excelbook = xlsxwriter.Workbook('Table21.xlsx')
  tab21sheet = excelbook.add_worksheet()
  
  worksheet = workbook.sheet_by_name(worksheetname[0])
  num_rows = worksheet.nrows - 1
  num_cells = worksheet.ncols -1
# Processing first worksheet in the input file
  
  curr_row = -1
  tab21sheet.add_table('A1:J'+str(num_rows),{'banded_rows': False})
  format = excelbook.add_format()
  format.set_pattern(1)  
  outFile.write('Node Name,Shelf/Slot/Port,Equipment(Type),Wave length(Lambda),Origin(Flag),Parameter,ANS Value,Card Reading,DHI\n')
  while curr_row < num_rows :
    curr_row += 1
    row = worksheet.row(curr_row)
# if there are values either for WDM-ANS or Actual Card
    if (is_number(row[9].value)) and (is_number(row[10].value)):
      # if there are values for both WDM-ANS and Actual Card, DHI is the difference between them  
      DHI = float(row[9].value) - float(row[10].value)  
      DHITotal = DHI+DHITotal
      # if only WDM-ANS value is available, then DHI is WDM-ANS

      
      # Since we are only picking Yellow and Red DHI, we are picking absolute value of DHI only if it is
      # not equal to 0
      DHI = math.fabs(DHI)
      if DHI == 0.0:
        greenDHI  = greenDHI + 1   
      if (DHI != 0.0): 
        if DHI > 1:
          DHIColor = 'RED'
          redDHI = redDHI + 1
          format.set_bg_color('#FF0000')
        elif DHI <= 1:
          DHIColor = 'YELLOW'
          yellowDHI = yellowDHI + 1    
          format.set_bg_color('#FFFF00')
        StrinDHI = str(DHI)
        row2 = str(int(row[1].value))+'/'+str(int(row[2].value))+'/'+str(int(row[8].value))
        outFile.write(row[0].value+','+row2+','+row[4].value+','+row[5].value+','+row[6].value+','+row[7].value+','+row[9].value+','+row[10].value+','+DHIColor+'\n')
        data = [row[0].value, row2, row[4].value, row[5].value, row[6].value, row[7].value, row[9].value, row[10].value, DHIColor]
        tab21sheet.write_row('A'+str(curr_row), data, format)
  
  print 'RED:', redDHI
  print 'Yellow:', yellowDHI
  print 'Green:', greenDHI
  totalNo = redDHI+yellowDHI+greenDHI
  total = (redDHI*1)+(yellowDHI*3)+(5*greenDHI) 
  print 'Total:', total
  average = total/float(totalNo)
  print average
  outFile.close()     
 
  
#================================================================================
if __name__ == "__main__":
  main()