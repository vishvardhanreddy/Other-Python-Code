'''
Created on Jul 30, 2014

@author: vvaka
'''


import re, sys, csv, xlrd, math, xlsxwriter
#================================================================================

#Input statement requires to include both input file name and output CSV file name    
def is_number(s):
# This function checks if the number is passed is number or character, if number returns True, else it will
# return False

  try:
    float(s)
    return True
  except ValueError:
    return False




def main():
    
    
  inputFile = '_WDM-ANS_2014-06-25-162433.xlsx'
# First argument is considered to be Input file name
  
# Second argument is considered to be Output file name
  
  try:
    workbook = xlrd.open_workbook(inputFile)
# Opening input file   
  except IOError:
    print "[ERROR] Cound't open file: " + inputFile

    
  worksheetname = workbook.sheet_names()
# Reading all the worksheet names  
  excelbook = xlsxwriter.Workbook('Table21.xlsx')
  tab21sheet = excelbook.add_worksheet('Table2.1')
  
  worksheet = workbook.sheet_by_name(worksheetname[0])
  num_rows = worksheet.nrows - 1
  num_cells = worksheet.ncols -1
# Processing first worksheet in the input file
  
  curr_row = 1

  tab21sheet.add_table('A1:H'+str(num_rows),{'banded_rows': False, 'autofilter': False })
  tab21sheet.set_column('A:A', 8)
  tab21sheet.set_column('B:B', 8)
  tab21sheet.set_column('C:C', 20)
  tab21sheet.set_column('D:D', 8)
  tab21sheet.set_column('E:E', 10)
  tab21sheet.set_column('F:F', 10)
  tab21sheet.set_column('G:G', 8)
  tab21sheet.set_column('H:H', 8)
  format = excelbook.add_format()
  format.set_pattern(1) 
  format.set_bg_color('#0000FF')
  format.set_bold()
  format.set_align('justify')
  format.set_font_size(8)
  format.set_text_wrap()
  data = ('NodeName','Shelf/Slot/Port','Equipment(Type)','Wavelength(Lambda)','Origin(Flag)','Parameter','ANS Value','CardReading') 
  tab21sheet.write_row('A'+str(curr_row), data, format)
  excel_row = 1
  while curr_row < num_rows :
    curr_row += 1
    row = worksheet.row(curr_row)
# if there are values either for WDM-ANS or Actual Card
    if (is_number(row[9].value)) or (is_number(row[10].value)):
      # if there are values for both WDM-ANS and Actual Card, DHI is the difference between them  
      if (is_number(row[9].value)) and (is_number(row[10].value)):
        DHI = float(row[9].value) - float(row[10].value)  
      # if only WDM-ANS value is available, then DHI is WDM-ANS
      elif (is_number(row[9].value)) and not (is_number(row[10].value)):
        DHI = float(row[9].value)
      # if only Card value is available, then DHI is Card value 
      elif (is_number(row[10].value)) and not (is_number(row[9].value)):
        DHI = float(row[10].value)
      
      # Since we are only picking Yellow and Red DHI, we are picking absolute value of DHI only if it is
      # not equal to 0
      DHI = math.fabs(DHI)
      if (DHI != 0.0):
        excel_row = excel_row + 1 
        if DHI > 1:
          DHIColor = 'red'
          format = excelbook.add_format()
          format.set_pattern(1) 
          format.set_bg_color('#FF0000')
        elif DHI <= 1:
          DHIColor = 'yellow' 
          format = excelbook.add_format()
          format.set_pattern(1)    
          format.set_bg_color('#FFFF00')
        format.set_text_wrap()
        row2 = str(row[1].value)+'/'+str(row[2].value)+'/'+str(row[8].value)
        data = [row[0].value, row2, row[4].value, row[5].value, row[6].value, row[7].value, row[9].value, row[10].value]
        format.set_font_size(8)
        tab21sheet.write_row('A'+str(excel_row), data, format)

 
  
#================================================================================
if __name__ == "__main__":
  main()