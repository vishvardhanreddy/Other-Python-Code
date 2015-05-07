'''
Created on Jul 30, 2014

@author: vvaka
'''

import re, sys, csv, xlrd, math, xlsxwriter

def is_number(s):
# This function checks if the number is passed is number or character, if number returns True, else it will
# return False

  try:
    float(s)
    return True
  except ValueError:
    return False

def create_table21(book):
  
  inputFile = '_WDM-ANS_2014-06-25-162433.xlsx'
# First argument is considered to be Input file name
  excelbook = book
# Second argument is considered to be Output file name
  
  try:
    workbook = xlrd.open_workbook(inputFile)
# Opening input file   
  except IOError:
    print "[ERROR] Cound't open file: " + inputFile

    
  worksheetname = workbook.sheet_names()
# Reading all the worksheet names  
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


def create_table22(book):
  excelbook = book
  inputFile = '_WDM-ANS_2014-06-25-162433.xlsx'
  try:
    workbook = xlrd.open_workbook(inputFile)
# Opening input file   
  except IOError:
    print "[ERROR] Cound't open file: " + inputFile

    
  worksheetname = workbook.sheet_names()
# Reading all the worksheet names  
  tab22sheet = excelbook.add_worksheet('Table2.2')
  
  worksheet = workbook.sheet_by_name(worksheetname[0])
  num_rows = worksheet.nrows - 1
  num_cells = worksheet.ncols -1
# Processing first worksheet in the input file
  
  curr_row = -1

  tab22sheet.add_table('A1:H'+str(num_rows),{'banded_rows': False, 'autofilter': False })
  tab22sheet.set_column('A:A', 8)
  tab22sheet.set_column('B:B', 8)
  tab22sheet.set_column('C:C', 20)
  tab22sheet.set_column('D:D', 8)
  tab22sheet.set_column('E:E', 10)
  tab22sheet.set_column('F:F', 10)
  tab22sheet.set_column('G:G', 8)
  tab22sheet.set_column('H:H', 8)
  format = excelbook.add_format()
  format.set_pattern(1) 
  format.set_bg_color('#0000FF')
  format.set_bold()
  format.set_align('justify')
  format.set_font_size(8)
  format.set_text_wrap()
  data = ('NodeName','Shelf/Slot/Port','Equipment(Type)','Wavelength(Lambda)','Origin(Flag)','Parameter','ANS Value','CTP Value') 
  excel_row = 1
  tab22sheet.write_row('A'+str(excel_row), data, format)
  while curr_row < num_rows :
    curr_row += 1
    row = worksheet.row(curr_row)

    if (is_number(row[13].value)) and (row[13].value != 0):
# Column N values in file are DHI values between Card level readings and CTP are captured in row[13] 
      DHI = math.fabs(float(row[13].value))
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
        data = [row[0].value, row2, row[4].value, row[5].value, row[6].value, row[7].value, row[9].value, row[12].value]
        format.set_font_size(8)
        tab22sheet.write_row('A'+str(excel_row), data, format)  

def create_table23(book):
    
  inputFile = '_WDM-ANS_2014-06-25-162433.xlsx'
  excelbook = book  
  try:
    workbook = xlrd.open_workbook(inputFile)
# Opening input file   
  except IOError:
    print "[ERROR] Cound't open file: " + inputFile

    
  worksheetname = workbook.sheet_names()
# Reading all the worksheet names  
  tab23sheet = excelbook.add_worksheet('Table2.3')
  
  worksheet = workbook.sheet_by_name(worksheetname[0])
  num_rows = worksheet.nrows - 1
  num_cells = worksheet.ncols -1
# Processing first worksheet in the input file
  
  curr_row = -1

  tab23sheet.add_table('A1:H'+str(num_rows),{'banded_rows': False, 'autofilter': False })
  tab23sheet.set_column('A:A', 8)
  tab23sheet.set_column('B:B', 8)
  tab23sheet.set_column('C:C', 20)
  tab23sheet.set_column('D:D', 8)
  tab23sheet.set_column('E:E', 10)
  tab23sheet.set_column('F:F', 10)
  tab23sheet.set_column('G:G', 8)
  tab23sheet.set_column('H:H', 8)
  format = excelbook.add_format()
  format.set_pattern(1) 
  format.set_bg_color('#0000FF')
  format.set_bold()
  format.set_align('justify')
  format.set_font_size(8)
  format.set_text_wrap()
  data = ('NodeName','Shelf/Slot/Port','Equipment(Type)','Wavelength(Lambda)','Origin(Flag)','Parameter','Actual Power','CTP Loss Threshold') 
  excel_row = 1
  tab23sheet.write_row('A'+str(excel_row), data, format)  
  
  while curr_row < num_rows :
    curr_row += 1
    row = worksheet.row(curr_row)

    if (is_number(row[15].value)) and (row[15].value != 0):
# Column N values in file are DHI values between Card level readings and CTP are captured in row[13] 
      DHI = float(row[15].value)
      if not (DHI >= 1):
        excel_row = excel_row + 1 
        if DHI < 0:
          DHIColor = 'red'
          format = excelbook.add_format()
          format.set_pattern(1) 
          format.set_bg_color('#FF0000')
        elif DHI < 1:
          DHIColor = 'yellow' 
          format = excelbook.add_format()
          format.set_pattern(1)    
          format.set_bg_color('#FFFF00')
        format.set_text_wrap()
        row2 = str(row[1].value)+'/'+str(row[2].value)+'/'+str(row[8].value)
        data = [row[0].value, row2, row[4].value, row[5].value, row[6].value, row[7].value, row[14].value, row[12].value]
        format.set_font_size(8)
        tab23sheet.write_row('A'+str(excel_row), data, format)  

def create_table24(book):

  inputFile = '_WDM-ANS_2014-06-25-162433.xlsx'
  excelbook = book  
  try:
    workbook = xlrd.open_workbook(inputFile)
# Opening input file   
  except IOError:
    print "[ERROR] Cound't open file: " + inputFile

    
  worksheetname = workbook.sheet_names()
# Reading all the worksheet names  
  tab24sheet = excelbook.add_worksheet('Table2.4')
  
  worksheet = workbook.sheet_by_name(worksheetname[0])
  num_rows = worksheet.nrows - 1
  num_cells = worksheet.ncols -1
# Processing first worksheet in the input file
  
  curr_row = -1

  tab24sheet.add_table('A1:H'+str(num_rows),{'banded_rows': False, 'autofilter': False })
  tab24sheet.set_column('A:A', 8)
  tab24sheet.set_column('B:B', 8)
  tab24sheet.set_column('C:C', 15)
  tab24sheet.set_column('D:D', 15)

  format = excelbook.add_format()
  format.set_pattern(1) 
  format.set_bg_color('#0000FF')
  format.set_bold()
  format.set_align('justify')
  format.set_font_size(8)
  format.set_text_wrap()
  data = ('NodeName','Side','Exp Min/Exp Max','Actual Loss') 
  excel_row = 1
  tab24sheet.write_row('A'+str(excel_row), data, format)  

  while curr_row < num_rows :
    curr_row += 1
    row = worksheet.row(curr_row)
    
    if ('Span Loss' in (row[7].value)) and (float(row[9].value != 'N/A')):
      
      a = re.findall("\d+.\d+",str(row[14].value))
      
      #a = filter(str.isdigit, str(row[14].value)) 
      if a:
        aValue = a.pop(0)
      else:
        aValue = 0.0   
      TC = math.fabs(float(aValue) - (float(row[9].value)))
      ThresholdCenterA = float(aValue) +((TC)*0.25)
      ThresholdCenterB = float(row[9].value)+((TC)*0.25)
      
      if  TC < ThresholdCenterA:
        DHI = 'Green'
        format = excelbook.add_format()
        format.set_pattern(1)    
        format.set_bg_color('#008000')
      elif TC >= ThresholdCenterB:
        DHI = 'Yellow'
        format = excelbook.add_format()
        format.set_pattern(1)    
        format.set_bg_color('#FFFF00')
      else:
        DHI = 'Red'
        format = excelbook.add_format()
        format.set_pattern(1) 
        format.set_bg_color('#FF0000')
      excel_row = excel_row + 1 
      format.set_text_wrap()
      data = [row[0].value, '', row[9].value, row[14].value]
      format.set_font_size(8)
      tab24sheet.write_row('A'+str(excel_row), data, format) 

def create_table31(book):

  inputFile ='_Report_2014-05-29-140642.csv'
  excelbook = book 
  try:
    inFile = open(inputFile,'r')
  except IOError:
    print "[ERROR] Cound't open file: " + inputFile
    return None
  tab31sheet = excelbook.add_worksheet('Table3.1')
  
# Processing first worksheet in the input file
  
  curr_row = -1

  tab31sheet.add_table('A1:H50',{'banded_rows': False, 'autofilter': False })
  tab31sheet.set_column('A:A', 8)
  tab31sheet.set_column('B:B', 8)

  format = excelbook.add_format()
  format.set_pattern(1) 
  format.set_bg_color('#0000FF')
  format.set_bold()
  format.set_align('justify')
  format.set_font_size(8)
  format.set_text_wrap()
  data = ('Attenuator/DCU','Loss(dB)',) 
  excel_row = 1
  tab31sheet.write_row('A'+str(excel_row), data, format)  
 
  count = 0
  for line in inFile:    
    line = re.sub('[\n\r]','',line)    
    flds = line.split(',')
    if is_number(flds[13]):
      excel_row = excel_row + 1 
      format = excelbook.add_format()
      format.set_pattern(1) 
      format.set_bg_color('#00FF00')
      format.set_text_wrap()
      data = [flds[12], flds[13]]
      format.set_font_size(8)
      tab31sheet.write_row('A'+str(excel_row), data, format)     

def create_table32(book):
  
  inputFile ='_Report_2014-05-29-140642.csv'
  excelbook = book 
  try:
    inFile = open(inputFile,'r')
  except IOError:
    print "[ERROR] Cound't open file: " + inputFile
    return None
  tab32sheet = excelbook.add_worksheet('Table3.2')
  
# Processing first worksheet in the input file
  
  curr_row = -1

  tab32sheet.add_table('A1:I30',{'banded_rows': False, 'autofilter': False })
  tab32sheet.set_column('A:A', 8)
  tab32sheet.set_column('B:B', 8)
  tab32sheet.set_column('C:C', 10)
  tab32sheet.set_column('D:D', 5)
  tab32sheet.set_column('E:E', 10)
  tab32sheet.set_column('F:F', 10)
  tab32sheet.set_column('G:G', 8)
  tab32sheet.set_column('H:H', 8)
  tab32sheet.set_column('I:I',5)
  format = excelbook.add_format()
  format.set_pattern(1) 
  format.set_bg_color('#0000FF')
  format.set_bold()
  format.set_align('justify')
  format.set_font_size(8)
  format.set_text_wrap()
  data = ('TX NODE','TX SHELF/SLOT/PORT','TX EQPT','RX NODE','RX SHELF/SLOT/PORT','RX EQPT','POWER DELTA','ATTENUATOR LOSS','DCU LOSS') 
  excel_row = 1
  tab32sheet.write_row('A'+str(excel_row), data, format)  
 
  count = 0
  for line in inFile:    
    line = re.sub('[\n\r]','',line)    
    flds = line.split(',')
    if (is_number(flds[6])) and ((math.fabs(float(flds[6]))) > 2.0):
      if not flds[12]: 
        flds[12] = 'Info Not Available'                 
      excel_row = excel_row + 1 
      format = excelbook.add_format()
      format.set_pattern(1) 
      format.set_bg_color('#FF0000')
      format.set_text_wrap()
      data = [flds[0],flds[4],flds[3],flds[8],flds[11],flds[10],flds[6],flds[12],flds[13]]
      format.set_font_size(8)
      tab32sheet.write_row('A'+str(excel_row), data, format)     
   

  
def main():
    
  if len(sys.argv) == 2:
    create_table = sys.argv[1]
  else:
    book = xlsxwriter.Workbook('DWDMTablesandFigures.xlsx')
    create_table21(book)
    create_table22(book)
    create_table23(book)
    create_table24(book)
    create_table31(book)
    create_table32(book)
#================================================================================
if __name__ == "__main__":
  main() 