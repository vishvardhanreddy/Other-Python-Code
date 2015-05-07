'''
Created on Jul 3, 2014

@author: vvaka
'''


import re, sys, csv, math, xlsxwriter
#================================================================================
def is_number(s):
  try:
    float(s)
    return True
  except ValueError:
    return False

def main():
    
  if len(sys.argv) != 1:
    sys.exit('Usage: python Figure3-1')
    return None
  
  inputFile = '_Report_2014-05-29-140642.csv'
  outputFile = 'fig31autooutput.csv'
  
  green = 0
  yellow = 0
  red = 0
  oldNodeName = 'X'
  try:
    inFile = open(inputFile,'r')
  except IOError:
    print "[ERROR] Cound't open file: " + "_Report_2014-05-29-140642.csv"
    return None
    
  try:
    outFile = open(outputFile,'w')
  except IOError:
    print "[ERROR] Cound't write test.csv"
    return False
  
  book = xlsxwriter.Workbook('Figure31.xlsx')  
  
  

  
  curr_row = -1

# Setting up size of data columns in Excel sheet
  sheetname = 'Patchcord DHI Values'
  sheet = book.add_worksheet(sheetname) 
  sheet.set_column('A:A', 15)
  sheet.set_column('B:B', 20)
  sheet.set_column('C:C', 20)

 
  
  # defining bold format
  bold = book.add_format({'bold':True})
  # Defining titles for Data

  sheet.write('A1', 'Green', bold)
  sheet.write('B1', 'Yellow', bold)
  sheet.write('C1', 'Red', bold)   
  
    # Setting size of the plot chart
  chart1 = book.add_chart({'type': 'column'})
  chart1.set_size({'width':720, 'height':586})
  chart1.set_plotarea({
    'layout':{
    'x': 0.3,
    'y': 0.4,
    'width' : 0.90,
    'height': 0.70,
     }
  }) 
  excel_row = 2
  
  for line in inFile:    
    line = re.sub('[\n\r]','',line)    
    flds = line.split(',')
    nodeName = flds[0]
    
    if (is_number(flds[6])):
      delta = math.fabs(float(flds[6]))
      if (delta <= 1.0):    
        green = green +1

      elif (delta <= 2.0):
        yellow = yellow +1

      else:
        red = red+1
    
  outFile.write(oldNodeName+','+str(green)+','+str(yellow)+','+str(red)+'\n')
  data = (green, yellow, red)
  sheet.write_row('A'+str(excel_row), data)
  totalPatchcordPowerLevels = green + yellow + red 
  print 'totalPatchcord Power Levels:', totalPatchcordPowerLevels
  average = (green+yellow+red)/3
  print 'average:', average  
   # Setting series for each data set 
  chart1.add_series({
    'name':       'Green',
    'categories': '='+sheetname+'!$D$2:$D$2',
    'values':     '='+sheetname+'!$A$2:$A$2',
    'fill': {'color':'green'},
    'border':{'color':'black'}
  })
  
  chart1.add_series({
    'name':       'Yellow',
    'categories': '='+sheetname+'!$B$2:$B$8',
    'values':     '='+sheetname+'!$B$2:$B$8',
    'fill': {'color':'yellow'},
    'border':{'color':'black'}
  })
  
  chart1.add_series({
    'name':       'Red',
    'categories': '='+sheetname+'!$C$2:$C$2',
    'values':     '='+sheetname+'!$C$2:$C$2',
    'fill': {'color':'red'},
    'border':{'color':'black'}
  })
  
  # Add a chart title and some axis labels.
  chart1.set_title ({'name': 'Count of Patchcord DHI Values'})
  chart1.set_x_axis({'name': 'DHI Values','interval_unit': 1, 'num_font': {'rotation':-45}})
  chart1.set_y_axis({'name': 'Count'})


# Insert the chart into the worksheet (with an offset).
  sheet.insert_chart('C9', chart1, {'x_offset': 7, 'y_offset': 7})


# exporting to Excel file
  book.close() 
  outFile.close()     
  inFile.close()
  
#================================================================================
if __name__ == "__main__":
  main()