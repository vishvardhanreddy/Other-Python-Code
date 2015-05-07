from graphlib import plot

import re, sys, csv, xlrd, math, xlsxwriter
#================================================================================


def main():
    
  if len(sys.argv) != 3:
    sys.exit('Usage: Figure2-1 infile outfile ')
    return None
  
  inputFile = '_WDM-ANS_2014-06-25-162433.xlsx'
  outputFile = 'fig21output.csv'
  
  yellow = 0
  red = 0
  green = 0
  greenCTP = 0
  yellowCTP = 0
  redCTP = 0
  greenCTPLos = 0
  yellowCTPLos = 0
  redCTPLos = 0
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
  book = xlsxwriter.Workbook('Figure21.xlsx')  
  worksheetname = workbook.sheet_names()
  
  
  worksheet = workbook.sheet_by_name(worksheetname[0])
  num_rows = worksheet.nrows - 1
  num_cells = worksheet.ncols -1
  
  curr_row = -1

# Setting up size of data columns in Excel sheet
  sheetname = 'Span loss Evaluation'
  sheet = book.add_worksheet(sheetname) 
  sheet.set_column('A:A', 15)
  sheet.set_column('B:B', 20)
  sheet.set_column('C:C', 20)
  sheet.set_column('D:D', 20)
 
  # defining bold format
  bold = book.add_format({'bold':True})
  # Defining titles for Data
  sheet.write('A1','Green', bold)
  sheet.write('B1', 'Yellow', bold)
  sheet.write('C1', 'Red', bold)
      
  outFile.write('NodeName,Green,Yellow,Red\n')
  
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
  green = 0
  yellow = 0
  red = 0
  while curr_row < num_rows :
    curr_row += 1
    row = worksheet.row(curr_row)
    nodeName = row[0].value
   
    if (is_number(row[9].value)) and (is_number(row[10].value)):
      DHI21 = float(row[9].value) - float(row[10].value)  
      DHI21 = math.fabs(DHI21)
         #outFile.write(nodeName+',Green,Yellow,Red'+'\n')
      if (DHI21 == 0.0):    
        green = green +1

      elif (math.fabs(DHI21) <= 1.0):
        yellow = yellow +1

      elif (math.fabs(DHI21) > 1.0):
        red = red+1
    
    if (is_number(row[13].value)):
          
      if ((float(row[13].value)) == 0):    
        green = green +1
      elif ((float(row[13].value)) <= 1.0):
        yellow = yellow +1
      elif ((row[13].value) > 1.0):
        red = red+1
        
    if (is_number(row[15].value)):
      if ((float(row[15].value)) >= 1.0):
        green = green + 1
      elif ((float(row[15].value)) >= 0.0):
        yellow = yellow + 1
      elif ((float(row[15].value)) < 0.0):
        red = red + 1
          
  
  totalANSSetPoints = green + yellow + red 
  print 'Total ANSSetPoints:', totalANSSetPoints
  average = (green+yellow+red)/3
  print 'average:', average     
  outFile.write(str(green)+','+str(yellow)+','+str(red)+'\n')    
  data = (green, yellow, red,'DHI Values')
  sheet.write_row('A'+str(excel_row), data)
  outFile.close()
 
 # Setting series for each data set 
  chart1.add_series({
    'name':       'Green',
    'categories': '='+sheetname+'!$D$2',
    'values':     '='+sheetname+'!$A$2:$A$2',
    'fill': {'color':'green'},
    'border':{'color':'black'}
  })
  
  chart1.add_series({
    'name':       'Yellow',
    'categories': '='+sheetname+'!$B$2',
    'values':     '='+sheetname+'!$B$2:$B$2',
    'fill': {'color':'yellow'},
    'border':{'color':'black'}
  })
  
  chart1.add_series({
    'name':       'Red',
    'categories': '='+sheetname+'!$C$2',
    'values':     '='+sheetname+'!$C$2:$C$2',
    'fill': {'color':'red'},
    'border':{'color':'black'}
  })
  
  # Add a chart title and some axis labels.
  chart1.set_title ({'name': 'Count of ANS DHI Values'})
  chart1.set_x_axis({'name': '','interval_unit': 1, 'num_font': {'rotation':-45}})
  chart1.set_y_axis({'name': 'Count'})


# Insert the chart into the worksheet (with an offset).
  sheet.insert_chart('C9', chart1, {'x_offset': 7, 'y_offset': 7})
  sample_attrs = {}
sample_attrs["plot_dir"] = "./plots"

sample_attrs["sort"] = 1
sample_attrs["title"] = "bar plot"
sample_attrs["plot_file"] = "bar_plot.png"
bar1 = plot.Plot('bar', "sample_data/bar.csv", sample_attrs)
bar1.gen_plot()

# exporting to Excel file
  book.close() 
#================================================================================
if __name__ == "__main__":
  main()