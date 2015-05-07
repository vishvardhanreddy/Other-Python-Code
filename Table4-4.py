'''
Created on Jul 29, 2014

@author: vvaka
'''
import re 
import sys
import math

holdIP = []
if len(sys.argv) != 1:
  sys.exit('Usage: python Table4-3.py')
  
inputFile = '_pRTRV-PM-ALL.csv'
outputFile = 'Table4.4.csv'

try:
  inFile = open(inputFile,'r')
except IOError:
  print "[ERROR] Cound't open file: " + inputFile
try:
  outFile = open(outputFile,'w')
except IOError:
  print "[ERROR] Cound't write Table4.3.csv"
   
for line in inFile:
  line = re.sub('[\n\r]','',line)    
  flds = line.split(',')
  if flds[0] == '#NODE':
    outFile.write('Node Name, Facility, Pre-FEC Bit Error Count, Uncorrected Word Count, DHI\n')
  else:
    if flds[4] == 'BIT-EC':
      if holdIP:
        outFile.write(holdIP[0]+','+holdIP[2]+','+holdIP[5]+','+UNCW+','+DHI+'\n') 
      
      holdIP = flds
      DHI = ''
      UNCW = '0'
    elif (flds[4] == 'UNC-Words') and (holdIP[1] == flds[1]):
      if flds[5] > 0:
        DHI =  'RED'
      elif (holdIP[5] > 0 and flds[5] == 0):
        DHI = 'YELLOW'
      else:
        DHI = ''
      UNCW = flds[5]
    elif holdIP:
      if holdIP[1] != flds[1]:
        outFile.write(holdIP[0]+','+holdIP[2]+','+holdIP[5]+','+UNCW+','+DHI+'\n')  
        holdIP = []
    
    