'''
Created on Jul 28, 2014

@author: vvaka
'''
import re 
import sys

if len(sys.argv) != 1:
  sys.exit('Usage: python Table4-3.py')
  
inputFile = '_pRTRV-OCH.csv'
outputFile = 'Table4.3.csv'

try:
  inFile = open(inputFile,'r')
except IOError:
  print "[ERROR] Cound't open file: " + inputFile
try:
  outFile = open(outputFile,'w')
except IOError:
  print "[ERROR] Cound't write Table4.3.csv"
outFile.write('Node Name,Facility,FEC Status,FEC Type\n')   
for line in inFile:
  line = re.sub('[\n\r]','',line)    
  flds = line.split(',')
  if len(flds) >= 17 and flds[17] != '' and flds[0] != '#NODE':
    if flds[17] == 'OFF':
      FECStatus = 'OFF'
      FECType = ''
    else:
      FECStatus = 'ON'
      FECType = flds[17]
        
    outFile.write(flds[0]+','+flds[2]+','+FECStatus+','+FECType+'\n')

                   