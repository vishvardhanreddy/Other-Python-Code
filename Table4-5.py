'''
Created on Jul 29, 2014

@author: vvaka
'''
import re 
import sys
import math

greenDHI = 0
yellowDHI = 0
redDHI = 0
holdIP = []
if len(sys.argv) != 1:
  sys.exit('Usage: python Table4-5.py')
  
inputFile = '_pRTRV-OCH.csv'
outputFile = 'Table4.5.csv'

try:
  inFile = open(inputFile,'r')
except IOError:
  print "[ERROR] Cound't open file: " + inputFile
try:
  outFile = open(outputFile,'w')
except IOError:
  print "[ERROR] Cound't write Table4.5.csv"
   
for line in inFile:
  line = re.sub('[\n\r]','',line)    
  flds = line.split(',')
  if flds[2]=='AID':
    outFile.write('#Node Name,Facility,WaveLength,Laser Bias%,DHI\n')    
  elif len(flds) > 18:
    if flds[18] and flds[18] != '0.0':
      if float(flds[18]) > 90:
        DHI = 'RED'
        redDHI = redDHI + 1
      elif float(flds[18]) < 75:
        DHI = 'GREEN'
        greenDHI = greenDHI + 1
      else:
        DHI = 'YELLOW'
        yellowDHI = yellowDHI + 1
      outFile.write(flds[0]+','+flds[2]+','+flds[4]+','+flds[18]+','+DHI+'\n')
      
print 'Red', redDHI
print yellowDHI
print greenDHI
totalNo = redDHI+yellowDHI+greenDHI
total = (redDHI*1)+(yellowDHI*3)+(5*greenDHI) 
print total
average = total/float(totalNo)
print average