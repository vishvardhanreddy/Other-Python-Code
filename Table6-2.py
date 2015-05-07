
'''
Created on Jul 10, 2014

@author: vvaka
'''
import re, math, csv, sys

def is_number(s):
  try:
    float(s)
    return True
  except ValueError:
    return False

def main():
  yellowDHI = 0 
  if len(sys.argv) != 3:
    sys.exit('Usage: 72TabALCStates infile outfile')
    return None
  
  inputFile = sys.argv[1]
  outputFile = sys.argv[2]  
  diff = []
  count = 0
  try:
    inFile = open(inputFile,'r')
  except IOError:
    print "[ERROR] Cound't open file: " + inputFile
    return None  
  try:
    outFile = open(outputFile,'w')
  except IOError:
    print "[ERROR] Cound't open file: " + outputFile
    return None
  outFile.write('Node Name,Facility,ALS Mode,DHI\n')
  for line in inFile:
    line = re.sub('[\n\r]','',line)    
    flds = line.split(',')
    if (flds[4] == 'DISABLED'):
      outFile.write(str(flds[0])+','+str(flds[2])+','+str(flds[4])+',YELLOW\n')
      yellowDHI = yellowDHI + 1
  #for i in range(0, 9):
  print yellowDHI
  #print flds
  inFile.close()  
  outFile.close() 
  
  
if __name__ == '__main__':
  main()
